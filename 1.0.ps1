<#
.SYNOPSIS
    Disk_Info_1.0 - Windows 磁碟健康狀態與容量檢測工具 (GUI)

.DESCRIPTION
    這是一支基於 PowerShell 與 Windows Forms 構建的圖形化工具，
    可即時監控本機硬碟、SSD、隨身碟的容量、SMART 健康度、溫度與剩餘壽命。
    支援無感自動刷新、純文字報告匯出與預覽。

.VERSION
    1.0 (Stable Release)
#>

# 靜默執行設定，防止終端機跳出紅字干擾 GUI
$ErrorActionPreference = 'SilentlyContinue'

# ---------------------------------------------------------
# 0. 啟用高解析度 High DPI 支援 (防止字體模糊)
# ---------------------------------------------------------
try {
    $DpiCode = @"
    using System;
    using System.Runtime.InteropServices;
    public class DpiHelper {
        [DllImport("user32.dll")]
        public static extern bool SetProcessDPIAware();
    }
"@
    Add-Type -TypeDefinition $DpiCode
    [void][DpiHelper]::SetProcessDPIAware()
} catch {}

# 載入 GUI 所需組件
[void][System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
[void][System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")
[System.Windows.Forms.Application]::EnableVisualStyles()

# ---------------------------------------------------------
# 1. 檢查並自動請求管理員權限 (讀取 SMART 需要)
# ---------------------------------------------------------
$IsAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $IsAdmin) {
    $exePath = [System.Diagnostics.Process]::GetCurrentProcess().MainModule.FileName
    Start-Process $exePath -Verb RunAs
    exit
}

# ---------------------------------------------------------
# 2. 初始化主視窗
# ---------------------------------------------------------
$form = New-Object System.Windows.Forms.Form
$form.Text = "Disk_Info_1.0"
$form.Size = New-Object System.Drawing.Size(750, 740)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "FixedSingle"
$form.MaximizeBox = $false
$form.BackColor = [System.Drawing.Color]::FromArgb(245, 245, 247)

# 自動讀取 EXE 圖示套用到視窗 (若未打包成 EXE 則忽略)
try {
    $exePath = [System.Diagnostics.Process]::GetCurrentProcess().MainModule.FileName
    $form.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon($exePath)
} catch {}

# 全域字型設定 (Sans-Serif: Microsoft JhengHei UI)
$sansSerifFont = "Microsoft JhengHei UI"
$fontTitle = New-Object System.Drawing.Font($sansSerifFont, 10.5, [System.Drawing.FontStyle]::Bold)
$fontNormal = New-Object System.Drawing.Font($sansSerifFont, 9.5, [System.Drawing.FontStyle]::Regular)
$fontGrid = New-Object System.Drawing.Font($sansSerifFont, 9.5, [System.Drawing.FontStyle]::Regular)
$fontStatus = New-Object System.Drawing.Font($sansSerifFont, 9, [System.Drawing.FontStyle]::Italic)

# ---------------------------------------------------------
# 3. 頂部 MenuStrip 功能表列
# ---------------------------------------------------------
$menuStrip = New-Object System.Windows.Forms.MenuStrip
$menuFile = New-Object System.Windows.Forms.ToolStripMenuItem("File (&F)")

$menuPreview = New-Object System.Windows.Forms.ToolStripMenuItem("預覽記錄 (&P)")
$menuExport = New-Object System.Windows.Forms.ToolStripMenuItem("匯出記錄 (.txt) (&E)")
$menuExit = New-Object System.Windows.Forms.ToolStripMenuItem("結束 (&X)")

[void]$menuFile.DropDownItems.Add($menuPreview)
[void]$menuFile.DropDownItems.Add($menuExport)
[void]$menuFile.DropDownItems.Add((New-Object System.Windows.Forms.ToolStripSeparator))
[void]$menuFile.DropDownItems.Add($menuExit)
[void]$menuStrip.Items.Add($menuFile)
$form.MainMenuStrip = $menuStrip
$form.Controls.Add($menuStrip)

# ---------------------------------------------------------
# 4. 控制項佈局 (選擇標籤、下拉選單、按鈕)
# ---------------------------------------------------------
$labelSelect = New-Object System.Windows.Forms.Label
$labelSelect.Text = "選擇要偵測的磁碟："
$labelSelect.Font = $fontTitle
$labelSelect.Location = New-Object System.Drawing.Point(20, 45)
$labelSelect.AutoSize = $true
$form.Controls.Add($labelSelect)

$comboDisks = New-Object System.Windows.Forms.ComboBox
$comboDisks.Font = $fontNormal
$comboDisks.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$comboDisks.Location = New-Object System.Drawing.Point(185, 43)
$comboDisks.Size = New-Object System.Drawing.Size(375, 28)
$form.Controls.Add($comboDisks)

$btnRefresh = New-Object System.Windows.Forms.Button
$btnRefresh.Text = "立即重新整理"
$btnRefresh.Font = $fontNormal
$btnRefresh.Location = New-Object System.Drawing.Point(575, 42)
$btnRefresh.Size = New-Object System.Drawing.Size(130, 30)
$btnRefresh.BackColor = [System.Drawing.Color]::FromArgb(0, 122, 255)
$btnRefresh.ForeColor = [System.Drawing.Color]::White
$btnRefresh.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$btnRefresh.FlatAppearance.BorderSize = 0
$form.Controls.Add($btnRefresh)

# ---------------------------------------------------------
# 5. DataGridView 表格設定 (雙重緩衝防閃爍)
# ---------------------------------------------------------
$grid = New-Object System.Windows.Forms.DataGridView
$grid.Location = New-Object System.Drawing.Point(20, 90)
$grid.Size = New-Object System.Drawing.Size(685, 540)
$grid.Font = $fontGrid
$grid.ColumnCount = 2
$grid.Columns[0].Name = "項目 (Property)"
$grid.Columns[1].Name = "詳細資訊 (Value)"
$grid.Columns[0].Width = 220
$grid.Columns[1].AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::Fill
$grid.ReadOnly = $true
$grid.AllowUserToAddRows = $false
$grid.RowHeadersVisible = $false
$grid.SelectionMode = "FullRowSelect"
$grid.AutoSizeRowsMode = "AllCells"
$grid.DefaultCellStyle.WrapMode = [System.Windows.Forms.DataGridViewTriState]::True
$grid.BackgroundColor = [System.Drawing.Color]::White
$grid.BorderStyle = [System.Windows.Forms.BorderStyle]::Fixed3D
$grid.EnableHeadersVisualStyles = $true
$grid.ColumnHeadersHeightSizeMode = [System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode]::DisableResizing
$grid.ColumnHeadersHeight = 35
$grid.ColumnHeadersDefaultCellStyle.Font = New-Object System.Drawing.Font($sansSerifFont, 9.5, [System.Drawing.FontStyle]::Bold)

# 開啟雙重緩衝以解決重繪閃爍問題
$gridType = $grid.GetType()
$bindingFlags = [System.Reflection.BindingFlags]::Instance -bor [System.Reflection.BindingFlags]::NonPublic
$doubleBufferedProp = $gridType.GetProperty("DoubleBuffered", $bindingFlags)
$doubleBufferedProp.SetValue($grid, $true, $null)

$form.Controls.Add($grid)

# 底部狀態標籤
$labelStatus = New-Object System.Windows.Forms.Label
$labelStatus.Text = "自動監控中... (每 5 秒刷新一次)"
$labelStatus.Font = $fontStatus
$labelStatus.ForeColor = [System.Drawing.Color]::DimGray
$labelStatus.Location = New-Object System.Drawing.Point(20, 640)
$labelStatus.Size = New-Object System.Drawing.Size(685, 20)
$form.Controls.Add($labelStatus)

# ---------------------------------------------------------
# 6. 核心功能函式庫
# ---------------------------------------------------------

# 安全獲取磁碟編號 (修復隨身碟抓不到 Number 的問題)
function Get-SafeDiskNumber($diskObj) {
    if ($null -ne $diskObj.Number -and $diskObj.Number -match '^\d+$') { return [int]$diskObj.Number }
    if ($null -ne $diskObj.DeviceId -and $diskObj.DeviceId -match '^\d+$') { return [int]$diskObj.DeviceId }
    return 0
}

# 載入物理磁碟至下拉選單
function Load-DiskList {
    $currentSelected = $comboDisks.SelectedIndex
    $comboDisks.Items.Clear()
    $script:disks = Get-PhysicalDisk | Sort-Object { Get-SafeDiskNumber $_ }
    foreach ($d in $script:disks) {
        $dNum = Get-SafeDiskNumber $d
        $sizeGB = [math]::Round($d.Size / 1GB, 2)
        [void]$comboDisks.Items.Add("Disk ${dNum}: $($d.FriendlyName) ($sizeGB GB)")
    }
    if ($currentSelected -ge 0 -and $currentSelected -lt $comboDisks.Items.Count) {
        $comboDisks.SelectedIndex = $currentSelected
    } elseif ($comboDisks.Items.Count -gt 0) {
        $comboDisks.SelectedIndex = 0
    }
}

# 原位更新 DataGridView (無感刷新)
function Set-GridRowValue {
    param([string]$key, [string]$val)
    
    # 防止 $null 寫入崩潰
    $safeVal = if ([string]::IsNullOrWhiteSpace($val)) { "無資料 / 不適用" } else { $val.Trim() }
    
    $foundRow = $null
    foreach ($r in $grid.Rows) {
        if ($r.Cells[0].Value -eq $key) {
            $foundRow = $r
            break
        }
    }
    
    if ($foundRow) {
        if ($foundRow.Cells[1].Value -ne $safeVal) {
            $foundRow.Cells[1].Value = $safeVal
        }
    } else {
        [void]$grid.Rows.Add($key, $safeVal)
    }
}

# 產生純文字報告 (供匯出與預覽)
function Get-DiskReportText {
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine("==========================================================================")
    [void]$sb.AppendLine(" Disk_Info_1.0 - 磁碟健康與容量資訊報告")
    [void]$sb.AppendLine(" 產生時間：$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')")
    [void]$sb.AppendLine("==========================================================================")
    [void]$sb.AppendLine()

    foreach ($row in $grid.Rows) {
        $key = if ($null -ne $row.Cells[0].Value) { [string]$row.Cells[0].Value } else { "" }
        $val = if ($null -ne $row.Cells[1].Value) { [string]$row.Cells[1].Value } else { "" }
        
        if ($key -like "[ * ]") {
            [void]$sb.AppendLine()
            [void]$sb.AppendLine("$key")
            [void]$sb.AppendLine("--------------------------------------------------------------------------")
        } else {
            # 確保對齊：補齊空格
            $paddedKey = $key.PadRight(35)
            [void]$sb.AppendLine("${paddedKey}: ${val}")
        }
    }
    return $sb.ToString()
}

# 讀取並顯示詳細資料
function Display-DiskDetails {
    if ($comboDisks.SelectedIndex -lt 0 -or $null -eq $script:disks) { return }
    
    $grid.SuspendLayout()

    $selectedDisk = $script:disks[$comboDisks.SelectedIndex]
    $diskNum = Get-SafeDiskNumber $selectedDisk

    # 若切換磁碟，清空目前表格避免殘留舊資料
    if ($script:currentDiskNum -ne $diskNum) {
        $grid.Rows.Clear()
        $script:currentDiskNum = $diskNum
    }

    # -------------------------------------
    # 區塊 1: 硬體基本資訊
    # -------------------------------------
    Set-GridRowValue "[ 1. 硬體基本資訊 ]" "----------------------------------------"
    Set-GridRowValue "磁碟編號 (Disk #)" "Disk $diskNum"
    Set-GridRowValue "型號名稱 (Friendly Name)" $selectedDisk.FriendlyName
    Set-GridRowValue "序號 (Serial Number)" $selectedDisk.SerialNumber
    Set-GridRowValue "介面類型 (Bus Type)" $selectedDisk.BusType
    Set-GridRowValue "介質類型 (Media Type)" $selectedDisk.MediaType
    
    $diskObj = Get-Disk -Number $diskNum -ErrorAction SilentlyContinue
    if ($diskObj) {
        Set-GridRowValue "分區樣式 (Partition Style)" $diskObj.PartitionStyle
    }

    # -------------------------------------
    # 區塊 2: 容量與分區資訊
    # -------------------------------------
    Set-GridRowValue "[ 2. 容量與分區資訊 ]" "----------------------------------------"
    $totalGB = [math]::Round($selectedDisk.Size / 1GB, 2)
    $totalTB = [math]::Round($selectedDisk.Size / 1TB, 2)
    Set-GridRowValue "總實體容量 (Total Size)" "$totalGB GB ($totalTB TB)"

    $partitions = Get-Partition -DiskNumber $diskNum -ErrorAction SilentlyContinue
    $hasVolume = $false
    
    if ($partitions) {
        foreach ($p in $partitions) {
            $vol = $p | Get-Volume -ErrorAction SilentlyContinue
            if ($vol -and $vol.Size -gt 0) {
                $hasVolume = $true
                $vSize = [math]::Round($vol.Size / 1GB, 2)
                $vFree = [math]::Round($vol.SizeRemaining / 1GB, 2)
                $vUsed = [math]::Round(($vol.Size - $vol.SizeRemaining) / 1GB, 2)
                $freePct = [math]::Round(($vol.SizeRemaining / $vol.Size) * 100, 1)
                $driveLabel = if ($vol.DriveLetter) { "$($vol.DriveLetter):" } else { "無盤符 (分區 $($p.PartitionNumber))" }
                
                $volInfo = "總計: ${vSize} GB | 已用: ${vUsed} GB | 剩餘: ${vFree} GB (${freePct}%) | 格式: $($vol.FileSystem)"
                Set-GridRowValue "磁碟區 [$driveLabel]" $volInfo
            }
        }
    }
    
    if (-not $hasVolume) {
        Set-GridRowValue "邏輯分區 (Volumes)" "未偵測到已掛載或可辨識的邏輯磁碟區"
    }

    # -------------------------------------
    # 區塊 3: 健康與運作狀態
    # -------------------------------------
    Set-GridRowValue "[ 3. 健康與運作狀態 ]" "----------------------------------------"
    Set-GridRowValue "健康狀態 (Health Status)" $selectedDisk.HealthStatus
    Set-GridRowValue "運作狀態 (Operational Status)" ($selectedDisk.OperationalStatus -join ", ")

    # -------------------------------------
    # 區塊 4: SMART 溫度與數據偵測
    # -------------------------------------
    Set-GridRowValue "[ 4. SMART 溫度與數據偵測 ]" "----------------------------------------"
    $reliability = Get-StorageReliabilityCounter -PhysicalDisk $selectedDisk -ErrorAction SilentlyContinue

    if ($reliability) {
        # 溫度
        if ($reliability.Temperature -and $reliability.Temperature -gt 0) {
            Set-GridRowValue "當前溫度 (Temperature)" "$($reliability.Temperature) °C"
            if ($reliability.TemperatureMax) {
                Set-GridRowValue "歷史最高溫度 (Max Temp)" "$($reliability.TemperatureMax) °C"
            }
        } else {
            Set-GridRowValue "當前溫度 (Temperature)" "硬體不支援讀取溫度"
        }

        # 壽命 (Wear)
        if ($null -ne $reliability.Wear) {
            $wearRemaining = 100 - $reliability.Wear
            Set-GridRowValue "SSD 剩餘壽命估計 (Health/Life)" "$wearRemaining% (磨損率: $($reliability.Wear)%)"
        }

        # 通電時間
        if ($reliability.PowerOnHours) {
            $days = [math]::Round($reliability.PowerOnHours / 24, 1)
            Set-GridRowValue "通電總時間 (Power-On Hours)" "$($reliability.PowerOnHours) 小時 (約 $days 天)"
        }

        # 錯誤紀錄
        Set-GridRowValue "讀取錯誤次數 (Read Errors)" "$($reliability.ReadErrorsTotal)"
        Set-GridRowValue "寫入錯誤次數 (Write Errors)" "$($reliability.WriteErrorsTotal)"
        
        if ($reliability.ReadErrorsUncorrected) {
            Set-GridRowValue "未修正讀取錯誤" "$($reliability.ReadErrorsUncorrected)"
        }
    } else {
        Set-GridRowValue "SMART 數據 (Reliability)" "無法獲取 (隨身碟或外接盒可能不支援 SMART 傳輸)"
    }

    # 分組標題列樣式渲染
    foreach ($row in $grid.Rows) {
        if ($row.Cells[0].Value -like "[ * ]") {
            $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(230, 235, 245)
            $row.DefaultCellStyle.Font = New-Object System.Drawing.Font($sansSerifFont, 9.5, [System.Drawing.FontStyle]::Bold)
            $row.DefaultCellStyle.ForeColor = [System.Drawing.Color]::FromArgb(0, 51, 102)
        }
    }

    $grid.ResumeLayout()

    $currentTime = Get-Date -Format "yyyy/MM/dd HH:mm:ss"
    $labelStatus.Text = "最後更新時間：$currentTime (每 5 秒自動刷新)"
}

# ---------------------------------------------------------
# 7. Menu 點擊事件綁定
# ---------------------------------------------------------

# 預覽記錄
$menuPreview.Add_Click({
    $reportText = Get-DiskReportText
    
    $previewForm = New-Object System.Windows.Forms.Form
    $previewForm.Text = "記錄預覽 - Disk_Info_1.0"
    $previewForm.Size = New-Object System.Drawing.Size(680, 580)
    $previewForm.StartPosition = "CenterParent"
    $previewForm.MinimizeBox = $false
    $previewForm.MaximizeBox = $true
    $previewForm.BackColor = [System.Drawing.Color]::White

    $txtPreview = New-Object System.Windows.Forms.TextBox
    $txtPreview.Multiline = $true
    $txtPreview.ReadOnly = $true
    $txtPreview.ScrollBars = "Both"
    $txtPreview.WordWrap = $false
    $txtPreview.Font = New-Object System.Drawing.Font("Consolas", 10.5)
    $txtPreview.Dock = "Fill"
    $txtPreview.Text = $reportText
    $txtPreview.BackColor = [System.Drawing.Color]::White

    $previewForm.Controls.Add($txtPreview)
    [void]$previewForm.ShowDialog($form)
})

# 匯出記錄 (.txt)
$menuExport.Add_Click({
    $saveDlg = New-Object System.Windows.Forms.SaveFileDialog
    $saveDlg.Filter = "文字檔案 (*.txt)|*.txt|所有檔案 (*.*)|*.*"
    $saveDlg.Title = "儲存磁碟資訊報告"
    $saveDlg.FileName = "Disk_Report_$(Get-Date -Format 'yyyyMMdd_HHmmss').txt"

    if ($saveDlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        try {
            $reportText = Get-DiskReportText
            [System.IO.File]::WriteAllText($saveDlg.FileName, $reportText, [System.Text.Encoding]::UTF8)
            [System.Windows.Forms.MessageBox]::Show("紀錄已成功匯出至：`n$($saveDlg.FileName)", "匯出成功", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        } catch {
            [System.Windows.Forms.MessageBox]::Show("匯出失敗：$($_)", "錯誤", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    }
})

# 結束
$menuExit.Add_Click({
    $form.Close()
})

# ---------------------------------------------------------
# 8. 刷新邏輯與視窗事件掛載
# ---------------------------------------------------------
$timer = New-Object System.Windows.Forms.Timer
$timer.Interval = 5000
$timer.Add_Tick({ Display-DiskDetails })

$comboDisks.Add_SelectedIndexChanged({
    $timer.Stop()
    $script:currentDiskNum = -1
    Display-DiskDetails
    $timer.Start()
})

$btnRefresh.Add_Click({ 
    $timer.Stop()
    Load-DiskList
    $script:currentDiskNum = -1
    Display-DiskDetails
    $timer.Start()
})

$form.Add_Shown({
    Load-DiskList
    Display-DiskDetails
    $timer.Start()
})

$form.Add_FormClosing({
    $timer.Stop()
    $timer.Dispose()
})

# 啟動應用程式
[void]$form.ShowDialog()