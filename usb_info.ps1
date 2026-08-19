<#
.SYNOPSIS
    USB_Info - Windows USB 裝置檢測工具 (GUI - PowerShell)
    專注篩選：外接 USB 滑鼠 (Mouse) 與隨身碟 (USB Drive)
#>

# 靜默執行設定
$ErrorActionPreference = 'SilentlyContinue'

# ---------------------------------------------------------
# 0. 啟用高解析度 High DPI 支援
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
# 1. 初始化主視窗
# ---------------------------------------------------------
$form = New-Object System.Windows.Forms.Form
$form.Text = "USB_Info"
$form.Size = New-Object System.Drawing.Size(900, 740)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "FixedSingle"
$form.MaximizeBox = $false
$form.BackColor = [System.Drawing.Color]::FromArgb(245, 245, 247)

try {
    $exePath = [System.Diagnostics.Process]::GetCurrentProcess().MainModule.FileName
    $form.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon($exePath)
} catch {}

$sansSerifFont = "Microsoft JhengHei UI"
$fontTitle = New-Object System.Drawing.Font($sansSerifFont, 10.5, [System.Drawing.FontStyle]::Bold)
$fontNormal = New-Object System.Drawing.Font($sansSerifFont, 9.5, [System.Drawing.FontStyle]::Regular)
$fontGrid = New-Object System.Drawing.Font($sansSerifFont, 9.5, [System.Drawing.FontStyle]::Regular)
$fontStatus = New-Object System.Drawing.Font($sansSerifFont, 9, [System.Drawing.FontStyle]::Italic)

# ---------------------------------------------------------
# 2. 頂部 MenuStrip 功能表列
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
# 3. 控制項佈局 (調整 X 軸座標，消除重疊遮擋)
# ---------------------------------------------------------
$labelSelect = New-Object System.Windows.Forms.Label
$labelSelect.Text = "選擇要偵測的 USB 裝置："
$labelSelect.Font = $fontTitle
$labelSelect.Location = New-Object System.Drawing.Point(20, 45)
$labelSelect.AutoSize = $true
$form.Controls.Add($labelSelect)

# 將 Combo 移至 X = 240，給左側文字充足 space
$comboUsb = New-Object System.Windows.Forms.ComboBox
$comboUsb.Font = $fontNormal
$comboUsb.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$comboUsb.Location = New-Object System.Drawing.Point(240, 42)
$comboUsb.Size = New-Object System.Drawing.Size(460, 28)
$comboUsb.DropDownWidth = 700
$form.Controls.Add($comboUsb)

$btnRefresh = New-Object System.Windows.Forms.Button
$btnRefresh.Text = "立即重新整理"
$btnRefresh.Font = $fontNormal
$btnRefresh.Location = New-Object System.Drawing.Point(715, 40)
$btnRefresh.Size = New-Object System.Drawing.Size(140, 32)
$btnRefresh.BackColor = [System.Drawing.Color]::FromArgb(0, 122, 255)
$btnRefresh.ForeColor = [System.Drawing.Color]::White
$btnRefresh.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$btnRefresh.FlatAppearance.BorderSize = 0
$form.Controls.Add($btnRefresh)

# ---------------------------------------------------------
# 4. DataGridView 表格設定
# ---------------------------------------------------------
$grid = New-Object System.Windows.Forms.DataGridView
$grid.Location = New-Object System.Drawing.Point(20, 90)
$grid.Size = New-Object System.Drawing.Size(835, 540)
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

# 雙重緩衝 (防止更新畫面時閃爍)
$gridType = $grid.GetType()
$bindingFlags = [System.Reflection.BindingFlags]::Instance -bor [System.Reflection.BindingFlags]::NonPublic
$doubleBufferedProp = $gridType.GetProperty("DoubleBuffered", $bindingFlags)
$doubleBufferedProp.SetValue($grid, $true, $null)

$form.Controls.Add($grid)

$labelStatus = New-Object System.Windows.Forms.Label
$labelStatus.Text = "自動監控中... (每 5 秒刷新一次)"
$labelStatus.Font = $fontStatus
$labelStatus.ForeColor = [System.Drawing.Color]::DimGray
$labelStatus.Location = New-Object System.Drawing.Point(20, 640)
$labelStatus.Size = New-Object System.Drawing.Size(835, 20)
$form.Controls.Add($labelStatus)

$tooltipUsb = New-Object System.Windows.Forms.ToolTip
$tooltipUsb.ShowAlways = $true

# ---------------------------------------------------------
# 5. 核心功能函式庫
# ---------------------------------------------------------

# 解析 VID / PID / 序號
function Parse-DeviceID($deviceString) {
    $vid = "未知"
    $pid = "未知"
    $serial = "未知 / 無資料"

    if ($deviceString -match "VID_([0-9A-Fa-f]{4})") { $vid = "0x" + $Matches[1] }
    if ($deviceString -match "(?:PID|PID&)_([0-9A-Fa-f]{4})") { $pid = "0x" + $Matches[1] }
    
    if ($deviceString -match "(?:USB|HID)\\[^\\]+\\([^\s\&\?]+)") {
        $serial = $Matches[1]
    }

    return @{ VID = $vid; PID = $pid; Serial = $serial }
}

# 抓取並嚴格過濾去重：僅保留「外接 USB 滑鼠」與「USB 隨身碟」
function Load-UsbList {
    $currentSelected = $comboUsb.SelectedIndex
    $comboUsb.Items.Clear()
    $script:devices = [System.Collections.ArrayList]::new()

    # 1. 抓取外接 USB 滑鼠 (排除內建，並透過 VID/PID 去重)
    $mousePnpDevs = Get-CimInstance Win32_PnPEntity | Where-Object { 
        ($_.PNPClass -eq "Mouse" -or $_.Service -match "mouhid") -and 
        $_.PNPDeviceID -like "*VID_*" -and
        $_.Present -eq $true
    }
    
    $seenVidPid = @{}

    foreach ($m in $mousePnpDevs) {
        $parsed = Parse-DeviceID $m.PNPDeviceID
        $vidPidKey = "$($parsed.VID)_$($parsed.PID)"

        if ($parsed.VID -ne "未知" -and -not $seenVidPid.ContainsKey($vidPidKey)) {
            $seenVidPid[$vidPidKey] = $true

            $rawName = if ($m.Name) { $m.Name } else { $m.Caption }
            $displayName = if ($rawName -match "HID-compliant|符合 HID") {
                "$rawName [VID: $($parsed.VID) PID: $($parsed.PID)]"
            } else {
                $rawName
            }

            [void]$script:devices.Add(@{
                Type = "滑鼠 (Mouse)"
                Name = $displayName
                PNPDeviceID = $m.PNPDeviceID
                Status = $m.Status
                Caption = $m.Caption
                RawObj = $m
            })
        }
    }

    # 2. 抓取 USB 隨身碟 (外接抽取式/USB 磁碟)
    $drives = Get-CimInstance Win32_DiskDrive | Where-Object { $_.InterfaceType -eq "USB" }
    foreach ($d in $drives) {
        $sizeGB = if ($d.Size) { [math]::Round($d.Size / 1GB, 2) } else { 0 }
        $displayName = if ($sizeGB -gt 0) { "$($d.Model) ($sizeGB GB)" } else { $d.Model }

        [void]$script:devices.Add(@{
            Type = "隨身碟 (USB Drive)"
            Name = $displayName
            PNPDeviceID = $d.PNPDeviceID
            Status = $d.Status
            Caption = $d.Caption
            RawObj = $d
        })
    }

    foreach ($dev in $script:devices) {
        $label = "[" + $dev.Type + "] " + $dev.Name
        [void]$comboUsb.Items.Add($label)
    }

    if ($currentSelected -ge 0 -and $currentSelected -lt $comboUsb.Items.Count) {
        $comboUsb.SelectedIndex = $currentSelected
    } elseif ($comboUsb.Items.Count -gt 0) {
        $comboUsb.SelectedIndex = 0
    }
}

function Set-GridRowValue {
    param([string]$key, [string]$val, [bool]$isHeader = $false)

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
        $rowIndex = $grid.Rows.Add($key, $safeVal)
        if ($isHeader) {
            $row = $grid.Rows[$rowIndex]
            $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(230, 235, 245)
            $row.DefaultCellStyle.Font = New-Object System.Drawing.Font($sansSerifFont, 9.5, [System.Drawing.FontStyle]::Bold)
            $row.DefaultCellStyle.ForeColor = [System.Drawing.Color]::FromArgb(0, 51, 102)
        }
    }
}

function Get-UsbReportText {
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine("==========================================================================")
    [void]$sb.AppendLine(" USB_Info - USB 裝置資訊報告 (僅外接滑鼠/隨身碟)")
    $reportTime = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    [void]$sb.AppendLine(" 產生時間：$reportTime")
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
            $paddedKey = $key.PadRight(35)
            [void]$sb.AppendLine("${paddedKey}: ${val}")
        }
    }
    return $sb.ToString()
}

function Display-UsbDetails {
    if ($comboUsb.SelectedIndex -lt 0 -or $null -eq $script:devices -or $comboUsb.SelectedIndex -ge $script:devices.Count) { return }

    $grid.SuspendLayout()

    $selectedDev = $script:devices[$comboUsb.SelectedIndex]
    $devIdx = $comboUsb.SelectedIndex

    if ($script:currentDevIdx -ne $devIdx) {
        $grid.Rows.Clear()
        $script:currentDevIdx = $devIdx
    }

    $parsedInfo = Parse-DeviceID $selectedDev.PNPDeviceID
    $pnpDev = Get-CimInstance Win32_PnPEntity | Where-Object { $_.PNPDeviceID -eq $selectedDev.PNPDeviceID }

    # 1. USB 裝置基本資訊
    Set-GridRowValue "[ 1. USB 裝置基本資訊 ]" "----------------------------------------" -isHeader $true
    Set-GridRowValue "裝置類別 (Category)" $selectedDev.Type
    Set-GridRowValue "裝置名稱 (Device Name)" $selectedDev.Name
    Set-GridRowValue "廠商識別碼 (VID)" $parsedInfo.VID
    Set-GridRowValue "產品識別碼 (PID)" $parsedInfo.PID
    Set-GridRowValue "裝置序號 (Serial Number)" $parsedInfo.Serial
    Set-GridRowValue "PNP 裝置 ID (PNPDeviceID)" $selectedDev.PNPDeviceID

    # 2. 通訊協定與連接狀態
    Set-GridRowValue "[ 2. 通訊協定與連接狀態 ]" "----------------------------------------" -isHeader $true
    Set-GridRowValue "通訊協定版本 (USB Protocol)" "USB 2.0 / 3.0+ (自動向下相容)"
    Set-GridRowValue "運作狀態 (Status)" (if ($pnpDev.Status) { $pnpDev.Status } else { "OK (正常運作)" })
    Set-GridRowValue "驅動服務名稱 (Service)" $pnpDev.Service
    Set-GridRowValue "製造商 (Manufacturer)" $pnpDev.Manufacturer

    # 3. 供電與電源狀態
    Set-GridRowValue "[ 3. 供電與電源狀態 ]" "----------------------------------------" -isHeader $true
    $powerState = if ($pnpDev.ConfigManagerErrorCode -eq 0) { "正常供電 (Power On - Operational)" } else { "電源或驅動異常 (Code: $($pnpDev.ConfigManagerErrorCode))" }
    Set-GridRowValue "電源供電狀態 (Power Status)" $powerState
    Set-GridRowValue "省電機制 (Power Saving)" "由系統 USB 選擇性暫停與 HID 節能管理"

    # 4. 類別專屬資訊 (隨身碟)
    if ($selectedDev.Type -like "*隨身碟*") {
        Set-GridRowValue "[ 4. 隨身碟儲存專屬資訊 ]" "----------------------------------------" -isHeader $true
        $disk = $selectedDev.RawObj
        Set-GridRowValue "介面類型 (Interface)" $disk.InterfaceType
        Set-GridRowValue "磁碟分割區數 (Partitions)" $disk.Partitions
        
        $partList = Get-CimInstance -Query "ASSOCIATORS OF {Win32_DiskDrive.DeviceID='$($disk.DeviceID)'} WHERE AssocClass = Win32_DiskDriveToDiskPartition"
        $driveLetters = @()
        foreach ($part in $partList) {
            $volList = Get-CimInstance -Query "ASSOCIATORS OF {Win32_DiskPartition.DeviceID='$($part.DeviceID)'} WHERE AssocClass = Win32_LogicalDiskToPartition"
            foreach ($vol in $volList) {
                $driveLetters += $vol.DeviceID
            }
        }
        $driveStr = if ($driveLetters.Count -gt 0) { $driveLetters -join ", " } else { "未配置磁碟機代號" }
        Set-GridRowValue "掛載代號 (Drive Letter)" $driveStr
    }

    $grid.ResumeLayout()

    $currentTime = Get-Date -Format "yyyy/MM/dd HH:mm:ss"
    $labelStatus.Text = "最後更新時間：$currentTime (每 5 秒自動刷新)"
    $tooltipUsb.SetToolTip($comboUsb, "完整名稱：" + $selectedDev.Name)
}

# ---------------------------------------------------------
# 6. 選單與事件掛載
# ---------------------------------------------------------
$menuPreview.Add_Click({
    $reportText = Get-UsbReportText
    $previewForm = New-Object System.Windows.Forms.Form
    $previewForm.Text = "記錄預覽 - USB_Info"
    $previewForm.Size = New-Object System.Drawing.Size(680, 580)
    $previewForm.StartPosition = "CenterParent"
    $previewForm.MinimizeBox = $false

    $txtPreview = New-Object System.Windows.Forms.TextBox
    $txtPreview.Multiline = $true
    $txtPreview.ReadOnly = $true
    $txtPreview.ScrollBars = "Both"
    $txtPreview.WordWrap = $false
    $txtPreview.Font = New-Object System.Drawing.Font("Consolas", 10.5)
    $txtPreview.Dock = "Fill"
    $txtPreview.Text = $reportText

    $previewForm.Controls.Add($txtPreview)
    [void]$previewForm.ShowDialog($form)
})

$menuExport.Add_Click({
    $saveDlg = New-Object System.Windows.Forms.SaveFileDialog
    try {
        $saveDlg.Filter = "文字檔案 (*.txt)|*.txt|所有檔案 (*.*)|*.*"
        $saveDlg.Title = "儲存 USB 資訊報告"
        $saveDlg.FileName = "USB_Report_" + (Get-Date -Format 'yyyyMMdd_HHmmss') + ".txt"

        if ($saveDlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $reportText = Get-UsbReportText
            [System.IO.File]::WriteAllText($saveDlg.FileName, $reportText, [System.Text.Encoding]::UTF8)
            [System.Windows.Forms.MessageBox]::Show("紀錄已成功匯出至：`n" + $saveDlg.FileName, "匯出成功", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        }
    } finally {
        $saveDlg.Dispose()
    }
})

$menuExit.Add_Click({ $form.Close() })

$script:currentDevIdx = -1

$timer = New-Object System.Windows.Forms.Timer
$timer.Interval = 5000
$timer.Add_Tick({ Display-UsbDetails })

$comboUsb.Add_SelectedIndexChanged({
    $timer.Stop()
    $script:currentDevIdx = -1
    Display-UsbDetails
    $timer.Start()
})

$btnRefresh.Add_Click({
    $timer.Stop()
    Load-UsbList
    $script:currentDevIdx = -1
    Display-UsbDetails
    $timer.Start()
})

$form.Add_Shown({
    Load-UsbList
    Display-UsbDetails
    $timer.Start()
})

$form.Add_FormClosing({
    $timer.Stop()
    $timer.Dispose()
})

# 啟動應用程式
[void]$form.ShowDialog()