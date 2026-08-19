<#
.SYNOPSIS
    CPU_Info_1.0 - Windows CPU 處理器資訊檢測工具 (GUI - Bug Fixed & Optimized)
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

# 效能優化：初始化全域效能計數器 (解決 Get-Counter 介面卡頓 Bug)
try {
    $script:cpuCounter = New-Object System.Diagnostics.PerformanceCounter("Processor", "% Processor Time", "_Total")
    [void]$script:cpuCounter.NextValue()
} catch {
    $script:cpuCounter = $null
}

# ---------------------------------------------------------
# 1. 初始化主視窗
# ---------------------------------------------------------
$form = New-Object System.Windows.Forms.Form
$form.Text = "CPU_Info_1.0 (Optimized)"
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
# 3. 控制項佈局
# ---------------------------------------------------------
$labelSelect = New-Object System.Windows.Forms.Label
$labelSelect.Text = "選擇要偵測的處理器："
$labelSelect.Font = $fontTitle
$labelSelect.Location = New-Object System.Drawing.Point(20, 45)
$labelSelect.AutoSize = $true
$form.Controls.Add($labelSelect)

$comboCpus = New-Object System.Windows.Forms.ComboBox
$comboCpus.Font = $fontNormal
$comboCpus.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$comboCpus.Location = New-Object System.Drawing.Point(210, 43)
$comboCpus.Size = New-Object System.Drawing.Size(500, 28)
$comboCpus.DropDownWidth = 700
$form.Controls.Add($comboCpus)

$btnRefresh = New-Object System.Windows.Forms.Button
$btnRefresh.Text = "立即重新整理"
$btnRefresh.Font = $fontNormal
$btnRefresh.Location = New-Object System.Drawing.Point(720, 42)
$btnRefresh.Size = New-Object System.Drawing.Size(130, 30)
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

# 雙重緩衝
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

$tooltipCpu = New-Object System.Windows.Forms.ToolTip
$tooltipCpu.ShowAlways = $true

# ---------------------------------------------------------
# 5. 核心功能函式庫
# ---------------------------------------------------------
function Get-ArchitectureName($archCode) {
    switch ($archCode) {
        0  { "x86 (32-bit)" }
        5  { "ARM" }
        6  { "Itanium (IA-64)" }
        9  { "x64 (64-bit)" }
        12 { "ARM64" }
        default { "未知 (Code: $archCode)" }
    }
}

function Get-CpuStatusName($status) {
    switch ($status) {
        "OK"          { "正常運作 (OK)" }
        "Degraded"    { "效能降低 (Degraded)" }
        "Error"       { "錯誤 (Error)" }
        "Stressed"    { "壓力過大 (Stressed)" }
        default       { $status }
    }
}

function Format-CacheSize($kb) {
    if ($null -eq $kb -or $kb -eq 0) { return "無資料 / 不適用" }
    $mb = [math]::Round($kb / 1024, 2)
    if ($mb -ge 1024) {
        $gb = [math]::Round($mb / 1024, 2)
        return "$kb KB ($mb MB / $gb GB)"
    }
    return "$kb KB ($mb MB)"
}

function Format-Temperature($kelvinTenths) {
    if ($null -eq $kelvinTenths -or $kelvinTenths -le 0) { return "無資料 / 不適用" }
    $celsius = [math]::Round(($kelvinTenths - 2732) / 10.0, 1)
    $fahrenheit = [math]::Round($celsius * 9 / 5 + 32, 1)
    return "$celsius °C ($fahrenheit °F)"
}

function Load-CpuList {
    $currentSelected = $comboCpus.SelectedIndex
    $comboCpus.Items.Clear()
    $script:cpus = @(Get-CimInstance Win32_Processor | Sort-Object DeviceID)
    foreach ($c in $script:cpus) {
        $idx = $script:cpus.IndexOf($c)
        $label = "CPU " + $idx + ": " + $c.Name
        if ($c.MaxClockSpeed) { $label += " (" + $c.MaxClockSpeed + " MHz)" }
        [void]$comboCpus.Items.Add($label)
    }
    if ($currentSelected -ge 0 -and $currentSelected -lt $comboCpus.Items.Count) {
        $comboCpus.SelectedIndex = $currentSelected
    } elseif ($comboCpus.Items.Count -gt 0) {
        $comboCpus.SelectedIndex = 0
    }
}

# [優化] 寫入表格同時設定樣式，避免重複全表格迴圈
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

function Get-CpuReportText {
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine("==========================================================================")
    [void]$sb.AppendLine(" CPU_Info_1.0 - CPU 處理器資訊報告")
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

# [優化] 高效且不阻塞介面的 CPU 使用率讀取
function Get-RealtimeCpuUsage {
    if ($null -ne $script:cpuCounter) {
        try {
            $usage = [math]::Round($script:cpuCounter.NextValue(), 1)
            return "$usage %"
        } catch {}
    }
    return "無法獲取"
}

function Get-RealtimeCpuTemperature {
    try {
        $temp = Get-CimInstance MSAcpi_ThermalZoneTemperature -Namespace root/wmi -ErrorAction Stop
        if ($temp -and $temp.Count -gt 0) {
            $maxTemp = $temp | Sort-Object CurrentTemperature -Descending | Select-Object -First 1
            return Format-Temperature $maxTemp.CurrentTemperature
        }
    } catch {}
    return "無法獲取 (需管理員權限或不支援)"
}

function Display-CpuDetails {
    # 防護機制：避免索引越界崩潰
    if ($comboCpus.SelectedIndex -lt 0 -or $null -eq $script:cpus -or $comboCpus.SelectedIndex -ge $script:cpus.Count) { return }

    $grid.SuspendLayout()

    $selectedCpu = $script:cpus[$comboCpus.SelectedIndex]
    $cpuIdx = $comboCpus.SelectedIndex

    if ($script:currentCpuIdx -ne $cpuIdx) {
        $grid.Rows.Clear()
        $script:currentCpuIdx = $cpuIdx
    }

    # 1. 處理器基本資訊
    Set-GridRowValue "[ 1. 處理器基本資訊 ]" "----------------------------------------" -isHeader $true
    Set-GridRowValue "處理器編號 (CPU #)" "CPU $cpuIdx"
    Set-GridRowValue "處理器名稱 (Name)" $selectedCpu.Name
    Set-GridRowValue "製造商 (Manufacturer)" $selectedCpu.Manufacturer
    Set-GridRowValue "架構 (Architecture)" (Get-ArchitectureName $selectedCpu.Architecture)
    Set-GridRowValue "插槽類型 (Socket)" $selectedCpu.SocketDesignation
    Set-GridRowValue "處理器序號 (Processor ID)" $selectedCpu.ProcessorId
    Set-GridRowValue "運作狀態 (Status)" (Get-CpuStatusName $selectedCpu.Status)
    Set-GridRowValue "裝置識別碼 (DeviceID)" $selectedCpu.DeviceID

    # 2. 核心與執行緒
    Set-GridRowValue "[ 2. 核心與執行緒 ]" "----------------------------------------" -isHeader $true
    $physCores = $selectedCpu.NumberOfCores
    $logCores = $selectedCpu.NumberOfLogicalProcessors
    Set-GridRowValue "實體核心數 (Physical Cores)" "$physCores 核心"
    Set-GridRowValue "邏輯執行緒數 (Logical Threads)" "$logCores 執行緒"
    if ($physCores -and $physCores -gt 0) {
        $threadsPerCore = [math]::Round($logCores / $physCores, 1)
        $htStatus = if ($threadsPerCore -gt 1) { "是" } else { "否" }
        Set-GridRowValue "每核心執行緒數" "$threadsPerCore (超執行緒: $htStatus)"
    }

    $totalCpus = $script:cpus.Count
    $systemPhysCores = ($script:cpus | Measure-Object NumberOfCores -Sum).Sum
    $systemLogCores = ($script:cpus | Measure-Object NumberOfLogicalProcessors -Sum).Sum
    Set-GridRowValue "系統處理器總數 (Total CPUs)" "$totalCpus 顆"
    Set-GridRowValue "系統實體核心總數" "$systemPhysCores 核心"
    Set-GridRowValue "系統邏輯執行緒總數" "$systemLogCores 執行緒"

    # 3. 時脈與頻率
    Set-GridRowValue "[ 3. 時脈與頻率 ]" "----------------------------------------" -isHeader $true
    $maxSpeed = $selectedCpu.MaxClockSpeed
    $curSpeed = $selectedCpu.CurrentClockSpeed
    $extClock = $selectedCpu.ExtClock

    Set-GridRowValue "最大時脈 (Max Clock)" "$maxSpeed MHz"
    Set-GridRowValue "目前時脈 (Current Clock)" "$curSpeed MHz"
    if ($extClock -and $extClock -gt 0) {
        Set-GridRowValue "外部匯流排時脈 (Ext Clock / FSB)" "$extClock MHz"
        if ($maxSpeed) {
            $multiplier = [math]::Round($maxSpeed / $extClock, 1)
            Set-GridRowValue "倍頻 (Multiplier)" "$multiplier x ($extClock MHz x $multiplier)"
        }
    }
    if ($maxSpeed -and $maxSpeed -gt 0 -and $curSpeed) {
        $speedPct = [math]::Round(($curSpeed / $maxSpeed) * 100, 1)
        Set-GridRowValue "目前頻率佔比" "$speedPct %"
    }

    # 4. 快取記憶體 (內部 Level 邏輯修正)
    Set-GridRowValue "[ 4. 快取記憶體 ]" "----------------------------------------" -isHeader $true
    $l1Cache = $null
    $l2Cache = $selectedCpu.L2CacheSize
    $l3Cache = $selectedCpu.L3CacheSize

    try {
        $cacheMem = Get-CimInstance Win32_CacheMemory -ErrorAction Stop
        if ($cacheMem) {
            $l1 = $cacheMem | Where-Object { $_.Level -eq 3 } | Select-Object -First 1
            if ($l1) { $l1Cache = $l1.InstalledSize }
        }
    } catch {}

    Set-GridRowValue "L1 快取 (Level 1)" (Format-CacheSize $l1Cache)
    Set-GridRowValue "L2 快取 (Level 2)" (Format-CacheSize $l2Cache)
    Set-GridRowValue "L3 快取 (Level 3)" (Format-CacheSize $l3Cache)

    # 5. 電壓與功率
    Set-GridRowValue "[ 5. 電壓與功率 ]" "----------------------------------------" -isHeader $true
    $voltage = $selectedCpu.CurrentVoltage
    if ($null -ne $voltage -and $voltage -gt 0) {
        $volts = [math]::Round($voltage / 10.0, 2)
        Set-GridRowValue "目前電壓 (Current Voltage)" "$volts V"
    } else {
        Set-GridRowValue "目前電壓 (Current Voltage)" "無資料 / 不適用"
    }

    $tdp = "無法獲取"
    try {
        $regKey = "HKLM:\HARDWARE\DESCRIPTION\System\CentralProcessor\$cpuIdx"
        $regVal = Get-ItemProperty -Path $regKey -ErrorAction Stop
        if ($regVal.PSObject.Properties.Name -contains 'TDP' -and $regVal.TDP) {
            $tdp = "$([math]::Round($regVal.TDP, 1)) W"
        }
    } catch {}
    Set-GridRowValue "功率設計 (TDP)" $tdp

    # 6. 即時狀態
    Set-GridRowValue "[ 6. 即時狀態 ]" "----------------------------------------" -isHeader $true
    Set-GridRowValue "CPU 使用率 (Usage)" (Get-RealtimeCpuUsage)
    Set-GridRowValue "CPU 負載 (Load %)" (if ($null -ne $selectedCpu.LoadPercentage) { "$($selectedCpu.LoadPercentage) %" } else { "無資料 / 不適用" })
    Set-GridRowValue "即時運轉時脈" "$curSpeed MHz"
    Set-GridRowValue "CPU 溫度 (Temperature)" (Get-RealtimeCpuTemperature)

    # 7. 虛擬化與其他特性
    Set-GridRowValue "[ 7. 虛擬化與特性 ]" "----------------------------------------" -isHeader $true
    $virtSupport = $selectedCpu.VirtualizationFirmwareEnabled
    $virtText = if ($virtSupport) { "已啟用" } elseif ($null -ne $virtSupport) { "未啟用" } else { "無資料 / 不適用" }
    Set-GridRowValue "虛擬化支援 (Virtualization)" $virtText
    Set-GridRowValue "資料寬度 (Data Width)" "$($selectedCpu.DataWidth) 位元"
    Set-GridRowValue "位址寬度 (Address Width)" "$($selectedCpu.AddressWidth) 位元"

    $grid.ResumeLayout()

    $currentTime = Get-Date -Format "yyyy/MM/dd HH:mm:ss"
    $labelStatus.Text = "最後更新時間：$currentTime (每 5 秒自動刷新)"
    $tooltipCpu.SetToolTip($comboCpus, "完整名稱：" + $selectedCpu.Name)
}

# ---------------------------------------------------------
# 6. 選單與事件掛載
# ---------------------------------------------------------
$menuPreview.Add_Click({
    $reportText = Get-CpuReportText
    $previewForm = New-Object System.Windows.Forms.Form
    $previewForm.Text = "記錄預覽 - CPU_Info_1.0"
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
        $saveDlg.Title = "儲存 CPU 資訊報告"
        $saveDlg.FileName = "CPU_Report_" + (Get-Date -Format 'yyyyMMdd_HHmmss') + ".txt"

        if ($saveDlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $reportText = Get-CpuReportText
            [System.IO.File]::WriteAllText($saveDlg.FileName, $reportText, [System.Text.Encoding]::UTF8)
            [System.Windows.Forms.MessageBox]::Show("紀錄已成功匯出至：`n" + $saveDlg.FileName, "匯出成功", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        }
    } finally {
        $saveDlg.Dispose() # 確保記憶體資源釋放
    }
})

$menuExit.Add_Click({ $form.Close() })

$script:currentCpuIdx = -1

$timer = New-Object System.Windows.Forms.Timer
$timer.Interval = 5000
$timer.Add_Tick({ Display-CpuDetails })

$comboCpus.Add_SelectedIndexChanged({
    $timer.Stop()
    $script:currentCpuIdx = -1
    Display-CpuDetails
    $timer.Start()
})

$btnRefresh.Add_Click({
    $timer.Stop()
    Load-CpuList
    $script:currentCpuIdx = -1
    Display-CpuDetails
    $timer.Start()
})

$form.Add_Shown({
    Load-CpuList
    Display-CpuDetails
    $timer.Start()
})

$form.Add_FormClosing({
    $timer.Stop()
    $timer.Dispose()
    if ($null -ne $script:cpuCounter) { $script:cpuCounter.Dispose() }
})

# 啟動應用程式
[void]$form.ShowDialog()