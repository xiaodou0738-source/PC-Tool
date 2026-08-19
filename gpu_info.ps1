<#
.SYNOPSIS
    GPU_Info - Windows GPU 顯示卡資訊檢測工具 (GUI - PowerShell)
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
$form.Text = "GPU_Info "
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
$labelSelect.Text = "選擇要偵測的顯示卡："
$labelSelect.Font = $fontTitle
$labelSelect.Location = New-Object System.Drawing.Point(20, 45)
$labelSelect.AutoSize = $true
$form.Controls.Add($labelSelect)

$comboGpus = New-Object System.Windows.Forms.ComboBox
$comboGpus.Font = $fontNormal
$comboGpus.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$comboGpus.Location = New-Object System.Drawing.Point(210, 43)
$comboGpus.Size = New-Object System.Drawing.Size(500, 28)
$comboGpus.DropDownWidth = 700
$form.Controls.Add($comboGpus)

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

$tooltipGpu = New-Object System.Windows.Forms.ToolTip
$tooltipGpu.ShowAlways = $true

# ---------------------------------------------------------
# 5. 核心功能函式庫
# ---------------------------------------------------------
function Format-BytesToMBorGB($bytes) {
    if ($null -eq $bytes -or $bytes -le 0) { return "無資料 / 不適用" }
    $mb = [math]::Round($bytes / 1MB, 2)
    if ($mb -ge 1024) {
        $gb = [math]::Round($bytes / 1GB, 2)
        return "$gb GB ($mb MB)"
    }
    return "$mb MB"
}

function Get-GpuStatusName($status) {
    switch ($status) {
        "OK"          { "正常運作 (OK)" }
        "Degraded"    { "效能降低 (Degraded)" }
        "Error"       { "錯誤 (Error)" }
        "Stressed"    { "壓力過大 (Stressed)" }
        default       { if ($status) { $status } else { "正常 (OK)" } }
    }
}

function Get-MemoryTypeName($typeCode) {
    switch ($typeCode) {
        1 { "Other" }
        2 { "Unknown" }
        3 { "VRAM" }
        4 { "DRAM" }
        5 { "SRAM" }
        6 { "WRAM" }
        7 { "EDO RAM" }
        8 { "Burst Synchronous DRAM" }
        9 { "Pipelined Burst SRAM" }
        10 { "CDRAM" }
        11 { "3D RAM" }
        12 { "SDRAM" }
        13 { "SGRAM" }
        default { "未知 / 自動識別" }
    }
}

function Load-GpuList {
    $currentSelected = $comboGpus.SelectedIndex
    $comboGpus.Items.Clear()
    $script:gpus = @(Get-CimInstance Win32_VideoController | Sort-Object DeviceID)
    foreach ($g in $script:gpus) {
        $idx = $script:gpus.IndexOf($g)
        $label = "GPU " + $idx + ": " + $g.Name
        if ($g.AdapterRAM -and $g.AdapterRAM -gt 0) {
            $ramGB = [math]::Round($g.AdapterRAM / 1GB, 2)
            if ($ramGB -ge 1) { $label += " (" + $ramGB + " GB)" }
        }
        [void]$comboGpus.Items.Add($label)
    }
    if ($currentSelected -ge 0 -and $currentSelected -lt $comboGpus.Items.Count) {
        $comboGpus.SelectedIndex = $currentSelected
    } elseif ($comboGpus.Items.Count -gt 0) {
        $comboGpus.SelectedIndex = 0
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

function Get-GpuReportText {
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine("==========================================================================")
    [void]$sb.AppendLine(" GPU_Info - GPU 顯示卡資訊報告")
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

function Get-NvidiaDynamicInfo($gpuName) {
    # 嘗試呼叫 nvidia-smi 獲取即時資料 (使用率、溫度、顯存)
    $smiPath = "$env:ProgramFiles\NVIDIA Corporation\NVSMI\nvidia-smi.exe"
    if (-not (Test-Path $smiPath)) {
        $smiPath = "nvidia-smi.exe"
    }
    try {
        $smiOutput = & $smiPath --query-gpu=utilization.gpu,temperature.gpu,memory.used,memory.total --format=csv,noheader,nounits 2>$null
        if ($smiOutput) {
            $info = $smiOutput.ToString().Split(",")
            return @{
                Usage = "$($info[0].Trim()) %"
                Temp  = "$($info[1].Trim()) °C"
                Mem   = "$($info[2].Trim()) MB / $($info[3].Trim()) MB"
            }
        }
    } catch {}
    return $null
}

function Display-GpuDetails {
    if ($comboGpus.SelectedIndex -lt 0 -or $null -eq $script:gpus -or $comboGpus.SelectedIndex -ge $script:gpus.Count) { return }

    $grid.SuspendLayout()

    $selectedGpu = $script:gpus[$comboGpus.SelectedIndex]
    $gpuIdx = $comboGpus.SelectedIndex

    if ($script:currentGpuIdx -ne $gpuIdx) {
        $grid.Rows.Clear()
        $script:currentGpuIdx = $gpuIdx
    }

    # 1. 顯示卡基本資訊
    Set-GridRowValue "[ 1. 顯示卡基本資訊 ]" "----------------------------------------" -isHeader $true
    Set-GridRowValue "顯示卡編號 (GPU #)" "GPU $gpuIdx"
    Set-GridRowValue "顯示卡名稱 (Name)" $selectedGpu.Name
    Set-GridRowValue "製造商/晶片商 (Adapter Compatibility)" $selectedGpu.AdapterCompatibility
    Set-GridRowValue "裝置識別碼 (DeviceID)" $selectedGpu.DeviceID
    Set-GridRowValue "PNP 裝置 ID (PNPDeviceID)" $selectedGpu.PNPDeviceID
    Set-GridRowValue "運作狀態 (Status)" (Get-GpuStatusName $selectedGpu.Status)

    # 2. 顯示記憶體 (VRAM)
    Set-GridRowValue "[ 2. 顯示記憶體 (VRAM) ]" "----------------------------------------" -isHeader $true
    Set-GridRowValue "獨立顯存容量 (Adapter RAM)" (Format-BytesToMBorGB $selectedGpu.AdapterRAM)
    Set-GridRowValue "記憶體類型 (Video Memory Type)" (Get-MemoryTypeName $selectedGpu.VideoMemoryType)
    Set-GridRowValue "架構處理器 (Video Processor)" $selectedGpu.VideoArchitecture

    # 3. 驅動程式資訊
    Set-GridRowValue "[ 3. 驅動程式資訊 ]" "----------------------------------------" -isHeader $true
    Set-GridRowValue "驅動程式版本 (Driver Version)" $selectedGpu.DriverVersion
    Set-GridRowValue "驅動程式日期 (Driver Date)" (if ($selectedGpu.DriverDate) { [Management.ManagementDateTimeConverter]::ToDateTime($selectedGpu.DriverDate).ToString("yyyy/MM/dd") } else { "無資料" })
    Set-GridRowValue "安裝的驅動檔案 (Installed Drivers)" $selectedGpu.InstalledDisplayDrivers

    # 4. 顯示與解析度設定
    Set-GridRowValue "[ 4. 顯示與解析度設定 ]" "----------------------------------------" -isHeader $true
    $resX = $selectedGpu.CurrentHorizontalResolution
    $resY = $selectedGpu.CurrentVerticalResolution
    $refreshRate = $selectedGpu.CurrentRefreshRate
    $colorBits = $selectedGpu.CurrentBitsPerPixel

    if ($resX -and $resY) {
        Set-GridRowValue "目前解析度 (Resolution)" "$resX x $resY"
    } else {
        Set-GridRowValue "目前解析度 (Resolution)" "未連接螢幕或無資料"
    }
    Set-GridRowValue "螢幕更新率 (Refresh Rate)" (if ($refreshRate) { "$refreshRate Hz" } else { "無資料" })
    Set-GridRowValue "色彩深度 (Color Depth)" (if ($colorBits) { "$colorBits Bit" } else { "無資料" })

    # 5. 即時動態狀態 (透過 nvidia-smi，適用於 NVIDIA 卡)
    Set-GridRowValue "[ 5. 即時動態監控 ]" "----------------------------------------" -isHeader $true
    $nvInfo = Get-NvidiaDynamicInfo -gpuName $selectedGpu.Name
    if ($null -ne $nvInfo) {
        Set-GridRowValue "GPU 使用率 (Usage)" $nvInfo.Usage
        Set-GridRowValue "GPU 溫度 (Temperature)" $nvInfo.Temp
        Set-GridRowValue "顯存使用量 (VRAM Usage)" $nvInfo.Mem
    } else {
        Set-GridRowValue "GPU 使用率 (Usage)" "需特定原廠工具 API 或權限"
        Set-GridRowValue "GPU 溫度 (Temperature)" "需特定原廠工具 API 或權限"
        Set-GridRowValue "顯存使用量 (VRAM Usage)" "系統保留 / 未透出"
    }

    $grid.ResumeLayout()

    $currentTime = Get-Date -Format "yyyy/MM/dd HH:mm:ss"
    $labelStatus.Text = "最後更新時間：$currentTime (每 5 秒自動刷新)"
    $tooltipGpu.SetToolTip($comboGpus, "完整名稱：" + $selectedGpu.Name)
}

# ---------------------------------------------------------
# 6. 選單與事件掛載
# ---------------------------------------------------------
$menuPreview.Add_Click({
    $reportText = Get-GpuReportText
    $previewForm = New-Object System.Windows.Forms.Form
    $previewForm.Text = "記錄預覽 - GPU_Info"
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
        $saveDlg.Title = "儲存 GPU 資訊報告"
        $saveDlg.FileName = "GPU_Report_" + (Get-Date -Format 'yyyyMMdd_HHmmss') + ".txt"

        if ($saveDlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $reportText = Get-GpuReportText
            [System.IO.File]::WriteAllText($saveDlg.FileName, $reportText, [System.Text.Encoding]::UTF8)
            [System.Windows.Forms.MessageBox]::Show("紀錄已成功匯出至：`n" + $saveDlg.FileName, "匯出成功", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        }
    } finally {
        $saveDlg.Dispose()
    }
})

$menuExit.Add_Click({ $form.Close() })

$script:currentGpuIdx = -1

$timer = New-Object System.Windows.Forms.Timer
$timer.Interval = 5000
$timer.Add_Tick({ Display-GpuDetails })

$comboGpus.Add_SelectedIndexChanged({
    $timer.Stop()
    $script:currentGpuIdx = -1
    Display-GpuDetails
    $timer.Start()
})

$btnRefresh.Add_Click({
    $timer.Stop()
    Load-GpuList
    $script:currentGpuIdx = -1
    Display-GpuDetails
    $timer.Start()
})

$form.Add_Shown({
    Load-GpuList
    Display-GpuDetails
    $timer.Start()
})

$form.Add_FormClosing({
    $timer.Stop()
    $timer.Dispose()
})

# 啟動應用程式
[void]$form.ShowDialog()