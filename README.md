# 💻 PC-Tool (電腦工具箱)

![PowerShell](https://img.shields.io/badge/PowerShell-5.1%2B-blue.svg)
![Platform](https://img.shields.io/badge/Platform-Windows-lightgrey.svg)
![License](https://img.shields.io/badge/License-MIT-green.svg)

`PC-Tool` 是一個基於 PowerShell 與 WinForms 開發的輕量級 Windows 系統資訊檢測工具箱。提供簡潔直覺的圖形化介面 (GUI)，協助使用者即時監控與檢測電腦硬體狀態。

---

## 🚀 主要功能

* **CPU 資訊檢測 (`CPU_Info`)**
  * **基本規格**：即時讀取處理器名稱、製造商、架構、腳位 (Socket) 及裝置識別碼。
  * **核心與執行緒**：顯示實體核心數、邏輯執行緒數及超執行緒 (Hyper-Threading) 支援狀態。
  * **即時效能監控**：高效率讀取 CPU 即時使用率 (%)、負載、目前運轉時脈與電壓。
  * **快取記憶體**：詳細列出 L1、L2、L3 快取大小與規格。
  * **報告匯出**：支援將檢測結果一鍵預覽或匯出為純文字 (.txt) 報告。

* **磁碟資訊檢測 (`Disk_Info`)**
  * **硬碟狀態監控**：讀取本機磁碟容量、剩餘空間、磁碟區標籤與檔案系統格式。
  * **實體磁碟詳細數據**：支援檢測硬碟介面類型與運作健康狀態。

---

## 🛠️ 技術特色與優化

* **零套件依賴**：完全採用 Windows 原生 PowerShell 與 .NET WinForms 打造，無需安裝額外框架。
* **高解析度 (High DPI) 支援**：內建 DPI Awareness 設定，在高解析度螢幕下畫面依然清晰不模糊。
* **雙重緩衝 (Double Buffered)**：控制項經過繪製優化，視窗滑動與數據更新時零閃爍。
* **非阻塞式架構**：使用原生 `.NET PerformanceCounter` 進行即時監控，UI 介面順暢不卡頓。

---

## 📦 下載與使用方式

你可以選擇以下兩種方式來使用此工具：

### 1. 直接下載執行檔 (推薦一般使用者)
前往本專案的 [Releases](../../releases) 頁面下載最新版本：
* 下載打包好的 `.exe` 執行檔，雙擊即可直接執行（無需安裝 PowerShell 環境）。

### 2. 執行原始碼 (適合開發者)
複製或下載本專案的 `.ps1` 腳本檔案，於 PowerShell 中執行：

```powershell
# 執行 CPU 檢測工具
.\CPU_Info.ps1

# 執行磁碟檢測工具
.\Disk_Info.ps1
