# ahk-nycuh

這個 repo 是一組以 **AutoHotkey v2** 撰寫的放射科工作流自動化腳本，主要針對 **NYCUH Radiology** 的 RIS / PACS / 報告撰寫場景整理。

目前內容不只是單一報告系統腳本，而是由多個可獨立執行的入口組成，涵蓋：

- RIS 報告視窗快捷鍵與 hotstrings
- INFINITT PACS 操作輔助與肺結節追蹤
- ShuttlePRO v2 控制器按鍵映射
- 肺結節追蹤 GUI 工具

## 主要入口

### `nycu.v2.ahk`

主力 RIS 腳本。

- 載入 `Lib/RisController.v2.ahk`
- 啟用大量報告用 hotstrings
- 提供 RIS 視窗內的編輯、格式整理、歷史報告操作、相似檢查切換等快捷鍵
- 整合 AI 相關功能，例如 indication / impression 生成與選取文字潤稿

### `G3PacsTools.v2.ahk`

INFINITT PACS 整合入口，可獨立編譯發布。

- 載入 PACS 閱片快捷鍵與 Calcium Score 影像 AI 功能
- 整合 `NoduleTracker.ahk` 的肺結節定位與整理工具
- 提供切換至 PACS 與關閉 PACS 的全域快捷鍵

編譯後的一般 PACS 與 NoduleTracker 功能可直接使用；Calcium Score AI 只有在目標電腦存在 `config/private.ini` 時才啟用，API key 不會包入 EXE。

### `NoduleTracker.ahk`

可獨立執行、亦可由 PACS 整合入口載入的肺結節定位與整理工具。

- 從 PACS 畫面抓取 Series / Image 編號
- 依肺葉分類整理
- 提供 GUI 匯總與複製
- 內含 probe / Acc / OCR 混合策略與 MPR 補償邏輯

### `ShuttlePROv2.v2.ahk`

ShuttlePRO v2 控制器整合腳本。

- 依目前作用中的應用程式切換按鍵映射
- 目前內建 EBM Web Viewer、RDP、INFINITT PACS、GE AWS、報告系統等情境
- 用於滾輪、同步、歷史切換與常用操作

## 目錄概覽

- `Lib/`: 共用函式庫，例如 `RisController`、`UIA`、`OCR`、`AHKHID`
- `Hotstrings/`: 各影像次專科與模板 hotstrings
- `Hotkeys/`: PACS / RIS 相關快捷鍵
- `Utilities/`: 驗證、測試與小工具腳本
- `config/`: 設定檔，目前包含相似檢查分組
- `assets/`: tray icon 與其他靜態資源
- `releases/`: 舊版或獨立釋出腳本
- `others/`: 其他輔助檔，例如 WebRIS userscript / style

## 環境需求

- Windows
- AutoHotkey v2
- 部分功能依賴醫院端實際安裝的應用程式與視窗結構，例如 RIS、INFINITT PACS、ShuttlePRO 裝置
- AI 功能需自行在腳本設定中填入 provider API key

## AI provider 設定

RIS 內建 AI 功能可在 `Lib/RisConfig.v2.ahk` 內針對 `Indication`、`Impression`、`Refine` 分別選擇 `Provider: "google"` 或 `Provider: "openai"`。

OpenAI API key 請放在 `config/private.ini` 的 `[OpenAI]` 區段，可共用 `APIKey`，或分別設定 `IndicationAPIKey`、`ImpressionAPIKey`、`RefineAPIKey`。目前預設策略為：

- `Indication` / `Refine`: `gpt-5.4-nano`
- `Impression`: `gpt-5.4-mini`

## 使用方式

依需求執行對應入口腳本即可，例如：

- `nycu.v2.ahk`
- `G3PacsTools.v2.ahk`
- `NoduleTracker.ahk`
- `ShuttlePROv2.v2.ahk`

若要做個人化調整，優先修改相鄰的設定區或 `config/` 內檔案，先沿用既有模式，不建議另外拆新架構。

## 驗證

此 repo 內建固定的 compile check，會依目前 git 變更自動挑選受影響的入口；若無法可靠判定，則會退回驗證全部入口。

若要執行 AHK unit tests：

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\Utilities\test.ps1
```

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\Utilities\compile-check.ps1
```

若要額外驗證實際編譯流程：

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\Utilities\compile-check.ps1 -Compile
```

目前預設驗證的入口為：

- `nycu.v2.ahk`
- `G3PacsTools.v2.ahk`
- `NoduleTracker.ahk`
- `ShuttlePROv2.v2.ahk`

## 注意事項

- 這些腳本高度依賴實際視窗標題、控制項名稱與院內工作流程。
- 修改 `Lib/RisController.v2.ahk`、`Hotstrings/` 或 `Hotkeys/` 後，通常會影響 `nycu.v2.ahk`。
