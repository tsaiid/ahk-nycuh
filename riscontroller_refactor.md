# RisController Refactor Notes

## 背景

`Lib\RisController.v2.ahk` 已接近 5000 行，且同時承擔多種責任：

- RIS 視窗與 UIA selector/cache/preload
- Shell hook focus、layout、font handling
- 編輯器操作、selection、formatting
- 報告文字解析與去識別化
- AI prompt orchestration
- Google/OpenAI payload、transport、response parsing
- Notify GUI、worklist/webhook

本次重構目標不是追求跨腳本重用，而是降低 `RisController` 的維護成本、讓純文字/純 helper 邏輯逐步離開 controller，同時維持 `nycu.v2.ahk` 對外呼叫方式不變。

## 已驗證不可取的方向

曾試過在 `class RisController { ... }` 內使用 `#Include` 做物理拆分，例如：

```ahk
class RisController {
    #Include .\RisController\Notify.v2.ahk
}
```

AHK compile-check 可以通過，但編輯器會把被 include 的單檔視為 class 外的孤立 method，造成語法錯誤提示。結論：

- 不採用 class 內 `#Include`。
- 不做單純物理拆分。
- 後續拆分應使用完整獨立 class/helper，並在 `RisController` top-level include 後直接委派或直接呼叫。

## 已採用原則

- 小步提交，每次只抽一個低風險責任。
- 優先抽純文字、純資料轉換、純 parser helper。
- 不先碰 UIA cache、ShellHook、Timer、Notify GUI、Clipboard、SendMessage selection mutation。
- 若 helper 已獨立，優先讓呼叫點直接使用新 class，避免保留無意義轉接 method。
- 每次 code 變更後執行：

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File Utilities\compile-check.ps1
```

## 目前已完成

目前分支：`refactor/riscontroller-physical-split`

已完成 commits：

- `6740416 refactor: 抽出 RIS 日期轉換 helper`
  - 新增 `Lib\RisDate.v2.ahk`
  - `RisDate.ConvertRISDate()` 接手 RIS 民國/西元日期轉換

- `dd8d5d2 refactor: 抽出 RIS 報告文字解析 helper`
  - 新增 `Lib\RisReportText.v2.ahk`
  - `RisReportText.FindContentRange()` 接手 Finding 內容範圍解析

- `b16cf13 refactor: 移出 RIS 文字去識別化邏輯`
  - `RisReportText.DeidentifyText()` 接手 AI prompt 前的文字去識別化

- `57ea493 refactor: 移出 RIS 檢查類型判斷`
  - `RisReportText.GetExamType()` 接手檢查名稱到 CT/MR/US/CR 的純文字分類

- `d3ed2aa refactor: 抽出 AI JSON 字串跳脫 helper`
  - 新增 `Lib\RisAIText.v2.ahk`
  - `RisAIText.EscapeJsonString()` 接手 AI payload JSON 字串跳脫

- `9ef1d4c refactor: 抽出 AI 回應 JSON 文字解碼`
  - `RisAIText.DecodeJsonEscapedText()` 接手 JSON escape 與 Unicode 還原

- `aad34b4 refactor: 抽出 AI Markdown 回應清理`
  - `RisAIText.StripMarkdownCodeFence()` 接手 AI 回應 Markdown code fence 清理

- `fba1737 refactor: 移出 AI provider 回應文字擷取`
  - `RisAIText.ExtractGoogleResponseText()`
  - `RisAIText.ExtractOpenAIResponseText()`

- `f4c8641 refactor: 移出 AI provider 回應解析流程`
  - `RisAIText.ParseGoogleResponse()`
  - `RisAIText.ParseOpenAIResponse()`

## 目前檔案邊界

### `Lib\RisDate.v2.ahk`

負責 RIS 日期字串轉換。

目前包含：

- `ConvertRISDate(inputString)`

後續可以考慮加入其他純日期格式 helper，但不要放入需要讀 UI 的邏輯。

### `Lib\RisReportText.v2.ahk`

負責 RIS 報告內容與臨床文字的純文字規則。

目前包含：

- `FindContentRange(text, mode)`
- `DeidentifyText(text)`
- `GetExamType(examName)`

適合後續放入純 report text transformation。不要放入會讀 UI control、selection、SendMessage 或 clipboard 的邏輯。

### `Lib\RisAIText.v2.ahk`

負責 AI payload/response 的純文字處理與 response parser。

目前包含：

- `EscapeJsonString(text)`
- `DecodeJsonEscapedText(text)`
- `StripMarkdownCodeFence(text)`
- `ExtractGoogleResponseText(responseText)`
- `ExtractOpenAIResponseText(responseText)`
- `ParseGoogleResponse(responseText)`
- `ParseOpenAIResponse(responseText)`

適合後續放入 AI text parsing 或 payload text helper。不要直接放 HTTP request、API key resolution、model fallback、Notify、debug GUI。

## 後續建議順序

### 1. 可繼續拆的低風險項目

#### `_ShouldRetryAIModel(status)`

可考慮移到 AI helper，例如：

```ahk
RisAIText.ShouldRetryModelStatus(status)
```

但它已經開始接近 transport policy。若要拆，命名可能不該放 `RisAIText`，而應考慮 `RisAITransportPolicy` 或 `RisAIProviderPolicy`。

#### `_GoogleAIModelSupportsThinkingLevel(modelName)`

可拆成純 policy helper，例如：

```ahk
RisAIProviderPolicy.GoogleModelSupportsThinkingLevel(modelName)
```

它是純字串/policy 判斷，風險低。不建議塞到 `RisAIText`，因為不是文字 parse，而是 provider capability。

#### `_GetAIProviderModels(aiConfig, providerName, fallbackModel := "")`

這是 config normalization，仍屬純資料處理，但和 provider config schema 有關。可晚一點拆到 `RisAIProviderPolicy` 或 `RisAIConfigResolver`。

### 2. 中等風險，建議晚一點

#### Google/OpenAI payload builder

可能目標：

```ahk
RisAIPayload.BuildGooglePayload(promptText, options)
RisAIPayload.BuildOpenAIPayload(promptText, options)
RisAIPayload.BuildGoogleUrl(options)
```

風險點：

- options schema 由 controller 內 resolver 建出。
- payload builder 與 provider config evolution 綁定。
- 應先確認 `RisAIText` 已穩定，再建立新 class。

#### AI config resolver

包含：

- `_GetGoogleAIConfig()`
- `_ResolveGoogleAIOptions()`
- `_ResolveOpenAIOptions()`
- `_Resolve...APIKey()`
- `_Resolve...ModelList()`

風險點：

- 讀 `config\private.ini`
- 依賴 `RisConfig`
- API key fallback 行為要保持精準

建議先不要急著拆。

### 3. 暫時不要拆

#### UIA cache / preload

包含：

- `_GetOrUpdateNode`
- `_ResetWindowScopedCaches`
- `_PreloadCache`
- `_PreloadStep`
- `_FinalizeUIPreload`
- `_MaybeStartIndicationPreload`

原因：

- 依賴 window lifecycle、timer、AI preload、cache invalidation。
- 行為回歸測試困難。
- 一旦拆錯，會出現 stale UIA object 或新視窗抓錯 node。

#### ShellHook focus / layout / font

包含：

- `EnableShellHookFocus`
- `_ShellMessage`
- `_FocusRisWindow`
- `_StartWindowInitialization`
- `_ApplyLayout`
- `EnableFontEnforcer`

原因：

- 事件驅動，和 RIS 視窗時序高度耦合。
- 先保持在 controller 較安全。

#### Editor selection mutation

包含：

- `_EditSetSel`
- `_EditReplaceSel`
- `_ReorderSelectedText`
- `_GetLogicalLineBoundaries`
- `_SelectLine`
- `_SelectLineForRemoval`
- `MoveCurrentLine`
- `DeleteCurrentLine`
- `CutLineOrSelection`

原因：

- 雖然部分是文字處理，但多數牽涉 `SendMessage`、scroll restoration、Win32 edit handle。
- 不應和純 `RisReportText` 混在一起。
- 若要拆，應另外設計 `RisEditControl`，但要一次處理 HWND/MSG ownership。

#### Notify GUI

原因：

- class 內 include 已證明編輯器體驗不好。
- 改成真正獨立 class 會牽涉 callback binding、GUI state、slot queue ownership。
- 可以拆，但不是下一步最安全選項。

## 建議下一步

若要繼續小步拆分，建議從 `RisAIProviderPolicy` 開始，而不是再把所有東西塞進 `RisAIText`。

候選第一步：

```ahk
Lib\RisAIProviderPolicy.v2.ahk

class RisAIProviderPolicy {
    static ShouldRetryModelStatus(status) {
        return status == 500 || status == 503
    }

    static GoogleModelSupportsThinkingLevel(modelName) {
        ...
    }
}
```

可先只拆 `_ShouldRetryAIModel(status)`，但更有價值的是一起拆 `_GoogleAIModelSupportsThinkingLevel(modelName)`。兩者都是 provider policy，不涉及 HTTP request 本身。

需要注意：

- `RisController._ShouldRetryGoogleAIModel(status)` 目前只是包 `_ShouldRetryAIModel(status)`，可一起移除或改呼叫新 helper。
- 替換後需搜尋：

```powershell
rg "_ShouldRetryAIModel|_ShouldRetryGoogleAIModel|_GoogleAIModelSupportsThinkingLevel"
```

再跑 compile-check。

## PR / review 注意事項

- 目前這個分支由多個小 commit 組成，適合 review。
- 每個 commit 都是純 refactor，理論上不改對外行為。
- 若要合併前整理 commit，可考慮保留多 commit 以方便 bisect；不一定要 squash。
- 合併前建議再跑一次 full compile-check。

## 2026-05-10 AI refactor 延伸紀錄

本段是在前述 notes 之後繼續追加的狀態整理。目標仍然是逐步把 AI 相關責任移出 `RisController`，但避免一次搬走完整 orchestration，讓每個 commit 都能獨立 review、獨立 bisect。

### 本輪新增 commits

- `b72b279 refactor: 抽出 AI provider policy`
  - 新增 `Lib\RisAIProviderPolicy.v2.ahk`
  - 移出 retry status、Google thinking level support 判斷

- `1ea53a0 refactor: 抽出 AI model 清單解析`
  - `RisAIProviderPolicy.ResolveModelList()` 接手 provider-specific / generic / fallback model list normalization

- `acc3675 refactor: 抽出 AI config 布林解析`
  - `RisAIProviderPolicy.ParseConfigBool()` 接手 `EnableGoogleSearch` 等 config bool normalization

- `17da50a refactor: 抽出 AI payload builder`
  - 新增 `Lib\RisAIPayload.v2.ahk`
  - 移出 Google URL、Google payload、OpenAI payload builder

- `7e95a94 refactor: 抽出 AI config resolver`
  - 新增 `Lib\RisAIConfigResolver.v2.ahk`
  - 移出 API key name override、INI lookup、fallback APIKey 規則
  - `RisAIProviderPolicy.ResolveProvider()` 接手 provider normalization

- `0f503f3 refactor: 抽出 AI config reader`
  - `RisAIConfigResolver.GetGoogleConfig()`
  - `RisAIConfigResolver.GetOpenAIConfig()`

- `b44029c refactor: 抽出 AI options resolver`
  - `RisAIConfigResolver.ResolveGoogleOptions()`
  - `RisAIConfigResolver.ResolveOpenAIOptions()`

- `2401326 refactor: 抽出 AI request builder`
  - 新增 `Lib\RisAIRequestBuilder.v2.ahk`
  - 移出 Google/OpenAI request object shape 與 payload build timing

- `4cec2bb refactor: 抽出 AI transport helper`
  - 新增 `Lib\RisAITransport.v2.ahk`
  - 移出 WinHttp async send、同步 wait/send response wrapper

- `dc63824 refactor: 抽出 AI debug helper`
  - 新增 `Lib\RisAIDebug.v2.ahk`
  - 移出 Google curl command builder、PowerShell string escape、blocking metrics log

- `eb3e2f5 refactor: 抽出 Google AI model health`
  - 新增 `Lib\RisAIModelHealth.v2.ahk`
  - 移出 Google model HTTP error window、degraded model 判斷與 fallback ordering

- `0536042 refactor: 抽出 AI orchestration helper`
  - 新增 `Lib\RisAIOrchestration.v2.ahk`
  - 移出 AI result normalization、notification text、request object、polish/refine result formatting

- `22678b2 refactor: 下放 AI refine 資料 helper`
  - 下放 AI response result object、refine config clone、provider refine request、retry 判斷

- `058c87c refactor: 下放 AI provider response 建構`
  - 下放 refine provider transport response object、provider response parser、provider response result object

- `e08a156 refactor: 移除 AI orchestration 轉接 wrapper`
  - 移除 controller 內只做一行委派的 AI orchestration wrappers

- `c6565fa refactor: 下放 AI provider call result`
  - blocking Google/OpenAI provider call 改用共用 provider response parser 與 call result builder

- `ee94d19 refactor: 移除 AI provider 轉接 wrapper`
  - 移除 `_GetAIProvider()` 與 `_LogGoogleAIBlockingMetrics()` thin wrappers
  - 呼叫點直接使用 `RisAIProviderPolicy` / `RisAIDebug`

- `xxxxxxx refactor: 抽出 AI debug 與比對 GUI`
  - 新增 `Lib\RisAIDebugGui.v2.ahk`
  - 移出 Google AI debug curl, AI debug error window, Polish comparison (single & three-column) GUIs
  - 移除 `RisController` 內相關的 layout/style helpers

- `xxxxxxx refactor: 建立 AI service facade`
  - 新增 `Lib\RisAIService.v2.ahk`
  - 移出 AI provider call loop, parallel refine tasks, config resolution cache
  - 移除 `RisController` 內大部份 AI 調度邏輯，僅保留與 UI 深度耦合的部分

### 目前新增檔案邊界

#### `Lib\RisAIProviderPolicy.v2.ahk`

負責 AI provider 相關 policy 與輕量 normalization。

目前包含：

- `ResolveProvider(aiConfig, defaultProvider := "google")`
- `ResolveModelList(aiConfig, providerName, fallbackModel := "")`
- `ShouldRetryModelStatus(status)`
- `ParseConfigBool(value, defaultValue := false)`
- `GoogleModelSupportsThinkingLevel(modelName)`

不要放入 HTTP request、GUI、Notify 或完整 orchestration。

#### `Lib\RisAIConfigResolver.v2.ahk`

負責 AI config file reader、API key lookup、provider options normalization。

目前包含：

- `GetGoogleConfig(configFile := "config\private.ini")`
- `GetOpenAIConfig(configFile := "config\private.ini")`
- `ResolveGoogleOptions(cfg, aiConfig := 0, modelOverride := "")`
- `ResolveOpenAIOptions(cfg, aiConfig := 0, modelOverride := "")`
- `ResolveAPIKey(cfg, aiConfig, sectionName)`

注意：Google config cache 仍留在 `RisController._googleAIConfig` wrapper。這是刻意保留的行為邊界。

#### `Lib\RisAIPayload.v2.ahk`

負責 provider payload/url string builder。

目前包含：

- `BuildGoogleUrl(options)`
- `BuildGooglePayload(promptText, options)`
- `BuildOpenAIPayload(promptText, options)`

依賴 `RisAIText.EscapeJsonString()`。不要放 transport 或 config resolver。

#### `Lib\RisAIRequestBuilder.v2.ahk`

負責 provider request object shape 與 payload build timing。

目前包含：

- `BuildGoogleRequest(promptText, options, configTime)`
- `BuildOpenAIRequest(promptText, options, configTime)`

`ConfigReadTime` 仍由 controller wrapper 計算後傳入。

#### `Lib\RisAITransport.v2.ahk`

負責 WinHttp transport。

目前包含：

- `WaitForResponse(req)`
- `SendGoogleAsync(url, payload)`
- `SendGoogle(url, payload)`
- `SendOpenAIAsync(request)`
- `SendOpenAI(request)`

不負責 provider fallback、Notify、debug GUI、model health。

#### `Lib\RisAIDebug.v2.ahk`

負責純 debug helper。

目前包含：
- `EscapePowerShellSingleQuotedString(text)`
- `BuildGoogleCurlCommand(url, payload)`
- `LogGoogleBlockingMetrics(metrics, status := "")`

#### `Lib\RisAIDebugGui.v2.ahk`

負責 AI 相關的 Debug 與比對 GUI。

目前包含：
- `ShowGoogleAIDebugCurl(url, payload, response, request, options)`
- `ShowDebugError(errMsg, options)`
- `ShowPolishComparisonGui(hEdit, original, refined, sel, debugInfo, options)`
- `ShowPolishProviderComparisonGui(hEdit, original, openAIResult, googleResult, sel, options)`

依賴 `Notify` 與 `ApplyFont` callback。

#### `Lib\RisAIService.v2.ahk`

負責 AI provider 的調度、重試與傳輸協調。

目前包含：
- `Call(promptText, aiConfig, options)`
- `CallGoogle(promptText, aiConfig, options)`
- `CallOpenAI(promptText, aiConfig, options)`
- `RunRefineProvidersParallel(selectedText, providerSpecs, trailingNewlines, options)`

並行任務調度邏輯已移入此 class。Google config cache 仍由 `RisController._googleAIConfig` 持有，透過 `GetGoogleConfig` callback 傳入，避免 service facade 擁有 controller 原本的設定快取生命週期。

`RisAIService` 不直接依賴 GUI class；Google curl debug 視窗透過 `ShowGoogleDebugCurl` callback 由 controller 端處理。

#### `Lib\RisAIModelHealth.v2.ahk`

負責 Google model health state 與 degraded model ordering。

目前包含：

- `GetGooglePolicy()`
- `PruneGoogleHttpErrors(modelName, nowTick, policy)`
- `RecordGoogleHttpError(modelName, policy)`
- `IsGoogleModelDegraded(modelName, policy)`
- `PrioritizeGoogleModels(models, policy)`

內部持有 `_googleModelHttpErrors`。controller 已不再持有 `_googleAIModelHttpErrors`。

#### `Lib\RisAIOrchestration.v2.ahk`

負責不直接操作 UI 的 AI orchestration data helper。

目前包含：

- `NormalizeResult(result)`
- `FormatCompleteNotify(title, apiKeyName, modelName, detail := "")`
- `NormalizePolishResult(result, trailingNewlines := "")`
- `CreateRequest(promptText, aiConfig, extraFields := 0)`
- `BuildRequestResult(response, apiTime)`
- `CloneConfigWithProvider(aiConfig, providerName)`
- `BuildRefineRequest(baseConfig, selectedText, providerName)`
- `ShouldRetryRefineProviderTask(task)`
- `BuildTransportResponse(req)`
- `ParseProviderResponse(providerName, responseText)`
- `BuildProviderResponseResult(parsed, request, apiTime, providerName)`
- `BuildProviderCallResult(parsed, request, providerName)`
- `FormatPolishComparisonDebugInfo(response)`
- `BuildRefineProviderSuccessResult(displayName, response, trailingNewlines := "")`
- `BuildRefineProviderFailureResult(displayName, message)`

不要把 GUI、cursor、Notify、clipboard、selection mutation 放進此 helper。

### `RisController` 目前仍保留的 AI 責任

以下仍在 controller，屬於目前比較高耦合或較高風險的部分：

- Foreground AI state：
  - `_BeginForegroundAIRequest()`
  - `_FinishForegroundAIRequest()`
  - `_isAIPending`

- Indication preload / pending insert / cache：
  - `_TryInsertCachedIndication()`
  - `_TryHandlePendingIndication()`
  - `_BeginIndicationRequest()`
  - `_CacheIndicationResult()`
  - `_FinishIndicationRequest()`
  - `_aiCache["_AI_Indication"]`
  - `_isIndicationPending`
  - `_pendingIndicationInsert`

- UI extraction / insertion / selection：
  - `_BuildIndicationRequest()`
  - `_BuildImpressionRequest()`
  - `_GetPolishSelectionContext()`
  - `_InsertAIResult()`
  - `_InsertAIResultToImpression()`
  - all edit control mutation

- AI service callback ownership：
  - `_CallAI()` thin wrapper
  - `_GetAIServiceOptions()`
  - `_GetGoogleAIConfig()`
  - `_GetOpenAIConfig()`
  - `_ShowGoogleAIDebugCurl()`
  - `_ShowPolishComparisonGui()`
  - `_ShowPolishProviderComparisonGui()`
  - `_ShowDebugError()`

### 建議下一步

目前低風險純 helper 大多已抽完。下一步若要繼續，需要開始選擇較大的責任邊界。

建議順序：

1. **穩定 `RisAIService` callback 邊界**
   - 避免 service 直接呼叫 GUI、Notify 或 controller mutation。
   - 若新增需要 UI side effect 的功能，優先從 controller 傳入 callback。
   - 保留 `_GetGoogleAIConfig()` 在 controller，除非另有明確理由改變 cache lifecycle。

2. **Indication / Impression orchestration 暫時不要拆**
   - 牽涉 UIA preload、cache、cursor、pending insert、editor insertion。
   - 若要拆，應等 AI transport/provider service 穩定後再處理。

3. **Selection/editor mutation 仍不要和 AI helper 混在一起**
   - `_GetPolishSelectionContext()` 雖然服務 AI，但它依賴 edit control selection、line boundary、通知訊息。
   - 未來若拆，應進 `RisEditControl` 或 `RisEditorSelection`，不是 `RisAI*`。

### 下次接續時建議先搜尋

```powershell
rg -n "_CallAI|_GetAIServiceOptions|_GetGoogleAIConfig|_ShowGoogleAIDebugCurl|RisAIService" Lib\RisController.v2.ahk Lib\RisAIService.v2.ahk
```

以及：

```powershell
rg -n "RisAI[A-Za-z]+\\." Lib\RisController.v2.ahk
```

每次 code 變更後仍必須執行：

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File Utilities\compile-check.ps1
```

## 2026-05-10 Editor refactor 接續紀錄

AI 主要責任已拆到 helper/service 後，下一個降低 `RisController` 維護成本的方向改為 editor control 操作邊界。

### 本輪新增狀態

- 新增 `Lib\RisEditControl.v2.ahk`
- `RisController` top-level include `RisEditControl`
- 先只下放 Win32 Edit control 的低階 primitive：
  - `GetSel(hCtrl)`
  - `SetSel(hCtrl, startPos, endPos)`
  - `ReplaceSel(hCtrl, text)`
  - `ScrollCaret(hCtrl)`
  - `ReplaceSelectionAndScroll(hCtrl, text)`
  - `GetFirstVisibleLine(hCtrl)`
  - `LineScroll(hCtrl, lineCount, columnCount := 0)`
  - `LineFromChar(hCtrl, charPos := -1)`
  - `GetLineCount(hCtrl)`
  - `GetLogicalLineBoundaries(hCtrl, specificPos := -1)`
  - `SelectLine(hCtrl)`
  - `SelectLineForRemoval(hCtrl)`
  - `InsertNewLine(hCtrl, mode := "Below")`
  - `KillLine(hCtrl)`
  - `DeleteCurrentLine(hCtrl)`
  - `CutLineOrSelection(hCtrl)`
  - `CopyLineOrSelection(hCtrl)`
  - `MoveCaret(hCtrl, mode)`

`RisController` 目前仍保留 `_EditGetSel()` / `_EditSetSel()` / `_EditReplaceSel()` / `_ReplaceSelectionAndScroll()` / `_GetLogicalLineBoundaries()` / `_SelectLine()` / `_SelectLineForRemoval()` wrappers，目的是維持既有呼叫點穩定，避免同一 commit 同時處理大量 call-site churn。

### `Lib\RisEditControl.v2.ahk`

負責 Win32 Edit control 的低階 SendMessage 封裝。

適合放入：
- Edit selection / replacement primitive
- scroll primitive
- line index / line count primitive
- 純 edit-control line boundary / select line helper
- 不依賴 focus/clipboard/Notify 的 edit mutation helper

暫時不要放入：
- RIS UIA control lookup
- Notify / GUI
- AI selection context
- clipboard workflow
- controller foreground/pending state

### 建議下一步

1. **再評估搬高階 editor command**
   - `MoveCurrentLine(direction)`

高階 command 會碰到 target focus、clipboard、Notify 或 WinActive 行為，建議等 primitive 與 line boundary 穩定後再處理。
