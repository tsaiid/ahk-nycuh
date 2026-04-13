/*
    AiPolish.ahk
    Author: 蔡依達 (NYCUH Radiology)
    Version: PoC Demo

    用途:
    針對目前畫面上已選取的文字，透過 AI 進行潤色，並在接受後貼回原本視窗。

    使用方式:
    1. 在任意應用程式中先選取一段文字。
    2. 按下本檔案設定的快捷鍵，預設為 Alt+R。
    3. 腳本會先送出 Ctrl+C，把選取文字複製到 clipboard。
    4. 若 clipboard 內有文字，腳本會呼叫 AI API 進行潤色。
    5. 顯示原文與潤色結果的比較視窗。
    6. 按 Accept 後，腳本會切回原視窗並送出 Ctrl+V 貼上結果。
    7. 按 Reject 或 Esc 則取消，不會貼回。

    設定方法:
    - `AI_POLISH_HOTKEY`
      設定啟動快捷鍵，預設為 `!r`。
    - `AI_POLISH_CONFIG.CopyShortcut`
      設定複製快捷鍵，預設為 `^c`。
    - `AI_POLISH_CONFIG.PasteShortcut`
      設定貼上快捷鍵，預設為 `^v`。
    - `AI_POLISH_CONFIG.Provider`
      選擇 AI 供應商，預設為 `google`。
    - `AI_POLISH_CONFIG.Providers`
      在這裡直接填入各 provider 的 API key、model 與參數。
      目前實作為 `google`，`openai` 與 `claude` 先保留擴充結構。
    - `AI_POLISH_CONFIG.Prompt`
      這是送給 AI 的核心 system prompt，決定潤色的方向與輸出風格。
      預設沿用 RIS 的 Refine prompt，重點是把輸入文字整理成專業、流暢、結構清楚的醫學英文。
      未來若要調整，建議優先從以下方向修改:
      1. `ROLE`
         定義 AI 扮演的角色，例如 radiologist、medical editor、translator。
         如果想讓語氣更像正式報告，可強化專業角色描述。
      2. `GOAL`
         定義這個功能的主要目的，例如純潤色、翻譯、摘要、改寫成 impression 風格。
         若用途改變，先改這段，讓模型知道任務邊界。
      3. `SPECIFIC INSTRUCTIONS`
         放具體規則，例如是否保留條列、換行、編號、單位、醫學術語、縮寫風格。
         如果你發現 AI 常把格式改掉，應優先在這裡補強限制。
      4. `CONSTRAINTS`
         放硬性要求，例如只能輸出結果本文、不能加前言、不能解釋、不能閒聊。
         若 AI 回傳太多多餘內容，通常要在這裡加強。
      5. 特定情境規則
         適合放像 pulmonary nodule 這種領域特化邏輯。
         若未來有其他 radiology 常見需求，也可比照加入，但建議保持簡短且明確。
      修改原則:
      - 先小幅調整，不要一次大改整份 prompt，這樣比較容易觀察效果差異。
      - 優先用明確、可驗證的規則，不要寫模糊形容詞。
      - 若希望保留原文格式，一定要明寫 preserve line breaks / numbering / bullets。
      - 若只是想改語氣或精簡程度，先改 `GOAL` 與 `CONSTRAINTS`，通常比重寫全部 prompt 更穩定。
    - `AI_POLISH_CONFIG.Debug`
      設為 `true` 時，會把 request/response 資訊寫入 clipboard 方便除錯。

    注意事項:
    - 這是 PoC demo，設定以方便操作為優先，API key 直接寫在檔案內。
    - 依賴目標應用程式支援標準 Ctrl+C / Ctrl+V。
    - 若目標應用程式在 review 視窗跳出後失去原本選取狀態，最終貼上行為以該應用程式實作為準。
*/

#Requires AutoHotkey v2.0
#SingleInstance Force

global AI_POLISH_HOTKEY := "!r"

global AI_POLISH_CONFIG := {
    Hotkey: AI_POLISH_HOTKEY,
    PasteShortcut: "^v",
    CopyShortcut: "^c",
    Provider: "google",
    NotifyDurationMs: 1800,
    ReviewGuiTitle: "AI 潤色結果比對",
    ReviewFontName: "Microsoft JhengHei UI",
    ReviewFontSize: 11,
    Debug: false,
    Prompt: "
    (
        # ROLE
        You are a professional Radiologist and Medical Editor specializing in clinical documentation for Radiology Reports and Electronic Health Records (EHR).

        # GOAL
        Refine or translate the input text into professional, fluent, and logically structured medical English.

        # SPECIFIC INSTRUCTIONS
        1. **Clinical Fluency**: Ensure the output uses standard medical terminology and professional reporting syntax.
        2. **Format Preservation**: You MUST strictly preserve all original bullet points, numbering, and line breaks.
        3. **Special Logic (Pulmonary Nodules)**:
        - If the input describes a "pulmonary nodule" and provides two dimensions (e.g., 10 x 8 mm).
        - ACTION: Calculate the mean diameter: $\frac{length + width}{2}$.
        - FORMAT: Include the result in the sentence, e.g., "(mean diameter: 9 mm)".

        # CONSTRAINTS
        - Output ONLY the refined medical text.
        - Do NOT provide any preamble, explanations, or conversational fillers.
        - Maintain the exact hierarchical structure of the original input.
        - Use metric units as provided in the source text.
    )",
    Providers: Map(
        "google", {
            APIKey: "",
            Model: "gemma-4-31b-it",
            Temperature: 0.3,
            TopP: 0.95
        },
        "openai", {
            APIKey: "",
            BaseUrl: "https://api.openai.com/v1/responses",
            Model: "gpt-5-mini",
            Temperature: 0.3
        },
        "claude", {
            APIKey: "",
            BaseUrl: "https://api.anthropic.com/v1/messages",
            Model: "claude-3-5-sonnet-latest",
            Temperature: 0.3,
            MaxTokens: 2048
        }
    )
}

Hotkey(AI_POLISH_CONFIG.Hotkey, PolishClipboardSelectionWithAI)

return

PolishClipboardSelectionWithAI(*) {
    context := CaptureWindowContext()
    selectedText := CopyCurrentSelection()

    if (Trim(selectedText) = "") {
        Notify("沒有選取到可潤色的文字", 2200)
        return
    }

    Notify("AI 潤色中...", 3000)

    try {
        result := RequestPolish(selectedText)
        result := NormalizeLineBreaks(result)
        ShowPolishComparisonGui(context, selectedText, result)
    } catch as err {
        Notify("AI 潤色失敗: " . err.Message, 3500)
    }
}

CaptureWindowContext() {
    context := {
        WindowHwnd: 0,
        WindowTitle: ""
    }

    try context.WindowHwnd := WinGetID("A")
    try context.WindowTitle := WinGetTitle("A")

    return context
}

CopyCurrentSelection() {
    savedClip := ClipboardAll()

    try {
        A_Clipboard := ""
        SendEvent(AI_POLISH_CONFIG.CopyShortcut)
        if !ClipWait(0.8) {
            return ""
        }
        return A_Clipboard
    } finally {
        A_Clipboard := savedClip
    }
}

RequestPolish(selectedText) {
    promptText := BuildPrompt(selectedText)
    return CallAI(promptText)
}

BuildPrompt(selectedText) {
    return AI_POLISH_CONFIG.Prompt . "`n`nInput Text:`n" . selectedText
}

CallAI(promptText) {
    providerName := StrLower(AI_POLISH_CONFIG.Provider)

    switch providerName {
        case "google":
            return CallGoogleAI(promptText)
        case "openai", "claude":
            throw Error("目前尚未實作 provider: " . providerName)
        default:
            throw Error("不支援的 provider: " . providerName)
    }
}

GetProviderRuntimeConfig(providerName) {
    if !AI_POLISH_CONFIG.Providers.Has(providerName) {
        throw Error("找不到 provider 設定: " . providerName)
    }

    provider := AI_POLISH_CONFIG.Providers[providerName]
    if (!provider.HasOwnProp("APIKey") || provider.APIKey = "") {
        throw Error("請先在 AiPolish.ahk 的 Providers 設定 APIKey")
    }

    runtime := {
        APIKey: provider.APIKey,
        Model: provider.Model,
        Temperature: provider.Temperature
    }

    if provider.HasOwnProp("TopP") {
        runtime.TopP := provider.TopP
    }

    if provider.HasOwnProp("BaseUrl") {
        runtime.BaseUrl := provider.BaseUrl
    }

    if provider.HasOwnProp("MaxTokens") {
        runtime.MaxTokens := provider.MaxTokens
    }

    return runtime
}

CallGoogleAI(promptText) {
    runtime := GetProviderRuntimeConfig("google")
    url := "https://generativelanguage.googleapis.com/v1beta/models/" . runtime.Model . ":generateContent?key=" . runtime.APIKey
    payload := BuildGooglePayload(promptText, runtime)
    response := SendJsonRequest(url, payload)

    if (AI_POLISH_CONFIG.Debug) {
        A_Clipboard := "URL: " . url . "`n`nPayload: " . payload . "`n`nResponse: " . response.ResponseText
    }

    if (response.Status != 200) {
        throw Error("HTTP " . response.Status . " - " . response.ResponseText)
    }

    return ParseGoogleResponse(response.ResponseText)
}

BuildGooglePayload(promptText, runtime) {
    escapedPrompt := EscapeJsonString(promptText)

    return '{'
        . '"contents":[{'
            . '"role":"user",'
            . '"parts":[{"text":"' . escapedPrompt . '"}]'
        . '}],'
        . '"generationConfig":{'
            . '"temperature":' . runtime.Temperature . ','
            . '"thinkingConfig":{"thinkingLevel":"MINIMAL"},'
            . '"topP":' . runtime.TopP
        . '},'
        . '"tools":[{"googleSearch":{}}]'
    . '}'
}

SendJsonRequest(url, payload) {
    req := ComObject("WinHttp.WinHttpRequest.5.1")
    req.Open("POST", url, true)
    req.SetRequestHeader("Content-Type", "application/json")
    req.Send(payload)

    while !req.WaitForResponse(0.01) {
        Sleep(10)
    }

    return {
        Status: req.Status,
        ResponseText: req.ResponseText
    }
}

ParseGoogleResponse(responseText) {
    text := ExtractGoogleResponseText(responseText)
    text := DecodeGoogleResponseText(text)
    return StripMarkdownCodeFence(text)
}

ExtractGoogleResponseText(responseText) {
    combinedText := ""
    searchPos := 1

    while (searchPos := RegExMatch(responseText, 's)"text":\s*"(.*?)(?<!\\)"', &match, searchPos)) {
        val := match[1]
        context := SubStr(responseText, searchPos + match.Len, 100)
        if !RegExMatch(context, '^\s*,\s*"thought":\s*true') {
            combinedText .= val
        }
        searchPos += match.Len
    }

    if (combinedText = "") {
        throw Error("無法從 Google AI 回應中提取有效文字")
    }

    return combinedText
}

DecodeGoogleResponseText(text) {
    val := text
    val := StrReplace(val, "\n", "`n")
    val := StrReplace(val, "\r", "`r")
    val := StrReplace(val, "\t", "`t")
    val := StrReplace(val, '\"', '"')
    val := StrReplace(val, "\\", "\")

    while RegExMatch(val, "i)\\u([0-9a-f]{4})", &m) {
        val := StrReplace(val, m[0], Chr(Integer("0x" . m[1])))
    }

    return val
}

StripMarkdownCodeFence(text) {
    text := Trim(text, " `t`r`n")

    if (RegExMatch(text, "s)^```(?:\w+)?\R?(.*?)\R?```$", &m)) {
        text := m[1]
    } else if (SubStr(text, 1, 3) = "```" && SubStr(text, -2) = "```") {
        text := SubStr(text, 4, StrLen(text) - 6)
    }

    return Trim(text, " `t`r`n")
}

EscapeJsonString(text) {
    escaped := StrReplace(text, "\", "\\")
    escaped := StrReplace(escaped, "`"", "\`"")
    escaped := StrReplace(escaped, "`n", "\n")
    escaped := StrReplace(escaped, "`r", "\r")
    escaped := StrReplace(escaped, "`t", "\t")
    return escaped
}

ShowPolishComparisonGui(context, original, refined) {
    myGui := Gui("+AlwaysOnTop", AI_POLISH_CONFIG.ReviewGuiTitle)
    myGui.SetFont("s" . AI_POLISH_CONFIG.ReviewFontSize, AI_POLISH_CONFIG.ReviewFontName)

    myGui.Add("Text", "w400", "原始文字 (Original):")
    myGui.Add("Text", "x+20 yp w400", "潤色結果 (Refined):")

    myGui.Add("Edit", "xm w400 r15 ReadOnly Multi -WantReturn", original)
    refinedEdit := myGui.Add("Edit", "x+20 yp w400 r15 ReadOnly Multi -WantReturn", refined)

    btnAccept := myGui.Add("Button", "Default w180 x220 y+20", "Accept (Enter)")
    btnReject := myGui.Add("Button", "w180 x+20", "Reject (Esc)")

    handleAccept(*) {
        finalText := NormalizeLineBreaks(refinedEdit.Value)

        try {
            AcceptPolishResult(context, finalText)
            myGui.Destroy()
            Notify("已更新文字", 1800)
        } catch as err {
            Notify("貼回失敗: " . err.Message, 3500)
        }
    }

    btnAccept.OnEvent("Click", handleAccept)
    btnReject.OnEvent("Click", (*) => myGui.Destroy())
    myGui.OnEvent("Escape", (*) => myGui.Destroy())

    myGui.Show("Center")
    refinedEdit.Focus()
    SendMessage(0x00B1, 0, 0, refinedEdit.Hwnd)
}

AcceptPolishResult(context, finalText) {
    if (TryPasteBackToWindow(context, finalText)) {
        return
    }

    throw Error("無法回到原始視窗貼上結果")
}

TryPasteBackToWindow(context, finalText) {
    if (!context.WindowHwnd) {
        return false
    }

    savedClip := ClipboardAll()

    try {
        A_Clipboard := ""
        A_Clipboard := finalText
        if !ClipWait(1) {
            throw Error("無法設定剪貼簿")
        }

        WinActivate("ahk_id " . context.WindowHwnd)
        WinWaitActive("ahk_id " . context.WindowHwnd, , 1)
        Sleep(120)
        SendEvent(AI_POLISH_CONFIG.PasteShortcut)
        Sleep(150)
        return true
    } catch {
        return false
    } finally {
        A_Clipboard := savedClip
    }
}

NormalizeLineBreaks(text) {
    text := StrReplace(text, "`r`n", "`n")
    return StrReplace(text, "`n", "`r`n")
}

Notify(text, duration := 0) {
    if (!duration) {
        duration := AI_POLISH_CONFIG.NotifyDurationMs
    }

    ToolTip(text)
    SetTimer(() => ToolTip(), -duration)
}
