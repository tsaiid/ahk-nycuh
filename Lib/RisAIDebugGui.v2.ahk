#Requires AutoHotkey v2.0

#Include .\RisAIDebug.v2.ahk
#Include .\RisDialog.v2.ahk

/**
 * 負責 AI 相關的 Debug 與比對 GUI
 */
class RisAIDebugGui {
    static comparisonGuiHwnd := 0
    static applyOriginalChoiceFunc := 0
    static applyFirstChoiceFunc := 0
    static applySecondChoiceFunc := 0
    static applyCustomChoiceFunc := 0

    /**
     * 套用三欄比對視窗中的選項 (0: 原始, 1: 第一個結果/OpenAI, 2: 第二個結果/Google AI, 3: 自訂所選組合)
     * @param {Integer} index
     */
    static ApplyPolishProviderChoice(index) {
        if (index == 0 && this.applyOriginalChoiceFunc) {
            (this.applyOriginalChoiceFunc)()
        } else if (index == 1 && this.applyFirstChoiceFunc) {
            (this.applyFirstChoiceFunc)()
        } else if (index == 2 && this.applySecondChoiceFunc) {
            (this.applySecondChoiceFunc)()
        } else if (index == 3 && this.applyCustomChoiceFunc) {
            (this.applyCustomChoiceFunc)()
        }
    }

    /**
     * 顯示完整 AI Prompt，並讓使用者確認是否繼續呼叫 API
     * @param title 視窗標題
     * @param promptText 完整 prompt
     * @param options 選項 { Notify: func, Header: string }
     */
    static ShowPromptConfirm(title, promptText, options := 0) {
        notify := (IsObject(options) && options.HasOwnProp("Notify")) ? options.Notify : (*) => 0
        header := (IsObject(options) && options.HasOwnProp("Header")) ? options.Header : "Prompt 已複製到剪貼簿。"
        result := false

        promptGui := RisDialog.Create(title, "+AlwaysOnTop +ToolWindow +Resize", {MarginX: 14, MarginY: 12})

        promptGui.Add("Text", "w860", header . "`r`n字元數: " . StrLen(promptText))
        promptEdit := promptGui.Add("Edit", "xm w860 h520 ReadOnly Multi -Wrap -WantReturn", promptText)

        btnContinue := promptGui.Add("Button", "Default w140 xm y+12", "繼續呼叫 API")
        btnContinue.OnEvent("Click", (*) => (
            result := true,
            promptGui.Destroy()
        ))

        btnCopy := promptGui.Add("Button", "w120 x+10 yp", "複製 Prompt")
        btnCopy.OnEvent("Click", (*) => (
            A_Clipboard := promptText,
            notify("已複製 Prompt", 1800)
        ))

        btnCancel := promptGui.Add("Button", "w100 x+10 yp", "取消")
        btnCancel.OnEvent("Click", (*) => promptGui.Destroy())
        promptGui.OnEvent("Close", (*) => promptGui.Destroy())
        promptGui.OnEvent("Escape", (*) => promptGui.Destroy())

        RisDialog.ShowCenter(promptGui)
        btnContinue.Focus()
        SendMessage(0x00B1, 0, 0, promptEdit.Hwnd)
        WinWaitClose("ahk_id " . promptGui.Hwnd)

        return result
    }

    /**
     * 顯示 Google AI Debug curl 視窗
     * @param url API URL
     * @param payload 請求內容
     * @param response 伺服器回應物件 {Status, ResponseText}
     * @param request 原始請求資訊物件
     * @param options 選項 { Notify: func, FontHwnd: handle }
     */
    static ShowGoogleAIDebugCurl(url, payload, response, request := 0, options := 0) {
        curlCommand := RisAIDebug.BuildGoogleCurlCommand(url, payload)
        modelText := IsObject(request) && request.HasOwnProp("Model") ? request.Model : "(unknown)"
        apiKeyName := IsObject(request) && request.HasOwnProp("APIKeyName") ? request.APIKeyName : "(unknown)"
        waitText := ""
        if (IsObject(request) && request.HasOwnProp("Metrics") && request.Metrics.HasOwnProp("WaitForResponseTime")) {
            waitText := "`r`nWaitForResponse: " . request.Metrics.WaitForResponseTime . " ms"
        }

        debugGui := RisDialog.Create("Google AI Debug - curl", "+AlwaysOnTop +ToolWindow +Resize", {MarginX: 14, MarginY: 12})

        debugGui.Add("Text", "w760", "Status: " . response.Status
            . "`r`nModel: " . modelText
            . "`r`nAPI Key: " . apiKeyName
            . waitText
            . "`r`nPayload: " . StrLen(payload) . " chars"
            . "`r`nResponse: " . StrLen(response.ResponseText) . " chars")
        curlEdit := debugGui.Add("Edit", "xm w760 h360 ReadOnly Multi -Wrap -WantReturn", curlCommand)

        notify := (IsObject(options) && options.HasOwnProp("Notify")) ? options.Notify : (*) => 0

        btnCopy := debugGui.Add("Button", "w140 xm y+12", "複製 curl")
        btnCopy.OnEvent("Click", (*) => (
            A_Clipboard := curlCommand,
            notify("已複製 curl 測試指令", 2000)
        ))

        btnCopyAll := debugGui.Add("Button", "w180 x+10 yp", "複製完整 debug")
        btnCopyAll.OnEvent("Click", (*) => (
            A_Clipboard := "URL: " . url . "`r`n`r`nPayload:`r`n" . payload . "`r`n`r`nResponse:`r`n" . response.ResponseText . "`r`n`r`nCurl:`r`n" . curlCommand,
            notify("已複製完整 debug 資訊", 2000)
        ))

        btnClose := debugGui.Add("Button", "Default w100 x+10 yp", "關閉")
        btnClose.OnEvent("Click", (*) => debugGui.Destroy())
        debugGui.OnEvent("Close", (*) => debugGui.Destroy())
        debugGui.OnEvent("Escape", (*) => debugGui.Destroy())

        RisDialog.ShowCenter(debugGui)
        btnClose.Focus()
        SendMessage(0x00B1, 0, 0, curlEdit.Hwnd)
    }

    /**
     * 顯示 AI 處理失敗的 Debug 視窗
     * @param errMsg 錯誤訊息
     * @param options 選項 { Notify: func }
     */
    static ShowDebugError(errMsg, options := 0) {
        notify := (IsObject(options) && options.HasOwnProp("Notify")) ? options.Notify : (*) => 0

        errGui := RisDialog.Create("AI Debug - 處理失敗", "+AlwaysOnTop +ToolWindow +Resize", {MarginX: 14, MarginY: 12})

        errGui.Add("Text", "w560", "API 呼叫或處理過程中發生例外錯誤：")
        errEdit := errGui.Add("Edit", "xm w560 h260 ReadOnly Multi -WantReturn", errMsg)

        btnCopy := errGui.Add("Button", "w160 xm y+12", "📋 複製完整訊息")
        btnCopy.OnEvent("Click", (*) => (
            A_Clipboard := errMsg,
            notify("已複製錯誤訊息至剪貼簿！", 2000)
        ))

        btnClose := errGui.Add("Button", "Default w100 x+10 yp", "關閉")
        btnClose.OnEvent("Click", (*) => errGui.Destroy())
        errGui.OnEvent("Close", (*) => errGui.Destroy())
        errGui.OnEvent("Escape", (*) => errGui.Destroy())

        RisDialog.ShowCenter(errGui, "AutoSize")
        btnClose.Focus()
        SendMessage(0x00B1, 0, 0, errEdit.Hwnd) ; 避免唯讀 Edit 在顯示時自動全選
    }

    /**
     * 顯示 AI 回傳結果之 Debug 視窗
     * @param title 視窗標題
     * @param resultText 回傳之文字結果
     * @param options 選項 { Notify: func, ExtractTime: int, ApiTime: int, Model: string, APIKeyName: string, Wait: bool }
     */
    static ShowResponseDebug(title, resultText, options := 0) {
        notify := (IsObject(options) && options.HasOwnProp("Notify")) ? options.Notify : (*) => 0
        extractTime := (IsObject(options) && options.HasOwnProp("ExtractTime")) ? options.ExtractTime : "-"
        apiTime := (IsObject(options) && options.HasOwnProp("ApiTime")) ? options.ApiTime : "-"
        modelName := (IsObject(options) && options.HasOwnProp("Model")) ? options.Model : "-"
        apiKeyName := (IsObject(options) && options.HasOwnProp("APIKeyName")) ? options.APIKeyName : "-"

        fullDebugText := Format(
            "【AI 資訊】`r`nModel: {1}`r`nAPI Key: {2}`r`n資料提取: {3} ms | API 耗時: {4} ms`r`n`r`n【回傳結果】`r`n{5}",
            modelName,
            apiKeyName,
            extractTime,
            apiTime,
            resultText
        )

        respGui := RisDialog.Create(title, "+AlwaysOnTop +ToolWindow +Resize", {MarginX: 14, MarginY: 12})

        statsHeader := Format(
            "Model: {1} | API Key: {2}`r`n資料提取: {3} ms | API 耗時: {4} ms",
            modelName,
            apiKeyName,
            extractTime,
            apiTime
        )
        respGui.Add("Text", "w700", statsHeader)
        respGui.Add("Text", "w700 y+8", "AI 回傳結果:")

        respEdit := respGui.Add("Edit", "xm w700 h260 ReadOnly Multi +Wrap -WantReturn", resultText)

        btnCopyAll := respGui.Add("Button", "w160 xm y+12", "📋 複製完整內容")
        btnCopyAll.OnEvent("Click", (*) => (
            A_Clipboard := fullDebugText,
            notify("已複製完整除錯資訊至剪貼簿", 1800)
        ))

        btnCopyResult := respGui.Add("Button", "w120 x+10 yp", "複製結果")
        btnCopyResult.OnEvent("Click", (*) => (
            A_Clipboard := resultText,
            notify("已複製結果至剪貼簿", 1800)
        ))

        btnClose := respGui.Add("Button", "Default w100 x+10 yp", "確定")
        btnClose.OnEvent("Click", (*) => respGui.Destroy())
        respGui.OnEvent("Close", (*) => respGui.Destroy())
        respGui.OnEvent("Escape", (*) => respGui.Destroy())

        RisDialog.ShowCenter(respGui)
        btnClose.Focus()
        SendMessage(0x00B1, 0, 0, respEdit.Hwnd)
        WinWaitClose("ahk_id " . respGui.Hwnd)
    }

    /**
     * 顯示單一 AI 潤色結果比對視窗
     */
    static ShowPolishComparisonGui(hEdit, original, refined, sel, debugInfo := "", options := 0) {
        notify := (IsObject(options) && options.HasOwnProp("Notify")) ? options.Notify : (*) => 0
        applyFont := (IsObject(options) && options.HasOwnProp("ApplyFont")) ? options.ApplyFont : (*) => 0
        onAccept := (IsObject(options) && options.HasOwnProp("OnAccept")) ? options.OnAccept : (*) => 0

        myGui := RisDialog.Create("AI 潤色結果比對", "+AlwaysOnTop +ToolWindow", {MarginX: 14, MarginY: 12})

        myGui.Add("Text", "w400", "原始文字 (Original):")
        myGui.Add("Text", "x+20 yp w400", "潤色結果 (Refined):")

        originalEdit := myGui.Add("Edit", "xm w400 r15 ReadOnly Multi -WantReturn", original)
        refinedEdit := myGui.Add("Edit", "x+20 yp w400 r15 ReadOnly Multi -WantReturn", refined)
        applyFont(originalEdit.Hwnd, refinedEdit.Hwnd)

        if IsObject(debugInfo) {
            myGui.SetFont("s10", "Consolas")
            myGui.Add("Text", "xm y+12 w260 Center", "API Key: " . debugInfo.APIKeyName)
            myGui.Add("Text", "x+20 yp w260 Center", "Model: " . debugInfo.Model)
            myGui.Add("Text", "x+20 yp w260 Center", "API Time: " . debugInfo.ApiTime)
            myGui.SetFont("s11", "Microsoft JhengHei UI")
        }

        btnAccept := myGui.Add("Button", "Default w180 x220 y+20", "✅ Accept (Enter)")
        btnReject := myGui.Add("Button", "w180 x+20", "❌ Reject (Esc)")

        handleAccept(*) {
            finalText := refinedEdit.Value
            onAccept(hEdit, finalText, sel)
            myGui.Destroy()
            notify("已更新文字")
        }

        btnAccept.OnEvent("Click", handleAccept)
        btnReject.OnEvent("Click", (*) => myGui.Destroy())
        myGui.OnEvent("Escape", (*) => myGui.Destroy())

        RisDialog.ShowCenter(myGui)

        refinedEdit.Focus()
        SendMessage(0x00B1, 0, 0, refinedEdit.Hwnd)
    }

    /**
     * 顯示雙 AI 潤色結果比對視窗 (三欄卡片式，支援段落挑選與保留原始文字)
     */
    static ShowPolishProviderComparisonGui(hEdit, original, openAIResult, googleResult, sel, options := 0) {
        notify := (IsObject(options) && options.HasOwnProp("Notify")) ? options.Notify : (*) => 0
        onAccept := (IsObject(options) && options.HasOwnProp("OnAccept")) ? options.OnAccept : (*) => 0

        MonitorGetWorkArea(, &left, &top, &right, &bottom)
        workWidth := right - left
        workHeight := bottom - top
        windowWidth := Min(1180, Floor(workWidth * 0.94))
        windowHeight := Min(680, Floor(workHeight * 0.88))

        myGui := RisDialog.Create("AI 潤色結果三欄比對", "+AlwaysOnTop +ToolWindow +Resize", {MarginX: 0, MarginY: 0})

        browserCtrl := myGui.Add("ActiveX", Format("x0 y0 w{1} h{2}", windowWidth, windowHeight), "Shell.Explorer")
        browser := browserCtrl.Value
        try {
            browser.Silent := true
        }

        monoFont := this._GetReportMonospaceFont()
        fontSize := this._GetReportFontSize()
        openAIText := (IsObject(openAIResult) && openAIResult.HasOwnProp("Text")) ? openAIResult.Text : ""
        googleText := (IsObject(googleResult) && googleResult.HasOwnProp("Text")) ? googleResult.Text : ""

        align := this._AlignParagraphs(original, openAIText, googleText)
        html := this._BuildProviderComparisonHtml(original, openAIResult, googleResult, align, monoFont, fontSize)

        browser.Navigate("about:blank")
        while (browser.Busy || browser.ReadyState < 4) {
            Sleep(10)
        }

        doc := browser.Document
        doc.Open()
        doc.Write(html)
        doc.Close()

        trailingNewlines := ""
        if RegExMatch(original, "(\r?\n)+$", &m) {
            trailingNewlines := m[0]
        }

        cleanupGui() {
            RisAIDebugGui.comparisonGuiHwnd := 0
            RisAIDebugGui.applyOriginalChoiceFunc := 0
            RisAIDebugGui.applyFirstChoiceFunc := 0
            RisAIDebugGui.applySecondChoiceFunc := 0
            RisAIDebugGui.applyCustomChoiceFunc := 0
        }

        closeGui(*) {
            cleanupGui()
            myGui.Destroy()
        }

        applyResult(finalText, label) {
            if (trailingNewlines != "" && !RegExMatch(finalText, "(\r?\n)+$")) {
                finalText .= trailingNewlines
            }
            closeGui()
            onAccept(hEdit, finalText, sel)
            notify("已套用 " . label . " 版本")
        }

        doc.parentWindow.ahkAccept := (finalText, label) => applyResult(finalText, label)
        doc.parentWindow.ahkClose := () => closeGui()

        myGui.OnEvent("Size", (gui, minMax, w, h) => (
            (minMax != -1 && browserCtrl) ? browserCtrl.Move(0, 0, w, h) : 0
        ))
        myGui.OnEvent("Close", closeGui)
        myGui.OnEvent("Escape", closeGui)

        applyOriginalChoice(*) {
            try doc.parentWindow.applyAll(0)
        }
        applyFirstChoice(*) {
            if (openAIResult.Success) {
                try doc.parentWindow.applyAll(1)
            }
        }
        applySecondChoice(*) {
            if (googleResult.Success) {
                try doc.parentWindow.applyAll(2)
            }
        }
        applyCustomChoice(*) {
            try doc.parentWindow.applyCustom()
        }

        this.comparisonGuiHwnd := myGui.Hwnd
        this.applyOriginalChoiceFunc := applyOriginalChoice
        this.applyFirstChoiceFunc := applyFirstChoice
        this.applySecondChoiceFunc := applySecondChoice
        this.applyCustomChoiceFunc := applyCustomChoice

        RisDialog.ShowCenter(myGui, Format("w{1} h{2}", windowWidth, windowHeight))
        try browserCtrl.Focus()
    }

    static _GetReportMonospaceFont() {
        try {
            risCtrl := (%("RisController")%)
            if (HasProp(risCtrl, "EnforcedFontName") && risCtrl.EnforcedFontName != "") {
                return risCtrl.EnforcedFontName
            }
        }
        return "Maple Mono Normal NF CN"
    }

    static _GetReportFontSize() {
        try {
            risCtrl := (%("RisController")%)
            if (HasProp(risCtrl, "EnforcedFontSize") && risCtrl.EnforcedFontSize > 0) {
                return risCtrl.EnforcedFontSize . "pt"
            }
        }
        return "11pt"
    }

    static _AlignParagraphs(original, openAIText, googleText) {
        origNorm := StrReplace(StrReplace(original, "`r`n", "`n"), "`r", "`n")
        openAINorm := StrReplace(StrReplace(openAIText, "`r`n", "`n"), "`r", "`n")
        googleNorm := StrReplace(StrReplace(googleText, "`r`n", "`n"), "`r", "`n")

        origTrim := Trim(origNorm, "`n")
        openAITrim := Trim(openAINorm, "`n")
        googleTrim := Trim(googleNorm, "`n")

        if (origTrim != "" && openAITrim != "" && googleTrim != "") {
            ; 1. 嘗試以雙換行分段 (\n\n+)
            if (RegExMatch(origTrim, "\n{2,}") || RegExMatch(openAITrim, "\n{2,}") || RegExMatch(googleTrim, "\n{2,}")) {
                origDbl := StrSplit(RegExReplace(origTrim, "\n{2,}", "`f"), "`f")
                openAIDbl := StrSplit(RegExReplace(openAITrim, "\n{2,}", "`f"), "`f")
                googleDbl := StrSplit(RegExReplace(googleTrim, "\n{2,}", "`f"), "`f")
                if (origDbl.Length == openAIDbl.Length && openAIDbl.Length == googleDbl.Length && origDbl.Length > 1) {
                    return { IsMulti: true, Separator: "`r`n`r`n", Orig: origDbl, OpenAI: openAIDbl, Google: googleDbl }
                }
            }

            ; 2. 嘗試以單換行分段 (\n)
            if (InStr(origTrim, "`n") || InStr(openAITrim, "`n") || InStr(googleTrim, "`n")) {
                origSingle := StrSplit(origTrim, "`n")
                openAISingle := StrSplit(openAITrim, "`n")
                googleSingle := StrSplit(googleTrim, "`n")
                if (origSingle.Length == openAISingle.Length && openAISingle.Length == googleSingle.Length && origSingle.Length > 1) {
                    return { IsMulti: true, Separator: "`r`n", Orig: origSingle, OpenAI: openAISingle, Google: googleSingle }
                }
            }
        }

        ; Fallback: 整篇模式
        return { IsMulti: false, Separator: "", Orig: [original], OpenAI: [openAIText], Google: [googleText] }
    }

    static _EscapeHtml(text) {
        text := StrReplace(text, "&", "&amp;")
        text := StrReplace(text, "<", "&lt;")
        text := StrReplace(text, ">", "&gt;")
        text := StrReplace(text, '"', "&quot;")
        text := StrReplace(text, "'", "&#39;")
        return text
    }

    static _EscapeJsString(str) {
        str := StrReplace(str, "\", "\\")
        str := StrReplace(str, "`r", "\r")
        str := StrReplace(str, "`n", "\n")
        str := StrReplace(str, '"', '\"')
        return '"' . str . '"'
    }

    static _BuildProviderComparisonHtml(original, openAIResult, googleResult, align, monoFont, fontSize := "11pt") {
        origParts := align.Orig
        openAIParts := align.OpenAI
        googleParts := align.Google
        rowCount := origParts.Length
        isMulti := align.IsMulti
        separator := align.Separator

        openAISuccess := (IsObject(openAIResult) && openAIResult.HasOwnProp("Success")) ? openAIResult.Success : false
        googleSuccess := (IsObject(googleResult) && googleResult.HasOwnProp("Success")) ? googleResult.Success : false

        defaultChoice := openAISuccess ? 1 : (googleSuccess ? 2 : 0)

        jsOrigParts := "["
        jsOpenAIParts := "["
        jsGoogleParts := "["
        jsSelections := "["
        for i, part in origParts {
            jsOrigParts .= (i > 1 ? "," : "") . this._EscapeJsString(part)
            jsOpenAIParts .= (i > 1 ? "," : "") . this._EscapeJsString(openAIParts[i])
            jsGoogleParts .= (i > 1 ? "," : "") . this._EscapeJsString(googleParts[i])
            jsSelections .= (i > 1 ? "," : "") . defaultChoice
        }
        jsOrigParts .= "]"
        jsOpenAIParts .= "]"
        jsGoogleParts .= "]"
        jsSelections .= "]"
        jsSeparator := this._EscapeJsString(separator)

        openAIDebug := this.FormatProviderDebugLine(openAIResult)
        googleDebug := this.FormatProviderDebugLine(googleResult)

        rowsHtml := ""
        loop rowCount {
            r := A_Index - 1
            rowNum := A_Index
            origText := this._EscapeHtml(origParts[A_Index])
            openAIText := this._EscapeHtml(openAIParts[A_Index])
            googleText := this._EscapeHtml(googleParts[A_Index])

            origSelected := (defaultChoice == 0)
            openAISelected := (defaultChoice == 1)
            googleSelected := (defaultChoice == 2)

            origClass := "card card-orig" . (origSelected ? " selected" : "")
            openAIClass := "card card-ai" . (!openAISuccess ? " disabled" : (openAISelected ? " selected" : ""))
            googleClass := "card card-ai" . (!googleSuccess ? " disabled" : (googleSelected ? " selected" : ""))

            rowTitleHtml := isMulti ? Format("<div class='row-title'>【段落 {1}】</div>", rowNum) : ""

            rowsHtml .= "<div class='row-container'>" . rowTitleHtml
                . "<table class='card-grid-table'><tr>"
                . "<td class='col-cell'>"
                . Format("<div class='{1}' id='card_0_{2}' onclick='selectCard({2}, 0)'>", origClass, r)
                . "<div class='card-head'>"
                . "<span class='provider-badge badge-orig'>📄 原始文字</span>"
                . Format("<span class='check-icon' id='check_0_{1}' style='visibility:{2};'>✔</span>", r, origSelected ? "visible" : "hidden")
                . "</div>"
                . Format("<div class='card-body'>{1}</div>", origText)
                . "</div>"
                . "</td>"
                . "<td class='col-cell'>"
                . Format("<div class='{1}' id='card_1_{2}' {3}>", openAIClass, r, openAISuccess ? Format("onclick='selectCard({1}, 1)'", r) : "")
                . "<div class='card-head'>"
                . "<span class='provider-badge badge-openai'>OpenAI</span>"
                . Format("<span class='check-icon' id='check_1_{1}' style='visibility:{2};'>✔</span>", r, openAISelected ? "visible" : "hidden")
                . "</div>"
                . Format("<div class='card-body'>{1}</div>", openAIText)
                . "</div>"
                . "</td>"
                . "<td class='col-cell'>"
                . Format("<div class='{1}' id='card_2_{2}' {3}>", googleClass, r, googleSuccess ? Format("onclick='selectCard({1}, 2)'", r) : "")
                . "<div class='card-head'>"
                . "<span class='provider-badge badge-google'>Google AI</span>"
                . Format("<span class='check-icon' id='check_2_{1}' style='visibility:{2};'>✔</span>", r, googleSelected ? "visible" : "hidden")
                . "</div>"
                . Format("<div class='card-body'>{1}</div>", googleText)
                . "</div>"
                . "</td>"
                . "</tr></table></div>"
        }

        html := "<!doctype html><html><head><meta http-equiv='X-UA-Compatible' content='IE=edge'>"
            . "<meta charset='utf-8'>"
            . "<style>"
            . "* { box-sizing: border-box; }"
            . "html { margin:0; padding:0; }"
            . "body { margin:0; padding:12px 18px 75px 18px; background:#F8FAFC; color:#1E293B; font-family:'Microsoft JhengHei UI','Segoe UI',sans-serif; font-size:13px; overflow-y:scroll; user-select:none; -ms-user-select:none; }"
            . ".col-header-wrapper { margin-bottom:26px; }"
            . ".col-header-box { padding:6px 10px; background:#F1F5F9; border:1px solid #E2E8F0; border-radius:6px; overflow:hidden; white-space:nowrap; text-overflow:ellipsis; }"
            . ".col-header-title { font-weight:700; color:#334155; margin-right:6px; font-size:12px; }"
            . ".col-header-meta { font-family:'Cascadia Mono',Consolas,monospace; font-size:11px; color:#64748B; }"
            . ".row-container { margin-bottom:14px; }"
            . ".row-title { font-size:12px; font-weight:700; color:#475569; margin-bottom:6px; padding-left:2px; }"
            . ".card-grid-table { width:100%; table-layout:fixed; border-collapse:separate; border-spacing:12px 0; margin:0; padding:0; }"
            . ".col-cell { width:33.333%; vertical-align:top; padding:0; }"
            . ".card { border:2px solid #CBD5E1; border-radius:8px; background:#FFFFFF; padding:10px 12px; cursor:pointer; margin:0; outline:none; }"
            . ".card:hover { border-color:#0F766E; }"
            . ".card.selected { border-color:#0F766E; background:#F0FDF4; }"
            . ".card.card-orig { background:#F8FAFC; border-color:#E2E8F0; }"
            . ".card.card-orig:hover { border-color:#64748B; }"
            . ".card.card-orig.selected { border-color:#475569; background:#F1F5F9; }"
            . ".card.disabled { opacity:0.45; filter:alpha(opacity=45); cursor:not-allowed; }"
            . ".card-head { height:22px; line-height:22px; margin-bottom:8px; overflow:hidden; clear:both; }"
            . ".provider-badge { float:left; font-size:11px; font-weight:700; padding:2px 7px; border-radius:4px; letter-spacing:0.3px; line-height:16px; margin-top:1px; }"
            . ".badge-orig { background:#E2E8F0; color:#475569; }"
            . ".badge-openai { background:#E0E7FF; color:#3730A3; }"
            . ".badge-google { background:#E0F2FE; color:#0369A1; }"
            . ".check-icon { float:right; font-size:11px; font-weight:bold; color:#0F766E; background:#CCFBF1; border-radius:9px; width:18px; height:18px; line-height:18px; text-align:center; }"
            . ".card.card-orig.selected .check-icon { color:#334155; background:#E2E8F0; }"
            . ".card-body { clear:both; font-family:'" . monoFont . "','Maple Mono CN','Cascadia Code','Consolas',monospace; font-size:" . fontSize . "; line-height:1.55; color:#1E293B; white-space:pre-wrap; word-wrap:break-word; word-break:break-word; user-select:text; -ms-user-select:text; }"
            . ".bottom-bar { position:fixed; bottom:0; left:0; right:0; width:100%; height:52px; background:#FFFFFF; border-top:1px solid #CBD5E1; z-index:9999; }"
            . ".bottom-table { width:100%; height:52px; border-collapse:collapse; }"
            . ".btn { display:inline-block; padding:6px 14px; border-radius:6px; font-size:13px; font-weight:500; font-family:'Microsoft JhengHei UI','Segoe UI',sans-serif; cursor:pointer; outline:none; margin-right:6px; }"
            . ".btn-secondary { background:#F1F5F9; border:1px solid #CBD5E1; color:#334155; }"
            . ".btn-secondary:hover { background:#E2E8F0; border-color:#94A3B8; }"
            . ".btn-secondary:disabled { opacity:0.4; filter:alpha(opacity=40); cursor:not-allowed; }"
            . ".btn-cancel { background:transparent; border:1px solid #CBD5E1; color:#64748B; }"
            . ".btn-cancel:hover { background:#F8FAFC; color:#334155; }"
            . ".btn-primary { background:#0F766E; border:1px solid #0F766E; color:#FFFFFF; font-weight:700; padding:6px 18px; }"
            . ".btn-primary:hover { background:#115E59; border-color:#115E59; }"
            . "</style>"
            . "<script>"
            . "var origParts = " . jsOrigParts . ";"
            . "var openAIParts = " . jsOpenAIParts . ";"
            . "var googleParts = " . jsGoogleParts . ";"
            . "var selections = " . jsSelections . ";"
            . "var separator = " . jsSeparator . ";"
            . "var openAISuccess = " . (openAISuccess ? "true" : "false") . ";"
            . "var googleSuccess = " . (googleSuccess ? "true" : "false") . ";"
            . "function selectCard(rowIdx, providerIdx) {"
            . "  selections[rowIdx] = providerIdx;"
            . "  for (var p = 0; p < 3; p++) {"
            . "    var card = document.getElementById('card_' + p + '_' + rowIdx);"
            . "    var check = document.getElementById('check_' + p + '_' + rowIdx);"
            . "    if (card) {"
            . "      if (p === providerIdx) {"
            . "        card.className = (p === 0 ? 'card card-orig selected' : 'card card-ai selected');"
            . "        if (check) check.style.visibility = 'visible';"
            . "      } else {"
            . "        card.className = (p === 0 ? 'card card-orig' : 'card card-ai');"
            . "        if (check) check.style.visibility = 'hidden';"
            . "      }"
            . "    }"
            . "  }"
            . "}"
            . "function applyAll(providerIdx) {"
            . "  if (providerIdx === 1 && !openAISuccess) return;"
            . "  if (providerIdx === 2 && !googleSuccess) return;"
            . "  var parts = (providerIdx === 0) ? origParts : (providerIdx === 1 ? openAIParts : googleParts);"
            . "  var label = (providerIdx === 0) ? '原始' : (providerIdx === 1 ? 'OpenAI' : 'Google AI');"
            . "  var fullText = parts.join(separator);"
            . "  if (window.ahkAccept) window.ahkAccept(fullText, label);"
            . "}"
            . "function applyCustom() {"
            . "  var resultParts = [];"
            . "  for (var i = 0; i < selections.length; i++) {"
            . "    var p = selections[i];"
            . "    if (p === 0) resultParts.push(origParts[i]);"
            . "    else if (p === 1) resultParts.push(openAIParts[i]);"
            . "    else if (p === 2) resultParts.push(googleParts[i]);"
            . "  }"
            . "  var fullText = resultParts.join(separator);"
            . "  if (window.ahkAccept) window.ahkAccept(fullText, '自訂組合');"
            . "}"
            . "function cancel() {"
            . "  if (window.ahkClose) window.ahkClose();"
            . "}"
            . "document.onkeydown = function(e) {"
            . "  e = e || window.event;"
            . "  var code = e.keyCode || 0;"
            . "  if (code === 27) { cancel(); return false; }"
            . "  if (code === 13) { applyCustom(); return false; }"
            . "  if (e.altKey) {"
            . "    if (e.key === '0' || code === 48 || code === 96) { applyAll(0); return false; }"
            . "    if (e.key === '1' || code === 49 || code === 97) { applyAll(1); return false; }"
            . "    if (e.key === '2' || code === 50 || code === 98) { applyAll(2); return false; }"
            . "  }"
            . "};"
            . "</script>"
            . "</head><body>"
            . "<div class='col-header-wrapper'>"
            . "<table class='card-grid-table'><tr>"
            . "<td class='col-cell'>"
            . "<div class='col-header-box'>"
            . "<span class='col-header-title'>原始文字 (Original)</span>"
            . "</div>"
            . "</td>"
            . "<td class='col-cell'>"
            . "<div class='col-header-box'>"
            . "<span class='col-header-title'>OpenAI</span>"
            . (openAIDebug != "" ? "<span class='col-header-meta'>" . this._EscapeHtml(openAIDebug) . "</span>" : "")
            . "</div>"
            . "</td>"
            . "<td class='col-cell'>"
            . "<div class='col-header-box'>"
            . "<span class='col-header-title'>Google AI</span>"
            . (googleDebug != "" ? "<span class='col-header-meta'>" . this._EscapeHtml(googleDebug) . "</span>" : "")
            . "</div>"
            . "</td>"
            . "</tr></table>"
            . "</div>"
            . rowsHtml
            . "<div class='bottom-bar'>"
            . "<table class='bottom-table'><tr>"
            . "<td style='text-align:left;vertical-align:middle;padding-left:18px;'>"
            . "<button class='btn btn-secondary' onclick='applyAll(0)'>保留原始 (Alt+0)</button>"
            . "<button class='btn btn-secondary' onclick='applyAll(1)' " . (!openAISuccess ? "disabled" : "") . ">全部 OpenAI (Alt+1)</button>"
            . "<button class='btn btn-secondary' onclick='applyAll(2)' " . (!googleSuccess ? "disabled" : "") . ">全部 Google (Alt+2)</button>"
            . "</td>"
            . "<td style='text-align:right;vertical-align:middle;padding-right:18px;'>"
            . "<button class='btn btn-cancel' onclick='cancel()'>取消 (Esc)</button>"
            . "<button class='btn btn-primary' onclick='applyCustom()'>✔ 套用所選組合 (Enter)</button>"
            . "</td>"
            . "</tr></table>"
            . "</div></body></html>"

        return html
    }

    static GetThreeColumnComparisonLayout() {
        MonitorGetWorkArea(, &left, &top, &right, &bottom)
        workWidth := right - left
        maxWindowWidth := Floor(workWidth * 0.9)
        marginX := 14
        columnGap := 16
        minColumnWidth := 280
        columnWidth := Floor((maxWindowWidth - (marginX * 2) - (columnGap * 2)) / 3)

        if (columnWidth < minColumnWidth) {
            columnWidth := minColumnWidth
        }

        return {
            MarginX: marginX,
            Gap: columnGap,
            ColumnWidth: columnWidth,
            WindowWidth: (columnWidth * 3) + (columnGap * 2) + (marginX * 2)
        }
    }

    static FormatProviderDebugLine(result) {
        if (!IsObject(result) || !result.HasOwnProp("DebugInfo")) {
            return ""
        }

        debugInfo := result.DebugInfo
        modelStr := (IsObject(debugInfo) && debugInfo.HasOwnProp("Model")) ? debugInfo.Model : ""
        timeStr := (IsObject(debugInfo) && debugInfo.HasOwnProp("ApiTime")) ? debugInfo.ApiTime : ""

        if (modelStr != "" && timeStr != "") {
            return "Model: " . modelStr . " | Time: " . timeStr
        } else if (modelStr != "") {
            return "Model: " . modelStr
        } else if (timeStr != "") {
            return "Time: " . timeStr
        }
        return ""
    }

    /**
     * 套用視窗樣式（Win10 移除 DWM 陰影與邊框，Win11 保留陰影但消除邊框）
     * @param hwnd 視窗控制代碼
     */
    static ApplyWindowStyle(hwnd) {
        RisDialog.ApplyWindowStyle(hwnd)
    }

    /**
     * 套用潤飾比對視窗樣式 (相容方法)
     * @param hwnd 視窗控制代碼
     */
    static ApplyPolishComparisonWindowStyle(hwnd) {
        RisDialog.ApplyWindowStyle(hwnd)
    }

    /**
     * 取得 Windows 系統 Build 編號
     * @returns {Integer}
     */
    static GetWindowsBuildNumber() {
        return RisDialog.GetWindowsBuildNumber()
    }
}

#HotIf (RisAIDebugGui.comparisonGuiHwnd && WinActive("ahk_id " . RisAIDebugGui.comparisonGuiHwnd))
!0::RisAIDebugGui.ApplyPolishProviderChoice(0)
!1::RisAIDebugGui.ApplyPolishProviderChoice(1)
!2::RisAIDebugGui.ApplyPolishProviderChoice(2)
Enter::RisAIDebugGui.ApplyPolishProviderChoice(3)
#HotIf
