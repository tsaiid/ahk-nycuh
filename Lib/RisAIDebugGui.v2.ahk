#Requires AutoHotkey v2.0

#Include .\RisAIDebug.v2.ahk
#Include .\RisDialog.v2.ahk

/**
 * 負責 AI 相關的 Debug 與比對 GUI
 */
class RisAIDebugGui {
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
     * 顯示雙 AI 潤色結果比對視窗 (三欄)
     */
    static ShowPolishProviderComparisonGui(hEdit, original, openAIResult, googleResult, sel, options := 0) {
        notify := (IsObject(options) && options.HasOwnProp("Notify")) ? options.Notify : (*) => 0
        applyFont := (IsObject(options) && options.HasOwnProp("ApplyFont")) ? options.ApplyFont : (*) => 0
        onAccept := (IsObject(options) && options.HasOwnProp("OnAccept")) ? options.OnAccept : (*) => 0

        layout := this.GetThreeColumnComparisonLayout()
        colW := layout.ColumnWidth
        gap := layout.Gap

        myGui := RisDialog.Create("AI 潤色結果三欄比對", "+AlwaysOnTop +ToolWindow +Resize", {MarginX: layout.MarginX, MarginY: 12})

        myGui.Add("Text", Format("w{}", colW), "原始文字 (Original):")
        myGui.Add("Text", Format("x+{} yp w{}", gap, colW), "OpenAI:")
        myGui.Add("Text", Format("x+{} yp w{}", gap, colW), "Google AI:")

        originalEdit := myGui.Add("Edit", Format("xm w{} r18 ReadOnly Multi -WantReturn", colW), original)
        openAIEdit := myGui.Add("Edit", Format("x+{} yp w{} r18 ReadOnly Multi -WantReturn", gap, colW), openAIResult.Text)
        googleEdit := myGui.Add("Edit", Format("x+{} yp w{} r18 ReadOnly Multi -WantReturn", gap, colW), googleResult.Text)
        applyFont(originalEdit.Hwnd, openAIEdit.Hwnd, googleEdit.Hwnd)

        myGui.SetFont("s9", "Consolas")
        myGui.Add("Text", Format("xm y+10 w{} Center", colW), "")
        myGui.Add("Text", Format("x+{} yp w{} Center", gap, colW), this.FormatProviderDebugLine(openAIResult))
        myGui.Add("Text", Format("x+{} yp w{} Center", gap, colW), this.FormatProviderDebugLine(googleResult))
        myGui.SetFont("s11", "Microsoft JhengHei UI")

        buttonWidth := Min(160, colW)
        buttonOffset := Floor((colW - buttonWidth) / 2)
        openAIButtonX := layout.MarginX + colW + gap + buttonOffset
        googleButtonX := layout.MarginX + (colW * 2) + (gap * 2) + buttonOffset
        btnUseOpenAI := myGui.Add("Button", Format("Default w{} x{} y+18", buttonWidth, openAIButtonX), "Use OpenAI")
        btnUseGoogle := myGui.Add("Button", Format("w{} x{} yp", buttonWidth, googleButtonX), "Use Google")

        if (!openAIResult.Success) {
            btnUseOpenAI.Enabled := false
        }
        if (!googleResult.Success) {
            btnUseGoogle.Enabled := false
        }

        applyResult(editCtrl, label, *) {
            finalText := editCtrl.Value
            onAccept(hEdit, finalText, sel)
            myGui.Destroy()
            notify("已套用 " . label . " 版本")
        }

        btnUseOpenAI.OnEvent("Click", applyResult.Bind(openAIEdit, "OpenAI"))
        btnUseGoogle.OnEvent("Click", applyResult.Bind(googleEdit, "Google AI"))
        myGui.OnEvent("Escape", (*) => myGui.Destroy())

        RisDialog.ShowCenter(myGui, Format("w{}", layout.WindowWidth))

        if (openAIResult.Success) {
            openAIEdit.Focus()
            SendMessage(0x00B1, 0, 0, openAIEdit.Hwnd)
        } else if (googleResult.Success) {
            googleEdit.Focus()
            SendMessage(0x00B1, 0, 0, googleEdit.Hwnd)
        }
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
        return "API Key: " . debugInfo.APIKeyName . " | Model: " . debugInfo.Model . " | Time: " . debugInfo.ApiTime
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
