#Requires AutoHotkey v2.0

/**
 * 負責 AI 相關的 Debug 與比對 GUI
 */
class RisAIDebugGui {
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

        debugGui := Gui("+AlwaysOnTop +Resize", "Google AI Debug - curl")
        debugGui.SetFont("s10", "Microsoft JhengHei UI")
        debugGui.Add("Text", "w760", "Status: " . response.Status
            . "`r`nModel: " . modelText
            . "`r`nAPI Key: " . apiKeyName
            . waitText
            . "`r`nPayload: " . StrLen(payload) . " chars"
            . "`r`nResponse: " . StrLen(response.ResponseText) . " chars")
        curlEdit := debugGui.Add("Edit", "w760 h360 ReadOnly Multi -Wrap", curlCommand)

        notify := (IsObject(options) && options.HasOwnProp("Notify")) ? options.Notify : (*) => 0

        btnCopy := debugGui.Add("Button", "w140 x10 y+12", "複製 curl")
        btnCopy.OnEvent("Click", (*) => (
            A_Clipboard := curlCommand,
            notify("已複製 curl 測試指令", 2000)
        ))

        btnCopyAll := debugGui.Add("Button", "w180 x+10 yp", "複製完整 debug")
        btnCopyAll.OnEvent("Click", (*) => (
            A_Clipboard := "URL: " . url . "`r`n`r`nPayload:`r`n" . payload . "`r`n`r`nResponse:`r`n" . response.ResponseText . "`r`n`r`nCurl:`r`n" . curlCommand,
            notify("已複製完整 debug 資訊", 2000)
        ))

        btnClose := debugGui.Add("Button", "w100 x+10 yp", "關閉")
        btnClose.OnEvent("Click", (*) => debugGui.Destroy())
        debugGui.Show()
    }

    /**
     * 顯示 AI 處理失敗的 Debug 視窗
     * @param errMsg 錯誤訊息
     * @param options 選項 { Notify: func }
     */
    static ShowDebugError(errMsg, options := 0) {
        notify := (IsObject(options) && options.HasOwnProp("Notify")) ? options.Notify : (*) => 0

        errGui := Gui("+AlwaysOnTop +Resize", "AI Debug - 處理失敗")
        errGui.SetFont("s10", "Microsoft JhengHei UI")

        errGui.Add("Text", "w500", "API 呼叫或處理過程中發生例外錯誤：")

        errEdit := errGui.Add("Edit", "w500 h250 ReadOnly Multi vErrText", errMsg)

        btnCopy := errGui.Add("Button", "w120 x10 y+15", "📋 複製完整訊息")
        btnCopy.OnEvent("Click", (*) => (
            A_Clipboard := errMsg,
            notify("已複製錯誤訊息至剪貼簿！", 2000)
        ))

        btnClose := errGui.Add("Button", "w100 x+270", "關閉")
        btnClose.OnEvent("Click", (*) => errGui.Destroy())

        errGui.Show("AutoSize Center")
        btnClose.Focus()
        SendMessage(0x00B1, 0, 0, errEdit.Hwnd) ; 避免唯讀 Edit 在顯示時自動全選
    }

    /**
     * 顯示單一 AI 潤色結果比對視窗
     */
    static ShowPolishComparisonGui(hEdit, original, refined, sel, debugInfo := "", options := 0) {
        notify := (IsObject(options) && options.HasOwnProp("Notify")) ? options.Notify : (*) => 0
        applyFont := (IsObject(options) && options.HasOwnProp("ApplyFont")) ? options.ApplyFont : (*) => 0
        onAccept := (IsObject(options) && options.HasOwnProp("OnAccept")) ? options.OnAccept : (*) => 0

        myGui := Gui("+AlwaysOnTop +ToolWindow", "AI 潤色結果比對")
        myGui.MarginX := 14
        myGui.MarginY := 12
        myGui.BackColor := "F4F5F7"
        myGui.SetFont("s11", "Microsoft JhengHei UI")

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

        myGui.Show("Center")
        this.ApplyPolishComparisonWindowStyle(myGui.Hwnd)

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

        myGui := Gui("+AlwaysOnTop +ToolWindow +Resize", "AI 潤色結果三欄比對")
        myGui.MarginX := layout.MarginX
        myGui.MarginY := 12
        myGui.BackColor := "F4F5F7"
        myGui.SetFont("s11", "Microsoft JhengHei UI")

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

        myGui.Show(Format("w{} Center", layout.WindowWidth))
        this.ApplyPolishComparisonWindowStyle(myGui.Hwnd)

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

    static ApplyPolishComparisonWindowStyle(hwnd) {
        try {
            if (this.GetWindowsBuildNumber() < 22000) {
                renderingPolicy := 1 ; DWMNCRP_DISABLED: remove Win10 DWM shadow/edge.
                DllCall("Dwmapi\DwmSetWindowAttribute", "Ptr", hwnd, "UInt", 2, "Int*", renderingPolicy, "UInt", 4)
                return
            }

            borderColor := 0xFFFFFFFE ; DWMWA_COLOR_NONE: remove Win11 DWM outline while keeping shadow.
            DllCall("Dwmapi\DwmSetWindowAttribute", "Ptr", hwnd, "UInt", 34, "UInt*", borderColor, "UInt", 4)
        }
    }

    static GetWindowsBuildNumber() {
        return RegExMatch(A_OSVersion, "^\d+\.\d+\.(\d+)", &match) ? Integer(match[1]) : 0
    }
}
