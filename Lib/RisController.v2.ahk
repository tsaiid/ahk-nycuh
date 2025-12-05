#Requires AutoHotkey v2.0
#Include <UIA.v2>

class RisController {
    ; =================================================================
    ; 0. 常數定義 (Constants)
    ; =================================================================
    static MSG := {
        SETFONT:       0x0030,
        GETSEL:        0x00B0,
        SETSEL:        0x00B1,
        LINESCROLL:    0x00B6, ; [新增] 用於還原捲動位置
        SCROLLCARET:   0x00B7,
        GETLINECOUNT:  0x00BA,
        LINEINDEX:     0x00BB,
        LINELENGTH:    0x00C1,
        REPLACESEL:    0x00C2,
        LINEFROMCHAR:  0x00C9,
        GETFIRSTVISIBLELINE: 0x00CE, ; [新增] 用於取得目前視窗最上方的行號
        CLEAR:         0x0303
    }

    ; =================================================================
    ; 1. 設定區 (Configuration)
    ; =================================================================
    static WinTitle := "報告作業(frmRISReport)"
    static AbnormalWinTitle := "檢查結果(frmPos)"

    static _AbnormalBtnMap := Map(
        1,          "WindowsForms10.BUTTON.app.0.2780b98_r24_ad13",
        2,          "WindowsForms10.BUTTON.app.0.2780b98_r24_ad15",
        3,          "WindowsForms10.BUTTON.app.0.2780b98_r24_ad16",
        4,          "WindowsForms10.BUTTON.app.0.2780b98_r24_ad14",
        "Save",     {AutomationId: "btnSave"},  ; Changed to UIA
        "Cancel",   {AutomationId: "btnBack"}   ; Changed to UIA
    )

    static Selectors := Map(
        "AutoNextCheckbox", { AutomationId: "chkAutoNext" },
        "ReportSaveButton", { AutomationId: "btnReportSave" },
        "FindingEdit",      { AutomationId: "txtReport" },
        "ImpressionEdit",   { AutomationId: "txtImpression" },
        "PastAllRadio",     { AutomationId: "rdoPastALL" },
        "PastModalityRadio",{ AutomationId: "rdoClassify" },
        "PastOnlyMyRadio",  { AutomationId: "rdoPastOnlyMy" },
        "ExamnameText",     { AutomationId: "txtExamName" },
        "PastFindingText",  { AutomationId: "rtxtPastReport" },
        "PastImpressionText", { AutomationId: "rtxtPastImpression" },
        "PastReportTable",  { AutomationId: "dgvPastReport" },
        "PathoDiagnosisText", { AutomationId: "txtDiagnosist" },
        "PathoDateText",    { AutomationId: "mtxtRcpDTM" },
        "ImpressionLabel",  { AutomationId: "label2" },
        "MedRecNoLabel",    { AutomationId: "txtMRNo" },
    )

    static _SimReportMap := Map(
        "CHEST PA/AP", Map("CHEST PA/AP+LAT", 1),
        "CHEST PA/AP+LAT", Map("CHEST PA/AP", 1),
        "KUB", Map("KUB+ABD LAT", 1),
        "KUB+L-SPINE LAT(supine)", Map("L-SPINE(AP+LAT)Standing", 1),
        "WHOLE  ABDOMEN CT WITH+ WITHOUT CONTRAST", Map("WHOLE  ABDOMEN CT WITHOUT CONTRAST", 1),
        "WHOLE  ABDOMEN CT WITHOUT CONTRAST", Map("WHOLE  ABDOMEN CT WITH+ WITHOUT CONTRAST", 1),
    )

    ; =================================================================
    ; 2. 內部狀態 (State)
    ; =================================================================
    static _cache := Map()
    static _currentNotifyGui := ""
    static _compContext := {MRN: "", Date: ""}
    static _hCustomFont := 0
    static _targetImpressionHeight := 95

    ; =================================================================
    ; 3. 公開屬性 (Getters)
    ; =================================================================
    static Ris => this._GetOrUpdateNode("Ris")
    static AutoNextCheckbox => this._GetOrUpdateNode("AutoNextCheckbox")
    static ReportSaveButton => this._GetOrUpdateNode("ReportSaveButton")
    static FindingEdit => this._GetOrUpdateNode("FindingEdit")
    static ImpressionEdit => this._GetOrUpdateNode("ImpressionEdit")
    static PastAllRadio => this._GetOrUpdateNode("PastAllRadio")
    static PastModalityRadio => this._GetOrUpdateNode("PastModalityRadio")
    static PastOnlyMyRadio => this._GetOrUpdateNode("PastOnlyMyRadio")
    static ExamnameText => this._GetOrUpdateNode("ExamnameText")
    static PastFindingText => this._GetOrUpdateNode("PastFindingText")
    static PastImpressionText => this._GetOrUpdateNode("PastImpressionText")
    static PastReportTable => this._GetOrUpdateNode("PastReportTable")
    static PathoDiagnosisText => this._GetOrUpdateNode("PathoDiagnosisText")
    static PathoDateText => this._GetOrUpdateNode("PathoDateText")

    ; =================================================================
    ; 4. 系統功能 (Notify & Focus)
    ; =================================================================
    static Notify(text, duration := 1500) {
        if (this._currentNotifyGui) {
            try {
                this._currentNotifyGui.Destroy()
            }
            this._currentNotifyGui := ""
        }
        g := Gui("+AlwaysOnTop -Caption +ToolWindow +E0x20")
        g.BackColor := "333333"
        g.SetFont("s16 cWhite bold", "微軟正黑體")
        g.MarginX := 20
        g.MarginY := 20
        g.Add("Text", , text)
        g.Show("NoActivate AutoSize Center")
        this._currentNotifyGui := g

        SetTimer () => (IsObject(g) ? g.Destroy() : ""), -duration
    }

    static IsTargetFocused() {
        try {
            focusedHwnd := ControlGetFocus("A")
            if !focusedHwnd {
                return false
            }

            if (this.FindingEdit.NativeWindowHandle == focusedHwnd) {
                return true
            }
            if (this.ImpressionEdit.NativeWindowHandle == focusedHwnd) {
                return true
            }
        }
        return false
    }

    static ActivateOrToggleFocus() {
        try {
            if !WinActive(this.WinTitle) {
                WinActivate(this.WinTitle)
                if !WinWaitActive(this.WinTitle, , 2) {
                    this.Notify("找不到或無法啟用 RIS 視窗")
                    return
                }
                try {
                    this.FindingEdit.SetFocus()
                }
            } else {
                try {
                    hFocus := ControlGetFocus("A")
                } catch {
                    hFocus := 0
                }

                hFind := 0
                try {
                    hFind := this.FindingEdit.NativeWindowHandle
                }

                if (hFocus == hFind) {
                    try {
                        this.ImpressionEdit.SetFocus()
                    }
                } else {
                    try {
                        this.FindingEdit.SetFocus()
                    }
                }
            }
        } catch as err {
            this.Notify("視窗切換失敗: " err.Message)
        }
    }

    ; =================================================================
    ; 5. 報告操作 (Paste, Append, Insert)
    ; =================================================================
    static PasteToFinding(text) => this.PasteTo(this.FindingEdit, text)
    static PasteToImpression(text) => this.PasteTo(this.ImpressionEdit, text)

    static PasteTo(targetEl, text) {
        if (text == "") {
            return
        }

        text := StrReplace(text, "`r`n", "`n")
        text := StrReplace(text, "`n", "`r`n")

        isNativeSuccess := false
        try {
            hwnd := targetEl.NativeWindowHandle
            if (hwnd && (targetEl.FrameworkId == "Win32" || targetEl.FrameworkId == "WinForm")) {
                EditPaste(text, hwnd)
                isNativeSuccess := true
            }
        }

        if (isNativeSuccess) {
            return
        }

        try {
            targetEl.SetFocus()
        } catch {
            this.Notify("無法聚焦目標欄位")
            return
        }

        if (StrLen(text) < 50) {
            SendText(text)
            return
        }

        SavedClip := ClipboardAll()
        A_Clipboard := text
        if !ClipWait(1) {
            this.Notify("複製失敗")
            A_Clipboard := SavedClip
            return
        }
        SetKeyDelay 50, 50
        SendEvent "^v"
        Sleep 300
        A_Clipboard := SavedClip
    }

    static AppendPreviousReport() {
        try {
            pastImp := ControlGetText(this.PastImpressionText.NativeWindowHandle)
            pastFind := ControlGetText(this.PastFindingText.NativeWindowHandle)
            hImpEdit := this.ImpressionEdit.NativeWindowHandle
            hFindEdit := this.FindingEdit.NativeWindowHandle
        } catch {
            return
        }

        rawDate := this._GetSelectedRowValue(1)
        if (rawDate != "") {
            this.SetComparisonContext(rawDate)
        }

        AppendToEdit(hEdit, textToAppend) {
            if (textToAppend == "") {
                return
            }

            currentText := ControlGetText(hEdit)
            currentLen := StrLen(currentText)

            this._EditSetSel(hEdit, currentLen, currentLen)
            if (currentLen > 0) {
                textToAppend := "`r`n" . textToAppend
            }

            this._EditReplaceSel(hEdit, textToAppend)
            this._EditScrollCaret(hEdit)
        }

        AppendToEdit(hImpEdit, pastImp)
        AppendToEdit(hFindEdit, pastFind)

        try {
            this.FindingEdit.SetFocus()
            ; [新增] 將游標移至 Finding 開頭 (位置 0) 並捲動畫面
            this._EditSetSel(hFindEdit, 0, 0)
            this._EditScrollCaret(hFindEdit)
        }
    }

    static InsertCopiedReportDate() {
        currentMRN := this._GetCurrentMRN()
        if (this._compContext.MRN != "" && this._compContext.MRN == currentMRN) {
            SendText(this._compContext.Date)
            return
        }
        this.InsertSelectedHistoryDate()
    }

    static InsertSelectedHistoryDate() => this._InsertFromSelectedRow(1, true)
    static InsertSelectedHistoryName() => this._InsertFromSelectedRow(3, false)

    static _InsertFromSelectedRow(colIndex, needDateConvert) {
        if !this.IsTargetFocused() {
            return
        }
        foundValue := this._GetSelectedRowValue(colIndex)
        if (foundValue != "") {
            if (needDateConvert) {
                foundValue := this._ConvertRISDate(foundValue)
            }
            SendText foundValue
        }
    }

    static InsertExamNameAtCaret() {
        if !this.IsTargetFocused() {
            return false
        }
        try {
            hEdit := ControlGetFocus("A")
            rawName := ControlGetText(this.ExamnameText.NativeWindowHandle)
        } catch {
            return false
        }

        if (rawName == "") {
            return true
        }

        cleanName := StrReplace(rawName, "檢查項目: ", "")
        this._EditReplaceSel(hEdit, cleanName . ":`r`n`r`n")
        return true
    }

    ; =================================================================
    ; 6. 編輯器指令 (KillLine, Delete, Format)
    ; =================================================================

    static ClearCurrentEdit() {
        try {
            hFocus := ControlGetFocus("A")
            ControlSetText("", hFocus)
        }
    }

    static KillLine() {
        try {
            hEdit := ControlGetFocus("A")

            sel := this._EditGetSel(hEdit)
            currentPos := sel.Start
            text := ControlGetText(hEdit)

            foundPos := InStr(text, "`r", , currentPos + 1)
            targetPos := (foundPos == 0) ? StrLen(text) : foundPos - 1

            if (targetPos > currentPos) {
                this._EditSetSel(hEdit, currentPos, targetPos)
                this._EditReplaceSel(hEdit, "")
            }
        }
    }

    static DeleteCurrentLine() {
        if !this.IsTargetFocused() {
            return false
        }
        try {
            hFocus := ControlGetFocus("A")
            this._SelectLine(hFocus)
            SendMessage(this.MSG.CLEAR, 0, 0, hFocus)
        }
        return true
    }

    static DeleteWordBackward() {
        if !this.IsTargetFocused() {
            return false
        }
        try {
            hCtrl := ControlGetFocus("A")
        } catch {
            return false
        }

        this._BashDeleteAlgo(hCtrl)
        return true
    }

    static MoveCaret(mode) {
        if !this.IsTargetFocused() {
            return false
        }
        try {
            hCtrl := ControlGetFocus("A")
        } catch {
            return false
        }

        lineIdx := SendMessage(this.MSG.LINEFROMCHAR, -1, 0, hCtrl)
        lineStart := SendMessage(this.MSG.LINEINDEX, lineIdx, 0, hCtrl)
        targetPos := 0

        if (mode = "Start") {
            targetPos := lineStart
        } else if (mode = "End") {
            lineLen := SendMessage(this.MSG.LINELENGTH, lineStart, 0, hCtrl)
            targetPos := lineStart + lineLen
        }

        this._EditSetSel(hCtrl, targetPos, targetPos)
        this._EditScrollCaret(hCtrl)
        return true
    }

    static MoveCaretWord(direction) {
        if !this.IsTargetFocused() {
            return false
        }
        if (direction = "Left") {
            Send "^{Left}"
        } else {
            Send "^{Right}"
        }
        return true
    }

    ; =================================================================
    ; 7. 格式化邏輯 (Format Finding/Impression)
    ; =================================================================
    static FormatFindingText() {
        if !this.IsTargetFocused() {
            return
        }
        try {
            hEdit := this.FindingEdit.NativeWindowHandle
            examType := this._GetCurrExamType()

            switch examType {
                case "CT", "MR": this._FormatFindingForAdvanced(hEdit)
                case "CR", "US": this._FormatFindingForBasic(hEdit)
            }
        } catch as err {
            this.Notify("格式化失敗: " err.Message)
        }
    }

    static FormatImpressionText() {
        if !this.IsTargetFocused() {
            return
        }
        try {
            hEdit := this.ImpressionEdit.NativeWindowHandle
            ControlFocus(hEdit)

            this._EditSetSel(hEdit, 0, -1)
            lineCount := this._CountNonEmptyLines(hEdit)

            if (lineCount > 1) {
                this._ReorderSelectedText(, , , , hEdit)
            } else {
                this._ReorderSelectedText(true, , , , hEdit)
            }
        } catch as err {
            this.Notify("Impression 格式化失敗: " err.Message)
        }
    }

    static ReorderSelection(options := {}) {
        if !this.IsTargetFocused() {
            return
        }
        try {
            hEdit := ControlGetFocus("A")
            deOrder := options.HasOwnProp("deOrder") ? options.deOrder : false
            keepEmpty := options.HasOwnProp("keepEmpty") ? options.keepEmpty : false
            itemChar := options.HasOwnProp("itemChar") ? options.itemChar : ""
            discardSeIm := options.HasOwnProp("discardSeIm") ? options.discardSeIm : true
            this._ReorderSelectedText(deOrder, keepEmpty, itemChar, discardSeIm, hEdit)
        }
    }

    ; =================================================================
    ; 8. 其他功能 (UI 互動, 字體, 排版, 滑鼠)
    ; =================================================================
    static EnableFontEnforcer(fontName := "Cascadia Code", fontSize := 12) {
        dpiRatio := 96 / A_ScreenDPI
        adjustedSize := Round(fontSize * dpiRatio, 1)

        if (this._hCustomFont != 0) {
            DllCall("DeleteObject", "Ptr", this._hCustomFont)
            this._hCustomFont := 0
        }

        dummyGui := Gui()
        dummyGui.SetFont("s" adjustedSize, fontName)
        dummyCtrl := dummyGui.Add("Text",, "Dummy")
        this._hCustomFont := SendMessage(0x0031, 0, 0, dummyCtrl.Hwnd)

        SetTimer(ObjBindMethod(this, "_EnforceFontTask"), 1000)
    }

    static _EnforceFontTask() {
        if !WinActive(this.WinTitle) {
            return
        }
        try {
            hFind := this.FindingEdit.NativeWindowHandle
            hImp  := this.ImpressionEdit.NativeWindowHandle
        } catch {
            return
        }

        if (this._hCustomFont) {
            try {
                SendMessage(this.MSG.SETFONT, this._hCustomFont, 1, , "ahk_id " hFind)
                SendMessage(this.MSG.SETFONT, this._hCustomFont, 1, , "ahk_id " hImp)
            }
        }
        this._ApplyLayout(hFind, hImp)
    }

    static _ApplyLayout(hFind, hImp) {
        ; 1. 安全檢查：如果 Handle 為 0 或空，直接離開
        if !hFind || !hImp
            return

        dpiScale := A_ScreenDPI / 96
        targetImpH  := this._targetImpressionHeight * dpiScale
        gap         := 30 * dpiScale
        labelOffset := 25 * dpiScale

        ; 2. 加上 Try-Catch 保護：避免視窗切換瞬間抓不到位置而報錯
        try {
            ControlGetPos(&fX, &fY, &fW, &fH, hFind)
            ControlGetPos(&iX, &iY, &iW, &iH, hImp)
        } catch {
            return ; 如果抓不到位置，這次就不調整
        }

        currentBottom := iY + iH
        targetImpY := currentBottom - targetImpH
        targetFindH := (targetImpY - gap) - fY

        tolerance := 5 * dpiScale
        if (Abs(iH - targetImpH) < tolerance && Abs(iY - targetImpY) < tolerance && Abs(fH - targetFindH) < tolerance) {
            return
        }

        try {
            ControlMove(,,, targetFindH, hFind) ; 先調整上面高度，避免重疊
            ControlMove(,, iW, targetImpH, hImp)
            ControlMove(, targetImpY,,, hImp)
        }

        try {
            elLabel := this._GetOrUpdateNode("ImpressionLabel")
            if (hLabel := elLabel.NativeWindowHandle) {
                ControlMove(, targetImpY - labelOffset,,, hLabel)
            }
        }
    }

    static CopyPathologyReport() {
        try {
            dateVal := this.PathoDateText.Value
            diagVal := this.PathoDiagnosisText.Value
            if (dateVal == "" && diagVal == "") {
                throw Error("找不到病理報告內容")
            }

            reportText := this._ConvertRISDate(dateVal) . ": " . diagVal
            A_Clipboard := reportText
            this.Notify("病理報告已複製")
        } catch as err {
            this.Notify("複製失敗: " err.Message)
        }
    }

    static SwitchHistoryFilter(modeName) {
        try {
            switch modeName {
                case "All":      this.PastAllRadio.ControlClick()
                case "Modality": this.PastModalityRadio.ControlClick()
                case "My":       this.PastOnlyMyRadio.ControlClick()
            }
        } catch as err {
            this.Notify("切換失敗: " err.Message) ; [修改] 取代 ToolTip
        }
    }

    static FindAndClickSimilarReport() {
        currExamName := this._GetCleanCurrentExamName()
        if (currExamName == "") {
            return
        }

        try {
            tableEle := this.PastReportTable
            rowElements := tableEle.FindAll({ Type: 'Custom' })
            if (rowElements.Length = 0) {
                throw Error("無資料")
            }

            for rowEle in rowElements {
                cellElements := rowEle.FindAll({ Type: 'DataItem' })
                if (cellElements.Length = 0) {
                    cellElements := rowEle.FindAll({ Type: 'Custom' })
                }

                if (cellElements.Length >= 3) {
                    historyExamName := cellElements[3].Value
                    if (this._IsRelatedReport(historyExamName, currExamName)) {
                        this._ClickUIAElement(cellElements[3])
                        this.Notify("已選取: " historyExamName)
                        return
                    }
                }
            }
            this.Notify("未找到相似報告", 1000)
        } catch as err {
            this.Notify("搜尋失敗: " err.Message)
        }
    }

    static SetAutoNextState(targetState) {
        try {
            if (!!targetState != !!this.AutoNextCheckbox.ToggleState) {
                this.AutoNextCheckbox.Toggle()
            }
        }
    }

    static ClickAbnormalButton(index) {
        if !WinActive(this.AbnormalWinTitle) {
            return
        }
        if !this._AbnormalBtnMap.Has(index) {
            return
        }

        target := this._AbnormalBtnMap[index]

        try {
            if IsObject(target) {
                ; UIA Logic for Save/Cancel
                hwnd := WinExist(this.AbnormalWinTitle)
                elWindow := UIA.ElementFromHandle(hwnd)

                ; FindElement 會根據傳入的 {AutomationId: "..."} 尋找
                elBtn := elWindow.FindElement(target)

                try {
                    elBtn.Invoke()
                } catch {
                    elBtn.Click()
                }
            } else {
                ; Legacy ControlClick for ClassNNs
                ControlClick(target, this.AbnormalWinTitle)
            }
        } catch as err {
            this.Notify("按鈕操作失敗: " . err.Message)
        }
    }

    static SaveReport() {
        try {
            this.ReportSaveButton.ControlClick()
        } catch as err {
            this.Notify("存檔失敗: " err.Message)
        }
    }

    static SmartExtendSelection(direction) {
        if !this.IsTargetFocused() {
            return false
        }
        try {
            hCtrl := ControlGetFocus("A")
        } catch {
            return false
        }

        lineIdx := SendMessage(this.MSG.LINEFROMCHAR, -1, 0, hCtrl)
        lineCount := SendMessage(this.MSG.GETLINECOUNT, 0, 0, hCtrl)

        if (direction == "Up") {
            SendInput (lineIdx == 0) ? "+{Home}" : "+{Up}"
        } else if (direction == "Down") {
            SendInput (lineIdx == lineCount - 1) ? "+{End}" : "+{Down}"
        }

        return true
    }

    static HandleTripleClick() {
        static clickCount := 0
        static lastClickTime := 0
        static DoubleClickTime := DllCall("GetDoubleClickTime")

        timeSinceLast := A_TickCount - lastClickTime
        if (timeSinceLast <= DoubleClickTime) {
            clickCount++
        } else {
            clickCount := 1
        }

        lastClickTime := A_TickCount

        if (clickCount == 3) {
            clickCount := 0
            MouseGetPos , , , &hCtrl, 2
            try {
                classNN := ControlGetClassNN(hCtrl)
                if (InStr(classNN, "Edit") && !InStr(classNN, "RichEdit")) {
                    this._SelectLine(hCtrl)
                }
            }
        }
    }

    ; =================================================================
    ; 9. 內部 Helper (Low-level Helpers)
    ; =================================================================

    ; --- Edit Control 底層操作 (封裝 SendMessage) ---

    static _EditSetSel(hCtrl, startPos, endPos) {
        SendMessage(this.MSG.SETSEL, startPos, endPos, hCtrl)
    }

    static _EditReplaceSel(hCtrl, text) {
        SendMessage(this.MSG.REPLACESEL, 1, StrPtr(text), hCtrl)
    }

    static _EditScrollCaret(hCtrl) {
        SendMessage(this.MSG.SCROLLCARET, 0, 0, hCtrl)
    }

    static _EditGetSel(hCtrl) {
        StartBuf := Buffer(4, 0), EndBuf := Buffer(4, 0)
        SendMessage(this.MSG.GETSEL, StartBuf.Ptr, EndBuf.Ptr, hCtrl)
        return {Start: NumGet(StartBuf, "UInt"), End: NumGet(EndBuf, "UInt")}
    }

    static _CountNonEmptyLines(hEdit) {
        try {
            text := ControlGetText(hEdit)
        } catch {
            return 0
        }

        if (text == "") {
            return 0
        }

        lines := StrSplit(text, "`n", "`r")
        count := 0

        for line in lines {
            if (Trim(line, " `t") != "") {
                count++
            }
        }
        return count
    }

    ; --- UIA 與 元件快取 ---

    static _GetOrUpdateNode(nodeName) {
        currentHwnd := WinExist(this.WinTitle)
        if !currentHwnd {
            throw TargetError("找不到 RIS 視窗")
        }

        if this._cache.Has(nodeName) {
            el := this._cache[nodeName]
            try {
                if (nodeName = "Ris" && el.WindowId != currentHwnd) {
                    throw Error("ID Mismatch")
                }
                _ := el.ControlType ; Probe
                return el
            } catch {
                this._cache.Delete(nodeName)
                if (nodeName = "Ris") {
                    this._cache := Map()
                }
            }
        }

        if (nodeName = "Ris") {
            try {
                this._cache["Ris"] := UIA.ElementFromHandle(currentHwnd)
                return this._cache["Ris"]
            } catch as err {
                throw Error("Root Error: " err.Message)
            }
        } else {
            parent := this.Ris
            if !this.Selectors.Has(nodeName) {
                throw Error("Undefined Selector: " nodeName)
            }
            try {
                this._cache[nodeName] := parent.FindElement(this.Selectors[nodeName])
                return this._cache[nodeName]
            } catch {
                throw TargetError("Node Not Found: " nodeName)
            }
        }
    }

    static GetText(el) {
        rawText := ""
        try {
            hwnd := el.NativeWindowHandle
            if (hwnd && (el.FrameworkId = "Win32" || el.FrameworkId = "WinForm")) {
                rawText := ControlGetText(hwnd)
            }
        }
        if (rawText == "") {
            try {
                rawText := el.Value
            }
            if (rawText == "" && el.IsPatternSupported("Text")) {
                try {
                    rawText := el.DocumentRange.GetText()
                }
            }
            if (rawText == "") {
                try {
                    rawText := el.Name
                }
            }
        }
        if (rawText != "") {
            temp := StrReplace(rawText, "`r`n", "`n")
            temp := StrReplace(temp, "`r", "`n")
            return StrReplace(temp, "`n", "`r`n")
        }
        return ""
    }

    ; --- 比較 Context 與 日期 ---

    static SetComparisonContext(targetDate) {
        currentMRN := this._GetCurrentMRN()
        formattedDate := this._ConvertRISDate(targetDate)
        this._compContext.MRN := currentMRN
        this._compContext.Date := formattedDate
    }

    static GetComparisonSuffix() {
        currentMRN := this._GetCurrentMRN()
        if (this._compContext.MRN != "" && this._compContext.MRN == currentMRN) {
            return " dated " . this._compContext.Date
        }
        return ""
    }

    static _GetCurrentMRN() {
        try {
            rawText := this.GetText(this._GetOrUpdateNode("MedRecNoLabel"))
            if RegExMatch(rawText, "\d+", &match) {
                return match[0]
            }
            return rawText
        }
        return ""
    }

    static _ConvertRISDate(inputString) {
        cleanString := StrReplace(StrReplace(StrReplace(inputString, "/"), ":"), " ")
        if RegExMatch(cleanString, "^((?:19|20)\d{2})(\d{2})(\d{2})", &m) {
            return Format("{:04}-{:02}-{:02}", m[1], m[2], m[3])
        }
        if RegExMatch(cleanString, "^(\d{3})(\d{2})(\d{2})", &m) {
            return Format("{:04}-{:02}-{:02}", Integer(m[1]) + 1911, m[2], m[3])
        }
        return inputString
    }

    ; --- 表格與文字處理 Helper ---

    static _GetSelectedRowValue(colIndex) {
        static STATE_SYSTEM_SELECTED := 0x2
        try {
            tableEle := this.PastReportTable
            rowElements := tableEle.FindAll({ Type: 'Custom' })
            for rowEle in rowElements {
                if IsObject(rowEle.LegacyIAccessiblePattern) && (rowEle.LegacyIAccessiblePattern.State & STATE_SYSTEM_SELECTED) {
                    targetCell := rowEle.FindElement({ ControlType: "DataItem" }, , colIndex)
                    if IsObject(targetCell) {
                        return targetCell.Value
                    }
                    break
                }
            }
        }
        return ""
    }

    static _ReorderSelectedText(deOrder := false, keepEmptyLine := false, itemChar := "", discardSeIm := true, targetHwnd := 0) {
        sel := this._EditGetSel(targetHwnd)
        if (sel.End <= sel.Start) {
            return
        }

        fullText := ControlGetText(targetHwnd)
        selectedText := SubStr(fullText, sel.Start + 1, sel.End - sel.Start)
        if (selectedText == "") {
            return
        }

        selectedText := StrReplace(selectedText, "`r`n", "`n")
        txtAry := StrSplit(selectedText, "`n")
        finalText := ""
        startLineNo := 1
        if (RegExMatch(selectedText, "^(\d+)", &existLineNo)) {
            startLineNo := existLineNo[1]
        }

        for index, line in txtAry {
            if (!RegExMatch(line, "^\s*$")) {
                tmpText := line
                isSpine := RegExMatch(line, "^\s*[-\+\*]*\s*([Vv]arying degree|[Mm]ild).+causing:")

                if (!deOrder) {
                    orderChar := (itemChar != "" ? itemChar : startLineNo++ . ".")
                    if (isSpine && RegExMatch(line, "^\s*([-\+\*]*|-->)\s*([CcTtLl]\d{1,2}-.+$)", &m)) {
                        finalText .= "--> "
                        tmpText := m[2]
                    } else {
                        finalText .= orderChar . " "
                    }
                }
                if (itemChar == "" && discardSeIm) {
                    tmpText := RegExReplace(tmpText, "\s*\((Srs|Ser)\/Img:[\s,-\/\d;]+\)", "")
                    tmpText := RegExReplace(tmpText, "Mark L\d+:\s*", "")
                }
                finalText .= RegExReplace(tmpText, "^(\s*)((\d+\.)|([-\+\*>=])|(\(?\d+\)))?(\s*)(\w?)(.*)", "$u{7}${8}")
                finalText .= "`r`n"
            } else {
                if (keepEmptyLine) {
                    finalText .= "`r`n"
                }
            }
        }
        finalText := RTrim(finalText, "`r`n")

        ; =========================================================
        ; [修改開始] 保持 Scroll 位置邏輯
        ; =========================================================

        ; 1. 記錄替換前，畫面最上方是第幾行
        firstVisibleLineBefore := SendMessage(this.MSG.GETFIRSTVISIBLELINE, 0, 0, targetHwnd)

        ; 2. 執行文字替換 (這通常會導致 Scroll 跳動以顯示 Caret)
        this._EditReplaceSel(targetHwnd, finalText)

        ; 3. 取得替換後，畫面現在最上方是第幾行
        firstVisibleLineAfter := SendMessage(this.MSG.GETFIRSTVISIBLELINE, 0, 0, targetHwnd)

        ; 4. 計算差距並滾動回去 (EM_LINESCROLL)
        ; 參數2: 水平滾動字元數 (0)
        ; 參數3: 垂直滾動行數 (負數往上，正數往下)
        linesToScroll := firstVisibleLineBefore - firstVisibleLineAfter
        if (linesToScroll != 0) {
            SendMessage(this.MSG.LINESCROLL, 0, linesToScroll, targetHwnd)
        }
        ; =========================================================
        ; [修改結束]
        ; =========================================================
    }

    static _FormatFindingForBasic(hEdit) {
        fullText := ControlGetText(hEdit)
        if RegExMatch(fullText, "im)FINDINGS:\r?\n|:\s*\r?\n\s*\r?\n", &match) {
            startPos := match.Pos + match.Len - 1
            this._EditSetSel(hEdit, startPos, -1)
            this._ReorderSelectedText(false, true, "-", false, hEdit)
        } else {
            this.Notify("報告格式不如預期，無法自動排版")
        }
    }

    static _FormatFindingForAdvanced(hEdit) {
        fullText := ControlGetText(hEdit)
        if RegExMatch(fullText, "im)FINDINGS:\r?\n|The study shows:\r?\n\r?\n|show the following findings:\r?\n\r?\n|which revealed:\r?\n\r?\n", &match) {
            startPos := match.Pos + match.Len - 1
            endPos := -1
            if RegExMatch(fullText, "im)REMARKS?:|RECOMMENDATION:", &endMatch, startPos + 1) {
                endPos := endMatch.Pos - 1 - (endMatch.Pos > 3 ? 2 : 0)
            }

            this._EditSetSel(hEdit, startPos, endPos)
            this._ReorderSelectedText(false, false, "-", true, hEdit)
        } else {
            this.Notify("報告格式不如預期，無法自動排版")
        }
    }

    static _BashDeleteAlgo(hCtrl) {
        try {
            fullText := ControlGetText(hCtrl)
        } catch {
            return
        }

        caretPos := this._EditGetSel(hCtrl).Start
        if (caretPos == 0) {
            return
        }

        i := caretPos
        while (i > 0 && this._IsSpace(SubStr(fullText, i, 1))) {
            i--
        }
        while (i > 0 && !this._IsSpace(SubStr(fullText, i, 1))) {
            i--
        }

        this._EditSetSel(hCtrl, i, caretPos)
        this._EditReplaceSel(hCtrl, "")
    }

    static _SelectLine(hCtrl) {
        try {
            fullText := ControlGetText(hCtrl)
        } catch {
            return
        }
        if (fullText = "") {
            return
        }

        caretPos := this._EditGetSel(hCtrl).Start
        ahkCaretPos := caretPos + 1

        prevLineBreak := InStr(fullText, "`n", , ahkCaretPos, -1)
        selStart := (prevLineBreak == 0) ? 0 : prevLineBreak

        nextR := InStr(fullText, "`r", , ahkCaretPos)
        nextN := InStr(fullText, "`n", , ahkCaretPos)
        selEnd := 0

        if (nextR == 0 && nextN == 0) {
            selEnd := StrLen(fullText)
        } else if (nextR > 0 && (nextN == 0 || nextR < nextN)) {
            selEnd := (SubStr(fullText, nextR + 1, 1) == "`n") ? nextR + 1 : nextR
        } else {
            selEnd := nextN
        }

        this._EditSetSel(hCtrl, selStart, selEnd)
    }

    static _GetCleanCurrentExamName() {
        try {
            return StrReplace(ControlGetText(this.ExamnameText.NativeWindowHandle), "檢查項目: ", "")
        }
        return ""
    }

    static _GetCurrExamType() {
        name := this._GetCleanCurrentExamName()
        if (InStr(name, "CT") || InStr(name, "電腦斷層")) {
            return "CT"
        }
        if (InStr(name, "MR") || InStr(name, "磁振造影")) {
            return "MR"
        }
        if (InStr(name, "US") || InStr(name, "超音波")) {
            return "US"
        }
        return "CR"
    }

    static _IsRelatedReport(prev, curr) {
        if (prev == curr) {
            return true
        }
        if (this._SimReportMap.Has(curr) && this._SimReportMap[curr].Has(prev)) {
            return true
        }
        return false
    }

    static _IsSpace(char) => (char == " " || char == "`t" || char == "`r" || char == "`n")

    static _ClickUIAElement(el) {
        try {
            r := el.BoundingRectangle
            MouseClick("Left", r.l + (el.Location.w / 2), r.t + (el.Location.h / 2))
        } catch {
            try {
                el.Invoke()
            }
        }
    }
}