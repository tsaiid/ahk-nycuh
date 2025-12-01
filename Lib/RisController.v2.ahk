#Requires AutoHotkey v2.0

#Include <UIA.v2>

class RisController {
    ; =================================================================
    ; 1. 設定區 (Configuration)
    ; =================================================================
    static WinTitle := "報告作業(frmRISReport)"
    static AbnormalWinTitle := "檢查結果(frmPos)"

    ; 定義按鈕對照表 (Key: 邏輯編號, Value: ClassNN)
    ; 注意：WindowsForms10 的 ClassNN 後綴 (如 _ad13) 可能會隨軟體更新而變動
    ; 如果未來失效，建議改用 UIA 抓 AutomationId
    static _AbnormalBtnMap := Map(
        1,      "WindowsForms10.BUTTON.app.0.2780b98_r24_ad13",
        2,      "WindowsForms10.BUTTON.app.0.2780b98_r24_ad15",
        3,      "WindowsForms10.BUTTON.app.0.2780b98_r24_ad16",
        4,      "WindowsForms10.BUTTON.app.0.2780b98_r24_ad14",
        "Save", "WindowsForms10.BUTTON.app.0.2780b98_r24_ad12"
    )

    ; 定義所有子元件的搜尋條件
    static Selectors := Map(
        "AutoNextCheckbox", { AutomationId: "chkAutoNext" },
        "ReportSaveButton", { AutomationId: "btnReportSave" },
        "FindingEdit", { AutomationId: "txtReport" },
        "ImpressionEdit", { AutomationId: "txtImpression" },
        "PastAllRadio", { AutomationId: "rdoPastALL" },
        "PastModalityRadio", { AutomationId: "rdoClassify" },
        "PastOnlyMyRadio", { AutomationId: "rdoPastOnlyMy" },
        "ExamnameText", { AutomationId: "txtExamName" },
        "PastFindingText", { AutomationId: "rtxtPastReport" },
        "PastImpressionText", { AutomationId: "rtxtPastImpression" },
        "PastReportTable", { AutomationId: "dgvPastReport" },
        "PathoDiagnosisText", { AutomationId: "txtDiagnosist" },
        "PathoDateText", { AutomationId: "mtxtRcpDTM" },
        "ImpressionLabel", { AutomationId: "label2" },
    )

    ; =================================================================
    ; 2. 內部狀態 (State)
    ; =================================================================
    static _cache := Map()
    static _currentNotifyGui := "" ; [新增] 用來追蹤當前的通知視窗

    ; =================================================================
    ; 3. 公開屬性 (Public Properties)
    ; =================================================================
    static Ris {
        get => this._GetOrUpdateNode("Ris")
    }
    static AutoNextCheckbox {
        get => this._GetOrUpdateNode("AutoNextCheckbox")
    }
    static ReportSaveButton {
        get => this._GetOrUpdateNode("ReportSaveButton")
    }
    static FindingEdit {
        get => this._GetOrUpdateNode("FindingEdit")
    }
    static ImpressionEdit {
        get => this._GetOrUpdateNode("ImpressionEdit")
    }
    static PastAllRadio {
        get => this._GetOrUpdateNode("PastAllRadio")
    }
    static PastModalityRadio {
        get => this._GetOrUpdateNode("PastModalityRadio")
    }
    static PastOnlyMyRadio {
        get => this._GetOrUpdateNode("PastOnlyMyRadio")
    }
    static ExamnameText {
        get => this._GetOrUpdateNode("ExamnameText")
    }
    static PastFindingText {
        get => this._GetOrUpdateNode("PastFindingText")
    }
    static PastImpressionText {
        get => this._GetOrUpdateNode("PastImpressionText")
    }
    static PastReportTable {
        get => this._GetOrUpdateNode("PastReportTable")
    }
    static PathoDiagnosisText {
        get => this._GetOrUpdateNode("PathoDiagnosisText")
    }
    static PathoDateText {
        get => this._GetOrUpdateNode("PathoDateText")
    }

    ; =================================================================
    ; [新增功能] 系統通知 (取代 ToolTip)
    ; =================================================================
    static Notify(text, duration := 1500) {
        ; 如果有舊的通知視窗，先銷毀它 (避免重疊)
        if (this._currentNotifyGui) {
            try this._currentNotifyGui.Destroy()
            this._currentNotifyGui := ""
        }

        g := Gui("+AlwaysOnTop -Caption +ToolWindow +E0x20") ; E0x20 讓滑鼠穿透
        g.BackColor := "333333"
        g.SetFont("s16 cWhite bold", "微軟正黑體")

        g.MarginX := 20
        g.MarginY := 20

        g.Add("Text", , text)
        g.Show("NoActivate AutoSize Center")

        ; 記錄當前 GUI
        this._currentNotifyGui := g

        ; 時間到自動銷毀
        SetTimer () => (IsObject(g) ? g.Destroy() : ""), -duration
    }

    ; =================================================================
    ; 4. 核心邏輯 (Core Logic)
    ; =================================================================
    static _GetOrUpdateNode(nodeName) {
        currentHwnd := WinExist(this.WinTitle)
        if !currentHwnd
            throw TargetError("找不到 RIS 視窗，請確認程式已開啟。")

        if this._cache.Has(nodeName) {
            el := this._cache[nodeName]
            try {
                if (nodeName = "Ris") {
                    if (el.WindowId != currentHwnd)
                        throw Error("視窗 ID 不匹配")
                }
                temp := el.ControlType
                return el
            }
            catch {
                this._cache.Delete(nodeName)
                if (nodeName = "Ris")
                    this._cache := Map()
            }
        }

        if (nodeName = "Ris") {
            try {
                newRoot := UIA.ElementFromHandle(currentHwnd)
                this._cache["Ris"] := newRoot
                return newRoot
            } catch as err {
                throw Error("無法取得 RIS Root Element: " err.Message)
            }
        }
        else {
            parent := this.Ris
            if !this.Selectors.Has(nodeName)
                throw Error("未定義: " nodeName)

            try {
                newChild := parent.FindElement(this.Selectors[nodeName])
                this._cache[nodeName] := newChild
                return newChild
            } catch {
                throw TargetError("找不到元件: " nodeName)
            }
        }
    }

    ; =================================================================
    ; 通用文字取得方法
    ; =================================================================
    static GetText(el) {
        rawText := ""
        isNativeSuccess := false

        try {
            hwnd := el.NativeWindowHandle
            if (hwnd && (el.FrameworkId = "Win32" || el.FrameworkId = "WinForm")) {
                rawText := ControlGetText(hwnd)
                isNativeSuccess := true
            }
        }

        if (!isNativeSuccess) {
            try {
                rawText := el.Value
            }
            if (rawText == "" && el.IsPatternSupported("Text")) {
                try rawText := el.DocumentRange.GetText()
            }
            if (rawText == "") {
                try rawText := el.Name
            }
        }

        if (rawText != "") {
            temp := StrReplace(rawText, "`r`n", "`n")
            temp := StrReplace(temp, "`r", "`n")
            return StrReplace(temp, "`n", "`r`n")
        }
        return ""
    }

    ; =================================================================
    ; 貼上功能
    ; =================================================================
    static PasteToFinding(text) {
        this.PasteTo(this.FindingEdit, text)
    }

    static PasteToImpression(text) {
        this.PasteTo(this.ImpressionEdit, text)
    }

    static PasteTo(targetEl, text)
    {
        if (text == "")
            return

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

        if (isNativeSuccess)
            return

        try {
            targetEl.SetFocus()
        } catch {
            this.Notify("無法聚焦目標欄位，貼上失敗")
            return
        }

        if (StrLen(text) < 50) {
            SendText(text)
            return
        }

        SavedClip := ClipboardAll()
        A_Clipboard := ""
        A_Clipboard := text

        if !ClipWait(1) {
            this.Notify("複製到剪貼簿失敗")
            A_Clipboard := SavedClip
            return
        }

        SetKeyDelay 50, 50
        SendEvent "^v"
        Sleep 300
        A_Clipboard := SavedClip
    }

    ; =================================================================
    ; 焦點檢查 helper
    ; =================================================================
    static IsTargetFocused()
    {
        try {
            focusedHwnd := ControlGetFocus("A")
        } catch {
            return false
        }

        try {
            if (this.FindingEdit.NativeWindowHandle == focusedHwnd)
                return true
        }
        try {
            if (this.ImpressionEdit.NativeWindowHandle == focusedHwnd)
                return true
        }
        return false
    }

    ; =================================================================
    ; 字體強制模組
    ; =================================================================
    static _hCustomFont := 0

    static EnableFontEnforcer(fontName := "Cascadia Code", fontSize := 12)
    {
        ; 1. 計算 DPI 修正係數
        ; A_ScreenDPI: 標準螢幕是 96，Retina 可能是 144, 192, 240...
        ; 如果是 Retina (192)，ratio 會變成 0.5，把 14pt 變成 7pt，
        ; 這樣 Windows 放大 200% 後，看起來就會剛好是原本 14pt 的大小。
        dpiRatio := 96 / A_ScreenDPI

        ; 計算實際要請求的字體大小 (取整數)
        adjustedSize := Round(fontSize * dpiRatio, 1)

        ; 2. 建立 Font Handle (若已存在先銷毀舊的，支援 RDP 重連後的解析度變更)
        if (this._hCustomFont != 0) {
            DllCall("DeleteObject", "Ptr", this._hCustomFont)
            this._hCustomFont := 0
        }

        ; 利用一個隱藏的 GUI 來產生合法的 HFONT
        dummyGui := Gui()

        ; 使用修正後的大小
        dummyGui.SetFont("s" adjustedSize, fontName)
        dummyCtrl := dummyGui.Add("Text",, "Dummy")

        ; 取得 Handle
        this._hCustomFont := SendMessage(0x31, 0, 0, dummyCtrl.Hwnd)

        ; 3. 啟動 Timer
        SetTimer(ObjBindMethod(this, "_EnforceFontTask"), 1000)

        ; 顯示除錯訊息 (確認是否有生效，之後可註解掉)
        ; ToolTip "DPI: " A_ScreenDPI "`nRatio: " dpiRatio "`nAdj Size: " adjustedSize
        ; SetTimer () => ToolTip(), -3000
    }

    ; =================================================================
    ; 自動排版模組
    ; =================================================================
    static _targetImpressionHeight := 95

    static _EnforceFontTask()
    {
        if !WinActive(this.WinTitle)
            return

        static WM_SETFONT := 0x30

        try {
            hFind := this.FindingEdit.NativeWindowHandle
            hImp  := this.ImpressionEdit.NativeWindowHandle
        } catch {
            return
        }

        if (this._hCustomFont) {
            try SendMessage(WM_SETFONT, this._hCustomFont, 1, , "ahk_id " hFind)
            try SendMessage(WM_SETFONT, this._hCustomFont, 1, , "ahk_id " hImp)
        }

        this._ApplyLayout(hFind, hImp)
    }

    static _ApplyLayout(hFind, hImp)
    {
        ; 1. 計算 DPI 縮放係數
        ; 96 DPI (100%) -> scale 1.0
        ; 192 DPI (200%) -> scale 2.0
        dpiScale := A_ScreenDPI / 96

        ; 2. 定義基礎間距 (這些是給標準螢幕用的數值)
        baseGap  := 30    ; Finding 與 Impression 之間的留白 (給標籤用的空間)
        baseLbl  := 25    ; 標籤要往上移多少距離

        ; 3. 【關鍵修正】所有數值都乘上縮放係數
        ; 這樣在 Retina 螢幕上，高度和間距會自動變大，才不會切到變大的字體
        targetImpH  := this._targetImpressionHeight * dpiScale
        gap         := baseGap * dpiScale
        labelOffset := baseLbl * dpiScale

        ; 4. 取得目前控制項的位置
        ControlGetPos(&fX, &fY, &fW, &fH, hFind)
        ControlGetPos(&iX, &iY, &iW, &iH, hImp)

        ; 5. 計算座標
        ; 以目前的底部為基準 (Anchor Bottom)
        currentBottom := iY + iH

        ; 計算 Impression 新的 Y 軸
        targetImpY := currentBottom - targetImpH

        ; 計算 Finding 新的高度 (填滿上方剩餘空間，並扣除變大後的 gap)
        targetFindH := (targetImpY - gap) - fY

        ; 6. 檢查是否需要移動 (容許誤差也隨 DPI 放大)
        tolerance := 5 * dpiScale
        if (Abs(iH - targetImpH) < tolerance && Abs(iY - targetImpY) < tolerance && Abs(fH - targetFindH) < tolerance)
            return

        ; 7. 開始移動

        ; (A) 移動 Impression Edit
        try ControlMove(,, iW, targetImpH, hImp) ; 先改高度
        try ControlMove(, targetImpY,,, hImp)    ; 再改位置

        ; (B) 移動 Finding Edit
        try ControlMove(,,, targetFindH, hFind)

        ; (C) 移動 Impression 標籤 (Label)
        try {
            elLabel := this._GetOrUpdateNode("ImpressionLabel")
            hLabel := elLabel.NativeWindowHandle

            if (hLabel) {
                ; 這裡使用縮放後的 labelOffset
                ; 確保在高解析度下，標籤不會跟輸入框重疊
                labelNewY := targetImpY - labelOffset
                ControlMove(, labelNewY,,, hLabel)
            }
        }
    }

    ; =================================================================
    ; 刪除與移動功能
    ; =================================================================
    static DeleteCurrentLine() {
        if !this.IsTargetFocused()
            return false

        try {
            hFocus := ControlGetFocus("A")
            if !hFocus
                return false
        } catch {
            return false
        }

        this._SelectLine(hFocus)
        SendMessage(0x0303, 0, 0, hFocus) ; WM_CLEAR

        return true
    }

    static DeleteWordBackward() {
        if !this.IsTargetFocused()
            return false

        try {
            hCtrl := ControlGetFocus("A")
            if !hCtrl
                return false
        } catch {
            return false
        }

        this._BashDeleteAlgo(hCtrl)
        return true
    }

    static _BashDeleteAlgo(hCtrl) {
        try {
            fullText := ControlGetText(hCtrl)
        } catch {
            return
        }

        caretPosRaw := SendMessage(0x00B0, 0, 0, hCtrl)
        caretPos := caretPosRaw & 0xFFFF

        if (caretPos == 0)
            return

        i := caretPos
        while (i > 0) {
            char := SubStr(fullText, i, 1)
            if (this._IsSpace(char)) {
                i--
            } else {
                break
            }
        }
        while (i > 0) {
            char := SubStr(fullText, i, 1)
            if (!this._IsSpace(char)) {
                i--
            } else {
                break
            }
        }

        selStart := i
        selEnd := caretPos
        SendMessage(0x00B1, selStart, selEnd, hCtrl)
        SendMessage(0x0303, 0, 0, hCtrl)
    }

    static _IsSpace(char) {
        return (char == " " || char == "`t" || char == "`r" || char == "`n")
    }

    ; =================================================================
    ; Emacs 風格游標移動
    ; =================================================================
    static MoveCaret(mode) {
        if !this.IsTargetFocused()
            return false

        try {
            hCtrl := ControlGetFocus("A")
            if !hCtrl
                return false
        } catch {
            return false
        }

        static EM_LINEFROMCHAR := 0x00C9
        static EM_LINEINDEX    := 0x00BB
        static EM_LINELENGTH   := 0x00C1
        static EM_SETSEL       := 0x00B1
        static EM_SCROLLCARET  := 0x00B7

        lineIdx := SendMessage(EM_LINEFROMCHAR, -1, 0, hCtrl)
        lineStart := SendMessage(EM_LINEINDEX, lineIdx, 0, hCtrl)
        targetPos := 0

        if (mode = "Start") {
            targetPos := lineStart
        }
        else if (mode = "End") {
            lineLen := SendMessage(EM_LINELENGTH, lineStart, 0, hCtrl)
            targetPos := lineStart + lineLen
        }

        SendMessage(EM_SETSEL, targetPos, targetPos, hCtrl)
        SendMessage(EM_SCROLLCARET, 0, 0, hCtrl)

        return true
    }

    ; =================================================================
    ; 選取邏輯行 (核心)
    ; =================================================================
    static _SelectLine(hCtrl) {
        try {
            fullText := ControlGetText(hCtrl)
        } catch {
            return
        }

        if (fullText = "")
            return

        caretPosRaw := SendMessage(0x00B0, 0, 0, hCtrl)
        caretPos := caretPosRaw & 0xFFFF

        ahkCaretPos := caretPos + 1
        prevLineBreak := InStr(fullText, "`n", , ahkCaretPos, -1)
        selStart := (prevLineBreak == 0) ? 0 : prevLineBreak

        nextR := InStr(fullText, "`r", , ahkCaretPos)
        nextN := InStr(fullText, "`n", , ahkCaretPos)
        selEnd := 0

        if (nextR == 0 && nextN == 0) {
            selEnd := StrLen(fullText)
        }
        else if (nextR > 0 && (nextN == 0 || nextR < nextN)) {
            if (SubStr(fullText, nextR + 1, 1) == "`n") {
                selEnd := nextR + 1
            } else {
                selEnd := nextR
            }
        }
        else {
            selEnd := nextN
        }

        SendMessage(0x00B1, selStart, selEnd, hCtrl)
    }

    ; =================================================================
    ; 單字移動
    ; =================================================================
    static MoveCaretWord(direction) {
        if !this.IsTargetFocused()
            return false

        if (direction = "Left")
            Send "^{Left}"
        else
            Send "^{Right}"

        return true
    }

    ; =================================================================
    ; 歷史報告過濾器
    ; =================================================================
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

    ; =================================================================
    ; 帶入前次報告
    ; =================================================================
    static AppendPreviousReport() {
        try {
            pastImp := ControlGetText(this.PastImpressionText.NativeWindowHandle)
            pastFind := ControlGetText(this.PastFindingText.NativeWindowHandle)
            hImpEdit := this.ImpressionEdit.NativeWindowHandle
            hFindEdit := this.FindingEdit.NativeWindowHandle
        } catch {
            return
        }

        static EM_SETSEL := 0x00B1
        static EM_REPLACESEL := 0x00C2
        static EM_SCROLLCARET := 0x00B7

        AppendToEdit(hEdit, textToAppend) {
            if (textToAppend == "")
                return

            try {
                currentText := ControlGetText(hEdit)
                currentLen := StrLen(currentText)
            } catch {
                currentLen := 0
            }

            SendMessage(EM_SETSEL, currentLen, currentLen, hEdit)

            if (currentLen > 0)
                textToAppend := "`r`n" . textToAppend

            SendMessage(EM_REPLACESEL, 1, StrPtr(textToAppend), hEdit)
            SendMessage(EM_SCROLLCARET, 0, 0, hEdit)
        }

        AppendToEdit(hImpEdit, pastImp)
        AppendToEdit(hFindEdit, pastFind)

        try this.FindingEdit.SetFocus()
    }

    ; =================================================================
    ; 檢查名稱與存檔
    ; =================================================================
    static InsertExamNameAtCaret() {
        if !this.IsTargetFocused()
            return false

        try {
            hEdit := ControlGetFocus("A")
            rawName := ControlGetText(this.ExamnameText.NativeWindowHandle)
        } catch {
            return false
        }

        if (rawName == "")
            return true

        cleanName := StrReplace(rawName, "檢查項目: ", "")
        textToInsert := cleanName . ":`r`n`r`n"

        static EM_REPLACESEL := 0x00C2
        SendMessage(EM_REPLACESEL, 1, StrPtr(textToInsert), hEdit)

        return true
    }

    static SetAutoNextState(targetState) {
        try {
            currentState := this.AutoNextCheckbox.ToggleState
            if (!!targetState != !!currentState) {
                this.AutoNextCheckbox.Toggle()
            }
        }
    }

    static SaveReport() {
        try {
            this.ReportSaveButton.ControlClick()
        } catch as err {
            this.Notify("存檔失敗: " err.Message)
        }
    }

    ; =================================================================
    ; 智慧選取 (Shift+Up/Down)
    ; =================================================================
    static SmartExtendSelection(direction) {
        if !this.IsTargetFocused()
            return false

        try {
            hCtrl := ControlGetFocus("A")
        } catch {
            return false
        }

        static EM_LINEFROMCHAR := 0x00C9
        static EM_GETLINECOUNT := 0x00BA

        currentLineIdx := SendMessage(EM_LINEFROMCHAR, -1, 0, hCtrl)
        lineCount      := SendMessage(EM_GETLINECOUNT, 0, 0, hCtrl)

        if (direction == "Up") {
            if (currentLineIdx == 0) {
                SendInput "+{Home}"
            } else {
                SendInput "+{Up}"
            }
        }
        else if (direction == "Down") {
            if (currentLineIdx == lineCount - 1) {
                SendInput "+{End}"
            } else {
                SendInput "+{Down}"
            }
        }

        return true
    }

    ; =================================================================
    ; 歷史報告互動
    ; =================================================================
    static _SimReportMap := Map(
        "CHEST PA/AP", Map("CHEST PA/AP+LAT", 1),
        "CHEST PA/AP+LAT", Map("CHEST PA/AP", 1),
        "KUB", Map("KUB+ABD LAT", 1),
        "KUB+L-SPINE LAT(supine)", Map("L-SPINE(AP+LAT)Standing", 1),
        "WHOLE  ABDOMEN CT WITH+ WITHOUT CONTRAST", Map("WHOLE  ABDOMEN CT WITHOUT CONTRAST", 1),
        "WHOLE  ABDOMEN CT WITHOUT CONTRAST", Map("WHOLE  ABDOMEN CT WITH+ WITHOUT CONTRAST", 1),
    )

    static FindAndClickSimilarReport() {
        currExamName := this._GetCleanCurrentExamName()
        if (currExamName == "")
            return

        SearchColumnIndex := 3

        try {
            tableEle := this.PastReportTable
            rowElements := tableEle.FindAll({ Type: 'Custom' })
            if (rowElements.Length = 0)
                throw Error("表格中找不到資料行 (Rows)")

            for rowEle in rowElements {
                cellElements := rowEle.FindAll({ Type: 'DataItem' })
                if (cellElements.Length = 0)
                    cellElements := rowEle.FindAll({ Type: 'Custom' })

                if (cellElements.Length < SearchColumnIndex)
                    continue

                targetCellEle := cellElements[SearchColumnIndex]
                historyExamName := targetCellEle.Value

                if (this._IsRelatedReport(historyExamName, currExamName)) {
                    this._ClickUIAElement(targetCellEle)
                    this.Notify("已選取相似報告: " historyExamName) ; [修改] 取代 ToolTip
                    return
                }
            }

            this.Notify("未找到相似的歷史報告", 1000) ; [修改] 取代 ToolTip，時間縮短

        } catch as err {
            this.Notify("搜尋失敗: " err.Message)
        }
    }

    static InsertSelectedHistoryDate() {
        this._InsertFromSelectedRow(1, true)
    }

    static InsertSelectedHistoryName() {
        this._InsertFromSelectedRow(3, false)
    }

    static _InsertFromSelectedRow(colIndex, needDateConvert) {
        if !this.IsTargetFocused()
            return

        static STATE_SYSTEM_SELECTED := 0x2
        try {
            tableEle := this.PastReportTable
            rowElements := tableEle.FindAll({ Type: 'Custom' })
            foundValue := ""

            for rowEle in rowElements {
                if IsObject(rowEle.LegacyIAccessiblePattern) {
                    if (rowEle.LegacyIAccessiblePattern.State & STATE_SYSTEM_SELECTED) {
                        targetCell := rowEle.FindElement({ ControlType: "DataItem" }, , colIndex)
                        if IsObject(targetCell) {
                            foundValue := targetCell.Value
                        }
                        break
                    }
                }
            }

            if (foundValue != "") {
                if (needDateConvert)
                    foundValue := this._ConvertRISDate(foundValue)
                SendText foundValue
            }
        }
    }

    static _GetCleanCurrentExamName() {
        try {
            rawName := ControlGetText(this.ExamnameText.NativeWindowHandle)
            return StrReplace(rawName, "檢查項目: ", "")
        } catch {
            return ""
        }
    }

    static _IsRelatedReport(prevName, currName) {
        if (prevName == currName)
            return true
        if this._SimReportMap.Has(currName) {
            similarExams := this._SimReportMap[currName]
            if similarExams.Has(prevName)
                return true
        }
        return false
    }

    static _ConvertRISDate(inputString) {
        cleanString := StrReplace(inputString, "/")
        if (StrLen(cleanString) < 7)
            return inputString

        minguoYear := SubStr(cleanString, 1, 3)
        month := SubStr(cleanString, 4, 2)
        day := SubStr(cleanString, 6, 2)

        gregorianYear := minguoYear + 1911
        return gregorianYear . "-" . month . "-" . day
    }

    static _ClickUIAElement(el) {
        try {
            rect := el.BoundingRectangle
            loc := el.Location
            ClickX := rect.l + (loc.w / 2)
            ClickY := rect.t + (loc.h / 2)
            MouseGetPos(&OrigX, &OrigY)
            MouseMove(ClickX, ClickY, 0)
            Click()
            MouseMove(OrigX, OrigY, 0)
        } catch {
            try el.Invoke()
        }
    }

    ; =================================================================
    ; 報告格式化與重排
    ; =================================================================
    static FormatFindingText() {
        if !this.IsTargetFocused()
            return

        try {
            examType := this._GetCurrExamType()
            hEdit := this.FindingEdit.NativeWindowHandle

            switch examType {
                case "CT", "MR":
                    this._FormatFindingForAdvanced(hEdit)
                case "CR", "US", "CR":
                    this._FormatFindingForBasic(hEdit)
            }
        } catch as err {
            this.Notify("格式化失敗: " err.Message)
        }
    }

    static FormatImpressionText() {
        if !this.IsTargetFocused()
            return

        try {
            hEdit := this.ImpressionEdit.NativeWindowHandle
            ControlFocus(hEdit)
            SendMessage(0x00B1, 0, -1, hEdit)

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
        if !this.IsTargetFocused()
            return

        try {
            hEdit := ControlGetFocus("A")
            deOrder := options.HasOwnProp("deOrder") ? options.deOrder : false
            keepEmpty := options.HasOwnProp("keepEmpty") ? options.keepEmpty : false
            itemChar := options.HasOwnProp("itemChar") ? options.itemChar : ""
            discardSeIm := options.HasOwnProp("discardSeIm") ? options.discardSeIm : true

            this._ReorderSelectedText(deOrder, keepEmpty, itemChar, discardSeIm, hEdit)
        }
    }

    static _GetCurrExamType() {
        name := this._GetCleanCurrentExamName()
        if (InStr(name, "CT") || InStr(name, "電腦斷層"))
            return "CT"
        if (InStr(name, "MR") || InStr(name, "磁振造影"))
            return "MR"
        if (InStr(name, "US") || InStr(name, "超音波"))
            return "US"
        return "CR"
    }

    static _FormatFindingForBasic(hEdit) {
        fullText := ControlGetText(hEdit)
        needle := "im)FINDINGS:\r?\n|:\s*\r?\n\s*\r?\n"

        if RegExMatch(fullText, needle, &match) {
            startPos := match.Pos + match.Len - 1
            SendMessage(0x00B1, startPos, -1, hEdit)
            this._ReorderSelectedText(false, true, "-", false, hEdit)
        }
    }

    static _FormatFindingForAdvanced(hEdit) {
        fullText := ControlGetText(hEdit)
        needle := "im)FINDINGS:\r?\n|The study shows:\r?\n\r?\n|show the following findings:\r?\n\r?\n|which revealed:\r?\n\r?\n"

        if RegExMatch(fullText, needle, &match) {
            startPos := match.Pos + match.Len - 1
            endNeedle := "im)REMARKS?:|RECOMMENDATION:"
            endPos := -1
            if RegExMatch(fullText, endNeedle, &endMatch, startPos + 1) {
                endPos := endMatch.Pos - 1
                if (endPos > 2)
                    endPos -= 2
            }
            SendMessage(0x00B1, startPos, endPos, hEdit)
            this._ReorderSelectedText(false, false, "-", true, hEdit)
        }
    }

    static _ReorderSelectedText(deOrder := false, keepEmptyLine := false, itemChar := "", discardSeIm := true, targetHwnd := 0) {
        selectedText := ""
        try {
            fullText := ControlGetText(targetHwnd)
            static EM_GETSEL := 0x00B0
            selRaw := SendMessage(EM_GETSEL, 0, 0, targetHwnd)
            start := selRaw & 0xFFFF
            end := (selRaw >> 16) & 0xFFFF

            if (end > start)
                selectedText := SubStr(fullText, start + 1, end - start)
        }

        if (selectedText == "")
            return

        selectedText := StrReplace(selectedText, "`r`n", "`n")
        txtAry := StrSplit(selectedText, "`n")
        finalText := ""
        startLineNo := 1

        if (RegExMatch(selectedText, "^(\d+)", &existLineNo))
            startLineNo := existLineNo[1]

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
                if (keepEmptyLine)
                    finalText .= "`r`n"
            }
        }
        finalText := RTrim(finalText, "`r`n")
        try EditPaste(finalText, targetHwnd)
    }

    static _CountNonEmptyLines(hEdit) {
        text := ControlGetText(hEdit)
        if (text == "")
            return 0
        lines := StrSplit(text, "`n", "`r")
        count := 0
        for line in lines {
            if (Trim(line, " `t") != "")
                count++
        }
        return count
    }

    ; =================================================================
    ; 病理報告複製
    ; =================================================================
    static CopyPathologyReport() {
        try {
            dateVal := this.PathoDateText.Value
            diagVal := this.PathoDiagnosisText.Value

            if (dateVal == "" && diagVal == "")
                throw Error("找不到病理報告內容")

            reportText := this._ConvertRISDate(dateVal) . ": " . diagVal
            A_Clipboard := reportText

            this.Notify("病理報告已複製") ; [修改] 取代 ToolTip
        } catch as err {
            this.Notify("複製失敗: " err.Message)
        }
    }

    ; =================================================================
    ; 滑鼠連點選取
    ; =================================================================
    static HandleTripleClick() {
        static clickCount := 0
        static lastClickTime := 0
        static DoubleClickTime := DllCall("GetDoubleClickTime")

        timeSinceLast := A_TickCount - lastClickTime
        if (timeSinceLast <= DoubleClickTime)
            clickCount++
        else
            clickCount := 1

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
    ; [新增功能] 全域視窗控制 (Global Window Control)
    ; =================================================================

    /**
     * 智慧切換視窗焦點
     * 1. 若 RIS 未啟用 -> 啟用視窗並聚焦 Finding
     * 2. 若 RIS 已啟用 -> 在 Finding 與 Impression 之間輪流切換
     */
    static ActivateOrToggleFocus() {
        try {
            ; 狀況 A: 視窗不在前景 -> 叫出來
            if !WinActive(this.WinTitle) {
                WinActivate(this.WinTitle)

                ; 等待視窗浮現，最多等 2 秒
                if !WinWaitActive(this.WinTitle, , 2) {
                    this.Notify("找不到或無法啟用 RIS 視窗")
                    return
                }

                ; 剛切換過來，預設聚焦 Finding
                try this.FindingEdit.SetFocus()
            }
            ; 狀況 B: 視窗已在前景 -> 切換焦點
            else {
                ; 取得目前焦點的 Handle (使用 Win32 API 遠比 UIA 快)
                try {
                    hFocus := ControlGetFocus("A")
                } catch {
                    hFocus := 0
                }

                ; 取得 Finding 的 Handle
                try {
                    hFind := this.FindingEdit.NativeWindowHandle
                } catch {
                    hFind := 0
                }

                ; 邏輯判斷
                if (hFocus == hFind) {
                    ; 如果現在在 Finding -> 切去 Impression
                    try this.ImpressionEdit.SetFocus()
                } else {
                    ; 如果在 Impression (或任何其他地方) -> 切回 Finding
                    try this.FindingEdit.SetFocus()
                }
            }
        } catch as err {
            this.Notify("視窗切換失敗: " err.Message)
        }
    }

    /**
     * 點擊異常報告視窗的選項按鈕
     * @param index 1, 2, 3, 4 或 "Save"
     */
    static ClickAbnormalButton(index) {
        ; 確保視窗是活躍的 (雙重保險)
        if !WinActive(this.AbnormalWinTitle)
            return

        if !this._AbnormalBtnMap.Has(index) {
            this.Notify("未定義的按鈕: " index)
            return
        }

        classNN := this._AbnormalBtnMap[index]

        try {
            ControlClick(classNN, this.AbnormalWinTitle)
        } catch as err {
            this.Notify("按鈕點擊失敗 (找不到元件)")
        }
    }
}