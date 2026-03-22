#Requires AutoHotkey v2.0
#Include .\UIA.v2.ahk

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
        CUT:           0x0300, ; [新增] 剪下
        COPY:          0x0301, ; [新增] 複製
        CLEAR:         0x0303
    }

    ; =================================================================
    ; 1. 設定區 (Configuration)
    ; =================================================================
    static WinTitle := "報告作業(frmRISReport)"
    static AbnormalWinTitle := "檢查結果(frmPos)"
    static ConsultationWinTitle := "會診資訊(frmReqCon)"
    static WorklistWinTitle := "工作清單(frmRIS)"

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
        "PhExamColumn",     { AutomationId: "goxExamine" },
        "PhExamDateText",   { AutomationId: "mtxtReportDTM" },
        "PhExamReportText", { AutomationId: "txtReport" },
        "PhExamImpChkBox",  { AutomationId: "chBoxImpression" },

        ; [新增] SOAP 與基本資料欄位
        "SubjectiveText",   { AutomationId: "rtxtSubjective" },
        "ObjectiveText",    { AutomationId: "rtxtObjective" },
        "AssessmentText",   { AutomationId: "rtxtICD10" },
        "PlanText",         { AutomationId: "rtxtAdmInICD" },
        "OrderDeptText",    { AutomationId: "txtAppSecName" },
        "GenderText",       { AutomationId: "txtGender" },
        "AgeText",          { AutomationId: "txtPtAge" },
    )

    ; [修改] 資料結構優化：使用「群組列表」代替繁瑣的手動 Mapping
    ; 只要將相似的檢查名稱放在同一個陣列 [] 裡，程式會自動建立雙向關聯。
    static _SimGroups := [
        ; Chest
        ["CHEST PA/AP", "CHEST PA/AP+LAT"],

        ; Abdomen / KUB / Spine
        ["KUB", "KUB+ABD LAT"],
        ["L-SPINE(AP+LAT)Standing", "KUB+L-SPINE LAT(supine)"],

        ; CT Abdomen (雙向互通)
        [
            "WHOLE  ABDOMEN CT WITH+ WITHOUT CONTRAST",
            "WHOLE  ABDOMEN CT WITHOUT CONTRAST"
        ],

        [
            "CT LUNG/ PLEURA/ CHEST WALL WITHOUT CONTRAST",
            "CT LUNG/ PLEURA/ CHEST WALL WITH+ WITHOUT CONTRAST"
        ],

        ; CT Brain (建立完整互聯：包含 Trauma, Non-con, With+Without)
        [
            "CT BRAIN (急診TRAUMA 專用)WITHOUT CONTRAST",
            "CT BRAIN WITHOUT CONTRAST",
            "CT BRAIN WITH+ WITHOUT CONTRAST"
        ]
    ]

    ; [新增] 執行期使用的快速查詢表 (由 __New 自動生成，無需手動維護)
    static _SimReportMap := Map()

    static _ConsultationCtrls := Map(
        "SourceTime", "WindowsForms10.EDIT.app.0.2780b98_r24_ad116", ; 原始時間 (國曆)
        "TargetTime", "WindowsForms10.EDIT.app.0.2780b98_r24_ad114"  ; 目標填入欄位
    )

    static _WorklistCtrls := Map(
        "RefreshButton", {AutomationId: "btnRefresh"},
        "ER",            {AutomationId: "dgvClassifyOPDE"}, ; 急診
        "ADM",           {AutomationId: "dgvClassifyADM"},  ; 住院
        "OPD",           {AutomationId: "dgvClassifyOPDR"}  ; 門診
    )

    ; =================================================================
    ; 1.1 初始化邏輯 (Initialization)
    ; =================================================================

    ; [新增] 類別載入時自動執行：將 _SimGroups 編譯為 _SimReportMap
    static __New() {
        this._SimReportMap.CaseSense := "Off" ; 設定為不分大小寫，增加比對容錯率

        for group in this._SimGroups {
            for item in group {
                ; 確保每個項目都有一個對應的 Map
                if !this._SimReportMap.Has(item) {
                    this._SimReportMap[item] := Map()
                    this._SimReportMap[item].CaseSense := "Off"
                }

                ; 將同群組內的其他項目加入該項目的關聯表
                for peer in group {
                    if (item != peer) {
                        this._SimReportMap[item][peer] := 1
                    }
                }
            }
        }
    }

    ; =================================================================
    ; 2. 內部狀態 (State)
    ; =================================================================
    static _cache := Map()
    static _currentNotifyGui := ""
    static _compContext := {MRN: "", Date: ""}
    static _hCustomFont := 0
    static _targetImpressionHeight := 95

    ; [自動更新相關狀態]
    static _lastUpdateTick := 0           ; 上次更新的時間
    static _updateInterval := 1800000     ; 30 分鐘 (標準生產環境設定)
    static _idleThreshold  := 300000      ; 5 分鐘
    static _isUpdating     := false       ; 防卡死旗標

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
    static PhExamColumn => this._GetOrUpdateNode("PhExamColumn")
    static PhExamDateText => this._GetOrUpdateNode("PhExamDateText")
    static PhExamReportText => this._GetOrUpdateNode("PhExamReportText")

    ; [新增] SOAP 與基本資料 Getters
    static SubjectiveText => this._GetOrUpdateNode("SubjectiveText")
    static ObjectiveText  => this._GetOrUpdateNode("ObjectiveText")
    static AssessmentText => this._GetOrUpdateNode("AssessmentText")
    static PlanText       => this._GetOrUpdateNode("PlanText")
    static OrderDeptText  => this._GetOrUpdateNode("OrderDeptText")
    static GenderText     => this._GetOrUpdateNode("GenderText")
    static AgeText        => this._GetOrUpdateNode("AgeText")

    ; [新增] 快捷文字存取屬性 (使用快速路徑並自動標準化換行)
    static FindingText => this.GetText(this.FindingEdit)
    static ImpressionText => this.GetText(this.ImpressionEdit)

    ; [新增] 取得過濾後的 Finding 內文 (去除標題與結尾)
    static GetFindingContent() {
        try {
            hEdit := this.FindingEdit.NativeWindowHandle
            fullText := ControlGetText(hEdit)

            ; 依照您的需求，這裡使用 Advanced (CT/MR) 的邏輯來剖析
            range := this._FindContentRange(fullText, "Advanced")

            if (!range) {
                return ""
            }

            length := (range.End == -1) ? StrLen(fullText) - range.Start : range.End - range.Start
            return SubStr(fullText, range.Start + 1, length)
        } catch {
            return ""
        }
    }

    ; =================================================================
    ; 4. 系統功能 (Notify & Focus)
    ; =================================================================

    ; [MODIFIED] 現代化 Notify UI
    static Notify(text, duration := 1500) {
        if (this._currentNotifyGui) {
            try {
                this._currentNotifyGui.Destroy()
            }
            this._currentNotifyGui := ""
        }

        ; +AlwaysOnTop: 置頂
        ; -Caption: 無標題列
        ; +ToolWindow: 不顯示在工作列
        ; +E0x20 (WS_EX_TRANSPARENT): 穿透滑鼠點擊
        ; +E0x08000000 (WS_EX_NOACTIVATE): 不搶奪焦點
        g := Gui("+AlwaysOnTop -Caption +ToolWindow +E0x20 +E0x08000000")

        ; 現代深色風格背景 (Deep Dark Gray)
        g.BackColor := "202020"

        ; 使用更現代的 UI 字體，微軟正黑體 UI 或 Segoe UI
        g.SetFont("s13 cWhite bold", "Microsoft JhengHei UI")

        ; 增加邊距讓文字呼吸
        g.MarginX := 25
        g.MarginY := 15

        g.Add("Text", "Center", text)

        ; 先顯示出來以計算尺寸 (NoActivate)
        g.Show("NoActivate AutoSize Center")

        ; --- 視覺特效處理 ---
        try {
            hwnd := g.Hwnd
            WinGetPos(,, &w, &h, hwnd)

            ; 1. 圓角效果 (WinSetRegion)
            ; r12-12 代表圓角半徑
            WinSetRegion("0-0 w" w " h" h " r12-12", hwnd)

            ; 2. 陰影效果 (CS_DROPSHADOW = 0x00020000)
            ; 透過設定 Class Style 讓無邊框視窗擁有系統陰影
            style := DllCall("GetClassLongPtr", "Ptr", hwnd, "Int", -26, "Ptr") ; GCL_STYLE = -26
            DllCall("SetClassLongPtr", "Ptr", hwnd, "Int", -26, "Ptr", style | 0x00020000)

            ; 3. 輕微透明度 (提升質感)
            WinSetTransparent(235, hwnd)
        }

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
        targetHwnd := 0 ; [新增] 用來記錄最後聚焦的視窗 Handle

        try {
            if !WinActive(this.WinTitle) {
                ; [修改] 1. 先確認視窗是否確實存在
                if !WinExist(this.WinTitle) {
                    this.Notify("找不到 RIS 視窗")
                    return
                }

                ; [修改] 2. 嘗試啟用視窗
                WinActivate(this.WinTitle)

                ; [修改] 3. 縮短第一次等待時間，若失敗則進行「強制喚醒」突破系統鎖定
                if !WinWaitActive(this.WinTitle, , 1) {
                    WinShow(this.WinTitle)     ; 確保視窗不是隱藏狀態
                    WinActivate(this.WinTitle) ; 再次要求焦點

                    if !WinWaitActive(this.WinTitle, , 1) {
                        this.Notify("無法啟用 RIS 視窗 (可能被系統阻擋)")
                        return
                    }
                }

                ; 如果目前焦點 "不在" Finding 或 Impression 上，才強制聚焦到 Finding
                if !this.IsTargetFocused() {
                    try {
                        this.FindingEdit.SetFocus()
                        targetHwnd := this.FindingEdit.NativeWindowHandle ; 記錄
                    }
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
                        targetHwnd := this.ImpressionEdit.NativeWindowHandle ; 記錄
                    }
                } else {
                    try {
                        this.FindingEdit.SetFocus()
                        targetHwnd := this.FindingEdit.NativeWindowHandle ; 記錄
                    }
                }
            }

            ; 如果上面沒抓到，就抓目前系統焦點
            if (!targetHwnd) {
                try {
                    targetHwnd := ControlGetFocus("A")
                }
            }

            SetTimer( () => this._HighlightCaret(targetHwnd), -10 )
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
        this._ShowWaitCursor()
        try {
            try {
                pastImp := ControlGetText(this.PastImpressionText.NativeWindowHandle)
                pastFind := ControlGetText(this.PastFindingText.NativeWindowHandle)
                hImpEdit := this.ImpressionEdit.NativeWindowHandle
                hFindEdit := this.FindingEdit.NativeWindowHandle
            } catch {
                return
            }

            ; [紀錄] 紀錄 Finding 目前的 Caret 位置
            initialFindSel := this._EditGetSel(hFindEdit)

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
                this._EditSetSel(hEdit, -1, -1)
                this._EditScrollCaret(hEdit)
            }

            AppendToEdit(hImpEdit, pastImp)
            AppendToEdit(hFindEdit, pastFind)

            try {
                this.FindingEdit.SetFocus()
                ; [修改] 不再強制移至 0，而是還原至 initialFindSel.Start
                this._EditSetSel(hFindEdit, initialFindSel.Start, initialFindSel.Start)
                this._EditScrollCaret(hFindEdit)
            }
        } finally {
            this._RestoreCursor()
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

    static CopyFindingToImpression() {
        ; 1. 檢查 Focus 是否在 Finding
        try {
            hFocus := ControlGetFocus("A")
            if (hFocus != this.FindingEdit.NativeWindowHandle) {
                return
            }
        } catch {
            return
        }

        ; 2. 記錄初始狀態與取得文字
        initialSel := this._EditGetSel(hFocus)
        textToCopy := ""
        wasSelectionEmpty := (initialSel.Start == initialSel.End)

        if (wasSelectionEmpty) {
            ; === 狀況 A: 原本無反白 -> 自動抓整行並還原 ===
            this._SelectLine(hFocus)

            newSel := this._EditGetSel(hFocus)
            fullText := ControlGetText(hFocus)
            if (newSel.End > newSel.Start) {
                textToCopy := SubStr(fullText, newSel.Start + 1, newSel.End - newSel.Start)
            }

            ; 還原 Caret 到原本位置
            this._EditSetSel(hFocus, initialSel.Start, initialSel.Start)
            this._EditScrollCaret(hFocus)
        } else {
            ; === 狀況 B: 原本有反白 -> 保持不動 ===
            fullText := ControlGetText(hFocus)
            textToCopy := SubStr(fullText, initialSel.Start + 1, initialSel.End - initialSel.Start)
        }

        if (textToCopy == "") {
            return
        }

        ; 3. Append 到 Impression
        try {
            hImp := this.ImpressionEdit.NativeWindowHandle
            currentImpText := ControlGetText(hImp)

            ; [修改重點] 判斷結尾是否已為換行
            ; SubStr(str, -1) 會取得最後一個字元
            ; 如果是空字串 OR 最後一字是 `n (換行)，就不用補前綴
            if (currentImpText == "" || SubStr(currentImpText, -1) == "`n") {
                prefix := ""
            } else {
                prefix := "`r`n"
            }

            impLen := StrLen(currentImpText)

            ; 移到 Impression 最後並貼上
            this._EditSetSel(hImp, impLen, impLen)
            this._EditReplaceSel(hImp, prefix . textToCopy)

            this.Notify("已複製至 Impression")
        } catch as err {
            this.Notify("複製失敗: " . err.Message)
        }
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

    static CutLineOrSelection() {
        if !this.IsTargetFocused() {
            return false
        }
        try {
            hFocus := ControlGetFocus("A")

            ; 檢查是否有選取文字
            sel := this._EditGetSel(hFocus)
            if (sel.Start == sel.End) {
                ; 沒有選取：選取目前所在的邏輯行 (準備剪下整行)
                this._SelectLineForRemoval(hFocus)
            }

            ; 執行剪下 (包含複製到 Clipboard 與刪除)
            SendMessage(this.MSG.CUT, 0, 0, hFocus)
            this._EditScrollCaret(hFocus)
        }
        return true
    }

    static CopyLineOrSelection() {
        if !this.IsTargetFocused() {
            return false
        }
        try {
            hFocus := ControlGetFocus("A")

            ; 檢查是否有選取文字
            sel := this._EditGetSel(hFocus)
            didAutoSelect := false

            if (sel.Start == sel.End) {
                ; 沒有選取：選取目前所在的邏輯行 (Copy 使用 _SelectLine，保留後方換行)
                this._SelectLine(hFocus)
                didAutoSelect := true
            }

            ; 執行複製
            SendMessage(this.MSG.COPY, 0, 0, hFocus)

            ; 如果是自動選取整行，複製完後還原游標位置，避免影響打字
            if (didAutoSelect) {
                this._EditSetSel(hFocus, sel.Start, sel.Start)
            }
        }
        return true
    }

    static DeleteCurrentLine() {
        if !this.IsTargetFocused() {
            return false
        }
        try {
            hFocus := ControlGetFocus("A")

            ; 重構：使用共用的選取邏輯
            this._SelectLineForRemoval(hFocus)

            ; 執行刪除 (Clear 不會影響 Clipboard)
            SendMessage(this.MSG.CLEAR, 0, 0, hFocus)
            this._EditScrollCaret(hFocus)
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
            bounds := this._GetLogicalLineBoundaries(hCtrl)

            targetPos := 0
            if (mode = "Start") {
                targetPos := bounds.Start
            } else if (mode = "End") {
                targetPos := bounds.ContentEnd
            }

            this._EditSetSel(hCtrl, targetPos, targetPos)
            this._EditScrollCaret(hCtrl)
            return true
        } catch {
            return false
        }
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

    static SmartPageMove(direction) {
        if !this.IsTargetFocused() {
            return
        }
        try {
            hEdit := ControlGetFocus("A")

            ; 取得目前 Scroll 位置 (最上方的行號)
            prevFirstLine := SendMessage(this.MSG.GETFIRSTVISIBLELINE, 0, 0, hEdit)

            if (direction == "Up") {
                Send "{PgUp}"
                Sleep 10 ; 等待 UI 更新
                currFirstLine := SendMessage(this.MSG.GETFIRSTVISIBLELINE, 0, 0, hEdit)

                ; 如果無法再往上捲 (前後行號一樣，且已在第 0 行)，則移到最前
                if (prevFirstLine == 0 && currFirstLine == 0) {
                    this._EditSetSel(hEdit, 0, 0)
                    this._EditScrollCaret(hEdit)
                }
            } else { ; Down
                Send "{PgDn}"
                Sleep 10 ; 等待 UI 更新
                currFirstLine := SendMessage(this.MSG.GETFIRSTVISIBLELINE, 0, 0, hEdit)

                ; 如果無法再往下捲 (前後行號一樣)，則移到最後
                if (prevFirstLine == currFirstLine) {
                    fullText := ControlGetText(hEdit)
                    len := StrLen(fullText)
                    this._EditSetSel(hEdit, len, len)
                    this._EditScrollCaret(hEdit)
                }
            }
        }
    }

    ; [新增] 移動目前所在行 (Alt+Up / Alt+Down 功能實作)
    static MoveCurrentLine(direction) {
        if !this.IsTargetFocused() {
            return
        }

        try {
            hCtrl := ControlGetFocus("A")
            sel := this._EditGetSel(hCtrl)

            ; 只有在沒有選取範圍 (Caret 狀態) 時才執行
            if (sel.Start != sel.End) {
                return
            }

            fullText := ControlGetText(hCtrl)

            ; 1. 取得目前行 (Line B)
            currLine := this._GetLogicalLineBoundaries(hCtrl, sel.Start)

            targetLine := {}

            if (direction == "Up") {
                if (currLine.Start == 0) ; 已經在第一行
                    return

                ; [關鍵修正]
                ; currLine.Start 指向的是前一行的結尾 \n 的位置。
                ; 我們只需要往回跨過這個換行符號 (可能是 \n 或 \r\n)，就能找到上一行的內容。
                ; 不要使用 while 迴圈，避免一次跨過多個空行。

                searchPos := currLine.Start - 1 ; 先跨過已知的 \n

                ; 如果前面還有 \r，也跨過去
                if (searchPos > 0 && SubStr(fullText, searchPos, 1) == "`r") {
                    searchPos -= 1
                }

                ; 此時 searchPos 位於「上一行內容」的最後一個字 (若是空行，則是再上一行的 \n)
                targetLine := this._GetLogicalLineBoundaries(hCtrl, searchPos)

                ; 定義交換順序：Target(上) . Current(下) -> Current(上) . Target(下)
                topLine := targetLine
                btmLine := currLine

            } else { ; Down
                if (currLine.FullEnd == StrLen(fullText)) ; 已經在最後一行
                    return

                ; 往下找比較簡單，直接從下一行的開頭 (FullEnd) 找即可
                targetLine := this._GetLogicalLineBoundaries(hCtrl, currLine.FullEnd)

                ; 定義交換順序：Current(上) . Target(下) -> Target(上) . Current(下)
                topLine := currLine
                btmLine := targetLine
            }

            ; 2. 取出文字
            txtTop := SubStr(fullText, topLine.Start + 1, topLine.FullEnd - topLine.Start)
            txtBtm := SubStr(fullText, btmLine.Start + 1, btmLine.FullEnd - btmLine.Start)

            ; 3. 處理最後一行可能沒有換行符號的邊界狀況
            ; 如果上面那行有換行，但下面那行沒有 (通常發生在與最後一行交換時)
            if (SubStr(txtTop, -1) == "`n" && SubStr(txtBtm, -1) != "`n") {
                txtTop := SubStr(txtTop, 1, StrLen(txtTop) - 2) ; 移除上行的 \r\n
                txtBtm .= "`r`n"                                 ; 補給下行 \r\n
            }

            ; 4. 執行交換
            newText := txtBtm . txtTop
            this._EditSetSel(hCtrl, topLine.Start, btmLine.FullEnd)
            this._EditReplaceSel(hCtrl, newText)

            ; 5. 還原 Caret 位置
            if (direction == "Up") {
                ; 往上移：currLine 跑到上面，Caret 相對位置不變
                offset := sel.Start - currLine.Start
                newPos := topLine.Start + offset
            } else {
                ; 往下移：currLine 跑到下面 (txtBtm 是原本的 Target，現在變上面了)
                offset := sel.Start - currLine.Start
                newPos := topLine.Start + StrLen(txtBtm) + offset
            }

            this._EditSetSel(hCtrl, newPos, newPos)
            this._EditScrollCaret(hCtrl)

        } catch as err {
            this.Notify("移動失敗: " err.Message)
        }
    }

    ; [整合] 邏輯行換行增強功能 (支援 Above / Below)
    static InsertNewLine(mode := "Below") {
        if !this.IsTargetFocused() {
            return
        }
        try {
            hEdit := ControlGetFocus("A")
            bounds := this._GetLogicalLineBoundaries(hEdit)

            if (mode = "Above") {
                ; Shift+Enter 邏輯：移至行首 -> 插入 -> 游標回新行行首
                this._EditSetSel(hEdit, bounds.Start, bounds.Start)
                this._EditReplaceSel(hEdit, "`r`n")
                this._EditSetSel(hEdit, bounds.Start, bounds.Start)
            } else {
                ; Ctrl+Enter 邏輯：移至行尾 -> 插入 -> 游標自然停在新行
                this._EditSetSel(hEdit, bounds.ContentEnd, bounds.ContentEnd)
                this._EditReplaceSel(hEdit, "`r`n")
            }

            this._EditScrollCaret(hEdit)
        }
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

            ; [新增] 自動偵測項目符號模式 (autoDetectItemChar)
            if (options.HasOwnProp("autoDetectItemChar") && options.autoDetectItemChar) {
                sel := this._EditGetSel(hEdit)
                if (sel.End > sel.Start) {
                    fullText := ControlGetText(hEdit)
                    selectedText := SubStr(fullText, sel.Start + 1, sel.End - sel.Start)

                    ; 逐行尋找第一行非空白文字
                    for line in StrSplit(selectedText, "`n", "`r") {
                        if (Trim(line, " `t") != "") {
                            ; 檢查行首是否為指定的常見項目符號 (支援 > - = + *)
                            if RegExMatch(line, "^\s*([>\-=\+\*])", &match) {
                                itemChar := match[1]
                            } else {
                                itemChar := "-" ; 若無符合的符號，預設使用 '-'
                            }
                            break
                        }
                    }
                }
            }

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

        ; [修改] 改為呼叫全面預載機制
        this._PreloadCache()
    }

    ; [修改] 全面靜默快取主畫面元件 (取代原本的 _PreloadAICache)
    static _PreloadCache() {
        ; 若已經執行過預載，就不再重複執行，直到換病人 (cache 被清空)
        if (this._cache.Has("_UI_Preloaded")) {
            return
        }

        this._cache["_UI_Preloaded"] := true ; 標記為已預載

        ; 遍歷所有定義在 Selectors 裡的元件名稱
        for key in this.Selectors {
            ; 略過不在主畫面的病理報告元件，等實際切換過去時再取
            if (key == "PathoDiagnosisText" || key == "PathoDateText") {
                continue
            }

            if (!this._cache.Has(key)) {
                try {
                    ; 動態觸發 Getter 抓取元件並存入 cache
                    _ := this.%key%
                } catch {
                    ; 負向快取 (Negative Cache)：
                    ; 找不到就存成 false。避免下次觸發時重複進行 1000ms 的 UIA 深度搜尋
                    this._cache[key] := false
                }
            }
        }
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
            elLabel := this._GetOrUpdateNode("PhExamImpChkBox")
            if (hLabel := elLabel.NativeWindowHandle) {
                ControlMove(, targetImpY - labelOffset,,, hLabel)
            }
        }
    }

    static CopyOtherReport() {
        isPhExam := false

        ; 1. 嘗試判斷是否為檢查報告頁面
        try {
            ; 存取 PhExamColumn。如果不在該頁面，Getter 可能會因為找不到元件而拋出 Error。
            ; 我們利用這個特性：如果拋錯，isPhExam 就維持 false，自然流向 else (病理報告)。
            if (InStr(this.PhExamColumn.Name, "檢查報告") == 1) {
                isPhExam := true
            }
        }

        ; 2. 根據判斷結果分流執行
        if (isPhExam) {
            this.CopyPhExamReport()
        } else {
            ; 如果判斷為 false，或是找不到檢查報告欄位，預設執行病理報告複製
            this.CopyPathologyReport()
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

    static CopyPhExamReport() {
        try {
            dateVal := this.PhExamDateText.Value
            repVal := this.PhExamReportText.Value
            if (dateVal == "" && repVal == "") {
                throw Error("找不到檢查報告內容")
            }

            reportText := this._ConvertRISDate(dateVal) . ": " . repVal
            A_Clipboard := reportText
            this.Notify("檢查報告已複製")
        } catch as err {
            this.Notify("複製失敗: " err.Message)
        }
    }

    static SwitchHistoryFilter(modeName) {
        ; [新增] 記住動作前的焦點
        savedFocus := ""
        try {
            savedFocus := ControlGetFocus("A")
        }

        try {
            switch modeName {
                case "All":      this.PastAllRadio.ControlClick()
                case "Modality": this.PastModalityRadio.ControlClick()
                case "My":       this.PastOnlyMyRadio.ControlClick()
            }

            ; [新增] 執行後把 focus 還回去
            if (savedFocus) {
                ControlFocus(savedFocus, "A")
            }
        } catch as err {
            this.Notify("切換失敗: " err.Message) ; [修改] 取代 ToolTip
        }
    }

    static FindAndClickSimilarReport() {
        this._ShowWaitCursor() ; [新增]
        try {
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
        } finally {
            this._RestoreCursor() ; [新增]
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

    static EnsureImpressionNotEmpty() {
        try {
            ; 1. 取得文字 (使用快捷屬性，自動處理 UIA 與 Win32 路徑)
            impText := this.ImpressionText
            
            ; 2. 檢查是否為空
            if (impText == "" || !RegExMatch(impText, "\S")) {
                this.PasteToImpression("As aforementioned.")
                Sleep(50) ; 縮短延遲，僅確保貼上動作發出
            }
        } catch as err {
            ; 靜默處理，不干擾存檔流程
            OutputDebug("[RisController] EnsureImpressionNotEmpty Error: " . err.Message)
        }
    }

    static SaveReport() {
        try {
            this.EnsureImpressionNotEmpty()
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

        ; 1. 取得滑鼠游標下的控制項 Handle (hMouseCtrl)
        MouseGetPos , , , &hMouseCtrl, 2

        ; 2. 驗證是否為目標欄位 (Finding 或 Impression)
        isTarget := false
        try {
            ; 嘗試取得目前物件中定義的兩個 Edit 的 Handle
            ; 注意：這裡會觸發 getter，如果 cache 有值會很快，沒值會透由 UIA 抓一次
            hFind := this.FindingEdit.NativeWindowHandle
            hImp  := this.ImpressionEdit.NativeWindowHandle

            ; 比對滑鼠下的 handle 是否等於其中之一
            if (hMouseCtrl && (hMouseCtrl == hFind || hMouseCtrl == hImp)) {
                isTarget := true
            }
        }

        ; 3. 如果不是目標欄位，重置計數並退出
        ;    這樣可以防止你在別的地方點兩下，移過來點一下就觸發全選
        if (!isTarget) {
            clickCount := 0
            return
        }

        ; 4. 計算點擊時間差
        timeSinceLast := A_TickCount - lastClickTime
        if (timeSinceLast <= DoubleClickTime) {
            clickCount++
        } else {
            clickCount := 1
        }

        lastClickTime := A_TickCount

        ; 5. 觸發三擊全選
        if (clickCount == 3) {
            clickCount := 0
            this._SelectLine(hMouseCtrl)
        }
    }

    ; 會診補時方法
    static AddConsultationTime(offsetMinutes := 20) {
        if !WinActive(this.ConsultationWinTitle) {
            this.Notify("請先切換至會診資訊視窗")
            return
        }

        try {
            ; 1. 取得原始時間字串 (格式: 115/01/17 08:47)
            rawText := ControlGetText(this._ConsultationCtrls["SourceTime"], this.ConsultationWinTitle)

            if !RegExMatch(rawText, "^(\d{3})/(\d{2})/(\d{2})\s+(\d{2}):(\d{2})", &m) {
                throw Error("無法解析原始時間格式")
            }

            ; 2. 轉換為 AHK 標準時間戳 (YYYYMMDDHH24MISS)
            ; 民國年 m[1] + 1911 = 西元年
            standardTime := Format("{:04}{:02}{:02}{:02}{:02}00", Integer(m[1]) + 1911, m[2], m[3], m[4], m[5])

            ; 3. 加上分鐘數
            newTime := DateAdd(standardTime, offsetMinutes, "Minutes")

            ; 4. 格式化回民國年字串 (YYY/MM/DD HH:MI)
            finalYear := SubStr(newTime, 1, 4) - 1911
            finalDate := Format("{:03}/{:02}/{:02} {:02}:{:02}",
                finalYear,
                SubStr(newTime, 5, 2),
                SubStr(newTime, 7, 2),
                SubStr(newTime, 9, 2),
                SubStr(newTime, 11, 2)
            )

            ; 5. 填入目標控制項
            ControlSetText(finalDate, this._ConsultationCtrls["TargetTime"], this.ConsultationWinTitle)
            this.Notify("時間已補上 " . offsetMinutes . " 分鐘")

        } catch as err {
            this.Notify("補時失敗: " . err.Message)
        }
    }

    ; [新增] 啟動背景自動更新機制 (請在腳本啟動時呼叫此方法)
    static EnableAutoWorklistUpdate() {
        ; 檢查頻率: 每 5 分鐘偵測一次 (300,000 ms)
        checkFrequency := 300000

        OutputDebug("[RisAuto] >>> 自動更新機制已啟動 <<<`n")
        OutputDebug("[RisAuto] 設定：每 " . (checkFrequency/60000) . " 分鐘偵測一次閒置狀況`n")

        SetTimer(ObjBindMethod(this, "_CheckAutoUpdate"), checkFrequency)
    }

    static _CheckAutoUpdate() {
        OutputDebug("[RisAuto] --- Timer 心跳檢查 (" . A_Hour . ":" . A_Min . ") ---`n")

        ; A. 防重疊鎖 (Mutex)
        if (this._isUpdating) {
            OutputDebug("[RisAuto] ⚠️ 跳過：上一次更新尚未完成 (可能卡住)`n")
            return
        }

        ; B. 環境預檢：檢查 RDP/Session 是否鎖定
        ; 如果 User32\OpenInputDesktop 失敗，代表畫面被鎖定 (RDP斷線 或 Win+L)
        ; 此時 UIA 必死無疑，直接跳過，等待下次連線恢復
        if !DllCall("User32\OpenInputDesktop", "uint", 0, "int", 0, "uint", 0, "ptr") {
            OutputDebug("[RisAuto] 💤 狀態：Session 已鎖定或無畫面 (RDP 斷線中)，暫停動作`n")
            return
        }

        ; C. 視窗存在檢查
        if !WinExist(this.WorklistWinTitle) {
            OutputDebug("[RisAuto] ℹ️ 狀態：找不到工作清單視窗`n")
            return
        }

        ; D. 冷卻時間檢查
        timeSinceLast := A_TickCount - this._lastUpdateTick
        if (this._lastUpdateTick != 0 && timeSinceLast < this._updateInterval) {
            minutesLeft := Round((this._updateInterval - timeSinceLast) / 60000, 1)
            OutputDebug("[RisAuto] ⏳ 狀態：冷卻中 (尚需 " . minutesLeft . " 分鐘)`n")
            return
        }

        ; E. 閒置檢查
        if (A_TimeIdle < this._idleThreshold) {
            idleSec := Round(A_TimeIdle / 1000, 1)
            OutputDebug("[RisAuto] ✋ 狀態：使用者活動中 (閒置 " . idleSec . "s < 門檻)`n")
            return
        }

        OutputDebug("[RisAuto] 🚀 條件全數吻合，開始執行更新...`n")

        ; F. 執行更新 (使用 Try-Finally 確保鎖定解除)
        this._isUpdating := true
        try {
            this.GetWorklistJson(true)
        } catch as err {
            OutputDebug("[RisAuto] ❌ 更新發生錯誤: " . err.Message . "`n")
        } finally {
            this._isUpdating := false
        }
    }

    ; 3. 讀取與上傳主程式 (回復為 UIA 版本)
    static GetWorklistJson(isAuto := false) {
        Log := (msg) => (isAuto ? OutputDebug("[RisAuto] " . msg . "`n") : this.Notify(msg))

        if !WinExist(this.WorklistWinTitle) {
            return
        }

        this._ShowWaitCursor()
        try {
            hwnd := WinExist(this.WorklistWinTitle)
            elWindow := UIA.ElementFromHandle(hwnd)

            ; 點擊更新按鈕
            try {
                btnSelector := this._WorklistCtrls["RefreshButton"]
                elBtn := elWindow.FindElement(btnSelector)
                try {
                    elBtn.Invoke()
                } catch {
                    elBtn.Click()
                }
            } catch {
                Log("⚠️ 無法點擊更新按鈕")
                return
            }

            Sleep(1500) ; 等待重新整理

            ; 讀取表格資料
            categories := ["ER", "ADM", "OPD"]
            jsonStr := "{"
            validDataCount := 0

            for i, cat in categories {
                gridData := this._ExtractGridData(elWindow, this._WorklistCtrls[cat])
                validDataCount += gridData.Count

                if (i > 1)
                    jsonStr .= ", "
                jsonStr .= '"' . cat . '": {'

                isFirstProp := true
                for k, v in gridData {
                    if (!isFirstProp)
                        jsonStr .= ", "
                    safeKey := StrReplace(k, "-", "_")
                    valStr := IsNumber(v) ? v : '"' . v . '"'
                    jsonStr .= Format('"{1}": {2}', safeKey, valStr)
                    isFirstProp := false
                }
                jsonStr .= "}"
            }
            jsonStr .= "}"

            if (validDataCount == 0) {
                Log("⚠️ 統計資料為空，略過上傳")
                return
            }

            this._lastUpdateTick := A_TickCount
            this.PostDataToWebhook(jsonStr, isAuto)

        } catch as err {
            Log("操作失敗: " . err.Message)
        } finally {
            this._RestoreCursor()
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
        ; 1. 取得目前實際視窗的 HWND (WinExist 速度極快，即使頻繁呼叫也無所謂)
        currentHwnd := WinExist(this.WinTitle)
        if !currentHwnd {
            throw TargetError("找不到 RIS 視窗")
        }

        ; =================================================================
        ; 2. [關鍵修改] 視窗身分驗證 (Window Identity Check)
        ; 如果 Cache 裡記錄的 HWND 與目前的 HWND 不同，代表視窗重開過。
        ; 此時必須「清空所有快取」，避免拿到上一個視窗的殭屍物件。
        ; =================================================================
        if (!this._cache.Has("_Hwnd") || this._cache["_Hwnd"] != currentHwnd) {
            this._cache := Map()          ; 清空所有快取
            this._cache["_Hwnd"] := currentHwnd ; 更新為新的 HWND
        }

        ; 3. 經過上面的檢查，如果 nodeName 還在 cache 裡，代表它屬於目前的視窗，可直接回傳
        if this._cache.Has(nodeName) {
            return this._cache[nodeName]
        }

        ; 4. 如果不在 cache 裡，則重新抓取 (Fetch Logic)
        if (nodeName = "Ris") {
            try {
                this._cache["Ris"] := UIA.ElementFromHandle(currentHwnd)
                return this._cache["Ris"]
            } catch as err {
                throw Error("Root Error: " err.Message)
            }
        } else {
            ; 這裡會遞迴呼叫 this.Ris，而 this.Ris 會走上面的邏輯，確保拿到新的 Root
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
        if !IsObject(el) {
            return ""
        }
        
        rawText := ""
        ; 優先嘗試使用 NativeWindowHandle (Win32 API)，這對 WinForms 編輯器最穩定
        try {
            hwnd := el.NativeWindowHandle
            if (hwnd) {
                rawText := ControlGetText(hwnd)
            }
        } catch {
            rawText := ""
        }

        ; 如果 Win32 API 失敗，才嘗試 UIA 屬性
        if (rawText == "") {
            try {
                rawText := el.Value
            } catch {
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

    static GetComparisonDate() {
        currentMRN := this._GetCurrentMRN()
        
        ; 只有當曾經執行過 Copy Report (AppendPreviousReport) 導致 Date 有值時才進行
        if (this._compContext.Date != "") {
            ; 寬鬆比對：只有在「兩者都有抓到 MRN」且「兩者不同」的情況下，才視為換病人並拒絕插入
            if (currentMRN != "" && this._compContext.MRN != "" && currentMRN != this._compContext.MRN) {
                return ""
            }
            return this._compContext.Date
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
        isSpine := false
        startLineNo := 1
        if (RegExMatch(selectedText, "^(\d+)", &existLineNo)) {
            startLineNo := existLineNo[1]
        }

        for index, line in txtAry {
            if (!RegExMatch(line, "^\s*$")) {
                tmpText := line
                if (RegExMatch(line, "^\s*[-\+\*]*\s*([Vv]arying degree|[Mm]ild).+causing:")) {
                    isSpine := true
                }

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
                    tmpText := RegExReplace(tmpText, "\s*\((Srs|Ser)\/Img:.+?\)", "")
                    tmpText := RegExReplace(tmpText, "Mark L\d+:\s*", "")
                }

                ; =========================================================
                ; [優化] 句尾標點自動補全機制 (符合正式英語文法)
                ; =========================================================
                ; 1. 徹底去除尾部空白與 Tab
                tmpText := RTrim(tmpText, " `t")

                if (tmpText != "") {
                    ; 2. 處理因正則刪除資訊而殘留的不合法結尾（逗號或分號），轉換為句號
                    if RegExMatch(tmpText, "[,;]$") {
                        tmpText := SubStr(tmpText, 1, -1) . "."
                    }
                    ; 3. 確保句尾有合法的斷句符號 (. : ? !)
                    ;    排除 [:] 以保留 "Liver:" 這類標題的正確性
                    else if !RegExMatch(tmpText, "[.:?!]$") {
                        tmpText .= "."
                    }
                }
                ; =========================================================

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
        ; 保持 Scroll 位置邏輯
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
    }

    static _FormatFindingForBasic(hEdit) {
        fullText := ControlGetText(hEdit)
        range := this._FindContentRange(fullText, "Basic")

        if (range) {
            this._EditSetSel(hEdit, range.Start, range.End)
            this._ReorderSelectedText(false, true, "-", false, hEdit)
        } else {
            this.Notify("報告格式不如預期，無法自動排版")
        }
    }

    static _FormatFindingForAdvanced(hEdit) {
        fullText := ControlGetText(hEdit)
        range := this._FindContentRange(fullText, "Advanced")

        if (range) {
            this._EditSetSel(hEdit, range.Start, range.End)
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

        ; 輔助函數：判斷字元類型 (1:空白, 2:單字, 3:符號)
        GetCharType(char) {
            if (this._IsSpace(char))
                return 1
            if (IsAlnum(char) || char == "_")
                return 2
            return 3
        }

        ; --- 階段一：先貪婪地吃掉所有緊鄰的空白 ---
        while (i > 0 && this._IsSpace(SubStr(fullText, i, 1))) {
            i--
        }

        ; --- 階段二：空白吃完後，接著吃掉緊鄰的「那一組」東西 ---
        ; 如果還沒刪到頭，就檢查現在停在什麼字元上 (單字 或是 符號)
        if (i > 0) {
            targetType := GetCharType(SubStr(fullText, i, 1))

            ; 繼續往回刪，直到遇到「不同類型」的東西 (例如遇到另一個空白，或是從單字變符號)
            while (i > 0) {
                currentChar := SubStr(fullText, i, 1)
                ; 如果遇到類型不同 (例如原本在刪單字，現在遇到 . 或 空白)，就停止
                if (GetCharType(currentChar) != targetType)
                    break
                i--
            }
        }

        this._EditSetSel(hCtrl, i, caretPos)
        this._EditReplaceSel(hCtrl, "")
    }

    ; ----------------------------------------------------------------------------------
    ; [新增] 邏輯行邊界計算 Helper
    ; 回傳 Map: {Start: 0-based索引, ContentEnd: 不含換行, FullEnd: 含換行}
    ; ----------------------------------------------------------------------------------
    ; [修改] 增加 specificPos 參數以支援查詢任意位置的行邊界
    static _GetLogicalLineBoundaries(hCtrl, specificPos := -1) {
        try {
            fullText := ControlGetText(hCtrl)
        } catch {
            return {Start: 0, ContentEnd: 0, FullEnd: 0}
        }

        ; 如果有指定位置則使用指定位置，否則抓取目前 Caret
        if (specificPos != -1) {
            caretPos := specificPos
        } else {
            sel := this._EditGetSel(hCtrl)
            caretPos := sel.Start
        }

        ; 1. 找開頭 (Start): 往前找 `n
        ; InStr 是 1-based，caretPos + 1 確保從游標處包含搜尋
        prevLineBreak := InStr(fullText, "`n", , caretPos + 1, -1)

        ; 換行符號在 prevLineBreak，下一字元(行首)的位置剛好等於 prevLineBreak 的數值 (因為 0-based 轉換關係)
        lineStart := (prevLineBreak == 0) ? 0 : prevLineBreak

        ; 2. 找結尾 (End): 往後找 `r
        nextLineBreak := InStr(fullText, "`r", , lineStart + 1)

        if (nextLineBreak == 0) {
            ; 最後一行，無換行符號
            contentEnd := StrLen(fullText)
            fullEnd := contentEnd
        } else {
            ; 找到 \r，ContentEnd 在 \r 之前 (index - 1)
            contentEnd := nextLineBreak - 1

            ; 判斷是否包含 \n (FullEnd)
            if (SubStr(fullText, nextLineBreak + 1, 1) == "`n") {
                fullEnd := nextLineBreak + 1 ; \r\n 之後
            } else {
                fullEnd := nextLineBreak ; 只有 \r 之後
            }
        }

        return {Start: lineStart, ContentEnd: contentEnd, FullEnd: fullEnd}
    }

    static _SelectLine(hCtrl) {
        bounds := this._GetLogicalLineBoundaries(hCtrl)
        if (bounds.FullEnd > bounds.Start) {
            this._EditSetSel(hCtrl, bounds.Start, bounds.FullEnd)
        }
    }

    ; [新增] 用於刪除整行時的選取邏輯 (共用於 CutLineOrSelection 與 DeleteCurrentLine)
    static _SelectLineForRemoval(hCtrl) {
        bounds := this._GetLogicalLineBoundaries(hCtrl)

        ; 判斷是否為最後一行
        ; 如果 FullEnd (含換行結尾) 等於 ContentEnd (內容結尾)，表示後面沒有換行符號
        isLastLine := (bounds.FullEnd == bounds.ContentEnd)

        if (!isLastLine) {
            ; === 情況 A: 普通行 ===
            ; 刪除行為：刪除整行 + 後方換行符號
            this._EditSetSel(hCtrl, bounds.Start, bounds.FullEnd)
        } else {
            ; === 情況 B: 最後一行 (標準編輯器行為) ===

            if (bounds.Start == 0) {
                ; 特例：文件只有這一行，直接清空
                this._EditSetSel(hCtrl, 0, bounds.FullEnd)
            } else {
                ; 標準刪除：刪除「前方」換行符號 + 整行內容
                ; 範圍：從 (Start - 2) 到 FullEnd
                ; 游標行為：Win32 Edit Control 清除後，游標會自然停在刪除點的位置 (即上一行的結尾)
                this._EditSetSel(hCtrl, bounds.Start - 2, bounds.FullEnd)
            }
        }
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

    ; =================================================================
    ; [精簡版] 既然檔案內含多重解析度，直接讀取並指定大小即可
    ; =================================================================
    static _ShowWaitCursor() {
        try {
            OCR_NORMAL := 32512
            OCR_IBEAM := 32513
            IMAGE_CURSOR := 2
            LR_LOADFROMFILE := 0x0010

            ; 1. 取得目前系統游標大小 (這是清晰的關鍵)
            cx := DllCall("GetSystemMetrics", "Int", 13)
            cy := DllCall("GetSystemMetrics", "Int", 14)

            ; 2. 指定游標檔案
            ;    由於現代 Windows 的 .ani 檔通常已內含所有尺寸，
            ;    直接讀取標準檔名，讓 LoadImage 去挑選最適合的圖層即可。
            cursorPath := A_WinDir . "\Cursors\aero_working.ani"

            ; 防呆：萬一檔案不存在
            if !FileExist(cursorPath) {
                cursorPath := A_WinDir . "\Cursors\wait.cur"
            }

            ; 3. 載入並指定大小 (LoadImage 會自動從檔案中抓出符合 cx/cy 的高清圖層)
            hCursorRaw := DllCall("LoadImage", "Ptr", 0, "Str", cursorPath, "UInt", IMAGE_CURSOR, "Int", cx, "Int", cy, "UInt", LR_LOADFROMFILE, "Ptr")

            if (!hCursorRaw)
                return

            ; 4. 複製與替換
            hCopyNormal := DllCall("CopyImage", "Ptr", hCursorRaw, "UInt", IMAGE_CURSOR, "Int", cx, "Int", cy, "UInt", 0, "Ptr")
            hCopyIBeam  := DllCall("CopyImage", "Ptr", hCursorRaw, "UInt", IMAGE_CURSOR, "Int", cx, "Int", cy, "UInt", 0, "Ptr")

            DllCall("SetSystemCursor", "Ptr", hCopyNormal, "Int", OCR_NORMAL)
            DllCall("SetSystemCursor", "Ptr", hCopyIBeam, "Int", OCR_IBEAM)

            ; 5. 清理
            DllCall("DestroyCursor", "Ptr", hCursorRaw)
        }
    }

    static _RestoreCursor() {
        ; SPI_SETCURSORS = 0x0057, 重置系統所有游標回預設值
        DllCall("SystemParametersInfo", "UInt", 0x0057, "UInt", 0, "Ptr", 0, "UInt", 0)
    }

    ; =================================================================
    ; [新增] 共用的內容範圍搜尋邏輯
    ; @param text  全文
    ; @param mode  "Basic" or "Advanced"
    ; @return      {Start: index, End: index} or false (if not found)
    ; =================================================================
    static _FindContentRange(text, mode) {
        startPos := 0
        endPos := -1

        if (mode == "Advanced") {
            ; CT/MR 的起始關鍵字
            if RegExMatch(text, "m)FINDINGS:\r?\n|The study shows:\r?\n\r?\n|show the following findings:\r?\n\r?\n|which revealed:\r?\n\r?\n", &match) {
                startPos := match.Pos + match.Len - 1

                ; CT/MR 特有的結尾偵測 (REMARKS/RECOMMENDATION)
                if RegExMatch(text, "m)(\r\n){1,2}REMARKS?:|RECOMMENDATION:", &endMatch, startPos + 1) {
                    endPos := endMatch.Pos - 1
                }
                return {Start: startPos, End: endPos}
            }
        } else {
            ; Basic (CR/US) 的起始關鍵字
            if RegExMatch(text, "m)FINDINGS:\r?\n|:\s*\r?\n\s*\r?\n", &match) {
                startPos := match.Pos + match.Len - 1
                return {Start: startPos, End: -1} ; Basic 預設選到最後
            }
        }

        return false
    }

    ; [紅色特效版] 紅色 + 半透明 + 圓形
    ; [最佳化順序版] 先裁切再顯示 (防閃爍)
    static _HighlightCaret(hTargetCtrl := 0) {
        try {
            ; 1. 設定座標模式 & 關閉 DPI 縮放
            CoordMode "Caret", "Screen"
            CoordMode "Mouse", "Screen"

            x := 0, y := 0
            isFound := false

            ; 2. 抓取座標
            if CaretGetPos(&cx, &cy) {
                x := cx
                y := cy
                isFound := true
            } else if (hTargetCtrl) {
                try {
                    WinGetPos(&wx, &wy, &ww, &wh, "ahk_id " hTargetCtrl)
                    x := wx + (ww / 2) - 20
                    y := wy + (wh / 2) - 20
                    isFound := true
                }
            }

            if (!isFound)
                return

            ; 3. 計算圓心位置
            if (x == cx) {
                finalX := x - 20
                finalY := y - 10
            } else {
                finalX := x
                finalY := y
            }

            ; 4. 建立 GUI (保留 -DPIScale)
            g := Gui("+AlwaysOnTop -Caption +ToolWindow +E0x20 +E0x08000000 -DPIScale")
            g.BackColor := "Red"

            ; =========================================================
            ; [優化] 在顯示之前，先設定好形狀與透明度
            ; 這樣顯示出來的瞬間就已經是完美的圓形，不會有方塊閃爍
            ; =========================================================
            try {
                ; 設定圓形 (注意：這裡的 w40 h40 要跟 Show 裡面的大小一致)
                WinSetRegion("0-0 w40 h40 E", g.Hwnd)

                ; 設定半透明
                WinSetTransparent(100, g.Hwnd)
            }

            ; 5. 最後才顯示 GUI
            g.Show("NA x" finalX " y" finalY " w40 h40")

            ; 6. 自動銷毀
            SetTimer () => (IsObject(g) ? g.Destroy() : ""), -400

        } catch {
            ; 靜默失敗
        }
    }

    ; [新增] 通用表格讀取 Helper (回傳 Map 物件)
    static _ExtractGridData(elWindow, gridSelector) {
        data := Map()
        try {
            ; 1. 尋找表格
            try {
                elGrid := elWindow.FindElement(gridSelector)
            } catch {
                return data ; 如果找不到該表格(例如該院區無此單位)，回傳空 Map
            }

            ; 2. 找出所有資料列 (Custom 容器)
            try {
                rowElements := elGrid.FindAll({Type: "Custom"})
            } catch {
                return data
            }

            if (rowElements.Length == 0)
                return data

            ; 3. 遍歷讀取
            walker := UIA.TreeWalkerTrue

            for row in rowElements {
                ; 抓 Key (第一個子物件)
                keyEl := walker.TryGetFirstChildElement(row)
                if (!keyEl)
                    continue

                ; 抓 Value (下一個兄弟)
                valEl := walker.TryGetNextSiblingElement(keyEl)
                if (!valEl)
                    continue

                ; 讀取數值
                k := "", v := "0"
                try k := keyEl.Value
                try v := valEl.Value

                if (k != "") {
                    k := Trim(k)
                    v := Trim(v)
                    if (v == "")
                        v := 0
                    data[k] := v
                }
            }
        } catch {
            ; 忽略單一表格的非預期錯誤，避免影響整體流程
        }
        return data
    }

    ; [新增] 發送 JSON 至 Webhook
    ; [修改] 新增 isSilent 參數
    static PostDataToWebhook(jsonStr, isSilent := false) {
        ; 定義內部 Log
        Log := (msg) => (isSilent ? OutputDebug("[RisPost] " . msg . "`n") : this.Notify(msg))

        configFile := "config.private.ini"
        url  := IniRead(configFile, "n8n", "WebhookURL", "")
        user := IniRead(configFile, "n8n", "Username", "")
        pass := IniRead(configFile, "n8n", "Password", "")

        if (url == "") {
            Log("❌ 錯誤：找不到 WebhookURL 設定")
            return
        }

        try {
            req := ComObject("WinHttp.WinHttpRequest.5.1")
            req.Open("POST", url, False)
            req.SetRequestHeader("Content-Type", "application/json")

            if (user != "" && pass != "") {
                authStr := this._Base64Encode(user . ":" . pass)
                req.SetRequestHeader("Authorization", "Basic " . authStr)
            }

            req.Send(jsonStr)

            if (req.Status == 200) {
                Log("✅ 資料已上傳至 n8n")
            } else {
                Log("❌ 上傳失敗 (Status: " . req.Status . ")")
            }
        } catch as err {
            Log("❌ 網路錯誤: " . err.Message)
        }
    }

    ; [新增] 用於 Basic Auth 的 Base64 編碼 Helper
    static _Base64Encode(text) {
        buf := Buffer(StrPut(text, "UTF-8"))
        StrPut(text, buf, "UTF-8")

        ; CRYPT_STRING_BASE64 = 0x00000001
        ; CRYPT_STRING_NOCRLF = 0x40000000 (不換行)
        flags := 0x40000001

        reqSize := 0
        DllCall("Crypt32\CryptBinaryToStringW", "Ptr", buf, "UInt", buf.Size - 1, "UInt", flags, "Ptr", 0, "UInt*", &reqSize)

        outBuf := Buffer(reqSize * 2)
        DllCall("Crypt32\CryptBinaryToStringW", "Ptr", buf, "UInt", buf.Size - 1, "UInt", flags, "Ptr", outBuf, "UInt*", &reqSize)

        return StrGet(outBuf)
    }

    ; =================================================================
    ; 10. AI 應用功能 (AI & NLP Integration)
    ; =================================================================

    ; [新增] 外部呼叫的主函式：產生並插入 Indication
    ; [修改] 增加 Benchmark 效能測量
    static GenerateAndInsertIndication(debugMode := false) {
        this._ShowWaitCursor()
        try {
            ; --- Benchmark: 記錄開始時間 ---
            t0 := A_TickCount

            ; 1. 取得並組合病歷資料
            clinicalData := this._GetAndFormatClinicalData()
            if (clinicalData == "") {
                this.Notify("無法取得病歷資料，請確認是否在正確視窗內")
                return
            }

            ; --- Benchmark: 計算取資料耗時 ---
            t1 := A_TickCount
            extractTime := t1 - t0

            ; 2. 準備 Prompt
            systemPrompt := "[Role]`nYou are a professional Radiologist assistant specialized in clinical data extraction.`n`n[Background]`nThe following is a patient's medical record in SOAP format, including demographics and the planned imaging study.`n`n[Task]`nSummarize the core clinical reason (indication) for the requested imaging study into one or two concise English sentences.`n`n[Input Data]`n"
            constraint := "`n`n[Constraint]`n1. Start the response strictly with the prefix `"INDICATION:`".`n2. Focus on the mechanism of injury (e.g., collision), symptoms (e.g., thigh pain), and suspected diagnosis (e.g., femur fracture).`n3. Do not include unrelated physical exam findings (like heart/lung sounds) unless abnormal.`n4. Output in professional medical English.`n5. Note: Dates and specific identifiers in the text have been replaced with placeholders like [DATE] or [PATIENT_NAME] for privacy. Please ignore the placeholders and focus on the clinical findings.`n`n[Output]`nINDICATION:"

            fullPrompt := systemPrompt . clinicalData . constraint

            ; 3. Debug 模式：顯示 Prompt 與取資料耗時並可中斷
            if (debugMode) {
                A_Clipboard := fullPrompt
                ans := MsgBox("Debug 模式開啟。`n【Benchmark】資料提取耗時: " . extractTime . " ms`n`nPrompt 已複製到剪貼簿。是否繼續呼叫 API？`n`n" . SubStr(fullPrompt, 1, 500) . "...", "AI Debug", "YesNo")
                if (ans == "No") {
                    return
                }
            }

            ; 4. 從設定檔讀取模型名稱
            configFile := "config.private.ini"
            modelName := IniRead(configFile, "GoogleAI", "Model", "gemini-2.5-flash")

            ; --- Benchmark: 記錄 API 呼叫前時間 ---
            t2 := A_TickCount

            ; 呼叫 Google AI
            result := this._CallGoogleAI(fullPrompt, modelName)

            ; --- Benchmark: 計算 API 耗時 ---
            t3 := A_TickCount
            apiTime := t3 - t2

            ; 由於 Prompt 要求嚴格以 "INDICATION:" 開頭，若 API 返回時缺少或格式異常，可做基本處理
            if (!InStr(result, "INDICATION:")) {
                result := "INDICATION: " . result
            }

            if (debugMode) {
                MsgBox("【Benchmark】`n資料提取: " . extractTime . " ms`nAPI 耗時: " . apiTime . " ms`n`n【API 回傳結果】`n" . result, "AI Debug")
            }

            ; 5. 插入結果至目標欄位
            if !WinActive(this.WinTitle) {
                WinActivate(this.WinTitle)
                WinWaitActive(this.WinTitle, , 2)
            }

            targetHwnd := 0

            if this.IsTargetFocused() {
                try {
                    targetHwnd := ControlGetFocus("A")
                }
            }

            if (!targetHwnd) {
                this.FindingEdit.SetFocus()
                targetHwnd := this.FindingEdit.NativeWindowHandle
                Sleep(50)
            }

            this._EditReplaceSel(targetHwnd, result . "`r`n`r`n")
            this._EditScrollCaret(targetHwnd)

            ; 將 Benchmark 數據顯示在完成的 Notify 中
            this.Notify(Format("已插入 Indication (取資:{}ms, API:{}ms)", extractTime, apiTime))
        } catch as err {
            if (debugMode) {
                fullErrorMsg := "【錯誤訊息】`n" . err.Message . "`n`n【發生位置】`n" . err.What . "`n`n【呼叫堆疊】`n" . err.Stack
                this._ShowDebugError(fullErrorMsg)
            } else {
                this.Notify("AI 處理失敗: " . err.Message)
            }
        } finally {
            this._RestoreCursor()
        }
    }

    ; [新增] 極速讀取 Helper：繞過 UIA Fallback，直接調用 Win32 API
    static _FastGetCtrlText(propName) {
        try {
            ; 利用 AHK v2 的動態屬性 (this.%propName%) 觸發 Getter
            ; 直接向底層 Handle 拿文字，不走 UIA 備用機制
            return ControlGetText(this.%propName%.NativeWindowHandle)
        } catch {
            return "" ; 欄位不存在或發生例外時，安全回傳空字串
        }
    }

    ; [修改] 使用 _FastGetCtrlText 取代原本的 this.GetText()
    static _GetAndFormatClinicalData() {
        ; [修改] 防呆：確保全面預載邏輯一定有跑過
        this._PreloadCache()

        ; 逐一嘗試抓取，即使某個欄位找不到也不會中斷整個字串的組合
        gender := this._FastGetCtrlText("GenderText")
        age    := this._FastGetCtrlText("AgeText")
        exam   := this._FastGetCtrlText("ExamnameText")
        dept   := this._FastGetCtrlText("OrderDeptText")

        sText  := this._FastGetCtrlText("SubjectiveText")
        oText  := this._FastGetCtrlText("ObjectiveText")
        aText  := this._FastGetCtrlText("AssessmentText")
        pText  := this._FastGetCtrlText("PlanText")

        ; 如果全部欄位都是空的，代表可能抓錯視窗
        if (gender == "" && sText == "" && oText == "") {
            return ""
        }

        rawText := Format("{1}`n{2}`n{3}`n{4}`n`nS: {5}`n`nO: {6}`n`nA: {7}`n`nP: {8}", gender, age, exam, dept, sText, oText, aText, pText)

        return this._DeidentifyText(rawText)
    }

    static _DeidentifyText(text) {
        if (text == "") {
            return ""
        }

        ; 1. 處理日期：民國/西元 (115/03/13, 2026-03-13, 115.3.13)
        text := RegExReplace(text, "i)(\d{3,4})[/\.\-]\d{1,2}[/\.\-]\d{1,2}", "[DATE]")

        ; 2. 處理連寫日期：(1150313), 20260313
        text := RegExReplace(text, "i)\(?\b(\d{7,8})\b\)?", "([DATE])")

        ; 3. 處理年齡：將 "29 歲 9 月" 轉換為 "20s-yo"
        ; 保留大致年齡段對放射科診斷有臨床價值，但移除精確月分
        if (RegExMatch(text, "(\d{1,2})\s*(歲|y|years?)", &m)) {
            age := Number(m[1])
            decade := Floor(age / 10) * 10
            text := RegExReplace(text, "\d{1,2}\s*(歲|y|years?)\s*(\d{1,2}\s*(月|m|months?))?", decade . "s-yo")
        }

        ; 4. 處理身分證字號 (台灣格式：首位字母 + 9位數字，涵蓋本國人 1/2 與外籍人士 8/9)
        text := RegExReplace(text, "i)[A-Z][1289]\d{8}", "[ID_REDACTED]")

        ; 5. 處理電話號碼 (09xx-xxx-xxx 或 02-xxxx-xxxx)
        text := RegExReplace(text, "i)0\d{1,2}-?\d{3,4}-?\d{3,4}", "[PHONE_REDACTED]")

        ; 6. 移除姓名標籤後的內容 (處理到行尾)
        text := RegExReplace(text, "i)(Name|姓名|Patient)\s*[:：]\s*\V+", "$1: [PATIENT_NAME]")

        return text
    }

    static _CallGoogleAI(promptText, modelName) {
        configFile := "config.private.ini"
        apiKey := IniRead(configFile, "GoogleAI", "APIKey", "")
        if (apiKey == "") {
            throw Error("請在 " . configFile . " 中設定 [GoogleAI] APIKey")
        }

        ; 組合 Google AI API 網址
        url := "https://generativelanguage.googleapis.com/v1beta/models/" . modelName . ":generateContent?key=" . apiKey

        ; JSON 逃脫處理 (針對雙引號與換行)
        escapedPrompt := StrReplace(promptText, "\", "\\")
        escapedPrompt := StrReplace(escapedPrompt, "`"", "\`"")
        escapedPrompt := StrReplace(escapedPrompt, "`n", "\n")
        escapedPrompt := StrReplace(escapedPrompt, "`r", "\r")
        escapedPrompt := StrReplace(escapedPrompt, "`t", "\t")

        payload := '{"contents": [{"parts": [{"text": "' . escapedPrompt . '"}]}]}'

        req := ComObject("WinHttp.WinHttpRequest.5.1")
        req.Open("POST", url, False)
        req.SetRequestHeader("Content-Type", "application/json")
        req.Send(payload)

        if (req.Status != 200) {
            throw Error("HTTP " . req.Status . " - " . req.ResponseText)
        }

        ; 簡易解析 Gemini API 回傳的 JSON (提取 content 內的 text)
        if (RegExMatch(req.ResponseText, 's)"text":\s*"(.*?)(?<!\\)"', &match)) {
            val := match[1]
            ; 還原 JSON 內的跳脫字元
            val := StrReplace(val, "\n", "`n")
            val := StrReplace(val, '\"', '"')
            val := StrReplace(val, "\\", "\")
            return val
        }

        throw Error("無法解析回傳的資料結構")
    }

    ; [新增] 專用於 Debug 的持久化錯誤視窗
    static _ShowDebugError(errMsg) {
        ; 建立新 GUI：置頂、有標題列、有外框
        errGui := Gui("+AlwaysOnTop +Resize", "AI Debug - 處理失敗")
        errGui.SetFont("s10", "Microsoft JhengHei UI")

        errGui.Add("Text", "w500", "API 呼叫或處理過程中發生例外錯誤：")

        ; 使用唯讀的 Edit 控制項來裝載可能很長的錯誤訊息，並啟用垂直捲軸
        errGui.Add("Edit", "w500 h250 ReadOnly Multi vErrText", errMsg)

        ; 一鍵複製按鈕
        btnCopy := errGui.Add("Button", "w120 x10 y+15", "📋 複製完整訊息")
        btnCopy.OnEvent("Click", (*) => (
            A_Clipboard := errMsg,
            this.Notify("已複製錯誤訊息至剪貼簿！", 2000)
        ))

        ; 關閉按鈕
        btnClose := errGui.Add("Button", "w100 x+270", "關閉")
        btnClose.OnEvent("Click", (*) => errGui.Destroy())

        ; 顯示在畫面中央
        errGui.Show("AutoSize Center")
    }
}