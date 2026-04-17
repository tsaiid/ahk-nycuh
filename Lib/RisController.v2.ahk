#Requires AutoHotkey v2.0
#Include .\UIA.v2.ahk
#Include .\RisConfig.v2.ahk

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
        SETREADONLY:   0x00CF, ; [新增] 用於設定唯讀狀態 (可用來切換底色)
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
        "Pos0",     {AutomationId: "rdoPos0"},
        "Pos1",     {AutomationId: "rdoPos1"},
        "Pos2",     {AutomationId: "rdoPos2"},
        "Pos3",     {AutomationId: "rdoPos3"},
        "Save",     {AutomationId: "btnSave"},
        "Cancel",   {AutomationId: "btnBack"}
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
        "PathologyOnlineDataRadio", { AutomationId: "rdoPathologyOnlineData" },
        "ImpressionLabel",  { AutomationId: "label2" },
        "MedRecNoLabel",    { AutomationId: "txtMRNo" },
        "AccessionNoText",  { AutomationId: "txtReqNo" },
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
        "RightTabControl",  { AutomationId: "tbRightList" },
        "ClinicalTabControl", { AutomationId: "tbcClinicalData" },
    )

    ; [修改] 相似檢查對應表，現在改為從外部檔案 (config/sim-groups.txt) 讀取
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
    ; 1.1 初始化與配置載入 (Initialization)
    ; =================================================================

    ; 類別載入時自動執行
    static __New() {
        this.LoadSimGroups()
    }

    /**
     * 從外部檔案載入相似檢查分組 (config/sim-groups.txt)
     * 支援「空白行」分隔群組，「#」或「;」作為註解
     */
    static LoadSimGroups() {
        ; 這裡路徑相對於主腳本目錄
        filePath := A_ScriptDir . "\config\sim-groups.txt"

        if !FileExist(filePath) {
            this.Notify("找不到相似分組設定檔`n" . filePath, 3000)
            return
        }

        try {
            content := FileRead(filePath, "UTF-8")

            ; 重置 Map
            this._SimReportMap := Map()
            this._SimReportMap.CaseSense := "Off" ; 設定為不分大小寫

            ; 解析邏輯：先依空白行切分區塊 (支援 \r\n 或 \n，且中間可有空格)
            ; 先將連續兩個以上的換行 (中間可夾雜空格) 替換為特殊標記字串 "[[BLOCK]]"
            tempContent := RegExReplace(content, "\R[ \t]*\R", "[[BLOCK]]")
            blocks := StrSplit(tempContent, "[[BLOCK]]")

            groupCount := 0
            debugInfo := ""

            for block in blocks {
                group := []
                ; 逐行處理區塊內的內容
                for line in StrSplit(block, "`n", "`r") {
                    line := Trim(line)
                    ; 略過空白行與註解行
                    if (line == "" || SubStr(line, 1, 1) == "#" || SubStr(line, 1, 1) == ";")
                        continue
                    group.Push(line)
                }

                ; 建立雙向關聯
                if (group.Length > 1) {
                    groupCount++
                    debugInfo .= "Group " . groupCount . ":`n"
                    for item in group {
                        debugInfo .= "  - " . item . "`n"
                        ; 確保每個項目都有一個對應的子 Map
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
                    debugInfo .= "`n"
                }
            }

            ; 將結果輸出到 Debug 視窗 (可用於追蹤)
            OutputDebug("`n[RisSimGroups] --- 對應表載入完成 ---`n" . debugInfo . "------------------------------`n")

            if (this.IsDebug)
                this.Notify("相似分組已載入 (" . groupCount . " 群組, " . this._SimReportMap.Count . " 項目)", 1500)

        } catch as err {
            this.Notify("載入分組失敗: " . err.Message)
        }
    }

    ; =================================================================
    ; 2. 內部狀態 (State)
    ; =================================================================
    static _uiCache := Map()
    static _stateCache := Map()
    static _aiCache := Map()
    static _googleAIConfig := 0
    static _notifyGui := 0
    static _notifyQueue := []
    static _notifySlots := []
    static _notifySlotItemIds := []
    static _notifySweepTimer := 0
    static _nextNotifyId := 0
    static _notifyMaxVisible := 5
    static _notifyDedupeWindow := 800
    static _notifyMinWidth := 320
    static _notifyWidth := 420
    static _notifyMaxWidth := 720
    static _notifyPaddingX := 24
    static _notifyPaddingY := 14
    static _notifySlotGap := 8
    static _notifySlotHeight := 36
    static _compContext := {ReqNo: "", Date: ""}
    static _workingCursorCache := {AppStarting: "", CursorBaseSize: "", Path: ""}
    static _hCustomFont := 0
    static _targetImpressionHeight := 95
    static _preloadQueue := [] ; [新增] 預載隊列
    static _preloadTask  := 0  ; [新增] 預載 Timer 參考
    static _indicationPreloadTask := 0 ; [新增] AI indication 預載排程
    static _isAIPending  := false ; [新增] AI 請求中旗標
    static _isIndicationPending := false ; [新增] indication 產生中旗標
    static _pendingIndicationInsert := false ; [新增] indication 完成後插入請求
    static _isShellHookEnabled := false ; [新增] ShellHook 狀態旗標
    static _shellTrack := Map()         ; [新增] 用於紀錄每個 HWND 的處理狀態 {time: TickCount, timer: Func}
    static _shellTrackTTL := 15000
    static IsDebug := false            ; [新增] Debug 模式切換

    ; [自動更新相關狀態]
    static _lastUpdateTick := 0           ; 上次更新的時間
    static _updateInterval := 1800000     ; 30 分鐘 (標準生產環境設定)
    static _idleThreshold  := 300000      ; 5 分鐘
    static _isUpdating     := false       ; 防卡死旗標

    ; =================================================================
    ; 2.1 責任邊界總覽 (Current Ownership)
    ; - UI cache / node resolve
    ; - 視窗焦點與 shell hook
    ; - 編輯器操作與文字插入
    ; - AI orchestration
    ; - AI transport / config
    ; - 工作清單與 webhook
    ; - Notify GUI
    ;
    ; 2.2 共享狀態歸屬 (Shared State)
    ; - Controller lifetime: _uiCache / _stateCache / _aiCache
    ; - UI preload / shell hook: _preloadQueue / _preloadTask / _indicationPreloadTask / _shellTrack
    ; - AI request gating: _isAIPending / _isIndicationPending / _pendingIndicationInsert
    ; - Notify runtime: _notifyGui / _notifyQueue / _notifySlots / _notifySweepTimer
    ; - Presentation state: _hCustomFont / _targetImpressionHeight / _compContext
    ;
    ; 2.3 後續可拆模組草案
    ; - RisUiCache: UI cache / node resolve / preload queue
    ; - RisEditorActions: editor actions, formatting, selection mutations
    ; - RisAiService: AI orchestration + transport facade
    ; - RisNotify: Notify GUI, queue, slot management
    ; =================================================================

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
    static PathologyOnlineDataRadio => this._GetOrUpdateNode("PathologyOnlineDataRadio")
    static PhExamColumn => this._GetOrUpdateNode("PhExamColumn")
    static PhExamDateText => this._GetOrUpdateNode("PhExamDateText")
    static PhExamReportText => this._GetOrUpdateNode("PhExamReportText")

    ; [新增] SOAP 與基本資料 Getters
    static MedRecNoLabel  => this._GetOrUpdateNode("MedRecNoLabel")
    static AccessionNoText => this._GetOrUpdateNode("AccessionNoText")
    static SubjectiveText => this._GetOrUpdateNode("SubjectiveText")
    static ObjectiveText  => this._GetOrUpdateNode("ObjectiveText")
    static AssessmentText => this._GetOrUpdateNode("AssessmentText")
    static PlanText       => this._GetOrUpdateNode("PlanText")
    static OrderDeptText  => this._GetOrUpdateNode("OrderDeptText")
    static GenderText     => this._GetOrUpdateNode("GenderText")
    static AgeText        => this._GetOrUpdateNode("AgeText")
    static RightTabControl => this._GetOrUpdateNode("RightTabControl")
    static ClinicalTabControl => this._GetOrUpdateNode("ClinicalTabControl")

    ; [新增] 快捷文字存取屬性 (使用快速路徑並自動標準化換行)
    static FindingText => this.GetText(this.FindingEdit)
    static ImpressionText => this.GetText(this.ImpressionEdit)

    ; [新增] 取得過濾後的 Finding 內文 (去除標題與結尾)
    static GetFindingContent() {
        try {
            fullText := this.FindingText

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

    static Notify(text, duration := 1500) {
        text := Trim(text)
        if (text == "")
            return

        this._PruneExpiredNotifications()

        existingId := this._FindRecentNotifyId(text)
        if (existingId) {
            this._RefreshNotifyDuration(existingId, duration)
            return
        }

        this._nextNotifyId += 1
        expiresAt := (duration > 0) ? (A_TickCount + duration) : 0
        this._notifyQueue.Push({
            id: this._nextNotifyId,
            text: text,
            createdAt: A_TickCount,
            duration: duration,
            expiresAt: expiresAt
        })

        if (this._notifyQueue.Length > this._notifyMaxVisible)
            this._notifyQueue.RemoveAt(1, this._notifyQueue.Length - this._notifyMaxVisible)

        this._EnsureNotifyGui()
        this._RenderNotifyQueue()
        this._UpdateNotifySweepTimer()
    }

    static _EnsureNotifyGui() {
        if this._notifyGui
            return this._notifyGui

        g := Gui("+AlwaysOnTop -Caption +ToolWindow +E0x08000000")
        g.BackColor := "202020"
        g.SetFont("s12 cWhite bold", "Microsoft JhengHei UI")
        g.MarginX := this._notifyPaddingX
        g.MarginY := this._notifyPaddingY

        this._notifyGui := g
        this._notifySlots := []
        this._notifySlotItemIds := []

        loop this._notifyMaxVisible {
            slot := g.Add("Text", Format("x{1} y{2} w{3} h{4} +0x200 Center Hidden", this._notifyPaddingX, this._notifyPaddingY, this._notifyWidth, this._notifySlotHeight), "")
            slot.OnEvent("Click", ObjBindMethod(this, "_HandleNotifySlotClick", A_Index))
            this._notifySlots.Push(slot)
            this._notifySlotItemIds.Push(0)
        }

        this._notifySweepTimer := ObjBindMethod(this, "_SweepNotifications")
        return g
    }

    static _RenderNotifyQueue() {
        g := this._EnsureNotifyGui()
        visibleCount := this._notifyQueue.Length
        innerWidth := this._GetNotifyContentWidth()
        height := this._notifyPaddingY * 2
        y := this._notifyPaddingY

        for index, slot in this._notifySlots {
            if (index <= visibleCount) {
                item := this._notifyQueue[index]
                displayText := this._WrapNotifyText(item.text, innerWidth)
                slotHeight := this._MeasureNotifyTextHeight(displayText)
                slot.Text := displayText
                slot.Move(this._notifyPaddingX, y, innerWidth, slotHeight)
                slot.Opt("-Hidden")
                this._notifySlotItemIds[index] := item.id
                height := y + slotHeight + this._notifyPaddingY
                y += slotHeight + this._notifySlotGap
            } else {
                slot.Text := ""
                slot.Opt("Hidden")
                this._notifySlotItemIds[index] := 0
            }
        }

        if (visibleCount = 0) {
            try g.Hide()
            return
        }

        totalWidth := innerWidth + this._notifyPaddingX * 2
        g.Show(Format("NoActivate Center w{1} h{2}", totalWidth, height))
        this._ApplyNotifyVisualStyle()
    }

    static _GetNotifyContentWidth() {
        width := this._notifyMinWidth

        for item in this._notifyQueue
            width := Max(width, this._MeasureNotifyNaturalWidth(item.text))

        return Min(this._notifyMaxWidth, width)
    }

    static _MeasureNotifyNaturalWidth(text) {
        hdcState := this._BeginNotifyTextMeasure()
        if !hdcState.hdc
            return this._notifyWidth

        width := 0
        sizeBuffer := Buffer(8, 0)

        for line in StrSplit(text, "`n", "`r") {
            lineText := (line = "") ? " " : line
            if DllCall("GetTextExtentPoint32", "Ptr", hdcState.hdc, "Str", lineText, "Int", StrLen(lineText), "Ptr", sizeBuffer.Ptr, "Int")
                width := Max(width, NumGet(sizeBuffer, 0, "Int"))
        }

        this._EndNotifyTextMeasure(hdcState)
        return Max(this._notifyMinWidth, width + 20)
    }

    static _MeasureNotifyTextHeight(text) {
        hdcState := this._BeginNotifyTextMeasure()
        if !hdcState.hdc
            return this._notifySlotHeight

        textHeight := 0
        sizeBuffer := Buffer(8, 0)
        for line in StrSplit(text, "`n", "`r") {
            lineText := (line = "") ? " " : line
            if DllCall("GetTextExtentPoint32", "Ptr", hdcState.hdc, "Str", lineText, "Int", StrLen(lineText), "Ptr", sizeBuffer.Ptr, "Int")
                textHeight += NumGet(sizeBuffer, 4, "Int")
        }
        this._EndNotifyTextMeasure(hdcState)
        return Max(this._notifySlotHeight, textHeight + 10)
    }

    static _WrapNotifyText(text, maxWidth) {
        wrappedLines := []
        for rawLine in StrSplit(text, "`n", "`r")
            this._AppendWrappedNotifyLine(wrappedLines, rawLine, maxWidth)

        return wrappedLines.Length ? this._JoinNotifyLines(wrappedLines) : text
    }

    static _AppendWrappedNotifyLine(lines, text, maxWidth) {
        if (text = "") {
            lines.Push("")
            return
        }

        current := ""
        tokenPattern := "(\s+|[^\s]+)"
        startPos := 1
        tokenFound := false

        while RegExMatch(text, tokenPattern, &match, startPos) {
            tokenFound := true
            token := match[1]
            candidate := current . token

            if (current = "" || this._MeasureNotifyLineWidth(candidate) <= maxWidth) {
                current := candidate
            } else {
                trimmed := Trim(current, " ")
                if (trimmed != "")
                    lines.Push(trimmed)
                this._AppendWrappedNotifyLongToken(lines, token, maxWidth, &current)
            }

            startPos := match.Pos + match.Len
        }

        if !tokenFound {
            this._AppendWrappedNotifyLongToken(lines, text, maxWidth, &current)
            tokenFound := true
        }

        finalLine := Trim(current, " ")
        if (tokenFound && finalLine != "")
            lines.Push(finalLine)
    }

    static _AppendWrappedNotifyLongToken(lines, token, maxWidth, &current) {
        token := Trim(token, " ")
        if (token = "") {
            current := ""
            return
        }

        if (this._MeasureNotifyLineWidth(token) <= maxWidth) {
            current := token
            return
        }

        current := ""
        chunk := ""
        loop parse token {
            char := A_LoopField
            candidate := chunk . char
            if (chunk = "" || this._MeasureNotifyLineWidth(candidate) <= maxWidth) {
                chunk := candidate
            } else {
                lines.Push(chunk)
                chunk := char
            }
        }
        current := chunk
    }

    static _MeasureNotifyLineWidth(text) {
        if (text = "")
            return 0

        hdcState := this._BeginNotifyTextMeasure()
        if !hdcState.hdc
            return this._notifyWidth

        sizeBuffer := Buffer(8, 0)
        width := 0
        if DllCall("GetTextExtentPoint32", "Ptr", hdcState.hdc, "Str", text, "Int", StrLen(text), "Ptr", sizeBuffer.Ptr, "Int")
            width := NumGet(sizeBuffer, 0, "Int")

        this._EndNotifyTextMeasure(hdcState)
        return width
    }

    static _JoinNotifyLines(lines) {
        result := ""
        for index, line in lines {
            if (index > 1)
                result .= "`n"
            result .= line
        }
        return result
    }

    static _BeginNotifyTextMeasure() {
        this._EnsureNotifyGui()

        if !this._notifySlots.Length
            return {hdc: 0, hwnd: 0, oldFont: 0}

        hwnd := this._notifySlots[1].Hwnd
        hdc := DllCall("GetDC", "Ptr", hwnd, "Ptr")
        if !hdc
            return {hdc: 0, hwnd: hwnd, oldFont: 0}

        hFont := SendMessage(0x0031, 0, 0, hwnd)
        oldFont := hFont ? DllCall("SelectObject", "Ptr", hdc, "Ptr", hFont, "Ptr") : 0
        return {hdc: hdc, hwnd: hwnd, oldFont: oldFont}
    }

    static _EndNotifyTextMeasure(hdcState) {
        if !hdcState.hdc
            return

        if hdcState.oldFont
            DllCall("SelectObject", "Ptr", hdcState.hdc, "Ptr", hdcState.oldFont, "Ptr")

        DllCall("ReleaseDC", "Ptr", hdcState.hwnd, "Ptr", hdcState.hdc)
    }

    static _ApplyNotifyVisualStyle() {
        if !this._notifyGui
            return

        hwnd := this._notifyGui.Hwnd
        WinGetPos(&x, &y, &w, &h, hwnd)

        try {
            WinSetRegion("0-0 w" w " h" h " r12-12", hwnd)

            style := DllCall("GetClassLongPtr", "Ptr", hwnd, "Int", -26, "Ptr")
            DllCall("SetClassLongPtr", "Ptr", hwnd, "Int", -26, "Ptr", style | 0x00020000)

            WinSetTransparent(235, hwnd)
        }
    }

    static _HandleNotifySlotClick(slotIndex, ctrl, *) {
        if (slotIndex < 1 || slotIndex > this._notifySlotItemIds.Length)
            return

        itemId := this._notifySlotItemIds[slotIndex]
        if !itemId
            return

        this._RemoveNotifyById(itemId)
    }

    static _RemoveNotifyById(itemId) {
        for index, item in this._notifyQueue {
            if (item.id = itemId) {
                this._notifyQueue.RemoveAt(index)
                this._RenderNotifyQueue()
                this._UpdateNotifySweepTimer()
                return
            }
        }
    }

    static _FindRecentNotifyId(text) {
        for index, item in this._notifyQueue {
            if (item.text == text && A_TickCount - item.createdAt <= this._notifyDedupeWindow)
                return item.id
        }
        return 0
    }

    static _RefreshNotifyDuration(itemId, duration) {
        for index, item in this._notifyQueue {
            if (item.id != itemId)
                continue

            item.createdAt := A_TickCount
            item.duration := duration
            item.expiresAt := (duration > 0) ? (A_TickCount + duration) : 0
            this._notifyQueue[index] := item
            this._RenderNotifyQueue()
            this._UpdateNotifySweepTimer()
            return
        }
    }

    static _PruneExpiredNotifications() {
        loop {
            removed := false
            for index, item in this._notifyQueue {
                if (item.expiresAt > 0 && A_TickCount >= item.expiresAt) {
                    this._notifyQueue.RemoveAt(index)
                    removed := true
                    break
                }
            }
        } until !removed
    }

    static _SweepNotifications() {
        this._PruneExpiredNotifications()
        this._RenderNotifyQueue()
        this._UpdateNotifySweepTimer()
    }

    static _UpdateNotifySweepTimer() {
        if !this._notifySweepTimer
            return

        SetTimer(this._notifySweepTimer, 0)

        for item in this._notifyQueue {
            if (item.expiresAt > 0) {
                SetTimer(this._notifySweepTimer, -150)
                return
            }
        }
    }

    /**
     * 自動搶回 RIS 報告視窗焦點 (封裝 ShellHook)
     * 當偵測到新視窗建立時，如果是 RIS 報告視窗，則延遲一小段時間後將其 Activate
     */
    static EnableShellHookFocus() {
        if (this._isShellHookEnabled)
            return

        DllCall("RegisterShellHookWindow", "Ptr", A_ScriptHwnd)
        ; 使用閉包確保參數對齊：(wParam, lParam, msg, hwnd)
        OnMessage(DllCall("RegisterWindowMessage", "Str", "SHELLHOOK"), (wp, lp, m, h) => this._ShellMessage(wp, lp, m, h))
        this._isShellHookEnabled := true

        if (this.IsDebug) {
            msg := "RIS 自動搶焦已啟動 (" . A_ScriptName . ")"
            this.Notify(msg, 1500)
        }
    }

    static _ShellMessage(wParam, lParam, msg, hwnd) {
        if (wParam = 1) { ; HSHELL_WINDOWCREATED
            this._PruneShellTrack()

            ; 這裡 lParam 是新視窗的 HWND
            if (this.IsDebug)
                OutputDebug("[RisShell] HSHELL_WINDOWCREATED received: HWND=" . lParam . " (Tick=" . A_TickCount . ")`n")

            ; [Map 防震邏輯]
            ; 如果 5 秒內已經處理過這個 HWND，且目前還有計時器正在跑，則跳過
            if (this._shellTrack.Has(lParam)) {
                track := this._shellTrack[lParam]
                if (A_TickCount - track.time < 5000) {
                    if (this.IsDebug)
                        OutputDebug("[RisShell] >>> Debounced (Map): HWND=" . lParam . "`n")
                    return
                }
            }

            try {
                title := WinGetTitle(lParam)
                ; 檢查是否為 RIS 報告視窗 (精確比對標題或包含關鍵字)
                if WinExist(lParam) && (title == this.WinTitle || InStr(title, "報告作業")) {
                    if (this.IsDebug)
                        OutputDebug("[RisShell] Target window detected. Title: " . title . "`n")

                    ; 建立固定的 Func 物件並存入 Map，確保 SetTimer 會正確重置而非重複啟動
                    timerFunc := ObjBindMethod(this, "_FocusRisWindow", lParam)
                    this._shellTrack[lParam] := {time: A_TickCount, timer: timerFunc}

                    ; 延遲 800ms 搶回焦點
                    SetTimer(timerFunc, -1800)
                }
            }
        }
    }

    static _FocusRisWindow(hwnd) {
        try {
            if !WinExist(hwnd)
                return

            ; [二次防震] 如果視窗已經是 Active 狀態且剛剛才處理過，就不再 Notify
            if WinActive(hwnd) && (this._shellTrack.Has(hwnd) && A_TickCount - this._shellTrack[hwnd].time < 1500) {
                return
            }

            ; 先確保視窗可見，再多次嘗試搶回前景焦點
            try WinShow(hwnd)

            activated := false
            Loop 3 {
                try WinActivate(hwnd)
                if WinWaitActive(hwnd, , 0.6) {
                    activated := true
                    break
                }
                Sleep 150
            }

            if !activated {
                try DllCall("SetForegroundWindow", "ptr", hwnd)
                try WinWaitActive(hwnd, , 0.5)
            }

            try this.FindingEdit.SetFocus()
            this._ScheduleWindowWarmup(hwnd)

            if (this.IsDebug) {
                msg := "已自動回歸 RIS 焦點 (" . A_ScriptName . " | " . hwnd . " | " . A_TickCount . ")"
                this.Notify(msg, 1500)
            }
        } finally {
            this._ClearShellTrack(hwnd)
        }
    }

    static _ScheduleWindowWarmup(hwnd, attempt := 1) {
        if !WinExist(hwnd)
            return

        currentHwnd := 0
        try currentHwnd := WinExist(this.WinTitle)
        if (currentHwnd != hwnd)
            return

        if !WinActive(hwnd) {
            if (attempt >= 4)
                return

            retryFunc := ObjBindMethod(this, "_ScheduleWindowWarmup", hwnd, attempt + 1)
            SetTimer(retryFunc, -250)
            return
        }

        this._StartWindowInitialization(hwnd)
    }

    static _StartWindowInitialization(hwnd) {
        if !WinExist(hwnd)
            return

        currentHwnd := 0
        try currentHwnd := WinExist(this.WinTitle)
        if (currentHwnd != hwnd)
            return

        if (this._stateCache.Has("_WindowInitialized") && this._stateCache["_WindowInitialized"] = hwnd)
            return

        try {
            hFind := this.FindingEdit.NativeWindowHandle
            hImp  := this.ImpressionEdit.NativeWindowHandle
        } catch {
            return
        }

        this._stateCache["_WindowInitialized"] := hwnd
        this._ApplyEditorFont(hFind, hImp)
        this._ApplyLayout(hFind, hImp)
        this._PreloadCache()
    }

    static _ApplyEditorFont(hFind, hImp) {
        if !hFind || !hImp || !this._hCustomFont
            return

        try {
            SendMessage(this.MSG.SETFONT, this._hCustomFont, 1, , "ahk_id " hFind)
            SendMessage(this.MSG.SETFONT, this._hCustomFont, 1, , "ahk_id " hImp)
        }
    }

    static _ClearShellTrack(hwnd) {
        if !this._shellTrack.Has(hwnd)
            return

        track := this._shellTrack[hwnd]
        try SetTimer(track.timer, 0)
        this._shellTrack.Delete(hwnd)
    }

    static _PruneShellTrack() {
        staleHwnds := []
        for trackedHwnd, track in this._shellTrack {
            if (!WinExist(trackedHwnd) || A_TickCount - track.time > this._shellTrackTTL)
                staleHwnds.Push(trackedHwnd)
        }

        for hwnd in staleHwnds
            this._ClearShellTrack(hwnd)

        if (this.IsDebug && staleHwnds.Length > 0)
            OutputDebug("[RisShell] Cleared stale shell track entries: " . staleHwnds.Length . "`n")
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
                pastImp := this.GetText(this.PastImpressionText)
                pastFind := this.GetText(this.PastFindingText)
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
        currentReqNo := this._GetCurrentReqNo()
        if (this._compContext.ReqNo != "" && this._compContext.ReqNo == currentReqNo) {
            SendText(this._compContext.Date)
        }
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
        this._ReplaceSelectionAndScroll(hEdit, cleanName . ":`r`n`r`n")
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
            fullText := this.FindingText
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
            currentImpText := this.ImpressionText

            ; [修改重點] 判斷結尾是否已為換行
            ; SubStr(str, -1) 會取得最後一個字元
            ; 如果是空字串 OR 最後一字是 `n (換行)，就不用補前綴
            if (currentImpText == "" || SubStr(currentImpText, -1) == "`n") {
                prefix := ""
            } else {
                prefix := "`r`n"
            }

            hImp := this.ImpressionEdit.NativeWindowHandle
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

    static SmartPageMove(direction, extend := false) {
        if !this.IsTargetFocused() {
            return
        }
        try {
            hEdit := ControlGetFocus("A")
            modifierKey := extend ? "+" : ""

            ; 取得目前 Scroll 位置 (最上方的行號)
            prevFirstLine := SendMessage(this.MSG.GETFIRSTVISIBLELINE, 0, 0, hEdit)

            if (direction == "Up") {
                SendInput(modifierKey "{PgUp}")
                Sleep 10 ; 等待 UI 更新
                currFirstLine := SendMessage(this.MSG.GETFIRSTVISIBLELINE, 0, 0, hEdit)

                ; 如果無法再往上捲 (前後行號一樣，且已在第 0 行)，則移到最前
                if (prevFirstLine == 0 && currFirstLine == 0) {
                    if (extend) {
                        SendInput("+^{Home}")
                    } else {
                        this._EditSetSel(hEdit, 0, 0)
                    }
                    this._EditScrollCaret(hEdit)
                }
            } else { ; Down
                SendInput(modifierKey "{PgDn}")
                Sleep 10 ; 等待 UI 更新
                currFirstLine := SendMessage(this.MSG.GETFIRSTVISIBLELINE, 0, 0, hEdit)

                ; 如果無法再往下捲 (前後行號一樣)，則移到最後
                if (prevFirstLine == currFirstLine) {
                    if (extend) {
                        SendInput("+^{End}")
                    } else {
                        fullText := ControlGetText(hEdit)
                        len := StrLen(fullText)
                        this._EditSetSel(hEdit, len, len)
                    }
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
            examType := this._GetCurrExamType()
            hEdit := this.FindingEdit.NativeWindowHandle

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
                this._ReorderSelectedText(hEdit, , , , , true)
            } else {
                this._ReorderSelectedText(hEdit, true, , , , true)
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
            forceStartFromOne := options.HasOwnProp("forceStartFromOne") ? options.forceStartFromOne : false

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

            this._ReorderSelectedText(hEdit, deOrder, keepEmpty, itemChar, discardSeIm, forceStartFromOne)
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

        currentHwnd := 0
        try currentHwnd := WinExist(this.WinTitle)
        if (!currentHwnd)
            return

        if !(this._stateCache.Has("_WindowInitialized") && this._stateCache["_WindowInitialized"] = currentHwnd) {
            this._StartWindowInitialization(currentHwnd)
            return
        }

        try {
            hFind := this.FindingEdit.NativeWindowHandle
            hImp  := this.ImpressionEdit.NativeWindowHandle
        } catch {
            return
        }

        this._ApplyEditorFont(hFind, hImp)
        this._ApplyLayout(hFind, hImp)
    }

    static _GetIndicationNodeKeys() {
        return [
            "ExamnameText",
            "GenderText", "AgeText", "OrderDeptText",
            "SubjectiveText", "ObjectiveText", "AssessmentText", "PlanText"
        ]
    }

    static _EnsureIndicationNodesReady() {
        for key in this._GetIndicationNodeKeys() {
            try {
                _ := this._GetOrUpdateNode(key)
            } catch {
                this._uiCache[key] := false
            }
        }
    }

    ; =================================================================
    ; 5. UI Preload Orchestration
    ; =================================================================
    ; [修改] 全面靜默快取主畫面元件 (改為非同步隊列模式)
    static _PreloadCache() {
        ; 1. 若已經執行過預載，就不再重複執行
        if (this._stateCache.Has("_UI_Preloaded")) {
            return
        }

        ; 2. 標記為已啟動預載，並建立待抓取隊列
        this._stateCache["_UI_Preloaded"] := true
        this._stateCache["_UI_Preloaded_Complete"] := false
        this._preloadQueue := []

        ; [優先順序優化] 嚴格依照使用者報告工作流順序
        ; 1. InsertExamNameAtCaret
        ; 2. AppendPreviousReport / InsertCopiedReportDate
        ; 3. GenerateAndInsertIndication (所需的病歷欄位)

        priorityKeys := [
            ; 1.
            "PastFindingText", "PastImpressionText", "PastReportTable",
            "FindingEdit", "ImpressionEdit",
        ]
        for key in this._GetIndicationNodeKeys() {
            priorityKeys.Push(key)
        }

        for key in priorityKeys {
            if (this.Selectors.Has(key) && !this._uiCache.Has(key)) {
                this._preloadQueue.Push(key)
            }
        }

        ; 處理剩餘的其他元件
        for key in this.Selectors {
            ; 略過不在主畫面的病理報告元件與臨床分頁
            if (key == "PathoDiagnosisText" || key == "PathoDateText" || key == "ClinicalTabControl")
                continue

            isPriority := false
            for pKey in priorityKeys {
                if (pKey == key) {
                    isPriority := true
                    break
                }
            }
            if (isPriority)
                continue

            ; 如果不在 cache 裡，就加入隊列
            if (!this._uiCache.Has(key)) {
                this._preloadQueue.Push(key)
            }
        }

        ; 3. 啟動非同步 Timer (每 50ms 處理一個元件)
        if (this._preloadQueue.Length > 0) {
            this._preloadTask := ObjBindMethod(this, "_PreloadStep")
            SetTimer(this._preloadTask, 50)
        }
    }

    static _StopPreloadTasks() {
        if (this._preloadTask) {
            SetTimer(this._preloadTask, 0)
            this._preloadTask := 0
        }
        if (this._indicationPreloadTask) {
            SetTimer(this._indicationPreloadTask, 0)
            this._indicationPreloadTask := 0
        }
    }

    static _FinalizeUIPreload() {
        this._StopPreloadTasks()
        this._stateCache["_UI_Preloaded_Complete"] := true

        try {
            hFind := this.FindingEdit.NativeWindowHandle
            hImp  := this.ImpressionEdit.NativeWindowHandle
            this._ApplyLayout(hFind, hImp)
        }

        this._MaybeStartIndicationPreload()
    }

    ; [新增] 非同步預載單步執行 (僅處理 UI 元件)
    static _PreloadStep() {
        if (this._preloadQueue.Length == 0) {
            this._StopPreloadTasks()
            return
        }

        key := this._preloadQueue.RemoveAt(1)

        try {
            ; [修改] 改用 _GetOrUpdateNode 確保即使沒有 Getter 也能抓到元件
            _ := this._GetOrUpdateNode(key)
        } catch {
            ; 負向快取
            this._uiCache[key] := false
        }

        ; 如果抽完最後一個，主動停止
        if (this._preloadQueue.Length == 0) {
            ; [新增] 標記預載完成，並立即觸發底色更新
            this._FinalizeUIPreload()
        }
    }

    ; [新增] UI 預載完成後，改由獨立排程啟動 AI indication 預載
    static _MaybeStartIndicationPreload() {
        if (!this._stateCache.Has("_UI_Preloaded_Complete") || !this._stateCache["_UI_Preloaded_Complete"]) {
            return
        }

        if (this._aiCache.Has("_AI_Indication") || this._isIndicationPending || this._indicationPreloadTask) {
            return
        }

        this._indicationPreloadTask := ObjBindMethod(this, "_StartIndicationPreload")
        SetTimer(this._indicationPreloadTask, -10)
    }

    static _StartIndicationPreload() {
        currentTask := this._indicationPreloadTask
        this._indicationPreloadTask := 0

        if (currentTask) {
            SetTimer(currentTask, 0)
        }

        if (this._aiCache.Has("_AI_Indication") || this._isIndicationPending) {
            return
        }

        try {
            this.GenerateAndInsertIndication(false, true)
        }
    }

    static _ApplyLayout(hFind, hImp) {
        ; 1. 安全檢查：如果 Handle 為 0 或空，直接離開
        if !hFind || !hImp
            return

        ; [新增] 底色反饋與狀態鎖定：預載完成前將 Impression 設為唯讀 (觸發灰色背景)，完成後解除。
        ; Finding 保持可寫，以免影響使用者操作體驗。
        isReady := this._stateCache.Has("_UI_Preloaded_Complete") && this._stateCache["_UI_Preloaded_Complete"]
        try {
            SendMessage(this.MSG.SETREADONLY, 0, 0, , "ahk_id " hFind) ; 確保 Finding 永遠可寫
            SendMessage(this.MSG.SETREADONLY, isReady ? 0 : 1, 0, , "ahk_id " hImp)
        }

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

    /**
     * 切換右側 Tab (AutomationId: "tbRightList")
     * @param index Tab 索引 (1-based, 1: 歷次報告, 2: 檢查清單, 3: 醫令/病理, 4: 會診)
     * Note: 具體順序可能依實際 RIS 介面而定
     */
    static SelectRightTab(index) {
        try {
            tabEle := this.RightTabControl
            hTab := tabEle.NativeWindowHandle

            ; 1. 最優先推薦：使用 AHK 內建的 ControlChooseIndex
            if (hTab) {
                try {
                    if (ControlGetIndex(hTab) == index)
                        return true
                }
                ControlChooseIndex(index, hTab)
                return true
            }

            ; 2. 備案：使用 UIA 的 Click()
            tabs := tabEle.FindAll({ Type: "TabItem" })
            if (index > 0 && index <= tabs.Length) {
                targetTab := tabs[index]
                try {
                    if (targetTab.IsSelected)
                        return true
                }
                targetTab.Click()
                return true
            }
        } catch {
            return false
        }
        return false
    }

    static SelectClinicalTab(index) {
        try {
            tabEle := this.ClinicalTabControl
            hTab := tabEle.NativeWindowHandle

            if (hTab) {
                try {
                    if (ControlGetIndex(hTab) == index)
                        return true
                }
                ControlChooseIndex(index, hTab)
                return true
            }

            tabs := tabEle.FindAll({ Type: "TabItem" })
            if (index > 0 && index <= tabs.Length) {
                targetTab := tabs[index]
                try {
                    if (targetTab.IsSelected)
                        return true
                }
                targetTab.Click()
                return true
            }
        } catch {
            return false
        }
        return false
    }

    static SelectPathologyOnlineData() {
        try {
            this.PathologyOnlineDataRadio.ControlClick()
            return true
        } catch as err {
            this.Notify("切換病理線上資料失敗: " err.Message)
            return false
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
                        targetCell := cellElements[3]
                        historyExamName := targetCell.Value
                        if (this._IsRelatedReport(historyExamName, currExamName)) {
                            targetCell.LegacyIAccessiblePattern.DoDefaultAction()
                            targetCell.ControlClick()
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
            hwnd := WinExist(this.AbnormalWinTitle)
            elWindow := UIA.ElementFromHandle(hwnd)
            elBtn := elWindow.FindElement(target)

            try {
                elBtn.Invoke()
            } catch {
                elBtn.Click()
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
        logFn := (msg) => (isAuto ? OutputDebug("[RisAuto] " . msg . "`n") : this.Notify(msg))

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
                logFn("⚠️ 無法點擊更新按鈕")
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
                logFn("⚠️ 統計資料為空，略過上傳")
                return
            }

            this._lastUpdateTick := A_TickCount
            this.PostDataToWebhook(jsonStr, isAuto)

        } catch as err {
            logFn("操作失敗: " . err.Message)
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

    static _ReplaceSelectionAndScroll(hCtrl, text) {
        this._EditReplaceSel(hCtrl, text)
        this._EditScrollCaret(hCtrl)
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

    ; =================================================================
    ; 9.1 UI Cache / Node Resolve
    ; =================================================================
    static _ResetWindowScopedCaches(currentHwnd) {
        this._uiCache := Map()
        this._stateCache := Map()
        this._aiCache := Map()
        this._stateCache["_Hwnd"] := currentHwnd
        this._preloadQueue := []
        this._StopPreloadTasks()
    }

    static _GetOrUpdateNode(nodeName) {
        ; 1. 取得目前實際視窗的 HWND (WinExist 速度極快，即使頻繁呼叫也無所謂)
        currentHwnd := WinExist(this.WinTitle)
        if !currentHwnd {
            throw TargetError("找不到 RIS 視窗")
        }

        ; =================================================================
        ; 2. [關鍵修改] 視窗身分驗證 (Window Identity Check)
        ; 如果 Cache 裡記錄的 HWND 與目前的 HWND 不同，代表視窗重開過。
        ; 此時必須清空所有「視窗生命週期」相關快取，避免拿到上一個視窗的殭屍物件。
        ; =================================================================
        if (!this._stateCache.Has("_Hwnd") || this._stateCache["_Hwnd"] != currentHwnd) {
            this._ResetWindowScopedCaches(currentHwnd)
        }

        ; 3. 經過上面的檢查，如果 nodeName 還在 cache 裡且有效，代表它屬於目前的視窗，可直接回傳
        ; [修改] 增加 IsObject 檢查，如果快取值是 false (預載失敗)，則強制重新抓取
        if this._uiCache.Has(nodeName) && IsObject(this._uiCache[nodeName]) {
            return this._uiCache[nodeName]
        }

        ; 4. 如果不在 cache 裡，則重新抓取 (Fetch Logic)
        if (nodeName = "Ris") {
            try {
                this._uiCache["Ris"] := UIA.ElementFromHandle(currentHwnd)
                return this._uiCache["Ris"]
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
                this._uiCache[nodeName] := parent.FindElement(this.Selectors[nodeName])
                return this._uiCache[nodeName]
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

    ; =================================================================
    ; 9.2 比較 Context 與 日期
    ; =================================================================

    static SetComparisonContext(targetDate) {
        currentReqNo := this._GetCurrentReqNo()
        formattedDate := this._ConvertRISDate(targetDate)
        this._compContext.ReqNo := currentReqNo
        this._compContext.Date := formattedDate
    }

    static GetComparisonSuffix() {
        currentReqNo := this._GetCurrentReqNo()
        if (this._compContext.ReqNo != "" && this._compContext.ReqNo == currentReqNo) {
            return " dated " . this._compContext.Date
        }
        return ""
    }

    static GetComparisonDate() {
        currentReqNo := this._GetCurrentReqNo()

        ; 只有當曾經執行過 Copy Report (AppendPreviousReport) 導致 Date 有值時才進行
        if (this._compContext.Date != "") {
            ; 寬鬆比對：只有在「兩者都有抓到 ReqNo」且「兩者不同」的情況下，才視為換報告並拒絕插入
            if (currentReqNo != "" && this._compContext.ReqNo != "" && currentReqNo != this._compContext.ReqNo) {
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

    static _GetCurrentReqNo() {
        try {
            return this.GetText(this.AccessionNoText)
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

    static _ReorderSelectedText(targetHwnd := 0, deOrder := false, keepEmptyLine := false, itemChar := "", discardSeIm := true, forceStartFromOne := false) {
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
        if (!forceStartFromOne && RegExMatch(selectedText, "^(\d+)", &existLineNo)) {
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
            this._ReorderSelectedText(hEdit, false, true, "-", false)
        } else {
            this.Notify("報告格式不如預期，無法自動排版")
        }
    }

    static _FormatFindingForAdvanced(hEdit) {
        fullText := ControlGetText(hEdit)
        range := this._FindContentRange(fullText, "Advanced")

        if (range) {
            this._EditSetSel(hEdit, range.Start, range.End)
            this._ReorderSelectedText(hEdit, false, false, "-", true)
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

        ; 2. 找結尾 (End): 往後找 `r 或 `n
        ; 先找第一個出現的換行符號 (可能是 \r 或 \n)
        rPos := InStr(fullText, "`r", , lineStart + 1)
        nPos := InStr(fullText, "`n", , lineStart + 1)

        if (rPos == 0 && nPos == 0) {
            nextLineBreak := 0
        } else if (rPos == 0) {
            nextLineBreak := nPos
        } else if (nPos == 0) {
            nextLineBreak := rPos
        } else {
            nextLineBreak := Min(rPos, nPos)
        }

        if (nextLineBreak == 0) {
            ; 最後一行，無換行符號
            contentEnd := StrLen(fullText)
            fullEnd := contentEnd
        } else {
            ; 找到換行符號，ContentEnd 在其之前 (index - 1)
            contentEnd := nextLineBreak - 1

            ; 判斷換行符號類型並計算 FullEnd
            char := SubStr(fullText, nextLineBreak, 1)
            if (char == "`r" && SubStr(fullText, nextLineBreak + 1, 1) == "`n") {
                fullEnd := nextLineBreak + 1 ; \r\n 之後
            } else {
                fullEnd := nextLineBreak ; \r 或 \n 之後
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

    ; =================================================================
    ; 優先依照使用者目前的 cursor scheme 與 CursorBaseSize 選擇 working 游標檔。
    ; 注意：SetSystemCursor 只會套用靜態游標內容，無法保留 .ani 動畫。
    ; =================================================================
    static _ShowWaitCursor() {
        try {
            OCR_NORMAL := 32512
            OCR_IBEAM := 32513
            cursorPath := this._ResolveWorkingCursorPath()

            if (cursorPath = "")
                return

            ; SetSystemCursor 會接手並銷毀 handle，因此需各自載入一份複本。
            hCopyNormal := DllCall("LoadCursorFromFile", "Str", cursorPath, "Ptr")
            hCopyIBeam  := DllCall("LoadCursorFromFile", "Str", cursorPath, "Ptr")

            if (!hCopyNormal || !hCopyIBeam)
                return

            DllCall("SetSystemCursor", "Ptr", hCopyNormal, "Int", OCR_NORMAL)
            DllCall("SetSystemCursor", "Ptr", hCopyIBeam, "Int", OCR_IBEAM)
        }
    }

    static _ResolveWorkingCursorPath() {
        try {
            appStarting := RegRead("HKCU\Control Panel\Cursors", "AppStarting", "")
        } catch {
            appStarting := ""
        }

        try {
            cursorBaseSizeRaw := RegRead("HKCU\Control Panel\Cursors", "CursorBaseSize", "32")
        } catch {
            cursorBaseSizeRaw := "32"
        }

        cache := this._workingCursorCache
        if (cache.AppStarting = appStarting
            && cache.CursorBaseSize = cursorBaseSizeRaw
            && cache.Path != ""
            && FileExist(cache.Path)) {
            return cache.Path
        }

        cursorBaseSize := Integer(cursorBaseSizeRaw)
        cursorPath := appStarting

        if (cursorPath = "" || !FileExist(cursorPath)) {
            cursorPath := A_WinDir . "\Cursors\aero_working.ani"
        }

        if (cursorBaseSize >= 64) {
            sizedPath := RegExReplace(cursorPath, "i)(\.[^\\.]*)$", "_xl$1")
            if (sizedPath != cursorPath && FileExist(sizedPath)) {
                this._workingCursorCache := {AppStarting: appStarting, CursorBaseSize: cursorBaseSizeRaw, Path: sizedPath}
                return sizedPath
            }
        }

        if (cursorBaseSize >= 48) {
            sizedPath := RegExReplace(cursorPath, "i)(\.[^\\.]*)$", "_l$1")
            if (sizedPath != cursorPath && FileExist(sizedPath)) {
                this._workingCursorCache := {AppStarting: appStarting, CursorBaseSize: cursorBaseSizeRaw, Path: sizedPath}
                return sizedPath
            }
        }

        resolvedPath := FileExist(cursorPath) ? cursorPath : ""
        this._workingCursorCache := {AppStarting: appStarting, CursorBaseSize: cursorBaseSizeRaw, Path: resolvedPath}
        return resolvedPath
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
        logFn := (msg) => (isSilent ? OutputDebug("[RisPost] " . msg . "`n") : this.Notify(msg))

        configFile := "config\private.ini"
        url  := IniRead(configFile, "n8n", "WebhookURL", "")
        user := IniRead(configFile, "n8n", "Username", "")
        pass := IniRead(configFile, "n8n", "Password", "")

        if (url == "") {
            logFn("❌ 錯誤：找不到 WebhookURL 設定")
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
                logFn("✅ 資料已上傳至 n8n")
            } else {
                logFn("❌ 上傳失敗 (Status: " . req.Status . ")")
            }
        } catch as err {
            logFn("❌ 網路錯誤: " . err.Message)
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
    ; controller 保留 orchestration；transport 已整理為單一 facade
    ; =================================================================

    ; 10.0 AI Orchestration Helpers
    static _BeginForegroundAIRequest() {
        if (this._isAIPending) {
            this.Notify("AI 正在背景產生中...")
            return false
        }

        this._ShowWaitCursor()
        this._isAIPending := true
        return true
    }

    static _FinishForegroundAIRequest() {
        this._isAIPending := false
        this._RestoreCursor()
    }

    static _NormalizeAIResult(result) {
        result := StrReplace(result, "`r`n", "`n")
        result := StrReplace(result, "`n", "`r`n")
        return result
    }

    ; 10.0.1 Indication
    static _TryInsertCachedIndication(isPreloadOnly) {
        if (!this._aiCache.Has("_AI_Indication") || isPreloadOnly) {
            return false
        }

        cached := this._aiCache["_AI_Indication"]
        this._InsertAIResult(cached.text)
        this.Notify(Format("已插入 Indication (來自快取, API:{}ms)", cached.apiTime))
        return true
    }

    static _TryHandlePendingIndication(isPreloadOnly) {
        if (!this._isIndicationPending) {
            return false
        }

        if (isPreloadOnly) {
            return true
        }

        this._pendingIndicationInsert := true
        this.Notify("AI 正在背景產生中，完成後將自動插入...")
        return true
    }

    static _BeginIndicationRequest(isPreloadOnly) {
        if (this._isAIPending) {
            if (!isPreloadOnly) {
                this.Notify("AI 正在背景產生中...")
            }
            return false
        }

        if (!isPreloadOnly) {
            this._ShowWaitCursor()
            this._pendingIndicationInsert := true
        }

        this._isAIPending := true
        this._isIndicationPending := true
        return true
    }

    static _NormalizeIndicationResult(result) {
        result := this._NormalizeAIResult(result)

        if (!InStr(result, "INDICATION:")) {
            result := "INDICATION: " . result
        }

        return result
    }

    static _CacheIndicationResult(result, apiTime, extractTime) {
        this._aiCache["_AI_Indication"] := {
            text: result,
            apiTime: apiTime,
            extractTime: extractTime
        }
    }

    static _HandleIndicationDebugPrompt(debugMode, isPreloadOnly, fullPrompt, extractTime) {
        if (!(debugMode && !isPreloadOnly)) {
            return true
        }

        A_Clipboard := fullPrompt
        ans := MsgBox("Debug 模式開啟。`n【Benchmark】資料提取耗時: " . extractTime . " ms`n`nPrompt 已複製到剪貼簿。是否繼續呼叫 API？`n`n" . SubStr(fullPrompt, 1, 500) . "...", "AI Debug", "YesNo")
        return ans != "No"
    }

    static _HandleIndicationSuccess(isPreloadOnly, result, apiTime, extractTime, debugMode := false) {
        normalized := this._NormalizeIndicationResult(result)
        this._CacheIndicationResult(normalized, apiTime, extractTime)

        if (debugMode && !isPreloadOnly) {
            MsgBox("【Benchmark】`n資料提取: " . extractTime . " ms`nAPI 耗時: " . apiTime . " ms`n`n【API 回傳結果】`n" . normalized, "AI Debug")
        }

        if (!isPreloadOnly) {
            this.Notify(Format("已產生 Indication (取資:{}ms, API:{}ms)", extractTime, apiTime))
        } else {
            OutputDebug("[RisController] AI Indication 已預載並快取`n")
        }
    }

    static _FinishIndicationRequest(requestMode, isPreloadOnly) {
        this._isAIPending := false
        this._isIndicationPending := false

        if (!isPreloadOnly) {
            this._RestoreCursor()
        }

        shouldInsert := this._pendingIndicationInsert && this._aiCache.Has("_AI_Indication")
        if (shouldInsert) {
            this._InsertAIResult(this._aiCache["_AI_Indication"].text)
            if (requestMode == "preload") {
                this.Notify("Indication 已完成並插入")
            }
        }

        this._pendingIndicationInsert := false
    }

    static _HandleImpressionDebugPrompt(debugMode, fullPrompt) {
        if (!debugMode) {
            return true
        }

        A_Clipboard := fullPrompt
        ans := MsgBox("Prompt 已複製。是否繼續？`n`n" . SubStr(fullPrompt, 1, 500) . "...", "AI Debug", "YesNo")
        return ans != "No"
    }

    static _HandleImpressionSuccess(result, extractTime, apiTime) {
        result := this._NormalizeAIResult(result)
        this._InsertAIResultToImpression(result)
        this.Notify(Format("已插入 Impression (取資:{}ms, API:{}ms)", extractTime, apiTime))
    }

    static _CreateAIRequest(promptText, aiConfig, extraFields := 0) {
        request := {
            Prompt: promptText,
            Config: aiConfig
        }

        if IsObject(extraFields) {
            for key, value in extraFields.OwnProps() {
                request.%key% := value
            }
        }

        return request
    }

    static _BuildAIRequestResult(promptText, aiConfig) {
        t0 := A_TickCount
        result := this._CallGoogleAI(promptText, aiConfig.Model, aiConfig.Temperature, aiConfig.TopP)

        return {
            Result: result,
            ApiTime: A_TickCount - t0
        }
    }

    static _BuildIndicationRequest() {
        t0 := A_TickCount
        this._EnsureIndicationNodesReady()
        clinicalData := this._GetAndFormatClinicalData()
        if (clinicalData == "") {
            return false
        }

        extractTime := A_TickCount - t0
        conf := RisConfig.AI.Indication
        fullPrompt := conf.SystemPrompt . clinicalData . conf.Constraint

        return this._CreateAIRequest(fullPrompt, conf, {
            ClinicalData: clinicalData,
            ExtractTime: extractTime
        })
    }

    static _BuildImpressionRequest() {
        this._PreloadCache()

        t0 := A_TickCount
        try {
            hFind := this.FindingEdit.NativeWindowHandle
            hImp  := this.ImpressionEdit.NativeWindowHandle
        } catch {
            return false
        }

        findingText := ControlGetText(hFind)
        if (findingText == "") {
            return {Error: "Findings 欄位為空，無法產生總結"}
        }

        findingText := this._DeidentifyText(findingText)
        extractTime := A_TickCount - t0
        conf := RisConfig.AI.Impression

        return this._CreateAIRequest(Format(conf.Prompt, findingText), conf, {
            ExtractTime: extractTime,
            ImpressionHwnd: hImp
        })
    }

    static _BuildRefineRequest(selectedText) {
        conf := RisConfig.AI.Refine
        prompt := conf.SystemPrompt . "`n`nInput Text:`n" . selectedText

        return this._CreateAIRequest(prompt, conf)
    }

    static _RunAIRequest(request) {
        return this._BuildAIRequestResult(request.Prompt, request.Config)
    }

    static _RunIndicationRequest(request, debugMode, isPreloadOnly) {
        response := this._RunAIRequest(request)
        this._HandleIndicationSuccess(isPreloadOnly, response.Result, response.ApiTime, request.ExtractTime, debugMode)
    }

    static _RunImpressionRequest(request) {
        response := this._RunAIRequest(request)
        this._HandleImpressionSuccess(response.Result, request.ExtractTime, response.ApiTime)
    }

    static _RunRefineRequest(request) {
        response := this._RunAIRequest(request)
        return response.Result
    }

    ; [新增] 外部呼叫的主函式：產生並插入 Indication
    ; [修改] 增加 Benchmark 效能測量
    static GenerateAndInsertIndication(debugMode := false, isPreloadOnly := false) {
        requestMode := isPreloadOnly ? "preload" : "manual"

        if (this._TryInsertCachedIndication(isPreloadOnly)) {
            return
        }

        if (this._TryHandlePendingIndication(isPreloadOnly)) {
            return
        }

        if (!this._BeginIndicationRequest(isPreloadOnly)) {
            return
        }

        try {
            request := this._BuildIndicationRequest()
            if (!request) {
                if (!isPreloadOnly)
                    this.Notify("無法取得病歷資料，請確認是否在正確視窗內")
                return
            }

            if (!this._HandleIndicationDebugPrompt(debugMode, isPreloadOnly, request.Prompt, request.ExtractTime)) {
                return
            }

            this._RunIndicationRequest(request, debugMode, isPreloadOnly)

        } catch as err {
            if (!isPreloadOnly) {
                fullErrorMsg := "【錯誤訊息】`n" . err.Message . "`n`n【發生位置】`n" . err.What . "`n`n【呼叫堆疊】`n" . err.Stack
                this._ShowDebugError(fullErrorMsg)
            }
        } finally {
            this._FinishIndicationRequest(requestMode, isPreloadOnly)
        }
    }

    ; [新增] 產生並插入 Impression (總結 Findings)
    static GenerateAndInsertImpression(debugMode := false) {
        if (!this._BeginForegroundAIRequest()) {
            return
        }

        try {
            request := this._BuildImpressionRequest()
            if (!request) {
                this.Notify("找不到編輯欄位，請確認視窗是否正確")
                return
            }
            if (request.HasOwnProp("Error")) {
                this.Notify(request.Error)
                return
            }

            if (!this._HandleImpressionDebugPrompt(debugMode, request.Prompt)) {
                return
            }

            this._RunImpressionRequest(request)

        } catch as err {
            fullErrorMsg := "【錯誤訊息】`n" . err.Message . "`n`n【發生位置】`n" . err.What . "`n`n【呼叫堆疊】`n" . err.Stack
            this._ShowDebugError(fullErrorMsg)
        } finally {
            this._FinishForegroundAIRequest()
        }
    }

    ; [新增] 文字潤色與翻譯 (Polishing)
    ; 使用 LLM 優化所選取的文字，並提供對照視窗供使用者確認是否採用
    static PolishSelectionWithAI() {
        if !this.IsTargetFocused() {
            this.Notify("請先點擊要處理的文字欄位")
            return
        }

        try {
            hEdit := ControlGetFocus("A")
            sel := this._EditGetSel(hEdit)

            ; 檢查是否有選取文字
            if (sel.Start == sel.End) {
                this.Notify("請先選取要潤色的文字")
                return
            }

            fullText := ControlGetText(hEdit)
            selectedText := SubStr(fullText, sel.Start + 1, sel.End - sel.Start)

            if (Trim(selectedText) == "") {
                this.Notify("選取的文字為空")
                return
            }

            this._ShowWaitCursor()
            this.Notify("AI 潤色中...", 3000)

            request := this._BuildRefineRequest(selectedText)
            result := this._RunRefineRequest(request)

            ; 格式化換行
            result := StrReplace(result, "`r`n", "`n")
            result := StrReplace(result, "`n", "`r`n")

            this._RestoreCursor()

            ; 顯示比對視窗
            this._ShowPolishComparisonGui(hEdit, selectedText, result, sel)

        } catch as err {
            this._RestoreCursor()
            this.Notify("AI 潤色失敗: " . err.Message)
        }
    }

    static _ShowPolishComparisonGui(hEdit, original, refined, sel) {
        ; 建立比對 GUI
        myGui := Gui("+AlwaysOnTop", "AI 潤色結果比對")
        myGui.SetFont("s11", "Microsoft JhengHei UI")

        ; --- 第一列：標題對齊 ---
        myGui.Add("Text", "w400", "原始文字 (Original):")
        myGui.Add("Text", "x+20 yp w400", "潤色結果 (Refined):")

        ; --- 第二列：內容對齊 ---
        ; 兩者皆設為 ReadOnly，並加入 -WantReturn 讓 Enter 鍵能觸發 Default 按鈕
        myGui.Add("Edit", "xm w400 r15 ReadOnly Multi -WantReturn", original)
        refinedEdit := myGui.Add("Edit", "x+20 yp w400 r15 ReadOnly Multi -WantReturn", refined)

        ; --- 第三列：按鈕區 ---
        btnAccept := myGui.Add("Button", "Default w180 x220 y+20", "✅ Accept (Enter)")
        btnReject := myGui.Add("Button", "w180 x+20", "❌ Reject (Esc)")

        ; 事件處理函式
        handleAccept(*) {
            finalText := refinedEdit.Value
            finalText := StrReplace(finalText, "`r`n", "`n")
            finalText := StrReplace(finalText, "`n", "`r`n")

            this._EditSetSel(hEdit, sel.Start, sel.End)
            this._ReplaceSelectionAndScroll(hEdit, finalText)
            myGui.Destroy()
            this.Notify("已更新文字")
        }

        btnAccept.OnEvent("Click", handleAccept)
        btnReject.OnEvent("Click", (*) => myGui.Destroy())

        ; 快捷鍵：Esc 關閉
        myGui.OnEvent("Escape", (*) => myGui.Destroy())

        myGui.Show("Center")

        ; 自動聚焦到結果框（方便按 Enter），並將游標移至開頭（防止自動全選）
        refinedEdit.Focus()
        SendMessage(0x00B1, 0, 0, refinedEdit.Hwnd) ; EM_SETSEL: Start=0, End=0
    }

    ; [內部 Helper] 專用於插入 Impression 欄位
    static _InsertAIResultToImpression(result) {
        if !WinActive(this.WinTitle) {
            WinActivate(this.WinTitle)
            WinWaitActive(this.WinTitle, , 2)
        }

        ; 強制聚焦到 ImpressionEdit
        this.ImpressionEdit.SetFocus()
        Sleep(50)

        targetHwnd := this.ImpressionEdit.NativeWindowHandle

        ; 插入文字 (清空原本內容還是附加？通常總結是全新的，這裡採附加但在最前面)
        ; 使用者可能希望直接覆蓋，或者在游標處插入。這裡遵循 _InsertAIResult 邏輯：在游標處插入。
        this._ReplaceSelectionAndScroll(targetHwnd, result)
    }

    ; [內部 Helper] 執行 AI 結果插入 UI
    static _InsertAIResult(result) {
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

        this._ReplaceSelectionAndScroll(targetHwnd, result . "`r`n`r`n")
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
        ; 逐一嘗試抓取，即使某個欄位找不到也不會中斷整個字串的組合
        gender := this._FastGetCtrlText("GenderText")
        ageText := this._FastGetCtrlText("AgeText")
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

        rawText := Format("{1}`n{2}`n{3}`n{4}`n`nS: {5}`n`nO: {6}`n`nA: {7}`n`nP: {8}", gender, ageText, exam, dept, sText, oText, aText, pText)

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
            matchedAge := Number(m[1])
            decade := Floor(matchedAge / 10) * 10
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

    ; =================================================================
    ; 10.1 AI Transport
    ; request prepare / transport wait / response parse
    ; 目前僅保留 Google AI 單一供應商 facade，不在此層擴充 provider switch
    ; =================================================================
    static _GetGoogleAIConfig() {
        if (this._googleAIConfig) {
            return this._googleAIConfig
        }

        configFile := "config\private.ini"
        apiKey := IniRead(configFile, "GoogleAI", "APIKey", "")
        if (apiKey == "") {
            throw Error("請在 " . configFile . " 中設定 [GoogleAI] APIKey")
        }

        this._googleAIConfig := {
            ConfigFile: configFile,
            APIKey: apiKey,
            Model: IniRead(configFile, "GoogleAI", "Model", "gemma-3-27b-it"),
            Temperature: IniRead(configFile, "GoogleAI", "Temperature", "0.2"),
            TopP: IniRead(configFile, "GoogleAI", "TopP", "0.95")
        }

        return this._googleAIConfig
    }

    static _ResolveGoogleAIOptions(modelName, temperature, topP) {
        cfg := this._GetGoogleAIConfig()

        return {
            APIKey: cfg.APIKey,
            Model: (modelName != "") ? modelName : cfg.Model,
            Temperature: (temperature != "") ? temperature : cfg.Temperature,
            TopP: (topP != "") ? topP : cfg.TopP
        }
    }

    static _BuildGoogleAIUrl(options) {
        return "https://generativelanguage.googleapis.com/v1beta/models/" . options.Model . ":generateContent?key=" . options.APIKey
    }

    static _BuildGoogleAIPayload(promptText, options) {
        escapedPrompt := this._EscapeJsonString(promptText)

        return '{'
            . '"contents": [{'
                . '"role": "user",'
                . '"parts": [{"text": "' . escapedPrompt . '"}]'
            . '}],'
            . '"generationConfig": {'
                . '"temperature": ' . options.Temperature . ','
                . '"thinkingConfig": {"thinkingLevel": "MINIMAL"},'
                . '"topP": ' . options.TopP
            . '},'
            . '"tools": [{"googleSearch": {}}]'
        . '}'
    }

    static _EscapeJsonString(text) {
        escaped := StrReplace(text, "\", "\\")
        escaped := StrReplace(escaped, "`"", "\`"")
        escaped := StrReplace(escaped, "`n", "\n")
        escaped := StrReplace(escaped, "`r", "\r")
        escaped := StrReplace(escaped, "`t", "\t")
        return escaped
    }

    static _BuildGoogleAIRequest(promptText, modelName := "", temperature := "", topP := "") {
        configStart := A_TickCount
        options := this._ResolveGoogleAIOptions(modelName, temperature, topP)
        configTime := A_TickCount - configStart

        payloadStart := A_TickCount

        return {
            Url: this._BuildGoogleAIUrl(options),
            Payload: this._BuildGoogleAIPayload(promptText, options),
            Metrics: {
                ConfigReadTime: configTime,
                PayloadBuildTime: A_TickCount - payloadStart
            }
        }
    }

    static _WaitForGoogleAIResponse(req) {
        ; 保留既有短輪詢行為，先集中在 transport helper 內，方便後續替換。
        while !req.WaitForResponse(0.01) {
            Sleep(10)
        }
    }

    static _SendGoogleAIRequest(url, payload) {
        req := ComObject("WinHttp.WinHttpRequest.5.1")

        ; 目前先保留既有非同步 request + wait 介面，後續可獨立抽成 transport/service。
        req.Open("POST", url, True)
        req.SetRequestHeader("Content-Type", "application/json")
        req.Send(payload)
        this._WaitForGoogleAIResponse(req)

        return {
            Status: req.Status,
            ResponseText: req.ResponseText
        }
    }

    static _ExtractGoogleAIResponseText(responseText) {
        combinedText := ""
        searchPos := 1

        ; 遍歷所有 "text": "..." 區段
        while (searchPos := RegExMatch(responseText, 's)"text":\s*"(.*?)(?<!\\)"', &match, searchPos)) {
            val := match[1]

            ; 如果這段文字不是 thought 區塊，才加入最終結果
            context := SubStr(responseText, searchPos + match.Len, 100)
            if !RegExMatch(context, '^\s*,\s*"thought":\s*true') {
                combinedText .= val
            }

            searchPos += match.Len
        }

        if (combinedText == "") {
            throw Error("無法從 API 回應中提取有效文字。")
        }

        return combinedText
    }

    static _DecodeGoogleAIResponseText(text) {
        val := text

        ; 1. 還原 JSON 內的跳脫字元
        val := StrReplace(val, "\n", "`n")
        val := StrReplace(val, "\r", "`r")
        val := StrReplace(val, "\t", "`t")
        val := StrReplace(val, '\"', '"')
        val := StrReplace(val, "\\", "\")

        ; 2. 解碼 \uXXXX (Unicode)
        while RegExMatch(val, "i)\\u([0-9a-f]{4})", &m) {
            val := StrReplace(val, m[0], Chr(Integer("0x" . m[1])))
        }

        return val
    }

    static _StripMarkdownCodeFence(text) {
        text := Trim(text, " `t`r`n")

        ; 3. 清理 Markdown 標記
        if (RegExMatch(text, "s)^``````(?:\w+)?\R?(.*?)\R?``````$", &m)) {
            text := m[1]
        } else if (SubStr(text, 1, 1) == "``" && SubStr(text, -1) == "``") {
            text := SubStr(text, 2, StrLen(text) - 2)
        }

        return Trim(text, " `t`r`n")
    }

    static _ParseGoogleAIResponse(responseText) {
        text := this._ExtractGoogleAIResponseText(responseText)
        text := this._DecodeGoogleAIResponseText(text)
        return this._StripMarkdownCodeFence(text)
    }

    static _DebugGoogleAIResponse(url, payload, response) {
        if (!this.IsDebug) {
            return
        }

        A_Clipboard := "URL: " . url . "`n`nPayload: " . payload . "`n`nResponse: " . response.ResponseText
        MsgBox("【API Debug】原始回應已複製到剪貼簿：`n`nStatus: " . response.Status . "`n`n" . SubStr(response.ResponseText, 1, 1000), "Google AI Debug")
    }

    static _LogGoogleAIBlockingMetrics(metrics, status := "") {
        statusText := (status != "") ? ", status=" . status : ""
        OutputDebug(Format(
            "[RisController] GoogleAI blocking metrics: config={}ms, payload={}ms, wait={}ms, parse={}ms{}`n",
            metrics.ConfigReadTime,
            metrics.PayloadBuildTime,
            metrics.WaitForResponseTime,
            metrics.ResponseParseTime,
            statusText
        ))
    }

    static _CallGoogleAI(promptText, modelName := "", temperature := "", topP := "") {
        request := this._BuildGoogleAIRequest(promptText, modelName, temperature, topP)
        waitStart := A_TickCount
        response := this._SendGoogleAIRequest(request.Url, request.Payload)
        request.Metrics.WaitForResponseTime := A_TickCount - waitStart
        this._DebugGoogleAIResponse(request.Url, request.Payload, response)

        if (response.Status != 200) {
            request.Metrics.ResponseParseTime := 0
            this._LogGoogleAIBlockingMetrics(request.Metrics, response.Status)
            throw Error("HTTP " . response.Status . " - " . response.ResponseText)
        }

        parseStart := A_TickCount
        parsed := this._ParseGoogleAIResponse(response.ResponseText)
        request.Metrics.ResponseParseTime := A_TickCount - parseStart
        this._LogGoogleAIBlockingMetrics(request.Metrics, response.Status)
        return parsed
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
