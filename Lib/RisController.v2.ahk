#Requires AutoHotkey v2.0
#Include .\UIA.v2.ahk
#Include .\RisConfig.v2.ahk
#Include .\RisAIConfigResolver.v2.ahk
#Include .\RisAIProviderPolicy.v2.ahk
#Include .\RisAIText.v2.ahk
#Include .\RisAIPayload.v2.ahk
#Include .\RisAIRequestBuilder.v2.ahk
#Include .\RisAITransport.v2.ahk
#Include .\RisAIDebug.v2.ahk
#Include .\RisAIDebugGui.v2.ahk
#Include .\RisAIService.v2.ahk
#Include .\RisAIModelHealth.v2.ahk
#Include .\RisAIOrchestration.v2.ahk
#Include .\RisEditControl.v2.ahk
#Include .\RisNotify.v2.ahk
#Include .\RisVisualFeedback.v2.ahk
#Include .\RisDate.v2.ahk
#Include .\RisReportText.v2.ahk

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
    static LdctReportWinTitle := "肺癌篩檢底劑量電腦斷層掃瞄結果報告  (frmLDCTReport)"
    static AbnormalWinTitle := "檢查結果(frmPos)"
    static ConsultationWinTitle := "會診資訊(frmReqCon)"

    ; [新增] 環境判定：是否為標準 RIS 報告視窗
    static IsStandardRis => WinActive(this.WinTitle)

    static _layoutShiftProp := "RisLayoutShiftApplied"
    static _bottomCtrlShiftProp := "RisBottomCtrlShiftApplied"

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
        "CurrentCountLabel", { AutomationId: "lblCurrentCount" },
        "BottomInfoLabel",  { AutomationId: "label5" },
        "ImpressionLabel",  { AutomationId: "label2" },
        "MedRecNoLabel",    { AutomationId: "txtMRNo" },
        "AccessionNoText",  { AutomationId: "txtReqNo" },
        "PhExamColumn",     { AutomationId: "goxExamine" },
        "PhExamDateText",   { AutomationId: "mtxtReportDTM" },
        "PhExamReportText", { AutomationId: "txtReport" },

        ; [新增] SOAP 與基本資料欄位
        "SubjectiveText",   { AutomationId: "rtxtSubjective" },
        "ObjectiveText",    { AutomationId: "rtxtObjective" },
        "AssessmentText",   { AutomationId: "rtxtICD10" },
        "PlanText",         { AutomationId: "rtxtAdmInICD" },
        "OrderDeptText",    { AutomationId: "txtAppSecName" },
        "BedInfoText",      { AutomationId: "txtBid" },
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

    ; =================================================================
    ; 1.1 初始化與配置載入 (Initialization)
    ; =================================================================

    ; 類別載入時自動執行
    static __New() {
        this.LoadSimGroups()
        GroupAdd "RisReportGroup", "報告作業(frmRISReport)"
        GroupAdd "RisReportGroup", "肺癌篩檢底劑量電腦斷層掃瞄結果報告  (frmLDCTReport)"

        ; 設定 RisNotify 偵測的目標視窗
        RisNotify.TargetTitles := [
            this.WinTitle,
            this.LdctReportWinTitle,
            this.AbnormalWinTitle
        ]
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
    static _compContext := {ReqNo: "", Date: ""}
    static _hCustomFont := 0
    static _targetImpressionHeight := 95
    static _preloadQueue := [] ; [新增] 預載隊列
    static _preloadTask  := 0  ; [新增] 預載 Timer 參考
    static _indicationPreloadTask := 0 ; [新增] AI indication 預載排程
    static _isAIPending  := false ; [新增] AI 請求中旗標
    static _isIndicationPending := false ; [新增] indication 產生中旗標
    static _pendingIndicationInsert := false ; [新增] indication 完成後插入請求
    static _pendingIndicationInsertAfter := ""
    static _isShellHookEnabled := false ; [新增] ShellHook 狀態旗標
    static _shellTrack := Map()         ; [新增] 用於紀錄每個 HWND 的處理狀態 {time: TickCount, timer: Func}
    static _shellTrackTTL := 15000
    static IsDebug := false            ; [新增] Debug 模式切換
    static ShowGoogleAICurlDebug := false

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
    ; - Presentation state: _hCustomFont / _targetImpressionHeight / _compContext
    ;
    ; 2.3 後續可拆模組草案
    ; - RisUiCache: UI cache / node resolve / preload queue
    ; - RisEditorActions: editor actions, formatting, selection mutations
    ; - RisAiService: AI orchestration + transport facade
    ; - RisNotify: Notify GUI, queue, slot management (done)
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
    static BedInfoText    => this._GetOrUpdateNode("BedInfoText")
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
            range := RisReportText.FindContentRange(fullText, "Advanced")

            if (!range) {
                return ""
            }

            length := (range.End == -1) ? StrLen(fullText) - range.Start : range.End - range.Start
            return SubStr(fullText, range.Start + 1, length)
        } catch {
            return ""
        }
    }

    static HasFindingContentRange(mode := "Advanced") {
        try {
            return RisReportText.FindContentRange(this.FindingText, mode) ? true : false
        } catch {
            return false
        }
    }

    static GetFindingSearchText(mode := "Advanced") {
        if (!this.IsStandardRis) {
            return ""
        }
        try {
            searchText := this.GetFindingContent()
            if (searchText != "" || this.HasFindingContentRange(mode)) {
                return searchText
            }

            return this.FindingText
        } catch {
            return ""
        }
    }

    ; =================================================================
    ; 4. 系統功能 (Notify & Focus)
    ; =================================================================

    static Notify(text, duration := 1500) {
        RisNotify.Show(text, duration)
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

        this._ApplyCustomFontToControls(hFind, hImp)
    }

    static _ApplyCustomFontToControls(hwnds*) {
        if !this._hCustomFont
            return

        try {
            for hwnd in hwnds {
                if hwnd
                    SendMessage(this.MSG.SETFONT, this._hCustomFont, 1, , "ahk_id " hwnd)
            }
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
        targetHwnd := 0

        ; [優先順序] 1. 檢查是否正在使用 LDCT 視窗，或是 LDCT 視窗存在
        if WinActive(this.LdctReportWinTitle) || WinExist(this.LdctReportWinTitle) {
            targetHwnd := WinExist(this.LdctReportWinTitle)

            if !WinActive(targetHwnd) {
                WinActivate(targetHwnd)
                WinWaitActive(targetHwnd, , 1)
            } else {
                ; 如果已經在 LDCT 視窗，嘗試在控制項間切換焦點 (簡單模擬)
                Send "{Tab}"
            }

            try targetHwnd := ControlGetFocus("A")
            SetTimer( () => RisController._ScrollAndHighlightCaret(targetHwnd), -10 )
            return
        }

        ; [優先順序] 2. 原本的標準 RIS 邏輯
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

            SetTimer( () => RisController._ScrollAndHighlightCaret(targetHwnd), -10 )
        } catch as err {
            this.Notify("視窗切換失敗: " err.Message)
        }
    }

    static _ScrollAndHighlightCaret(targetHwnd) {
        if (targetHwnd) {
            try RisEditControl.ScrollCaret(targetHwnd)
            Sleep 30
        }

        RisVisualFeedback.HighlightCaret(targetHwnd)
    }

    ; =================================================================
    ; 5. 報告操作 (Paste, Append, Insert)
    ; =================================================================
    static PasteToFinding(text) {
        if (!this.IsStandardRis) {
            SendText(text)
            return
        }
        this.PasteTo(this.FindingEdit, text)
    }

    static PasteToImpression(text) {
        if (!this.IsStandardRis) {
            SendText(text)
            return
        }
        this.PasteTo(this.ImpressionEdit, text)
    }

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
        RisVisualFeedback.ShowWaitCursor()
        try {
            try {
                pastImp := this.GetText(this.PastImpressionText)
                pastFind := this.GetText(this.PastFindingText)
                pastFind := RisReportText.ExtractPastFinding(pastFind)
                hImpEdit := this.ImpressionEdit.NativeWindowHandle
                hFindEdit := this.FindingEdit.NativeWindowHandle
            } catch {
                return
            }

            ; [紀錄] 紀錄 Finding 目前的 Caret 位置
            initialFindSel := RisEditControl.GetSel(hFindEdit)

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

                RisEditControl.SetSel(hEdit, currentLen, currentLen)
                if (currentLen > 0 && SubStr(currentText, -1) != "`n") {
                    textToAppend := "`r`n" . textToAppend
                }

                RisEditControl.ReplaceSel(hEdit, textToAppend)
                RisEditControl.SetSel(hEdit, -1, -1)
                RisEditControl.ScrollCaret(hEdit)
            }

            AppendToEdit(hImpEdit, pastImp)
            AppendToEdit(hFindEdit, pastFind)

            try {
                this.FindingEdit.SetFocus()
                ; [修改] 不再強制移至 0，而是還原至 initialFindSel.Start
                RisEditControl.SetSel(hFindEdit, initialFindSel.Start, initialFindSel.Start)
                RisEditControl.ScrollCaret(hFindEdit)
            }
        } finally {
            RisVisualFeedback.RestoreCursor()
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
                foundValue := RisDate.ConvertRISDate(foundValue)
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
        RisEditControl.ReplaceSelectionAndScroll(hEdit, cleanName . ":`r`n`r`n")
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
        initialSel := RisEditControl.GetSel(hFocus)
        textToCopy := ""
        wasSelectionEmpty := (initialSel.Start == initialSel.End)

        if (wasSelectionEmpty) {
            ; === 狀況 A: 原本無反白 -> 自動抓整行並還原 ===
            RisEditControl.SelectLine(hFocus)

            newSel := RisEditControl.GetSel(hFocus)
            fullText := this.FindingText
            if (newSel.End > newSel.Start) {
                textToCopy := SubStr(fullText, newSel.Start + 1, newSel.End - newSel.Start)
            }

            ; 還原 Caret 到原本位置
            RisEditControl.SetSel(hFocus, initialSel.Start, initialSel.Start)
            RisEditControl.ScrollCaret(hFocus)
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
            RisEditControl.SetSel(hImp, impLen, impLen)
            RisEditControl.ReplaceSel(hImp, prefix . textToCopy)

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
            RisEditControl.KillLine(hEdit)
        }
    }

    static CutLineOrSelection() {
        if !this.IsTargetFocused() {
            return false
        }
        try {
            hFocus := ControlGetFocus("A")
            RisEditControl.CutLineOrSelection(hFocus)
        }
        return true
    }

    static CopyLineOrSelection() {
        if !this.IsTargetFocused() {
            return false
        }
        try {
            hFocus := ControlGetFocus("A")
            RisEditControl.CopyLineOrSelection(hFocus)
        }
        return true
    }

    static DeleteCurrentLine() {
        if !this.IsTargetFocused() {
            return false
        }
        try {
            hFocus := ControlGetFocus("A")
            RisEditControl.DeleteCurrentLine(hFocus)
        }
        return true
    }

    static DeleteWordBackward() {
        ; 如果是標準 RIS，維持原本精確的欄位檢查
        if (this.IsStandardRis) {
            if !this.IsTargetFocused() {
                return false
            }
        }
        
        try {
            hCtrl := ControlGetFocus("A")
            if !hCtrl {
                return false
            }
            RisEditControl.DeleteWordBackward(hCtrl)
            return true
        } catch {
            return false
        }
    }

    static MoveCaret(mode) {
        if !this.IsTargetFocused() {
            return false
        }
        try {
            hCtrl := ControlGetFocus("A")
            RisEditControl.MoveCaret(hCtrl, mode)
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
            moved := RisEditControl.SmartPageMove(hEdit, direction, extend)
            if (moved) {
                SetTimer( () => RisVisualFeedback.HighlightCaret(hEdit), -10 )
            }
            return moved
        } catch {
            return false
        }
    }

    ; [新增] 移動目前所在行 (Alt+Up / Alt+Down 功能實作)
    static MoveCurrentLine(direction) {
        if !this.IsTargetFocused() {
            return
        }

        try {
            hCtrl := ControlGetFocus("A")
            RisEditControl.MoveCurrentLine(hCtrl, direction)
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
            RisEditControl.InsertNewLine(hEdit, mode)
        }
    }

    static SmartListEnter() {
        if !this.IsTargetFocused() {
            return false
        }

        try {
            hEdit := ControlGetFocus("A")
            return RisEditControl.SmartListEnter(hEdit)
        } catch {
            return false
        }
    }

    static ShouldSmartListEnter() {
        if !this.IsTargetFocused() {
            return false
        }

        try {
            hEdit := ControlGetFocus("A")
            return RisEditControl.ShouldSmartListEnter(hEdit)
        } catch {
            return false
        }
    }

    static SmartListBackspace() {
        if !this.IsTargetFocused() {
            return false
        }

        try {
            hEdit := ControlGetFocus("A")
            return RisEditControl.SmartListBackspace(hEdit)
        } catch {
            return false
        }
    }

    static ShouldSmartListBackspace() {
        if !this.IsTargetFocused() {
            return false
        }

        try {
            hEdit := ControlGetFocus("A")
            return RisEditControl.ShouldSmartListBackspace(hEdit)
        } catch {
            return false
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
            ControlFocus(hEdit)

            switch examType {
                case "CT", "MR": this._FormatFindingForAdvanced(hEdit)
                case "CR", "US": this._FormatFindingForBasic(hEdit)
            }
            SetTimer( () => RisController._ScrollAndHighlightCaret(hEdit), -10 )
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

            RisEditControl.SetSel(hEdit, 0, -1)
            lineCount := RisEditControl.CountNonEmptyLines(hEdit)

            if (lineCount > 1) {
                this._ReorderSelectedText(hEdit, , , , , true)
            } else {
                this._ReorderSelectedText(hEdit, true, , , , true)
            }
            SetTimer( () => RisController._ScrollAndHighlightCaret(hEdit), -10 )
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
                selectedText := RisEditControl.GetSelectedText(hEdit)
                if (selectedText != "") {
                    itemChar := RisReportText.DetectItemChar(selectedText)
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
            if (this.IsDebug) {
                OutputDebug(Format("[RisLayout] _FinalizeUIPreload hFind={} hImp={}`n", hFind, hImp))
            }
            this._ApplyLayout(hFind, hImp)
        } catch as err {
            if (this.IsDebug) {
                OutputDebug(Format("[RisLayout] _FinalizeUIPreload failed: {}`n", err.Message))
            }
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

        if (!this._AreReportEditorsBlankForIndicationPreload()) {
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

        if (!this._AreReportEditorsBlankForIndicationPreload()) {
            return
        }

        try {
            this.GenerateAndInsertIndication(false, true)
        }
    }

    static _AreReportEditorsBlankForIndicationPreload() {
        try {
            findingText := ControlGetText(this.FindingEdit.NativeWindowHandle)
            impressionText := ControlGetText(this.ImpressionEdit.NativeWindowHandle)
        } catch {
            return false
        }

        return Trim(findingText, " `t`r`n") == "" && Trim(impressionText, " `t`r`n") == ""
    }

    static _ApplyLayout(hFind, hImp) {
        ; 1. 安全檢查：如果 Handle 為 0 或空，直接離開
        if !hFind || !hImp {
            if (this.IsDebug) {
                OutputDebug("[RisLayout] _ApplyLayout skipped: missing edit hwnd`n")
            }
            return
        }

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
        layoutShift := 15 * dpiScale

        ; 2. 加上 Try-Catch 保護：避免視窗切換瞬間抓不到位置而報錯
        try {
            ControlGetPos(&fX, &fY, &fW, &fH, hFind)
            ControlGetPos(&iX, &iY, &iW, &iH, hImp)
        } catch as err {
            if (this.IsDebug) {
                OutputDebug(Format("[RisLayout] ControlGetPos failed: {}`n", err.Message))
            }
            return ; 如果抓不到位置，這次就不調整
        }

        winHwnd := WinExist(this.WinTitle)
        layoutApplied := !!DllCall("GetProp", "ptr", winHwnd, "str", this._layoutShiftProp, "ptr")
        currentBottom := iY + iH
        if (layoutApplied) {
            currentBottom -= layoutShift
        }

        targetImpY := currentBottom - targetImpH + layoutShift
        targetFindH := (targetImpY - gap) - fY

        tolerance := 5 * dpiScale
        mainLayoutReady := layoutApplied && Abs(iH - targetImpH) < tolerance && Abs(iY - targetImpY) < tolerance && Abs(fH - targetFindH) < tolerance

        if (!mainLayoutReady) {
            try {
                ControlMove(,,, targetFindH, hFind) ; 先調整上面高度，避免重疊
                ControlMove(,, iW, targetImpH, hImp)
                ControlMove(, targetImpY,,, hImp)
                if (winHwnd) {
                    DllCall("SetProp", "ptr", winHwnd, "str", this._layoutShiftProp, "ptr", 1)
                }
            } catch as err {
                if (this.IsDebug) {
                    OutputDebug(Format("[RisLayout] main layout move failed: {}`n", err.Message))
                }
            }
        }

        try {
            elLabel := this._GetOrUpdateNode("ImpressionLabel")
            if (hLabel := elLabel.NativeWindowHandle) {
                ControlMove(, targetImpY - labelOffset,,, hLabel)
            }
        } catch as err {
            if (this.IsDebug) {
                OutputDebug(Format("[RisLayout] ImpressionLabel failed: {}`n", err.Message))
            }
        }

        try {
            for ctrlName in ["BottomInfoLabel", "CurrentCountLabel", "AutoNextCheckbox"] {
                hLabel := 0
                try {
                    elLabel := this._GetOrUpdateNode(ctrlName)
                    hLabel := elLabel.NativeWindowHandle
                } catch as err {
                    if (this.IsDebug) {
                        OutputDebug(Format("[RisLayout] {} resolve failed: {}`n", ctrlName, err.Message))
                    }
                }
                if (this.IsDebug) {
                    OutputDebug(Format("[RisLayout] {} hwnd={}`n", ctrlName, hLabel))
                }
                if (hLabel) {
                    ControlGetPos(, &ctrlY,,, hLabel)
                    ctrlLayoutApplied := !!DllCall("GetProp", "ptr", hLabel, "str", this._bottomCtrlShiftProp, "ptr")
                    targetCtrlY := ctrlLayoutApplied ? ctrlY : ctrlY + layoutShift
                    ControlMove(, targetCtrlY,,, hLabel)
                    if (!ctrlLayoutApplied) {
                        DllCall("SetProp", "ptr", hLabel, "str", this._bottomCtrlShiftProp, "ptr", 1)
                    }
                }
            }
        } catch as err {
            if (this.IsDebug) {
                OutputDebug(Format("[RisLayout] bottom controls failed: {}`n", err.Message))
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

            reportText := RisDate.ConvertRISDate(dateVal) . ": " . diagVal
            A_Clipboard := reportText
            this.Notify("病理報告已複製")
        } catch as err {
            this.Notify("複製失敗: " err.Message)
        }
    }

    static CopyPathologyReportOrMRN() {
        if (this.IsPathologyOnlineDataPage()) {
            this.CopyPathologyReport()
        } else {
            this.CopyCurrentMRN()
        }
    }

    static CopyCurrentMRN() {
        try {
            mrn := this._NormalizeMRN(this._GetCurrentMRN())
            if (mrn == "") {
                throw Error("找不到病歷號")
            }

            A_Clipboard := mrn
            this.Notify("病歷號已複製")
        } catch as err {
            this.Notify("複製失敗: " err.Message)
        }
    }

    static _NormalizeMRN(mrn) {
        normalized := RegExReplace(mrn, "^0+")
        return normalized != "" ? normalized : mrn
    }

    static IsPathologyOnlineDataPage() {
        return this.GetSelectedRightTabIndex() == 3
    }

    static GetSelectedRightTabIndex() {
        try {
            tabEle := this.RightTabControl
            hTab := tabEle.NativeWindowHandle
            if (hTab) {
                return ControlGetIndex(hTab)
            }

            tabs := tabEle.FindAll({ Type: "TabItem" })
            for index, tab in tabs {
                try {
                    if (tab.IsSelected) {
                        return index
                    }
                }
            }
        }
        return 0
    }

    static CopyPhExamReport() {
        try {
            dateVal := this.PhExamDateText.Value
            repVal := this.PhExamReportText.Value
            if (dateVal == "" && repVal == "") {
                throw Error("找不到檢查報告內容")
            }

            reportText := RisDate.ConvertRISDate(dateVal) . ": " . repVal
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
        RisVisualFeedback.ShowWaitCursor() ; [新增]
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
                        anchorCell := cellElements[1]
                        targetCell := cellElements[3]
                        historyExamName := targetCell.Value
                        if (this._IsRelatedReport(historyExamName, currExamName)) {
                            clickTarget := IsObject(anchorCell) ? anchorCell : targetCell
                            successAttempt := this._ClickPastReportCellAndWait(rowEle, clickTarget)
                            if !successAttempt {
                                this.Notify("選取失敗: " historyExamName)
                                return
                            }
                            notifyText := this.IsDebug
                                ? "已選取 (第 " successAttempt " 次 click): " historyExamName
                                : "已選取: " historyExamName
                            this.Notify(notifyText)
                            return
                        }
                    }
                }
                this.Notify("未找到相似報告", 1000)
            } catch as err {
                this.Notify("搜尋失敗: " err.Message)
            }
        } finally {
            RisVisualFeedback.RestoreCursor() ; [新增]
        }
    }

    static _ClickPastReportCellAndWait(rowEle, clickTarget) {
        try WinActivate(this.WinTitle)
        try {
            hTable := this.PastReportTable.NativeWindowHandle
            if (hTable) {
                ControlFocus(hTable)
            }
        }

        loop 3 {
            clickTarget.LegacyIAccessiblePattern.DoDefaultAction()
            clickTarget.ControlClick()
            if this._WaitPastReportRowSelected(rowEle, clickTarget, 300) {
                Sleep 60
                return A_Index
            }

            Sleep 80
        }

        return 0
    }

    static _WaitPastReportRowSelected(rowEle, clickTarget, timeoutMs) {
        startTime := A_TickCount
        loop {
            if this._IsPastReportRowSelected(rowEle, clickTarget) {
                return true
            }
            if (A_TickCount - startTime >= timeoutMs) {
                break
            }
            Sleep 30
        }
        return false
    }

    static _IsPastReportRowSelected(rowEle, clickTarget) {
        static STATE_SYSTEM_SELECTED := 0x2

        try {
            if (rowEle.LegacyIAccessiblePattern.State & STATE_SYSTEM_SELECTED) {
                return true
            }
        }
        try {
            if (clickTarget.LegacyIAccessiblePattern.State & STATE_SYSTEM_SELECTED) {
                return true
            }
        }

        return false
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

    static _NormalizePatientSex(sexText) {
        sexText := Trim(sexText)

        if (InStr(sexText, "男") || RegExMatch(sexText, "i)\b(M|Male)\b")) {
            return "M"
        }
        if (InStr(sexText, "女") || RegExMatch(sexText, "i)\b(F|Female)\b")) {
            return "F"
        }
        return ""
    }

    static _GetSexSpecificTermRules(sex) {
        if (sex == "M") {
            return [
                {Label: "uterus / uterine", Pattern: "\b(uterus|uteri|uterine)\b"},
                {Label: "ovary / ovarian", Pattern: "\b(ovary|ovaries|ovarian)\b"},
                {Label: "fallopian tube", Pattern: "\b(fallopian\s+tubes?)\b"},
                {Label: "adnexa / adnexal", Pattern: "\b(adnexa|adnexae|adnexal)\b"},
                {Label: "endometrium / endometrial", Pattern: "\b(endometrium|endometrial)\b"},
                {Label: "myometrium / myometrial", Pattern: "\b(myometrium|myometrial)\b"},
                {Label: "cervix / gynecologic cervical", Pattern: "\b(cervix|uterine\s+cervix|cervical\s+(canal|os|mass|cancer|carcinoma|lesion))\b"},
                {Label: "hysterectomy", Pattern: "\b(hysterectomy|hysterectomies)\b"},
                {Label: "vagina / vaginal", Pattern: "\b(vagina|vaginal)\b"},
                {Label: "vulva / vulvar", Pattern: "\b(vulva|vulvar)\b"}
            ]
        }

        if (sex == "F") {
            return [
                {Label: "prostate / prostatic", Pattern: "\b(prostate|prostatic)\b"},
                {Label: "prostatectomy", Pattern: "\b(prostatectomy|prostatectomies)\b"},
                {Label: "seminal vesicle", Pattern: "\b(seminal\s+vesicles?)\b"},
                {Label: "testis / testicular", Pattern: "\b(testis|testes|testicular)\b"},
                {Label: "scrotum / scrotal", Pattern: "\b(scrotum|scrotal)\b"},
                {Label: "penis / penile", Pattern: "\b(penis|penile)\b"},
                {Label: "epididymis / epididymal", Pattern: "\b(epididymis|epididymal)\b"},
                {Label: "vas deferens", Pattern: "\b(vas\s+deferens|deferential\s+ducts?)\b"},
                {Label: "spermatic cord", Pattern: "\b(spermatic\s+cords?)\b"}
            ]
        }

        return []
    }

    static _FindFirstSexSpecificReportConflict(sex) {
        rules := this._GetSexSpecificTermRules(sex)
        targets := [
            {Name: "Finding", Hwnd: this.FindingEdit.NativeWindowHandle, Text: this.FindingText},
            {Name: "Impression", Hwnd: this.ImpressionEdit.NativeWindowHandle, Text: this.ImpressionText}
        ]

        for target in targets {
            firstMatch := 0

            for rule in rules {
                if RegExMatch(target.Text, "i)" . rule.Pattern, &match) {
                    matchPos := match.Pos - 1
                    if (!firstMatch || matchPos < firstMatch.Pos) {
                        firstMatch := {
                            Field: target.Name,
                            Hwnd: target.Hwnd,
                            Pos: matchPos,
                            Term: match[0],
                            Label: rule.Label
                        }
                    }
                }
            }

            if (firstMatch) {
                return firstMatch
            }
        }

        return 0
    }

    static _SelectReportLineAt(hCtrl, charPos) {
        try {
            ControlFocus(hCtrl)
            RisEditControl.SetSel(hCtrl, charPos, charPos)
            RisEditControl.SelectLine(hCtrl)
            RisEditControl.ScrollCaret(hCtrl)
        }
    }

    static ValidateReportSexSpecificTerms() {
        try {
            sex := this._NormalizePatientSex(this._FastGetCtrlText("GenderText"))

            if (sex == "") {
                this.Notify("無法判讀病人性別，已取消存檔。請確認性別欄位後再存檔。")
                return false
            }

            conflict := this._FindFirstSexSpecificReportConflict(sex)
            if (!conflict) {
                return true
            }

            sexName := (sex == "M") ? "男性" : "女性"
            this._SelectReportLineAt(conflict.Hwnd, conflict.Pos)
            this.Notify(Format("病人性別為 {1}，{2} 含不符性別字詞: {3}`n已取消存檔，請修正後再存檔。", sexName, conflict.Field, conflict.Term))
            return false
        } catch as err {
            this.Notify("報告安全檢查失敗，已取消存檔: " . err.Message)
            return false
        }
    }

    static SaveReport() {
        try {
            this.EnsureImpressionNotEmpty()
            if (!this.ValidateReportSexSpecificTerms()) {
                return
            }
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
            return RisEditControl.SmartExtendSelection(hCtrl, direction)
        } catch {
            return false
        }
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
            RisEditControl.SelectLine(hMouseCtrl)
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

    ; =================================================================
    ; 9. 內部 Helper (Low-level Helpers)
    ; =================================================================

    ; --- Edit Control 底層操作 (封裝 SendMessage) ---

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
        formattedDate := RisDate.ConvertRISDate(targetDate)
        this._compContext.ReqNo := currentReqNo
        this._compContext.Date := formattedDate
    }

    static GetComparisonSuffix() {
        if (!this.IsStandardRis) {
            return ""
        }
        currentReqNo := this._GetCurrentReqNo()
        if (this._compContext.ReqNo != "" && this._compContext.ReqNo == currentReqNo) {
            return " dated " . this._compContext.Date
        }
        return ""
    }

    static GetComparisonDate() {
        if (!this.IsStandardRis) {
            return ""
        }
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

    static GetIndicationFollowupSuffix() {
        comparisonDate := this.GetComparisonDate()
        if (comparisonDate != "") {
            return "COMPARISON: " . comparisonDate . "`r`n`r`nFINDINGS:`r`n"
        }
        return "FINDINGS:`r`n"
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
        selectedText := RisEditControl.GetSelectedText(targetHwnd)
        if (selectedText == "") {
            return
        }

        finalText := RisReportText.ReorderSelectedText(selectedText, deOrder, keepEmptyLine, itemChar, discardSeIm, forceStartFromOne)
        RisEditControl.ReplaceSelectionPreserveFirstVisibleLine(targetHwnd, finalText)
    }

    static _FormatFindingForBasic(hEdit) {
        fullText := ControlGetText(hEdit)
        range := RisReportText.FindContentRange(fullText, "Basic")

        if (range) {
            RisEditControl.SetSel(hEdit, range.Start, range.End)
            this._ReorderSelectedText(hEdit, false, true, "-", false)
        } else {
            this.Notify("報告格式不如預期，無法自動排版")
        }
    }

    static _FormatFindingForAdvanced(hEdit) {
        fullText := ControlGetText(hEdit)
        range := RisReportText.FindContentRange(fullText, "Advanced")

        if (range) {
            RisEditControl.SetSel(hEdit, range.Start, range.End)
            if (range.HasOwnProp("TrailingNewlines")) {
                selectedText := RisEditControl.GetSelectedText(hEdit)
                if (selectedText == "") {
                    return
                }

                finalText := RisReportText.ReorderSelectedText(selectedText, false, false, "-", true)
                finalText := RTrim(finalText, "`r`n") . range.TrailingNewlines
                RisEditControl.ReplaceSelectionPreserveFirstVisibleLine(hEdit, finalText)
            } else {
                this._ReorderSelectedText(hEdit, false, false, "-", true)
            }
        } else {
            this.Notify("報告格式不如預期，無法自動排版")
        }
    }

    static _GetCleanCurrentExamName() {
        try {
            return StrReplace(ControlGetText(this.ExamnameText.NativeWindowHandle), "檢查項目: ", "")
        }
        return ""
    }

    static _GetCurrExamType() {
        return RisReportText.GetExamType(this._GetCleanCurrentExamName())
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

        RisVisualFeedback.ShowWaitCursor()
        this._isAIPending := true
        return true
    }

    static _FinishForegroundAIRequest() {
        this._isAIPending := false
        RisVisualFeedback.RestoreCursor()
    }

    ; 10.0.1 Indication
    static _TryInsertCachedIndication(isPreloadOnly, insertAfter := "") {
        if (!this._aiCache.Has("_AI_Indication") || isPreloadOnly) {
            return false
        }

        cached := this._aiCache["_AI_Indication"]
        apiKeyName := cached.HasOwnProp("apiKeyName") ? cached.apiKeyName : "APIKey"
        modelName := cached.HasOwnProp("modelName") ? cached.modelName : "Model"
        this._InsertAIResult(cached.text, insertAfter)
        this.Notify(RisAIOrchestration.FormatCompleteNotify("已插入 Indication (來自快取)", apiKeyName, modelName, Format("API:{}ms", cached.apiTime)), 2500)
        return true
    }

    static _TryHandlePendingIndication(isPreloadOnly, insertAfter := "") {
        if (!this._isIndicationPending) {
            return false
        }

        if (isPreloadOnly) {
            return true
        }

        this._pendingIndicationInsert := true
        this._pendingIndicationInsertAfter := insertAfter
        this.Notify("AI 正在背景產生中，完成後將自動插入...")
        return true
    }

    static _BeginIndicationRequest(isPreloadOnly, insertAfter := "") {
        if (this._isAIPending) {
            if (!isPreloadOnly) {
                this.Notify("AI 正在背景產生中...")
            }
            return false
        }

        if (!isPreloadOnly) {
            RisVisualFeedback.ShowWaitCursor()
            this._pendingIndicationInsert := true
            this._pendingIndicationInsertAfter := insertAfter
        }

        this._isAIPending := true
        this._isIndicationPending := true
        return true
    }

    static _NormalizeIndicationResult(result) {
        result := RisAIOrchestration.NormalizeResult(result)

        if (!InStr(result, "INDICATION:")) {
            result := "INDICATION: " . result
        }

        return result
    }

    static _CacheIndicationResult(result, apiTime, extractTime, apiKeyName, modelName) {
        this._aiCache["_AI_Indication"] := {
            text: result,
            apiTime: apiTime,
            extractTime: extractTime,
            apiKeyName: apiKeyName,
            modelName: modelName
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

    static _HandleIndicationSuccess(isPreloadOnly, result, apiTime, extractTime, apiKeyName, modelName, debugMode := false) {
        normalized := this._NormalizeIndicationResult(result)
        this._CacheIndicationResult(normalized, apiTime, extractTime, apiKeyName, modelName)

        if (debugMode && !isPreloadOnly) {
            MsgBox("【Benchmark】`n資料提取: " . extractTime . " ms`nAPI 耗時: " . apiTime . " ms`n`n【API 回傳結果】`n" . normalized, "AI Debug")
        }

        if (!isPreloadOnly) {
            this.Notify(RisAIOrchestration.FormatCompleteNotify("已產生 Indication", apiKeyName, modelName, Format("取資:{}ms, API:{}ms", extractTime, apiTime)), 2500)
        } else {
            OutputDebug("[RisController] AI Indication 已預載並快取`n")
        }
    }

    static _FinishIndicationRequest(requestMode, isPreloadOnly) {
        this._isAIPending := false
        this._isIndicationPending := false

        if (!isPreloadOnly) {
            RisVisualFeedback.RestoreCursor()
        }

        shouldInsert := this._pendingIndicationInsert && this._aiCache.Has("_AI_Indication")
        if (shouldInsert) {
            insertAfter := this.HasOwnProp("_pendingIndicationInsertAfter") ? this._pendingIndicationInsertAfter : ""
            this._InsertAIResult(this._aiCache["_AI_Indication"].text, insertAfter)
            if (requestMode == "preload") {
                cached := this._aiCache["_AI_Indication"]
                apiKeyName := cached.HasOwnProp("apiKeyName") ? cached.apiKeyName : "APIKey"
                modelName := cached.HasOwnProp("modelName") ? cached.modelName : "Model"
                this.Notify(RisAIOrchestration.FormatCompleteNotify("Indication 已完成並插入", apiKeyName, modelName, Format("API:{}ms", cached.apiTime)), 2500)
            }
        }

        this._pendingIndicationInsert := false
        this._pendingIndicationInsertAfter := ""
    }

    static _HandleImpressionDebugPrompt(debugMode, fullPrompt) {
        if (!debugMode) {
            return true
        }

        A_Clipboard := fullPrompt
        return RisAIDebugGui.ShowPromptConfirm("AI Debug - Impression Prompt", fullPrompt, {
            Notify: this.Notify.Bind(this),
            Header: "Prompt 已複製到剪貼簿。確認後才會繼續呼叫 API。"
        })
    }

    static _HandleImpressionSuccess(result, extractTime, apiTime, apiKeyName, modelName) {
        result := RisAIOrchestration.NormalizeImpressionResult(result)
        this._InsertAIResultToImpression(result)
        this.Notify(RisAIOrchestration.FormatCompleteNotify("已插入 Impression", apiKeyName, modelName, Format("取資:{}ms, API:{}ms", extractTime, apiTime)), 2500)
    }

    static _BuildAIRequestResult(promptText, aiConfig) {
        t0 := A_TickCount
        response := this._CallAI(promptText, aiConfig)

        return RisAIOrchestration.BuildRequestResult(response, A_TickCount - t0)
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

        return RisAIOrchestration.CreateRequest(fullPrompt, conf, {
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

        findingText := RisReportText.DeidentifyText(findingText)
        clinicalContext := this._GetImpressionClinicalContext()
        extractTime := A_TickCount - t0
        conf := RisConfig.AI.Impression

        return RisAIOrchestration.CreateRequest(Format(conf.Prompt, clinicalContext, findingText), conf, {
            ExtractTime: extractTime,
            ImpressionHwnd: hImp
        })
    }

    static _BuildRefineRequest(selectedText) {
        conf := RisConfig.AI.Refine
        prompt := conf.SystemPrompt . "`n`nInput Text:`n" . selectedText

        return RisAIOrchestration.CreateRequest(prompt, conf)
    }

    static _RunAIRequest(request) {
        return this._BuildAIRequestResult(request.Prompt, request.Config)
    }

    static _RunIndicationRequest(request, debugMode, isPreloadOnly) {
        response := this._RunAIRequest(request)
        this._HandleIndicationSuccess(isPreloadOnly, response.Result, response.ApiTime, request.ExtractTime, response.APIKeyName, response.Model, debugMode)
    }

    static _RunImpressionRequest(request) {
        response := this._RunAIRequest(request)
        this._HandleImpressionSuccess(response.Result, request.ExtractTime, response.ApiTime, response.APIKeyName, response.Model)
    }

    static _RunRefineRequest(request) {
        response := this._RunAIRequest(request)
        return response
    }

    ; [新增] 外部呼叫的主函式：產生並插入 Indication
    ; [修改] 增加 Benchmark 效能測量
    static GenerateAndInsertIndication(debugMode?, isPreloadOnly := false, insertAfter := "") {
        debugMode := IsSet(debugMode) ? debugMode : this.IsDebug
        requestMode := isPreloadOnly ? "preload" : "manual"

        if (this._TryInsertCachedIndication(isPreloadOnly, insertAfter)) {
            return
        }

        if (this._TryHandlePendingIndication(isPreloadOnly, insertAfter)) {
            return
        }

        if (!this._BeginIndicationRequest(isPreloadOnly, insertAfter)) {
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
    static GenerateAndInsertImpression(debugMode?) {
        debugMode := IsSet(debugMode) ? debugMode : this.IsDebug
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
    static _GetPolishSelectionContext(selectCurrentLineIfEmpty := false) {
        if !this.IsTargetFocused() {
            this.Notify("請先點擊要處理的文字欄位")
            return false
        }

        hEdit := ControlGetFocus("A")
        sel := RisEditControl.GetSel(hEdit)
        fullText := ControlGetText(hEdit)

        if (selectCurrentLineIfEmpty && sel.Start == sel.End) {
            bounds := RisEditControl.GetLogicalLineBoundaries(hEdit)
            lineText := (bounds.ContentEnd > bounds.Start)
                ? SubStr(fullText, bounds.Start + 1, bounds.ContentEnd - bounds.Start)
                : ""
            if (Trim(lineText, " `t") == "") {
                this.Notify("目前行為空")
                return false
            }

            RisEditControl.SetSel(hEdit, bounds.Start, bounds.FullEnd)
            sel := {Start: bounds.Start, End: bounds.FullEnd}
        }

        ; 檢查是否有選取文字
        if (sel.Start == sel.End) {
            this.Notify("請先選取要潤色的文字")
            return false
        }

        selectedText := SubStr(fullText, sel.Start + 1, sel.End - sel.Start)
        trailingNewlines := ""

        if (Trim(selectedText, " `t`r`n") == "") {
            this.Notify("選取的文字為空")
            return false
        }

        trailingNewlines := RisReportText.GetNormalizedTrailingNewlines(selectedText)

        return {
            Hwnd: hEdit,
            Selection: sel,
            SelectedText: selectedText,
            TrailingNewlines: trailingNewlines
        }
    }

    static PolishSelectionWithAI() {
        try {
            context := this._GetPolishSelectionContext(true)
            if (!context) {
                return
            }

            RisVisualFeedback.ShowWaitCursor()
            this.Notify("AI 潤色中...", 3000)

            request := this._BuildRefineRequest(context.SelectedText)
            response := this._RunRefineRequest(request)
            result := RisAIOrchestration.NormalizePolishResult(response.Result, context.TrailingNewlines)

            RisVisualFeedback.RestoreCursor()

            ; 顯示比對視窗
            debugInfo := RisAIOrchestration.FormatPolishComparisonDebugInfo(response)
            this._ShowPolishComparisonGui(context.Hwnd, context.SelectedText, result, context.Selection, debugInfo)

        } catch as err {
            RisVisualFeedback.RestoreCursor()
            this.Notify("AI 潤色失敗: " . err.Message)
        }
    }

    static CompareSelectionWithAI() {
        if (!this._BeginForegroundAIRequest()) {
            return
        }

        try {
            context := this._GetPolishSelectionContext(true)
            if (!context) {
                return
            }

            this.Notify("OpenAI / Google AI 潤色中...", 3000)
            results := RisAIService.RunRefineProvidersParallel(context.SelectedText, [
                {Provider: "openai", DisplayName: "OpenAI"},
                {Provider: "google", DisplayName: "Google"}
            ], context.TrailingNewlines, this._GetAIServiceOptions())
            openAIResult := results["openai"]
            googleResult := results["google"]

            if (!openAIResult.Success && !googleResult.Success) {
                this.Notify("AI 對比失敗：兩家 provider 都沒有可用結果")
                return
            }

            this._ShowPolishProviderComparisonGui(context.Hwnd, context.SelectedText, openAIResult, googleResult, context.Selection)

        } catch as err {
            this.Notify("AI 對比失敗: " . err.Message)
        } finally {
            this._FinishForegroundAIRequest()
        }
    }

    static _ShowPolishComparisonGui(hEdit, original, refined, sel, debugInfo := "") {
        options := {
            Notify: this.Notify.Bind(this),
            ApplyFont: this._ApplyCustomFontToControls.Bind(this),
            OnAccept: (hEdit, finalText, sel) => (
                finalText := StrReplace(finalText, "`r`n", "`n"),
                finalText := StrReplace(finalText, "`n", "`r`n"),
                RisEditControl.SetSel(hEdit, sel.Start, sel.End),
                RisEditControl.ReplaceSelectionAndScroll(hEdit, finalText)
            )
        }
        RisAIDebugGui.ShowPolishComparisonGui(hEdit, original, refined, sel, debugInfo, options)
    }

    static _ShowPolishProviderComparisonGui(hEdit, original, openAIResult, googleResult, sel) {
        options := {
            Notify: this.Notify.Bind(this),
            ApplyFont: this._ApplyCustomFontToControls.Bind(this),
            OnAccept: (hEdit, finalText, sel) => (
                finalText := StrReplace(finalText, "`r`n", "`n"),
                finalText := StrReplace(finalText, "`n", "`r`n"),
                RisEditControl.SetSel(hEdit, sel.Start, sel.End),
                RisEditControl.ReplaceSelectionAndScroll(hEdit, finalText)
            )
        }
        RisAIDebugGui.ShowPolishProviderComparisonGui(hEdit, original, openAIResult, googleResult, sel, options)
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
        RisEditControl.ReplaceSelectionAndScroll(targetHwnd, result)
    }

    ; [內部 Helper] 執行 AI 結果插入 UI
    static _InsertAIResult(result, insertAfter := "") {
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

        RisEditControl.ReplaceSelectionAndScroll(targetHwnd, result . "`r`n`r`n" . insertAfter)
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

    static _GetImpressionClinicalContext() {
        dept := this._FastGetCtrlText("OrderDeptText")
        bedInfo := this._FastGetCtrlText("BedInfoText")
        visitType := this._HasBedNumber(bedInfo) ? "急診/住院" : "門診"

        rawText := Format("{1}`n檢查來源: {2}", dept, visitType)
        return RisReportText.DeidentifyText(rawText)
    }

    static _HasBedNumber(bedInfo) {
        bedInfo := StrReplace(bedInfo, "病床號:", "")
        bedInfo := Trim(StrReplace(bedInfo, "病床號：", ""), " `t`r`n")
        return bedInfo != ""
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

        return RisReportText.DeidentifyText(rawText)
    }

    ; =================================================================
    ; 10.1 AI Transport
    ; request prepare / transport wait / response parse
    ; =================================================================
    static _CallAI(promptText, aiConfig := 0) {
        return RisAIService.Call(promptText, aiConfig, this._GetAIServiceOptions())
    }

    static _GetAIServiceOptions() {
        return {
            Notify: this.Notify.Bind(this),
            ShowCurl: this.ShowGoogleAICurlDebug,
            ShowGoogleDebugCurl: this._ShowGoogleAIDebugCurl.Bind(this),
            GetGoogleConfig: this._GetGoogleAIConfig.Bind(this),
            GetOpenAIConfig: this._GetOpenAIConfig.Bind(this)
        }
    }

    static _GetGoogleAIConfig() {
        if (this._googleAIConfig) {
            return this._googleAIConfig
        }

        this._googleAIConfig := RisAIConfigResolver.GetGoogleConfig()
        return this._googleAIConfig
    }

    static _GetOpenAIConfig() {
        return RisAIConfigResolver.GetOpenAIConfig()
    }

    static _ResolveGoogleAIModelList(aiConfig := 0) {
        return RisAIService._ResolveGoogleAIModelList(aiConfig, this._GetAIServiceOptions())
    }

    static _BuildGoogleAIRequest(promptText, aiConfig := 0, modelOverride := "") {
        return RisAIService._BuildGoogleAIRequest(promptText, aiConfig, modelOverride, this._GetAIServiceOptions())
    }

    static _ResolveOpenAIModelList(aiConfig := 0) {
        return RisAIService._ResolveOpenAIModelList(aiConfig, this._GetAIServiceOptions())
    }

    static _BuildOpenAIRequest(promptText, aiConfig := 0, modelOverride := "") {
        return RisAIService._BuildOpenAIRequest(promptText, aiConfig, modelOverride, this._GetAIServiceOptions())
    }

    static _ShowGoogleAIDebugCurl(url, payload, response, request := 0) {
        RisAIDebugGui.ShowGoogleAIDebugCurl(url, payload, response, request, { Notify: this.Notify.Bind(this) })
    }

    ; [新增] 專用於 Debug 的持久化錯誤視窗
    static _ShowDebugError(errMsg) {
        RisAIDebugGui.ShowDebugError(errMsg, { Notify: this.Notify.Bind(this) })
    }
}
