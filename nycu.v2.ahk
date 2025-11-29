#Requires AutoHotkey v2.0
#SingleInstance Force
ProcessSetPriority "High"  ; 提高優先級

; 先設定選項：SendInput (SI) 和 忽略終止符 (O)
#Hotstring SI O
; 再設定終止字元：只有 Tab
#Hotstring EndChars `t

; === 解決 DPI 座標偏移問題 ===
; 1. 宣告 DPI 感知 (讓 BoundingRectangle 回報實體座標)
DllCall("SetThreadDpiAwarenessContext", "ptr", -3, "ptr")
; 2. 確保 AHK 的 Click 使用螢幕座標
CoordMode "Mouse", "Screen"
; ================================
SetMouseDelay -1
SetKeyDelay -1

global PRESERVE_CLIPBOARD := 0

#Include <UIA.v2>
#Include <Paste.v2>
#Include <Edit.v2>

#Include MyScripts\hotkey\reorder-selected-text.v2.ahk
#Include MyScripts\hotkey\remapping-original-hotkeys-infinitt.v2.ahk

global RISReportWinTitle := "報告作業(frmRISReport)" ; 替換成您的程式標題
global RISCTMRAbnormalWinTitle := "檢查結果(frmPos)"
global UIA_PastReportTable := { AutomationId: "dgvPastReport" }
global UIA_AutoNextCheckbox := { AutomationId: "chkAutoNext" }
global UIA_ReportSaveButton := { AutomationId: "btnReportSave" }
global UIA_ExamNameTxt := { AutomationId: "txtExamName" }
global UIA_FindingEdit := { AutomationId: "txtReport" }
global UIA_ImpressionEdit := { AutomationId: "txtImpression" }
global UIA_PastAllRadio := { AutomationId: "rdoPastALL" }
global UIA_PastModalityRadio := { AutomationId: "rdoClassify" }
global UIA_PastOnlyMyRadio := { AutomationId: "rdoPastOnlyMy" }
global UIA_PastReportFindingTxt := { AutomationId: "rtxtPastReport" }
global UIA_PastReportImpressionTxt := { AutomationId: "rtxtPastImpression" }
global UIA_PathoDiagnosisTxt := { AutomationId: "txtDiagnosist" }
global UIA_PathoDateTxt := { AutomationId: "mtxtRcpDTM" }
global ABNORMAL_VALUE_1_CONTROL := "WindowsForms10.BUTTON.app.0.2780b98_r24_ad13"
global ABNORMAL_VALUE_2_CONTROL := "WindowsForms10.BUTTON.app.0.2780b98_r24_ad15"
global ABNORMAL_VALUE_3_CONTROL := "WindowsForms10.BUTTON.app.0.2780b98_r24_ad16"
global ABNORMAL_VALUE_4_CONTROL := "WindowsForms10.BUTTON.app.0.2780b98_r24_ad14"
global ABNORMAL_VALUE_SAVE_BUTTON_CONTROL := "WindowsForms10.BUTTON.app.0.2780b98_r24_ad12"
global EXAMNAME_HWND := 0
global FINDING_CONTROL_HWND := 0
global IMPRESSION_CONTROL_HWND := 0
global PAST_ALL_RADIO_HWND := 0
global PAST_MODALITY_RADIO_HWND := 0
global PAST_ONLY_MY_RADIO_HWND := 0
global PAST_FINDING_HWND := 0
global PAST_IMPRESSION_HWND := 0
global risEle := ""
global autoNextEle := ""
global reportSaveEle := ""

; 腳本啟動時執行，設定為 Fira Code, 大小 14 (根據您的螢幕解析度調整)
RisController.EnableFontEnforcer("Cascadia Code", 12)

#HotIf WinActive(RISReportWinTitle)

#Include MyScripts\regex-hotstrings.v2.ahk
#Include MyScripts\others.v2.ahk
#Include MyScripts\chest-ct.v2.ahk
#Include MyScripts\abdomen-ct.v2.ahk
#Include MyScripts\abdomen-mr.v2.ahk
#Include MyScripts\ct-guide.v2.ahk
#Include MyScripts\ms-ct.v2.ahk
#Include MyScripts\ms-mri.v2.ahk
#Include MyScripts\neuro.v2.ahk
#Include MyScripts\abbreviations.v2.ahk
#Include MyScripts\mri.v2.ahk
#Include MyScripts\chest-x-ray.v2.ahk
#Include MyScripts\kub.v2.ahk
#Include MyScripts\bone-x-ray.v2.ahk
#Include MyScripts\other-x-ray.v2.ahk
#Include MyScripts\sono-guide.v2.ahk
#Include MyScripts\angio.v2.ahk
#Include MyScripts\sono.v2.ahk

#Include MyScripts\lib\ris-common.v2.ahk
^9:: {
}

; --- Emacs Word Movement ---
; 加上防呆機制：只有在打字時，Alt+F 才是向右移動
; 否則它會保留 Windows 預設行為 (開啟檔案選單)
!f:: {
    if !RisController.MoveCaretWord("Right")
        Send "!f" ; 透傳
}

!b:: {
    if !RisController.MoveCaretWord("Left")
        Send "!b" ; 透傳
}

; --- History Filter Switching ---
; 這裡不需要透傳，因為 Ctrl+數字鍵通常沒有其他重要功能
^1:: RisController.SwitchHistoryFilter("All")
^2:: RisController.SwitchHistoryFilter("Modality")
^3:: RisController.SwitchHistoryFilter("My")

; --- Business Logic ---
^Esc:: RisController.AppendPreviousReport()

; Ctrl+A: Emacs 行首 (若不在目標框則為全選)
^a:: {
    ; 嘗試移動到行首
    didMove := RisController.MoveCaret("Start")

    ; 如果不在 Finding/Impression 框框內，則送出原本的 Ctrl+A (全選)
    if (!didMove) {
        Send "^a"
    }
}

; Ctrl+E: Emacs 行尾
^e:: {
    didMove := RisController.MoveCaret("End")

    ; 如果不在目標框，送出原生的 Ctrl+E (有些系統可能是置中對齊或其他功能)
    if (!didMove) {
        Send "^e"
    }
}

; Ctrl+D 刪除整行
^d:: {
    ; 呼叫 Controller 執行刪除
    isDeleted := RisController.DeleteCurrentLine()

    ; 如果 Controller 回傳 false (代表沒在目標框內，或執行失敗)
    ; 則將 Ctrl+D 原封不動送回給系統
    if (!isDeleted) {
        Send("^d")
    }
}

; Ctrl+W: 刪除前一個字 (Bash Style)
^w:: {
    ; 嘗試呼叫 Controller 執行刪除
    didDelete := RisController.DeleteWordBackward()

    ; 如果 Controller 回傳 false (代表焦點不在 Finding/Impression 框框內)
    ; 則透傳 Ctrl+W 給系統 (例如：醫師在瀏覽器想關分頁，或在其他地方想用原生的 Ctrl+W)
    if (!didDelete) {
        Send "^w"
    }
}

; ==============================================================================
; 3. Benchmark 工具函數 (高精確度)
; ==============================================================================
; 參數:
; funcObj: 要測試的函數物件 (使用 CallbackCreate 或 Func.Bind)
; times: 執行的次數
; 回傳: 執行總耗時 (毫秒, ms)
; ==============================================================================
Benchmark(funcObj, times := 1) {
    ; 1. 獲取計時器頻率 (每秒多少 ticks)
    DllCall("QueryPerformanceFrequency", "Int64*", &freq := 0)

    ; 2. 記錄開始時間
    DllCall("QueryPerformanceCounter", "Int64*", &start := 0)

    ; 3. 執行迴圈
    loop times {
        funcObj()
    }

    ; 4. 記錄結束時間
    DllCall("QueryPerformanceCounter", "Int64*", &end := 0)

    ; 5. 計算耗時 ( (結束-開始) / 頻率 * 1000 轉換為毫秒 )
    return (end - start) / freq * 1000
}

; Alt+E: 在游標處插入檢查名稱
!e:: {
    ; 如果焦點在編輯框內，執行插入
    ; 如果焦點在其他地方，送出 Alt+E (開啟系統選單)
    if !RisController.InsertExamNameAtCaret() {
        Send "!e"
    }
}

; Alt+C: 取消 AutoNext 並存檔 (Save Only)
!c:: {
    ; 1. 設定 AutoNext 為 False (取消勾選)
    RisController.SetAutoNextState(false)
    ; 2. 存檔
    RisController.SaveReport()
}

; Ctrl+S: 勾選 AutoNext 並存檔 (Save & Next)
^s:: {
    ; 1. 設定 AutoNext 為 True (勾選)
    RisController.SetAutoNextState(true)
    ; 2. 存檔
    RisController.SaveReport()
}

!q:: {
    Send "^e"
}

^k:: {
    Send "+{End}"
    Send "{Del}"
}

#a:: {
    Send "^a"
}

#d:: {
    Send "^a"
    Sleep 100
    Send "{Del}"
}

; Shift + Up: 智慧向上選取
+Up:: {
    ; 如果沒執行智慧選取 (不在編輯框內)，則送出原生的 Shift+Up
    if !RisController.SmartExtendSelection("Up") {
        SendInput "+{Up}"
    }
}

; Shift + Down: 智慧向下選取
+Down:: {
    if !RisController.SmartExtendSelection("Down") {
        SendInput "+{Down}"
    }
}

; Alt+Esc: 根據目前的檢查名稱，自動搜尋並選取歷史報告中的相似項目
!Esc:: RisController.FindAndClickSimilarReport()

; Alt+D: 插入目前選取的歷史報告日期 (自動轉西元)
!d:: {
    if !RisController.IsTargetFocused()
        Send "!d" ; 透傳
    else
        RisController.InsertSelectedHistoryDate()
}

; Ctrl+Alt+E: 插入目前選取的歷史報告名稱
^!e:: {
    if !RisController.IsTargetFocused()
        Send "^!e" ; 透傳
    else
        RisController.InsertSelectedHistoryName()
}

ConvertRISDate(inputString) {
    ; 1. 標準化輸入：移除 "/" 符號
    ;    這樣 "114/10/14" 會變成 "1141014"
    ;    而 "1141014" 則不受影響
    cleanString := StrReplace(inputString, "/")

    ; 2. 從已清理的字串 (yyymmdd...) 中提取各個部分
    minguoYear := SubStr(cleanString, 1, 3)  ; yyy (例如: 114)
    month := SubStr(cleanString, 4, 2)        ; mm (例如: 10)
    day := SubStr(cleanString, 6, 2)          ; dd (例如: 14)

    ; 3. 將民國年轉換為西元年 (民國年 + 1911 = 西元年)
    gregorianYear := minguoYear + 1911

    ; 4. 組合並返回 yyyy-mm-dd 格式的字串
    ;    (維持您原本的 . 串接風格)
    outputDate := gregorianYear . "-" . month . "-" . day

    return outputDate
}

GetCurrExamName() {
    try {
        hExamname := RisController.ExamnameText.NativeWindowHandle
        examname := StrReplace(ControlGetText(hExamname), "檢查項目: ", "")
        return examname
    } catch as err {
        MsgBox "操作失敗: " err.Message
        return
    }
}

GetCurrExamType() {
    examname := GetCurrExamName()
    if (InStr(examname, "CT") || InStr(examname, "電腦斷層")) {
        return "CT"
    } else if (InStr(examname, "MR") || InStr(examname, "磁振造影")) {
        return "MR"
    } else if (InStr(examname, "US") || InStr(examname, "超音波")) {
        return "US"
    }
    return "CR"
}

OrderListForFindings() {
    examtype := GetCurrExamType()
    ;MsgBox(examtype)
    switch examtype {
        case "CT", "MR":
            UnorderListForFindingsOfCtOrMr()

        case "CR", "US":
            UnorderListForFindingsOfCrOrUs()
    }
}

UnorderListForFindingsOfCrOrUs() {
    ; 取得 Handle
    if (hEdit := RisController.FindingEdit.NativeWindowHandle) {

        ; 搜尋並選取文字 (這部分保持原本邏輯)
        startSel := Edit_FindText(hEdit, "FINDINGS:`r`n|:\s*`r`n\s*`r`n", , , "RegEx", &matchedText)
        if (startSel > -1) {
            startSel += matchedText.Len
            Edit_SetFocus(hEdit)
            Edit_SetSel(hEdit, startSel, -1)

            ; --- 修改點：直接將 hEdit 傳入 ---
            ReorderSelectedText(false, true, "-", false, hEdit)
        }
    }
}

UnorderListForFindingsOfCtOrMr() {
    if (hEdit := RisController.FindingEdit.NativeWindowHandle) {
        startSel := Edit_FindText(hEdit,
            "FINDINGS:`r`n|The study shows:`r`n`r`n|show the following findings:`r`n`r`n|which revealed:`r`n`r`n", , ,
            "RegEx", &matchedText)

        if (startSel > -1) {
            ;startSel += StrLen(matchedText)
            startSel += matchedText.Len
            loop 3 {
                newStartSel := startSel
                startText := Edit_GetTextRange(hEdit, newStartSel, newStartSel + 1)
                ;MsgBox % startText
                if (startText = "* ") {
                    newStartSel := Edit_FindText(hEdit, "`r`n", newStartSel)
                    ;MsgBox % startSel
                    if (newStartSel > -1) {
                        startSel := newStartSel + 2
                    }
                } else {
                    break
                }
            }

            endSel := Edit_FindText(hEdit, "REMARKS?:|RECOMMENDATION:", , , "RegEx")  ; -1 if not found
            if (endSel > -1) {
                endSel -= 2
            }
            Edit_SetFocus(hEdit)
            Edit_SetSel(hEdit, startSel, endSel)
            ReorderSelectedText(false, false, "-", , hEdit)
        }
    } else {
        MsgBox("FINDING_CONTROL_HWND is invalid!")
    }
}

;;; Formatting FINDINGS
;;;; Reorder Seleted Text And Keep SeIm
SC079:: {
    OrderListForFindings()
}
^!,:: {
    OrderListForFindings()
}

;;; Formatting IMPRESSION
;;;; Reorder Seleted Text And Discard SeIm
FormatImpressionText() {
    if (hEdit := RisController.ImpressionEdit.NativeWindowHandle) {
        Edit_SetFocus(hEdit)
        Edit_SetSel(hEdit)
        if (Edit_CountNonEmptyLines(hEdit) > 1) {
            ReorderSelectedText(, , , , hEdit)
        } else {
            ReorderSelectedText(true, , , , hEdit)
        }
    }
}

SC070:: {
    FormatImpressionText()
}
^!.:: {
    FormatImpressionText()
}

; Reorder Seleted Text And Discard SeIm
^!o:: {
    hEdit := UIA.GetFocusedElement().NativeWindowHandle
    ReorderSelectedText(, , , , hEdit)
}

; Reorder Seleted Text And Keep SeIm
^!+o:: {
    hEdit := UIA.GetFocusedElement().NativeWindowHandle
    ReorderSelectedText(, , , false, hEdit)
}

; Unorder Seleted Text
^+*:: {
    hEdit := UIA.GetFocusedElement().NativeWindowHandle
    ReorderSelectedText(false, true, "*", , hEdit)
}

^+-:: {
    hEdit := UIA.GetFocusedElement().NativeWindowHandle
    ReorderSelectedText(false, true, "-", , hEdit)
}

^+>:: {
    hEdit := UIA.GetFocusedElement().NativeWindowHandle
    ReorderSelectedText(false, true, ">", , hEdit)
}

CopyPathologyReport() {
    try {
        reportText := ConvertRISDate(RisController.PathoDateText.Value) . ": " . RisController.PathoDiagnosisText.Value
        if (reportText != "") {
            A_Clipboard := reportText
            if !ClipWait(0.8) {
                throw Error("複製文字失敗 (逾時)。")
            }
            Notify("病理報告已複製到剪貼簿。")
        } else {
            throw Error("找不到病理報告內容。")
        }
    } catch as err {
        MsgBox("操作失敗: " . err.Message)
    }
}

^+c:: {
    CopyPathologyReport()
}

; 取得系統設定的滑鼠連點時間 (通常是 500ms)，讓判定更符合你的手感
DoubleClickTime := DllCall("GetDoubleClickTime")

~LButton:: {
    static clickCount := 0
    static lastClickTime := 0

    ; 計算當前點擊與上次點擊的時間差
    timeSinceLast := A_TickCount - lastClickTime

    ; 如果時間差在連點允許範圍內，增加計數，否則重置為 1
    if (timeSinceLast <= DoubleClickTime) {
        clickCount++
    } else {
        clickCount := 1
    }

    ; 更新最後點擊時間
    lastClickTime := A_TickCount

    ; 偵測到第三次點擊
    if (clickCount = 3) {
        ; 重置計數器，避免連續點第 4 下又觸發
        clickCount := 0

        ; 取得滑鼠下的 Control 資訊 (hCtrl 是控制項的 Handle/ID)
        MouseGetPos , , , &hCtrl, 2

        try {
            ; 取得 ClassNN 用於過濾
            classNN := ControlGetClassNN(hCtrl)

            ; 過濾條件：包含 "Edit" 且 不包含 "RichEdit"
            if (InStr(classNN, "Edit") && !InStr(classNN, "RichEdit")) {

                ; 呼叫自定義函數來選取邏輯行
                SelectLogicalLine(hCtrl)
            }
        }
    }
}

SelectLogicalLine(hCtrl) {
    ; 1. 取得 Control 內的全部文字
    try {
        fullText := ControlGetText(hCtrl)
    } catch {
        return ; 如果無法取得文字則放棄
    }

    if (fullText = "")
        return

    ; 2. 取得當前游標位置 (EM_GETSEL = 0x00B0)
    ; 這裡回傳的是一個 DWORD，低位元組是起始位置，我們只需要知道游標在哪即可
    caretPosRaw := SendMessage(0x00B0, 0, 0, hCtrl)
    caretPos := caretPosRaw & 0xFFFF ; 這是 0-based 的索引

    ; 3. 計算邏輯行的開始 (Start)
    ; AHK 的字串索引是 1-based，所以計算時要小心轉換
    ; InStr 尋找換行符號 `n (Line Feed)
    ; 從游標位置往前找 (參數 -1 代表反向搜尋)

    ; 轉換 caretPos 到 AHK 的 1-based 視角
    ahkCaretPos := caretPos + 1

    ; 往回找上一個換行符號的位置
    prevLineBreak := InStr(fullText, "`n", , ahkCaretPos, -1)

    ; 如果找到了換行，起始點應該是換行符號的"下一個字"
    ; 如果沒找到 (prevLineBreak 為 0)，代表在第一行，起始點就是 0 (0-based)
    selStart := (prevLineBreak == 0) ? 0 : prevLineBreak

    ; --- 計算 End (修改邏輯：包含換行符號) ---

    ; 找尋游標後的下一個 `r 或 `n
    nextR := InStr(fullText, "`r", , ahkCaretPos)
    nextN := InStr(fullText, "`n", , ahkCaretPos)

    selEnd := 0

    ; 狀況 1: 後面完全沒有換行符號 -> 選到文字最後
    if (nextR == 0 && nextN == 0) {
        selEnd := StrLen(fullText)
    }
    ; 狀況 2: 先遇到 `r (通常是 Windows 的 `r`n 結構)
    else if (nextR > 0 && (nextN == 0 || nextR < nextN)) {
        ; 檢查這個 `r 後面是不是緊接著 `n
        if (SubStr(fullText, nextR + 1, 1) == "`n") {
            ; 是 `r`n 結構，選取範圍要包含這兩個字元
            ; 數學計算：
            ; `r 在位置 nextR (例如 5)
            ; `n 在位置 nextR+1 (例如 6)
            ; 我們要選到 6 結束 (包含 0~5 共 6 個字元)
            selEnd := nextR + 1
        } else {
            ; 只有 `r (罕見，但也算換行)，選取範圍包含 `r
            selEnd := nextR
        }
    }
    ; 狀況 3: 先遇到 `n (Unix 格式換行)，選取範圍包含 `n
    else {
        selEnd := nextN
    }

    ; 發送選取指令
    SendMessage(0x00B1, selStart, selEnd, hCtrl)
}

#HotIf ; WinActive(RISReportWinTitle)

;; for JIS keyboard
SC029:: {
    try {
        if (!WinActive(RISReportWinTitle)) {
            WinActivate(RISReportWinTitle)
            WinWaitActive(RISReportWinTitle)
            focusedEle := UIA.GetFocusedElement()
            if (focusedEle.AutomationId != UIA_FindingEdit.AutomationId && focusedEle.AutomationId !=
                UIA_ImpressionEdit.AutomationId) {
                ControlFocus(RisController.FindingEdit.NativeWindowHandle)
            }
        } else {
            focusedEle := UIA.GetFocusedElement()
            if (focusedEle.AutomationId = UIA_FindingEdit.AutomationId) {
                ControlFocus(RisController.ImpressionEdit.NativeWindowHandle)
            } else {
                ControlFocus(RisController.FindingEdit.NativeWindowHandle)
            }
        }
    } catch TargetError as err {
        MsgBox "操作失敗: " err.Message
    }
}

UpdateRisElements() {
    global risEle, autoNextEle, reportSaveEle
    global FINDING_CONTROL_HWND, IMPRESSION_CONTROL_HWND, EXAMNAME_HWND
    global PAST_ALL_RADIO_HWND, PAST_MODALITY_RADIO_HWND, PAST_ONLY_MY_RADIO_HWND
    global PAST_FINDING_HWND, PAST_IMPRESSION_HWND

    try
    {
        risEle := UIA.ElementFromHandle(WinGetID(RISReportWinTitle))

        ele := risEle.FindFirst(UIA_FindingEdit)
        FINDING_CONTROL_HWND := IsObject(ele) ? ele.NativeWindowHandle : 0

        ele := risEle.FindFirst(UIA_ImpressionEdit)
        IMPRESSION_CONTROL_HWND := IsObject(ele) ? ele.NativeWindowHandle : 0

        ele := risEle.FindFirst(UIA_ExamNameTxt)
        EXAMNAME_HWND := IsObject(ele) ? ele.NativeWindowHandle : 0

        ele := risEle.FindFirst(UIA_PastAllRadio)
        PAST_ALL_RADIO_HWND := IsObject(ele) ? ele.NativeWindowHandle : 0

        ele := risEle.FindFirst(UIA_PastModalityRadio)
        PAST_MODALITY_RADIO_HWND := IsObject(ele) ? ele.NativeWindowHandle : 0

        ele := risEle.FindFirst(UIA_PastOnlyMyRadio)
        PAST_ONLY_MY_RADIO_HWND := IsObject(ele) ? ele.NativeWindowHandle : 0

        ele := risEle.FindFirst(UIA_PastReportFindingTxt)
        PAST_FINDING_HWND := IsObject(ele) ? ele.NativeWindowHandle : 0

        ele := risEle.FindFirst(UIA_PastReportImpressionTxt)
        PAST_IMPRESSION_HWND := IsObject(ele) ? ele.NativeWindowHandle : 0

        autoNextEle := risEle.FindFirst(UIA_AutoNextCheckbox)
        reportSaveEle := risEle.FindElement(UIA_ReportSaveButton)
    }
    catch as e {
        ; 如果發生錯誤 (例如視窗找不到), 我們的快取就會是 stale
        ;MsgBox("UIA Error in timer: " . e.Message) ; (可選: 除錯用)

        ; 7. 如果視窗或元素找不到, 或發生錯誤, 就清除 global 變數
        risEle := ""
    }
}

global simReportMap := Map(
    "CHEST PA/AP", Map("CHEST PA/AP+LAT", 1),
    "CHEST PA/AP+LAT", Map("CHEST PA/AP", 1),
    "KUB", Map("KUB+ABD LAT", 1),
    "KUB+L-SPINE LAT(supine)", Map("L-SPINE(AP+LAT)Standing", 1),
    "WHOLE  ABDOMEN CT WITH+ WITHOUT CONTRAST", Map("WHOLE  ABDOMEN CT WITHOUT CONTRAST", 1),
    "WHOLE  ABDOMEN CT WITHOUT CONTRAST", Map("WHOLE  ABDOMEN CT WITH+ WITHOUT CONTRAST", 1),
)

;UpdateRisElements()
;SetTimer(UpdateRisElements, 60000)

Notify(text, duration := 1500) {
    g := Gui("+AlwaysOnTop -Caption +ToolWindow +E0x20") ; E0x20 讓滑鼠穿透
    g.BackColor := "333333" ; 背景色
    g.SetFont("s16 cWhite bold", "微軟正黑體") ; 字型設定

    ; === 修正點：設定內距 (Padding) ===
    g.MarginX := 20 ; 左右留白 20px
    g.MarginY := 20 ; 上下留白 20px

    ; 新增文字 (不需額外設定位置，會自動置中)
    g.Add("Text", , text)

    ; 顯示 (AutoSize 會根據 Margin 自動調整視窗大小)
    g.Show("NoActivate AutoSize Center")

    ; 時間到自動銷毀
    SetTimer () => g.Destroy(), -duration
}

;-----------------------------------------------------------
; Mouse Remapping
#HotIf WinActive(RISReportWinTitle)

XButton1:: {
    ReorderSelectedText(false, true, "-")
}
XButton2:: {
    ReorderSelectedText()
}

#HotIf ; WinActive(RISReportWinTitle)

;-----------------------------------------------------------
#HotIf WinActive(RISCTMRAbnormalWinTitle)
!1:: {
    global ABNORMAL_VALUE_1_CONTROL
    ControlClick(ABNORMAL_VALUE_1_CONTROL)
}
!2:: {
    global ABNORMAL_VALUE_2_CONTROL
    ControlClick(ABNORMAL_VALUE_2_CONTROL)
}
!3:: {
    global ABNORMAL_VALUE_3_CONTROL
    ControlClick(ABNORMAL_VALUE_3_CONTROL)
}
!4:: {
    global ABNORMAL_VALUE_4_CONTROL
    ControlClick(ABNORMAL_VALUE_4_CONTROL)
}

#HotIf ; WinActive(RISCTMRAbnormalWinTitle)

;
; Global Remap
;
^!r:: Reload

#^p:: {
    ProcessClose("G3PACS.exe")
}

SC07B::LButton