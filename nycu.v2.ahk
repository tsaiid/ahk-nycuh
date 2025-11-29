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

; --- Findings Formatting ---
SC079:: RisController.FormatFindingText() ; 日文鍵盤的轉換鍵?
^!,::   RisController.FormatFindingText()

; --- Impression Formatting ---
SC070:: RisController.FormatImpressionText() ; 日文鍵盤的無變換鍵?
^!.::   RisController.FormatImpressionText()

; --- Selection Reordering (Manual) ---
; 重排選取文字 (預設：自動編號 1. 2. 3.)
^!o:: RisController.ReorderSelection()

; 重排選取文字 (保留 Series/Image 標記)
^!+o:: RisController.ReorderSelection({discardSeIm: false})

; 移除編號，改用 "*"
^+*:: RisController.ReorderSelection({deOrder: false, keepEmpty: true, itemChar: "*"})

; 移除編號，改用 "-"
^+-:: RisController.ReorderSelection({deOrder: false, keepEmpty: true, itemChar: "-"})

; 移除編號，改用 ">"
^+>:: RisController.ReorderSelection({deOrder: false, keepEmpty: true, itemChar: ">"})

; --- Pathology Copy ---
^!c:: RisController.CopyPathologyReport()

; --- Triple Click Handler ---
; 讓滑鼠左鍵通過，同時觸發連點檢查
~LButton:: RisController.HandleTripleClick()

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