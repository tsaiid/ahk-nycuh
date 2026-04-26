#Requires AutoHotkey v2.0
A_MaxHotkeysPerInterval := 200

;; for INFINITT PACS
#HotIf MouseIsOverG3Pacs()
$WheelUp::FocusG3PacsUnderMouseAndScroll("WheelUp")
$WheelDown::FocusG3PacsUnderMouseAndScroll("WheelDown")
#HotIf

MouseIsOverG3Pacs() {
    MouseGetPos(,, &hwnd)
    try return WinGetProcessName("ahk_id " hwnd) = "G3PACS.exe"
    return false
}

FocusG3PacsUnderMouseAndScroll(direction) {
    MouseGetPos(,, &hwnd)
    if (hwnd && !WinActive("ahk_id " hwnd)) {
        WinActivate("ahk_id " hwnd)
    }
    Click(direction)
}

#HotIf WinActive("ahk_exe G3PACS.exe")
w::
{
    try {
        ; v2 的 ControlGetFocus 回傳 HWND，需轉為 ClassNN 才能做字串比對
        hCtl := ControlGetFocus("A")
        FocusedControl := ControlGetClassNN(hCtl)
    } catch {
        FocusedControl := ""
    }

    OutputVar := WinGetTitle("A")
    ;MsgBox(OutputVar)

    if (OutputVar = "INFINITT PACS" && SubStr(FocusedControl, 1, 3) = "Afx") {
        DiffSyncBtns := ["Button1", "Button85", "Button90", "Button102"]
        for idx, btn in DiffSyncBtns {
            try {
                t := ControlGetText(btn)
                if (t == " Auto sync" || t == "自動同步") {
                    ControlClick(btn)
                    break
                }
            }
        }
    } else {
        Send("w")
    }
}

f::
{
    try {
        hCtl := ControlGetFocus("A")
        FocusedControl := ControlGetClassNN(hCtl)
    } catch {
        FocusedControl := ""
    }

    OutputVar := WinGetTitle("A")
    ;MsgBox(OutputVar)

    if (OutputVar = "INFINITT PACS" && SubStr(FocusedControl, 1, 3) = "Afx") {
        DiffSyncBtns := ["Button2", "Button86", "Button91", "Button103"]
        for idx, btn in DiffSyncBtns {
            try {
                t := ControlGetText(btn)
                if (t == " Sync with other exams" || t == "不同檢查同步 ") {
                    ControlClick(btn)
                    break
                }
            }
        }
    } else {
        Send("f")
    }
}

e::
{
    try {
        hCtl := ControlGetFocus("A")
        FocusedControl := ControlGetClassNN(hCtl)
    } catch {
        FocusedControl := ""
    }
    ;MsgBox(FocusedControl)

    OutputVar := WinGetTitle("A")
    if (OutputVar = "INFINITT PACS" && SubStr(FocusedControl, 1, 3) = "Afx") {
        DiffSyncBtns := ["Button4", "Button78"]
        for idx, btn in DiffSyncBtns {
            try {
                t := ControlGetText(btn)
                if (t = " Scout lines") { ; 注意這裡原代碼有一個前導空白
                    ControlClick(btn)
                    break
                }
            }
        }
    } else {
        Send("e")
    }
}

/*
;; activate RIS and insert exam name
!e::
{
    global FINDING_CONTROL ; 引用外部全域變數
    if (!WinActive("報告作業(frmRISReport)")) {
        ; 檢查視窗是否存在，避免報錯
        if (WinExist("報告作業(frmRISReport)")) {
            WinActivate("報告作業(frmRISReport)")
            WinWaitActive("報告作業(frmRISReport)")

            ; 確保 FINDING_CONTROL 已定義且控制項存在
            if IsSet(FINDING_CONTROL) {
                try ControlFocus(FINDING_CONTROL)
            }

            ; 假設 InsertExamname() 是一個已定義的函數
            try InsertExamname()
        }
    }
}
*/
#HotIf ; end for INFINITT PACS
