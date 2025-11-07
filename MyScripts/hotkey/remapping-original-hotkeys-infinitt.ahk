;; for INFINITT PACS
#IfWinActive ahk_exe G3PACS.exe
w::
  ControlGetFocus, FocusedControl
  WinGetTitle, OutputVar
  ;MsgBox, "%OutputVar%"
  If (OutputVar = "INFINITT PACS" && SubStr(FocusedControl, 1, 3) == "Afx") {
    DiffSyncBtns := ["Button1", "Button85"]
    For idx, btn in DiffSyncBtns {
      ControlGetText, t, %btn%
      if (t = "自動同步") {
        ControlClick, %btn%
        Break
      }
    }
  } Else {
    Send, w
  }
Return

f::
  ControlGetFocus, FocusedControl
  WinGetTitle, OutputVar
;MsgBox, "%OutputVar%"
  If (OutputVar = "INFINITT PACS" && SubStr(FocusedControl, 1, 3) == "Afx") {
    DiffSyncBtns := ["Button2", "Button86", "Button91"]
    For idx, btn in DiffSyncBtns {
      ControlGetText, t, %btn%
      if (t = "不同檢查同步 ") {
        ControlClick, %btn%
        Break
      }
    }
  } Else {
    Send, f
  }
Return

e::
  ControlGetFocus, FocusedControl
  ;MsgBox, "%FocusedControl%"

  WinGetTitle, OutputVar
  If (OutputVar = "INFINITT PACS" && SubStr(FocusedControl, 1, 3) == "Afx") {
    DiffSyncBtns := ["Button4", "Button78"]
    For idx, btn in DiffSyncBtns {
      ControlGetText, t, %btn%
      if (t = " Scout lines") {
        ControlClick, %btn%
        Break
      }
    }
  } Else {
    Send, e
  }
Return

;; activate RIS and insert exam name
!e::
  global FINDING_CONTROL
  If (!WinActive("報告作業(frmRISReport)")) {
    WinActivate, 報告作業(frmRISReport)
    WinWaitActive, 報告作業(frmRISReport)
    ControlFocus, %FINDING_CONTROL%
    InsertExamname()
  }
Return
#IfWinActive  ; for INFINITT PACS
