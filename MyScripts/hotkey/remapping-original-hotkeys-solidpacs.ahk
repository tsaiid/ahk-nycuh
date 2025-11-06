;
; Solid PACS Viewer
;
#IfWinActive ahk_exe WEBVIE~1.EXE
;; zoom in/out
^WheelUp::
  Send {NumpadAdd}
Return
^WheelDown::
  Send {NumpadSub}
Return

;; activate WebRIS and copy report
^Esc::
  If (!WinActive("WebRIS")) {
    WinActivate, WebRIS
    Send ^0
  }
Return

;; activate WebRIS and insert exam name
!e::
  If (!WinActive("WebRIS")) {
    WinActivate, WebRIS
    Send ^!e
  }
Return
#IfWinActive  ; for ahk_exe WEBVIE~1.EXE
