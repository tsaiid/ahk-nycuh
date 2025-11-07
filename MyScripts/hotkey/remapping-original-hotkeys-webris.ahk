;
;
; Only for WebRIS
;
;
#IfWinActive WebRIS
#c::
  Send ^a
  Sleep 100
  Send ^c
Return

^w::
  Send ^{BS}
Return

#d::
  Send ^a
  Sleep 100
  Send {DEL}
Return

#a::
  Send ^a
Return

!d::
  Send ^!d
Return

!+d::
  Send ^!+d
Return

!p::
  Send ^!p
Return

;; Alt + E
;; Paste Examname
!e::
;  PasteExamname()
  Send ^!e
Return

;; Ctrl + Alt + Shift + E
;; Paste Examname with contrast informtion
^!+e::
;  PasteExamnameAndContrast()
  Send ^!f
Return

;; Insert Indication
;; Because Quill editor has a hotkey of Ctrl+I to italic
^i::
  Send ^!i
Return

^Esc::
;  CopyReportPath := "4.1.1.4.3.1.1.3.2.1.3.1.1.1.8.1.1.2.3.1"
;  btnObj := Acc_Get("Object", CopyReportPath, 0, "WebRIS")
;  btnObj.accDoDefaultAction(0)
  Send ^0
  Sleep, 500
  Send ^{Home}
Return

!Esc::
;  CopyReportPath := "4.1.1.4.3.1.1.3.2.1.3.1.1.1.8.1.1.2.3.1"
;  btnObj := Acc_Get("Object", CopyReportPath, 0, "WebRIS")
;  btnObj.accDoDefaultAction(0)
  Send ^9
Return

!q::
  Send {F4}   ; Quit without Save
Return

;; Remap Kana Key
;;; Formatting IMPRESSION
;;;; Reorder Seleted Text And Discard SeIm
SC070::
  ReorderSelectedText()
Return

;;; Formatting FINDINGS
;;;; Reorder Seleted Text And Keep SeIm
SC079::
  ReorderSelectedText(false, true, "-", false)
Return
#IfWinActive  ; end of WebRIS
