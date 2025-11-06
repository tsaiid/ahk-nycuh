; HotKey

#Include MyScripts\hotkey\remapping-original-hotkeys-webris.ahk
#Include MyScripts\hotkey\remapping-original-hotkeys-solidpacs.ahk
#Include MyScripts\hotkey\remapping-original-hotkeys-infinitt.ahk

;; For all RIS related window
#IfWinActive 報告作業(frmRISReport)
;; Ctrl + A
;; Go to start of line
^a::
  Send {Home}
Return

;; Ctrl + E
;; Go to end of line
^e::
  Send {End}
  /*
  ControlGetFocus, FocusedControl
  If (FocusedControl = "TMemo6" || FocusedControl = "TMemo7") {
    ControlGet, hEdit, Hwnd,, %FocusedControl%
    Edit_GetSel(hEdit, startSel)
    ;MsgBox % startSel
    l_text := Edit_GetTextRange(hEdit, startSel)
    ;MsgBox % l_text
    l_FoundPos:=InStr(l_Text, "`r`n")
    ;a:=Edit_FindText(hEdit, "`n")
    ;MsgBox % l_FoundPos
    If (l_FoundPos > 0) {
      endSel := startSel + l_FoundPos - 1
      Edit_SetSel(hEdit, endSel, endSel)
    } Else {
      endSel := Edit_GetTextLength(hEdit)
      Edit_SetSel(hEdit, endSel, -1)
    }
    ;p_LineIdx:=Edit_LineFromChar(hEdit,Edit_LineIndex(hEdit))
    ;l_StartSelPos:=Edit_LineIndex(hEdit,p_LineIdx)
    ;l_EndSelPos  :=l_StartSelPos+Edit_LineLength(hEdit,p_LineIdx)
    ;Edit_SetSel(hEdit,l_EndSelPos,l_EndSelPos)
  }
  */
Return

;; Ctrl + D
;; Delete a character
^d::
  Send {Del}
Return

^k::
  Send +{End}
  Send {Del}
Return

;^y::
;  Send {Home}
;  Send +{Down}
;  Send {Del}
;Return

;; Ctrl + Y
;; Delete current line
^y::
  global FINDING_CONTROL, IMPRESSION_CONTROL
  ControlGetFocus, FocusedControl
  If (FocusedControl = FINDING_CONTROL || FocusedControl = IMPRESSION_CONTROL) {
    ControlGet, hEdit, Hwnd,, %FocusedControl%
    ;p_LineIdx:=Edit_LineFromChar(hEdit,Edit_LineIndex(hEdit))
    ;l_StartSelPos:=Edit_LineIndex(hEdit,p_LineIdx)
    Edit_GetSel(hEdit, currStartSel)
    l_text := Edit_GetTextRange(hEdit, 0, currStartSel)
    l_FoundPos:=InStr(l_Text, "`r`n",, 0)
    If (l_FoundPos > 0) {
      startSel := l_FoundPos + 1
    } Else {
      startSel := 0
    }
    r_text := Edit_GetTextRange(hEdit, currStartSel, -1)
    r_FoundPos:=InStr(r_Text, "`r`n")
    If (r_FoundPos > 0) {
      endSel := currStartSel + r_FoundPos + 1
    } Else {
      endSel := -1
      ;MsgBox, %currStartSel% %r_FoundPos%
    }
    Edit_SetSel(hEdit, startSel, endSel)
    ;text_len := Edit_GetTextLength(hEdit)
    ;MsgBox, %startSel% %endSel% %text_len%
    Edit_Clear(hEdit)
  }
Return


#a::
  Send ^a
Return

#d::
  Send ^a
  Sleep 100
  Send {DEL}
Return

/*
!c::
  global SEND_REPORT_BTN_CONTROL
  NextReportChkboxPath := "4.1.4.1.4.1.4.1.4.2.4.6.4"
  action := Acc_Get("Action", NextReportChkboxPath, 0, ahk_exe XRay.exe)
  if (action == "取消核取") {
      chkBoxObj := Acc_Get("Object", NextReportChkboxPath, 0, ahk_exe XRay.exe)
      chkBoxObj.accDoDefaultAction(0)
  }
  ;ControlGet, isChecked, Checked, , WindowsForms10.BUTTON.app.0.2780b98_r24_ad13, ahk_exe XRay.exe
  ;ControlGet, isChecked, Checked, , %NEXT_REPORT_CHECKBOX_CONTROL%
  ;ControlGet, isChecked, Choice, , %NEXT_REPORT_CHECKBOX_CONTROL%
  ;ControlGetText, isChecked, %NEXT_REPORT_CHECKBOX_CONTROL%
  ;MsgBox, %isChecked%
  ;    ControlClick %NEXT_REPORT_CHECKBOX_CONTROL%
  ;if (isChecked = 1) {
      ;ControlClick %NEXT_REPORT_CHECKBOX_CONTROL%
  ;}
  ;Send ^s
  ControlClick, %SEND_REPORT_BTN_CONTROL%
Return

^s::
  global SEND_REPORT_BTN_CONTROL
  NextReportChkboxPath := "4.1.4.1.4.1.4.1.4.2.4.6.4"
  action := Acc_Get("Action", NextReportChkboxPath, 0, ahk_exe XRay.exe)
  if (action == "核取") {
      chkBoxObj := Acc_Get("Object", NextReportChkboxPath, 0, ahk_exe XRay.exe)
      chkBoxObj.accDoDefaultAction(0)
  }
  ControlClick, %SEND_REPORT_BTN_CONTROL%
  ;Send ^s
Return
*/


!1::
  global ABNORMAL_VALUE_1_CONTROL
  If (WinActive("檢查結果(frmPos)")) {
    ControlClick %ABNORMAL_VALUE_1_CONTROL%
  }
Return

!2::
  global ABNORMAL_VALUE_2_CONTROL
  If (WinActive("檢查結果(frmPos)")) {
    ControlClick %ABNORMAL_VALUE_2_CONTROL%
  }
Return

!3::
  global ABNORMAL_VALUE_3_CONTROL
  If (WinActive("檢查結果(frmPos)")) {
    ControlClick %ABNORMAL_VALUE_3_CONTROL%
  }
Return

!4::
  global ABNORMAL_VALUE_4_CONTROL
  If (WinActive("檢查結果(frmPos)")) {
    ControlClick %ABNORMAL_VALUE_4_CONTROL%
  }
Return

GetExamnameFromRIS()
{
  global EXAMNAME_CONTROL
  ControlGetText, t, %EXAMNAME_CONTROL%
  ;MsgBox, %t%
  examname := StrReplace(t, "檢查項目: ", "")
  Return examname
}

InsertExamname()
{
  global FINDING_CONTROL, IMPRESSION_CONTROL
  If WinActive("報告作業(frmRISReport)") {
    ControlGetFocus, FocusedControl
    If (FocusedControl = FINDING_CONTROL || FocusedControl = IMPRESSION_CONTROL) {
      ControlGet, hFindingEdit, Hwnd,, %FINDING_CONTROL%
      Edit_GetSel(hFindingEdit, currStartSel, currEndSel)
      examname := GetExamnameFromRIS()
      If (examname) {
        examname_text := Examname . ":`r`n`r`n"
        Edit_SetText(hFindingEdit, examname_text . Edit_GetText(hFindingEdit))
        newStartSel := currStartSel + StrLen(examname_text)
        newEndSel := currEndSel + StrLen(examname_text)
        Edit_SetSel(hFindingEdit, newStartSel, newEndSel)
      } Else {
        MsgBox, No examname found.
      }
    }
  }
}


;; Insert Exam Name
!e::
  InsertExamname()
  ;ExamNamePath := "4.1.4.1.4.2.4.1.4.8.4"
  ;examname := Acc_Get("Value", ExamNamePath, 0, ahk_exe XRay.exe)
  ;examname := GetExamnameFromRIS()
  ;Paste(examname . ":`n`n")
  ;Paste(examname)

  ;ControlGetText, t, WindowsForms10.EDIT.app.0.2780b98_r24_ad114, 報告作業(frmRISReport)
  ;ControlGetText, t, WindowsForms10.EDIT.app.0.2780b98_r24_ad114
  ;t := GetExamnameFromRIS()
  ;msgbox, %t%
Return

;; Remap Kana Key
;;; Formatting IMPRESSION
;;;; Reorder Seleted Text And Discard SeIm
SC070::
  ;ReorderSelectedText()
  ControlGet, hEdit, Hwnd,, %IMPRESSION_CONTROL%
  impStr := Edit_GetText(hEdit)
  If (StrLen(impStr)) {
    impStrArr := StrSplit(impStr, "`r`n")
    CRLFCount := 0
    For i, v in impStrArr {
      If (StrLen(v)) {
        CRLFCount++
      }
    }
    Edit_SetFocus(hEdit)
    Edit_SetSel(hEdit)
    If (CRLFCount > 1) {
      ReorderSelectedText()
    } Else {
      ReorderSelectedText(true)
    }
  }
Return

GetExamTypeFromRIS()
{
  examname := GetExamnameFromRIS()
  If (InStr(examname, "CT")) {
    Return "CT"
  } Else If (InStr(examname, "MR")) {
    Return "MR"
  } Else If (InStr(examname, "US")) {
    Return "US"
  }
  Return "CR"
}

OrderListForFindings()
{
  examtype := GetExamTypeFromRIS()
  ;MsgBox, "%examtype%"
  Switch examtype
  {
    case "CT", "MR":
      UnorderListForFindingsOfCtOrMr()

    case "CR", "US":
      UnorderListForFindingsOfCrUs()
  }
}

UnorderListForFindingsOfCrUs()
{
  examtype := GetExamTypeFromRIS()
  ;MsgBox, %examtype%
  If (examtype = "CR" || examtype = "US") {
    global FINDING_CONTROL
    ;ControlFocus, %FINDING_CONTROL%  ; need to get focus. ReorderSelectedText() use focused edit.
    ControlGet, hEdit, Hwnd,, %FINDING_CONTROL%
    ;startSel := Edit_FindText(hEdit, "FINDINGS:`r`n", , , "RegEx", matchedText)
    ;If (startSel == -1) {
      startSel := Edit_FindText(hEdit, "FINDINGS:`r`n|:\s*`r`n\s*`r`n", , , "RegEx", matchedText)
    ;}
    If (startSel > -1) {
      ;startSel += 5
      startSel += StrLen(matchedText)
      ;MsgBox % startSel
      Edit_SetFocus(hEdit)
      Edit_SetSel(hEdit, startSel, -1)
      ;Sleep, 10
      ReorderSelectedText(false, true, "-", false)
    }
  }
}

UnorderListForFindingsOfCtOrMr()
{
  examtype := GetExamTypeFromRIS()
  If (examtype = "CT" || examtype = "MR") {
    ControlGet, hEdit, Hwnd,, %FINDING_CONTROL%
    startSel := Edit_FindText(hEdit, "FINDINGS:`r`n|The study shows:`r`n`r`n|show the following findings:`r`n`r`n|which revealed:`r`n`r`n", , , "RegEx", matchedText)
    ;startSel := Edit_FindText(hEdit, "(FINDINGS:`r`n|COMPARISON:)", , , "RegEx", matchedText)
    If (startSel > -1) {
      ;startSel += 11
      startSel += StrLen(matchedText)
      Loop, 3 {
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
      ReorderSelectedText(false, true, "-")
      ;MsgBox % regex_out
    }
  }
}

;;; Formatting FINDINGS
;;;; Reorder Seleted Text And Keep SeIm
SC079::
  OrderListForFindings()
Return

;;; Append previous report to FINDINGS
^ESC::
  ;global FINDING_CONTROL
  ;PrevReportControl := "WindowsForms10.RichEdit20W.app.0.2780b98_r24_ad16"
  ;ControlFocus, %FindingControl%  ; need to get focus. ReorderSelectedText() use focused edit.
  ControlGet, hEdit, Hwnd,, %PREV_REPORT_CONTROL%
  if (hEdit) {
    prevReportTxt := Edit_GetText(hEdit)
    ;MsgBox, %prevReportTxt%
    ControlGet, hEdit, Hwnd,, %FINDING_CONTROL%
    if (hEdit) {
      ControlFocus, , ahk_id %hEdit%
      ; Append to finding
      Edit_SetSel(hEdit, Edit_GetTextLength(hEdit))
      Edit_ReplaceSel(hEdit, prevReportTxt)
      Edit_SetSel(hEdit, 0, 0)
    }
  }
Return

FindPrevCRLF(text) {
  found_pos := InStr(text, "`r`n",, 0)
  If (found_pos > 0) {
    found_pos := found_pos + 1
  } Else {
    found_pos := 0
  }
  Return found_pos
}

FindPrevText(text_to_find, needle_text, start_pos) {
  found_pos_space := InStr(text_to_find, needle_text,, 0)
  If (found_pos_space > 0) {
    If (found_pos_space = start_pos) {
      sub_text := SubStr(text_to_find, 1, found_pos_space - 1)
      found_pos_space := FindPrevText(sub_text, needle_text, found_pos_space - 1)
    }
  }

  ; should not cross to previous line
  found_pos_crlf := FindPrevCRLF(text_to_find)
  If (found_pos_crlf >= found_pos_space) {
    found_pos_space := found_pos_crlf
  }

  Return found_pos_space
}

;; Ctrl + W
;; delete previous word
^w::
  global FINDING_CONTROL, IMPRESSION_CONTROL
  ControlGetFocus, FocusedControl
  If (FocusedControl = FINDING_CONTROL || FocusedControl = IMPRESSION_CONTROL) {
    ControlGet, hEdit, Hwnd,, %FocusedControl%
    Edit_GetSel(hEdit, currStartSel)
      ;MsgBox % currStartSel
    If (currStartSel > 0) { ; if at the beginning of text, do nothing
      l_text := Edit_GetTextRange(hEdit, 0, currStartSel - 1)
      l_FoundPos := FindPrevText(l_text, " ", currStartSel)
      ;MsgBox, %currStartSel% %l_FoundPos%
      Edit_SetSel(hEdit, l_FoundPos, currStartSel)
      Edit_Clear(hEdit)
    }
  }
Return

/*

;; Insert Selected Prev Exam Date
!d::
  global FINDING_CONTROL, IMPRESSION_CONTROL, PREV_REPORT_TABLE_PATH
  ControlGetFocus, FocusedControl
  If (FocusedControl = FINDING_CONTROL || FocusedControl = IMPRESSION_CONTROL) {
    ;AccPathToTable := "4.1.4.2.4.1.4.1.4.6.4.1.4"
    oTableAcc := Acc_Get("Object", PREV_REPORT_TABLE_PATH, 0, "報告作業(frmRISReport)")
    if IsObject(oTableAcc) {
      RowCount := oTableAcc.accChildCount
      if (RowCount < 1) {
          MsgBox, 表格有 %RowCount% 行，沒有資料。
          return
      }
      str := ""
      Loop, % RowCount
      {
          RowIndex := A_Index + 0
          oRowAcc := oTableAcc.accChild(RowIndex) ; 獲取「行」物件
          CurrentState := oRowAcc.accState(0)
          STATE_SYSTEM_SELECTED := 0x2
          str .= CurrentState . "`n"
          If (CurrentState & STATE_SYSTEM_SELECTED) {
            oFirstCell := oRowAcc.accChild(1)
            FirstCellText := oFirstCell.accValue(0)
            MsgBox, % FirstCellText
            Paste(ConvertRISDate(FirstCellText))
            break
          }
      }
      ;MsgBox, % str
    }
  }
Return

*/

ClickOnSelectedRow(){
    oTableAcc := Acc_Get("Object", PREV_REPORT_TABLE_PATH, 0, "報告作業(frmRISReport)")
    if IsObject(oTableAcc) {
      RowCount := oTableAcc.accChildCount
      if (RowCount < 1) {
          MsgBox, 表格有 %RowCount% 行，沒有資料。
          return
      }
      Loop, % RowCount
      {
          RowIndex := A_Index + 0
          oRowAcc := oTableAcc.accChild(RowIndex) ; 獲取「行」物件
          CurrentState := oRowAcc.accState(0)
          STATE_SYSTEM_SELECTED := 0x2
          str .= CurrentState . "`n"
          ;If (CurrentState & STATE_SYSTEM_SELECTED) {
          If (RowIndex = 3) {
            oFirstCell := oRowAcc.accChild(1)
            oFirstCell.accDoDefaultAction(0)
            ;oRowAcc.accDoDefaultAction(0)
            ;MsgBox, % oFirstCell.accValue(0)
            break
          }
      }
    }
}

/*
SC07B::
  global PREV_REPORT_TABLE_CONTROL
  ControlGetFocus, FocusedControl
  ;MsgBox, %FocusedControl%
  If (FocusedControl = PREV_REPORT_TABLE_CONTROL) {
    ClickOnSelectedRow()
  } Else {
    Click
  }
Return
*/
#IfWinActive  ; end of ahk_group RIS



;
; Global Remap
;
#^p::
  Process, Close, G3PACS.exe
Return

;; for JIS keyboard
SC029::
  global FINDING_CONTROL, IMPRESSION_CONTROL
  If (!WinActive("報告作業(frmRISReport)")) {
    WinActivate, 報告作業(frmRISReport)
    WinWaitActive, 報告作業(frmRISReport)
    ;ControlFocus, %FINDING_CONTROL%
  }
  ControlGetFocus, FocusedControl
  ;MsgBox, %FocusedControl%
  If (FocusedControl = FINDING_CONTROL) {
    ControlFocus, %IMPRESSION_CONTROL%
  } Else {
    ControlFocus, %FINDING_CONTROL%
  }
Return

SC07B::LButton

#x::^x

;; for global windows environment
#Space::
  SendEvent ^{Space}  ; Need to send event to work in VirtualBox
Return






;
;
; Probably useless?
;
;
CopyPidAndLaunchPacsWorklist()
{
  ;MsgBox, 1
  Clipboard := ""
  Send, ^c
  ClipWait, 2
  if ErrorLevel
  {
    MsgBox, The attempt to copy text onto the clipboard failed.
    return
  }
  ;load patient in INFINITT PACS
  pacs_api =
  ( LTrim Join
    http://10.2.2.30/pkg_pacs/external_interface.aspx?
      TYPE=W&
      LID=A60076&
      LPW=A60076&
      PID=%Clipboard%
  )
  ;MsgBox %pacs_api%
  Run msedge.exe --app=%pacs_api%
}

;; for 檢查排程系統
#IfWinActive ahk_exe ExmSchSys.EXE
#c::
  CopyPidAndLaunchPacsWorklist()
Return
#IfWinActive  ; for ahk_exe ExmSchSys.EXE

;; for LibreOffice Calc
#IfWinActive - LibreOffice Calc
#c::
  CopyPidAndLaunchPacsWorklist()
Return
#IfWinActive  ; for LibreOffice Calc

