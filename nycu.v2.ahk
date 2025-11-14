#Requires AutoHotkey v2.0+

; === 解決 DPI 座標偏移問題 ===
; 1. 宣告 DPI 感知 (讓 BoundingRectangle 回報實體座標)
DllCall("SetThreadDpiAwarenessContext", "ptr", -3, "ptr")
; 2. 確保 AHK 的 Click 使用螢幕座標
CoordMode "Mouse", "Screen"
; ================================
SetMouseDelay -1
SetKeyDelay -1

Global PRESERVE_CLIPBOARD := 0

#Include <UIA.v2>
#Include <Paste.v2>
#Include <Edit.v2>

Global RISReportWinTitle := "報告作業(frmRISReport)" ; 替換成您的程式標題
Global RISCTMRAbnormalWinTitle := "檢查結果(frmPos)"
Global UIA_PastReportTable := {AutomationId: "dgvPastReport"}
Global UIA_AutoNextCheckbox := {AutomationId: "chkAutoNext"}
Global UIA_ReportSaveButton := {AutomationId: "btnReportSave"}
Global UIA_ExamNameTxt := {AutomationId: "txtExamName"}
Global UIA_FindingEdit := {AutomationId: "txtReport"}
Global UIA_ImpressionEdit := {AutomationId: "txtImpression"}
Global UIA_PastAllRadio := {AutomationId: "rdoPastALL"}
Global UIA_PastModalityRadio := {AutomationId: "rdoClassify"}
Global UIA_PastOnlyMyRadio := {AutomationId: "rdoPastOnlyMy"}
Global UIA_PastReportFindingTxt := {AutomationId: "rtxtPastReport"}
Global UIA_PastReportImpressionTxt := {AutomationId: "rtxtPastImpression"}
Global UIA_PathoDiagnosisTxt := {AutomationId: "txtDiagnosist"}
Global UIA_PathoDateTxt := {AutomationId: "mtxtRcpDTM"}
Global FINDING_CONTROL := "WindowsForms10.EDIT.app.0.2780b98_r24_ad12"
Global IMPRESSION_CONTROL := "WindowsForms10.EDIT.app.0.2780b98_r24_ad11"
Global ABNORMAL_VALUE_1_CONTROL := "WindowsForms10.BUTTON.app.0.2780b98_r24_ad13"
Global ABNORMAL_VALUE_2_CONTROL := "WindowsForms10.BUTTON.app.0.2780b98_r24_ad15"
Global ABNORMAL_VALUE_3_CONTROL := "WindowsForms10.BUTTON.app.0.2780b98_r24_ad16"
Global ABNORMAL_VALUE_4_CONTROL := "WindowsForms10.BUTTON.app.0.2780b98_r24_ad14"
Global ABNORMAL_VALUE_SAVE_BUTTON_CONTROL := "WindowsForms10.BUTTON.app.0.2780b98_r24_ad12"
Global EXAMNAME_HWND := 0
Global FINDING_CONTROL_HWND := 0
Global IMPRESSION_CONTROL_HWND := 0
Global PAST_ALL_RADIO_HWND := 0
Global PAST_MODALITY_RADIO_HWND := 0
Global PAST_ONLY_MY_RADIO_HWND := 0
Global PAST_FINDING_HWND := 0
Global PAST_IMPRESSION_HWND := 0

#HotIf WinActive(RISReportWinTitle)
^9::{
}

^1::{
  global PAST_ALL_RADIO_HWND
  if (PAST_ALL_RADIO_HWND) {
    ControlClick(PAST_ALL_RADIO_HWND)
  }
}
^2::{
  global PAST_MODALITY_RADIO_HWND
  if (PAST_MODALITY_RADIO_HWND) {
    ControlClick(PAST_MODALITY_RADIO_HWND)
  }
}
^3::{
  global PAST_ONLY_MY_RADIO_HWND
  if (PAST_ONLY_MY_RADIO_HWND) {
    ControlClick(PAST_ONLY_MY_RADIO_HWND)
  }
}
;;; Append previous report to FINDINGS and IMPRESSION
AppendPrevReport() {
  global PAST_FINDING_HWND, PAST_IMPRESSION_HWND
  global FINDING_CONTROL_HWND, IMPRESSION_CONTROL_HWND
  ;winEle := UIA.ElementFromHandle(WinGetID(RISReportWinTitle))
  ;winEle := risEle
  ;pFdEle := winEle.FindFirst(UIA_PastReportFindingTxt)
  if (hEdit := PAST_FINDING_HWND) {
    pastFinding := Edit_GetText(hEdit)
    ;fdEle := winEle.FindFirst(UIA_FindingEdit)
    if (hFdEdit := FINDING_CONTROL_HWND) {
      Edit_SetSel(hFdEdit, Edit_GetTextLength(hFdEdit))
      Edit_ReplaceSel(hFdEdit, pastFinding)
      Edit_SetSel(hFdEdit, 0, 0)
      Edit_ScrollCaret(hFdEdit)
    }
  }

  ;pImpEle := winEle.FindFirst(UIA_PastReportImpressionTxt)
  if (hEdit := PAST_IMPRESSION_HWND) {
    pastImpression := Edit_GetText(hEdit)
    ;impEle := winEle.FindFirst(UIA_ImpressionEdit)
    if (hImpEdit := IMPRESSION_CONTROL_HWND) {
      Edit_SetSel(hImpEdit, Edit_GetTextLength(hImpEdit))
      Edit_ReplaceSel(hImpEdit, pastImpression)
    }
  }
}
^ESC::{
  AppendPrevReport()
}
;; Ctrl + Y
;; Delete current line
^y::{
  ;local cacheRequest := UIA.CreateCacheRequest(["AutomationId", "NativeWindowHandle"])
  ;local focusedEle := UIA.GetFocusedElement(cacheRequest)
  local focusedEle := UIA.GetFocusedElement()
  ;if (focusedEle.CachedAutomationId = UIA_FindingEdit.AutomationId || focusedEle.CachedAutomationId = UIA_ImpressionEdit.AutomationId) {
  if (focusedEle.AutomationId = UIA_FindingEdit.AutomationId || focusedEle.AutomationId = UIA_ImpressionEdit.AutomationId) {
    ;hEdit := focusedEle.CachedNativeWindowHandle
    hEdit := focusedEle.NativeWindowHandle
    Edit_GetSel(hEdit, &currStartSel)
    l_text := Edit_GetTextRange(hEdit, 0, currStartSel)
    l_FoundPos := InStr(l_Text, "`r`n",, -1)
    if (l_FoundPos > 0) {
      startSel := l_FoundPos + 1
    } else {
      startSel := 0
    }
    r_text := Edit_GetTextRange(hEdit, currStartSel, -1)
    r_FoundPos := InStr(r_Text, "`r`n")
    if (r_FoundPos > 0) {
      endSel := currStartSel + r_FoundPos + 1
    } else {
      endSel := -1
      ;MsgBox, %currStartSel% %r_FoundPos%
    }
    Edit_SetSel(hEdit, startSel, endSel)
    ;text_len := Edit_GetTextLength(hEdit)
    ;MsgBox, %startSel% %endSel% %text_len%
    Edit_Clear(hEdit)
  }
}
FindPrevCRLF(text) {
  found_pos := InStr(text, "`r`n",, -1)
  if (found_pos > 0) {
    found_pos := found_pos + 1
  } else {
    found_pos := 0
  }
  return found_pos
}

FindPrevText(text_to_find, needle_text, start_pos) {
  found_pos_space := InStr(text_to_find, needle_text,, -1)
  if (found_pos_space > 0) {
    if (found_pos_space = start_pos) {
      sub_text := SubStr(text_to_find, 1, found_pos_space - 1)
      found_pos_space := FindPrevText(sub_text, needle_text, found_pos_space - 1)
    }
  }

  ; should not cross to previous line
  found_pos_crlf := FindPrevCRLF(text_to_find)
  if (found_pos_crlf >= found_pos_space) {
    found_pos_space := found_pos_crlf
  }

  return found_pos_space
}
;; Ctrl + W
;; delete previous word
^w::{
  focusedEle := UIA.GetFocusedElement()
  if (focusedEle.AutomationId = UIA_FindingEdit.AutomationId || focusedEle.AutomationId = UIA_ImpressionEdit.AutomationId) {
    hEdit := focusedEle.NativeWindowHandle
    Edit_GetSel(hEdit, &currStartSel)
    If (currStartSel > 0) { ; if at the beginning of text, do nothing
      l_text := Edit_GetTextRange(hEdit, 0, currStartSel - 1)
      l_FoundPos := FindPrevText(l_text, " ", currStartSel)
      ;MsgBox, %currStartSel% %l_FoundPos%
      Edit_SetSel(hEdit, l_FoundPos, currStartSel)
      Edit_Clear(hEdit)
    }
  }
}


InsertExamname() {
  ;winEle := UIA.ElementFromHandle(WinGetID(RISReportWinTitle))
  ;winEle := risEle
  ;fdEle := winEle.FindFirst(UIA_FindingEdit)
  ;impEle := winEle.FindFirst(UIA_ImpressionEdit)
  focusedEle := UIA.GetFocusedElement()
  if (focusedEle.AutomationId = UIA_FindingEdit.AutomationId || focusedEle.AutomationId = UIA_ImpressionEdit.AutomationId) {
    hEdit := focusedEle.NativeWindowHandle
    currStartSel := 0
    currEndSel := 0
    Edit_GetSel(hEdit, &currStartSel, &currEndSel)
    examname_text := GetCurrExamName() . ":`r`n`r`n"
    Edit_SetText(hEdit, examname_text . Edit_GetText(hEdit))
    newStartSel := currStartSel + StrLen(examname_text)
    newEndSel := currEndSel + StrLen(examname_text)
    Edit_SetSel(hEdit, newStartSel, newEndSel)
    ;MsgBox(examname)
  }
}

;; Insert Exam Name
!e::{
  InsertExamname()
}
!c::{
  ;MsgBox("Auto Next & Save Report")
  CheckNextAuto(false)
  SaveReport()
}

^s::{
  CheckNextAuto(true)
  SaveReport()
}

CheckNextAuto(checked := true){
  global autoNextEle
  if (!IsObject(autoNextEle)) {
    try {
        global risEle
        winEle := risEle
        if !IsObject(winEle)
            throw Error("找不到視窗: " . RISReportWinTitle)

        autoNextEle := winEle.FindFirst(UIA_AutoNextCheckbox)
        if !IsObject(autoNextEle)
            throw Error("找不到 '下一筆自動勾選框' 物件！`n請檢查您的 UIA_AutoNextCheckbox 查詢條件。`n`n目前條件: " . UIA_AutoNextCheckbox)

    }
    catch as e {
        MsgBox("UIA 發生錯誤:`n" . e.Message . "`n`n行: " . e.Line, "UIA Error", 16)
    }
  }

  if (checked ^ autoNextEle.ToggleState) {
    autoNextEle.Toggle()
  }
  ;MsgBox(autoNextEle.ToggleState)
}

SaveReport(){
  try {
      winEle := risEle
      if !IsObject(winEle)
          throw Error("找不到視窗: " . RISReportWinTitle)

      saveBtnEle := winEle.FindFirst(UIA_ReportSaveButton)
      if !IsObject(saveBtnEle)
          throw Error("找不到 '報告存檔按鈕' 物件！`n請檢查您的 UIA_ReportSaveButton 查詢條件。`n`n目前條件: " . UIA_ReportSaveButton)

      saveBtnEle.Click()
  }
  catch as e {
      MsgBox("UIA 發生錯誤:`n" . e.Message . "`n`n行: " . e.Line, "UIA Error", 16)
  }
}

!ESC::{
  examName := GetCurrExamName()
  FindSimilarReport(examName)
  ;MsgBox(GetCurrExamName())
}

GetCurrExamName() {
  Global EXAMNAME_HWND
    ;MsgBox(EXAMNAME_HWND)
  if (EXAMNAME_HWND) {
    examName := StrReplace(ControlGetText(EXAMNAME_HWND), "檢查項目: ", "")
    ;MsgBox(examName)
    return examName
  } else {
    MsgBox(EXAMNAME_HWND . " is invalid!")
  }
  try {
      winEle := risEle
      if !IsObject(winEle)
          throw Error("找不到視窗: " . RISReportWinTitle)

      examNameEle := winEle.FindFirst(UIA_ExamNameTxt)
      if !IsObject(examNameEle)
          throw Error("找不到 '檢查名稱' 輸入框！")

      examName := StrReplace(examNameEle.Value, "檢查項目: ", "")
      return examName
  }
  catch as e {
      MsgBox("UIA 發生錯誤:`n" . e.Message . "`n`n行: " . e.Line, "UIA Error", 16)
      return ""
  }
}

isSameExam(prevExamName, currExamName) {
  return (prevExamName = currExamName)
}
isSimilarExam(prevExamName, currExamName) {
  global simReportMap
  if !simReportMap.Has(currExamName)
    return false

  similarExams := simReportMap[currExamName]
  return similarExams.Has(prevExamName)
}
isRelatedReport(prevExamName, currExamName) {
  return isSameExam(prevExamName, currExamName) || isSimilarExam(prevExamName, currExamName)
}

FindSimilarReport(examName := "")
{
    ; --- 設定搜尋目標 ---
    ;Local SearchText := "CHEST PA/AP"
    Local SearchText := examName
    Local SearchColumnIndex := 3 ; 1=簽收日, 2=儀器, 3=檢查項目

    try
    {
        ; 2. 獲取視窗元素
        winEle := UIA.ElementFromHandle(WinGetID(RISReportWinTitle))
        if !IsObject(winEle)
            throw Error("找不到視窗: " . RISReportWinTitle)

        ; 3. 尋找「表格」元素
        tableEle := winEle.FindFirst(UIA_PastReportTable)
        if !IsObject(tableEle)
            throw Error("找不到 Table 物件！`n請檢查您的 UIA_PastReportTable 查詢條件。`n`n目前條件: " . UIA_PastReportTable)

        ; 4. 尋找表格中所有的「行」(Row)
        ; (WinForms 中, 行的 ControlType 通常是 'DataItem')
        ;rowElements := tableEle.FindAll({Type: 'DataItem'})
        rowElements := tableEle.FindAll({Type: 'Custom'})
        if (rowElements.Length = 0)
            throw Error("表格找到了，但裡面沒有 'DataItem' (Row)。")

        ;MsgBox(rowElements.Length)
        ;str := ""
        ; 5. 遍歷每一行
        for rowEle in rowElements
        {
            ; 6. 尋找該行中所有的「儲存格」(Cell)
            ; (Cell 的 ControlType 可能是 'Text', 'Edit' 或 'Custom')
            ; 您需要用 Accessibility Insights 檢查確認
            cellElements := rowEle.FindAll({Type: 'DataItem'})

            ; 如果 'Text' 找不到, 試試 'Custom'
            if (cellElements.Length = 0)
                cellElements := rowEle.FindAll({Type: 'Custom'})

            ; 檢查這行是否有足夠的欄位
            if (cellElements.Length < SearchColumnIndex)
                continue

            ; 7. 獲取目標儲存格 (UIA.ahk 陣列從 1 開始)
            targetCellEle := cellElements[SearchColumnIndex]

            ;str .= targetCellEle.Value . "`t"
            ; 8. 檢查文字
            ;if InStr(targetCellEle.Value, SearchText)
            if isRelatedReport(targetCellEle.Value, SearchText)
            {
                ; *** 找到了！ ***

                ; 9. 獲取儲存格的 BoundingRectangle (邊界矩形)
                ; 這是 UIA 的 accLocation, 幾乎一定有效
                rect := targetCellEle.BoundingRectangle
                loc := targetCellEle.Location

                ;if (rect.Width = 0 && rect.Height = 0)
                {
                    ;MsgBox("找到了 %SearchText%，但它的 BoundingRectangle 座標是 0。`n嘗試使用邏輯點擊 Click()...")
                    ;targetCellEle.Click() ; 嘗試邏輯點擊 (可能無法觸發第二視窗)
                    ;return
                }

                ; 10. 計算中心點並執行「真實滑鼠點擊」
                ; (這 100% 會觸發您要的 Click 事件)
                ClickX := rect.l + (loc.w / 2)
                ClickY := rect.t + (loc.h / 2)

                MouseGetPos(&OrigX, &OrigY)
                MouseMove(ClickX, ClickY, 0)
                Click()
                ;Sleep(3)
                ;Click(ClickX, ClickY)
                MouseMove(OrigX, OrigY, 0)

                ;WinActivate(RISReportWinTitle)
                ;Sleep(100)
                ;MsgBox("UIA 點擊成功！`n在 " . ClickX . ", " . ClickY . ", " . rect.l . ", " . rect.t . ", " . rect.r . ", " . rect.b)

                ;MsgBox("UIA 點擊成功！`n在 %ClickX%, %ClickY% 點擊了: " . targetCellEle.Value)
                return ; 任務完成, 退出
            }
        }

        MsgBox("掃描完畢，找不到包含 " . SearchText . " 的儲存格。")
    }
    catch as e
    {
        MsgBox("UIA 發生錯誤:`n" . e.Message . "`n`n行: " . e.Line, "UIA Error", 16)
    }
}

/*
^2::{
  Send "!q"
}

^3::{
  Send "!a"
}
  */
!q::{
  Send "^e"
}

/*
SC029::{
  If (!WinActive(RISReportWinTitle)) {
    WinActivate(RISReportWinTitle)
    WinWaitActive(RISReportWinTitle)
    ;ControlFocus, %FINDING_CONTROL%
  }
  FocusedControl := ControlGetFocus()
  ;MsgBox, %FocusedControl%
  If (FocusedControl = FINDING_CONTROL) {
    ControlFocus(IMPRESSION_CONTROL)
  } Else {
    ControlFocus(FINDING_CONTROL)
  }
}
*/

;; Insert Selected Prev Exam Date
!d::{
  InsertSelectedPrevExamDate()
}

;; Insert Selected Prev Exam Name
^!e::{
  InsertSelectedPrevExamName()
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

  Return outputDate
}

InsertSelectedPrevExamDate() {
  global risEle
  Local STATE_SYSTEM_SELECTED := 0x2
  try {
    ;local cacheRequest := UIA.CreateCacheRequest(["ControlType", "Value"], ["LegacyIAccessiblePattern"], UIA.TreeScope.Descendants)
    local cacheRequest := UIA.CreateCacheRequest()
    cacheRequest.TreeScope := UIA.TreeScope.Subtree
    cacheRequest.AddProperty("ControlType")
    cacheRequest.AddProperty("Value")
    cacheRequest.AddProperty("Name")
    cacheRequest.AddPattern("LegacyIAccessible")

    ; 2. 獲取視窗元素
    winEle := risEle
    ;winEle := UIA.ElementFromHandle(WinGetID(RISReportWinTitle))
    if !IsObject(winEle)
      throw Error("找不到視窗: " . RISReportWinTitle)

    ; 3. 尋找「表格」元素
    tableEle := winEle.FindFirst(UIA_PastReportTable)
    ;tableEle := winEle.FindElement(UIA_PastReportTable, , , , , cacheRequest)
    if !IsObject(tableEle)
      throw Error("找不到 Table 物件！`n請檢查您的 UIA_PastReportTable 查詢條件。`n`n目前條件: " . UIA_PastReportTable)

    rowElements := tableEle.FindElements({ ControlType: 'Custom' })
    ;rowElements := tableEle.FindElements({ ControlType: 'Custom' }, , , , cacheRequest)
    ;rowElements := tableEle.FindCachedElements({ ControlType: 'Custom' })
    ;MsgBox(rowElements.Length)
    if (rowElements.Length = 0)
      throw Error("表格找到了，但裡面沒有 'DataItem' (Row)。")

    for i, rowEle in rowElements {
      ;MsgBox(rowEle.CachedControlType)
      ;if IsObject(rowEle.CachedLegacyIAccessiblePattern) {
      if IsObject(rowEle.LegacyIAccessiblePattern) {
        Local legacyState := rowEle.LegacyIAccessiblePattern.State
       ; Local legacyState := rowEle.CachedLegacyIAccessiblePattern.State
        ;MsgBox(legacyState)
        if (legacyState & STATE_SYSTEM_SELECTED) {
          dateText := ""
          ;MsgBox(rowEle.CachedChildren.Length)
          ;for cell in rowEle.CachedChildren {
          ;  MsgBox(cell.CachedValue)
          ;  if (cell.CachedControlType = UIA.ControlType.DataItem) {
          ;      dateText := cell.CachedValue
          ;      break
          ;  }
          ;}
          dateCellEle := rowEle.FindElement({ControlType: "DataItem"}, , 1)
          ;dateCellEle := rowEle.FindCachedElement({ControlType: "DataItem"}, , 1)
          ;dateCellEle := rowEle.FindCachedElement({ControlType: "DataItem"})
          ;MsgBox(dateCellEle.CachedValue)
          if IsObject(dateCellEle) {
            ;dateText := dateCellEle.CachedValue
            dateText := dateCellEle.Value
          }
          ;MsgBox("找到反白的行！ (透過 Legacy 狀態)`n`n行號 (邏輯): " . i . "`n內容: " . dateText)
          ;MsgBox(dateText)
          Paste(ConvertRISDate(dateText))
          return
        }
      }
    }
    ;MsgBox("掃描完畢，沒有找到任何 'Selected' (反白) 的行。")
  }
  catch as e {
    MsgBox("UIA 發生錯誤:`n" . e.Message . "`n`n行: " . e.Line, "UIA Error", 16)
  }
}

InsertSelectedPrevExamName() {
  global risEle
  Local STATE_SYSTEM_SELECTED := 0x2
  try {
    ; 2. 獲取視窗元素
    ;winEle := UIA.ElementFromHandle(WinGetID(RISReportWinTitle))
    winEle := risEle
    if !IsObject(winEle)
      throw Error("找不到視窗: " . RISReportWinTitle)

    ; 3. 尋找「表格」元素
    tableEle := winEle.FindFirst(UIA_PastReportTable)
    if !IsObject(tableEle)
      throw Error("找不到 Table 物件！`n請檢查您的 UIA_PastReportTable 查詢條件。`n`n目前條件: " . UIA_PastReportTable)

    rowElements := tableEle.FindAll({ Type: 'Custom' })
    if (rowElements.Length = 0)
      throw Error("表格找到了，但裡面沒有 'DataItem' (Row)。")

    for i, rowEle in rowElements {
      if IsObject(rowEle.LegacyIAccessiblePattern) {
        Local legacyState := rowEle.LegacyIAccessiblePattern.State
        if (legacyState & STATE_SYSTEM_SELECTED) {
          examnameText := ""
          dateCellEle := rowEle.FindElement({ControlType: "DataItem"}, , 3)
          if IsObject(dateCellEle) {
            examnameText := dateCellEle.Value
          }
          ;MsgBox("找到反白的行！ (透過 Legacy 狀態)`n`n行號 (邏輯): " . i . "`n內容: " . dateText)
          Paste(examnameText)
          return
        }
      }
    }
    ;MsgBox("掃描完畢，沒有找到任何 'Selected' (反白) 的行。")
  }
  catch as e {
    MsgBox("UIA 發生錯誤:`n" . e.Message . "`n`n行: " . e.Line, "UIA Error", 16)
  }
}

GetCurrExamType()
{
  examname := GetCurrExamName()
  If (InStr(examname, "CT") || InStr(examname, "電腦斷層")) {
    Return "CT"
  } Else If (InStr(examname, "MR") || InStr(examname, "磁振造影")) {
    Return "MR"
  } Else If (InStr(examname, "US")) {
    Return "US"
  }
  Return "CR"
}

OrderListForFindings()
{
  examtype := GetCurrExamType()
  ;MsgBox(examtype)
  Switch examtype
  {
    case "CT", "MR":
      UnorderListForFindingsOfCtOrMr()

    case "CR", "US":
      ;UnorderListForFindingsOfCrUs()
  }
}

UnorderListForFindingsOfCtOrMr()
{
  examtype := GetCurrExamType()
  If (examtype = "CT" || examtype = "MR") {
    ;ControlGet, hEdit, Hwnd,, %FINDING_CONTROL%
    ;fdEle := risEle.FindFirst(UIA_FindingEdit)
    ;hEdit := fdEle.NativeWindowHandle
    ;MsgBox(FINDING_CONTROL_HWND)
    if (hEdit := FINDING_CONTROL_HWND) {
      startSel := Edit_FindText(hEdit, "FINDINGS:`r`n|The study shows:`r`n`r`n|show the following findings:`r`n`r`n|which revealed:`r`n`r`n", , , "RegEx", &matchedText)

      If (startSel > -1) {
        ;startSel += StrLen(matchedText)
        startSel += matchedText.Len
        Loop 3 {
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
}

;;; Formatting FINDINGS
;;;; Reorder Seleted Text And Keep SeIm
SC079::{
  OrderListForFindings()
}

;;; Formatting IMPRESSION
;;;; Reorder Seleted Text And Discard SeIm
FormatImpressionText() {
  global IMPRESSION_CONTROL_HWND
  if (hEdit := IMPRESSION_CONTROL_HWND) {
    Edit_SetFocus(hEdit)
    Edit_SetSel(hEdit)
    if (Edit_GetLogicalLineCount(hEdit) > 1) {
      ReorderSelectedText()
    } else {
      ReorderSelectedText(true)
    }
  }
}

SC070::{
  FormatImpressionText()
}

; Reorder Seleted Text And Discard SeIm
^!o::{
  ReorderSelectedText()
}

; Reorder Seleted Text And Keep SeIm
^!+o::{
  ReorderSelectedText(,,, false)
}

; Unorder Seleted Text
^+*::{
  ReorderSelectedText(false, true, "*")
}

^+-::{
  ReorderSelectedText(false, true, "-")
}

^+>::{
  ReorderSelectedText(false, true, ">")
}

CopyPathologyReport() {
  global risEle
  if (!IsObject(risEle)) {
    UpdateRisElements()
    MsgBox("請先打開病理報告視窗。")
  }
  pathoEle := risEle.FindFirst(UIA_PathoDiagnosisTxt)
  pathoDateEle := risEle.FindFirst(UIA_PathoDateTxt)
  reportText := ""
  if IsObject(pathoDateEle) && IsObject(pathoEle) {
    reportText .= ConvertRISDate(pathoDateEle.Value) . ": " . pathoEle.Value
  }
  if (reportText != "") {
    A_Clipboard := reportText
    ;MsgBox("病理報告已複製到剪貼簿。")
  } else {
    MsgBox("找不到病理報告內容。")
  }
}

^+c::{
  CopyPathologyReport()
}

#HotIf ; WinActive(RISReportWinTitle)

;; for JIS keyboard
SC029::{
  If (!WinActive(RISReportWinTitle)) {
    WinActivate(RISReportWinTitle)
    WinWaitActive(RISReportWinTitle)
    focusedEle := UIA.GetFocusedElement()
    if (focusedEle.AutomationId != UIA_FindingEdit.AutomationId && focusedEle.AutomationId != UIA_ImpressionEdit.AutomationId) {
      if (FINDING_CONTROL_HWND) {
        ControlFocus(FINDING_CONTROL_HWND)
      } else {
        fdEle := risEle.FindFirst(UIA_FindingEdit)
        if IsObject(fdEle) {
          fdEle.SetFocus()
        }
      }
    }
  } Else {
    focusedEle := UIA.GetFocusedElement()
    ;MsgBox(focusedEle.AutomationId)
    if (focusedEle.AutomationId = UIA_FindingEdit.AutomationId) {
      if (IMPRESSION_CONTROL_HWND) {
        ControlFocus(IMPRESSION_CONTROL_HWND)
      } else {
        impEle := risEle.FindFirst(UIA_ImpressionEdit)
        if IsObject(impEle) {
          impEle.SetFocus()
        }
      }
    } else {
      if (FINDING_CONTROL) {
        ControlFocus(FINDING_CONTROL_HWND)
      } else {
        fdEle := risEle.FindFirst(UIA_FindingEdit)
        if IsObject(fdEle) {
          fdEle.SetFocus()
        }
      }
    }
  }
}

UpdateRisElements()
{
  global risEle, autoNextEle
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
  }
  catch as e
  {
    ; 如果發生錯誤 (例如視窗找不到), 我們的快取就會是 stale
    ;MsgBox("UIA Error in timer: " . e.Message) ; (可選: 除錯用)

    ; 7. 如果視窗或元素找不到, 或發生錯誤, 就清除 global 變數
    risEle := ""
  }
}

ReorderSelectedText(deOrder := false, keepEmptyLine := false, itemChar := "", discardSeIm := true) {
    global PRESERVE_CLIPBOARD

    local isSpine := false
    local ClipSaved := "" ; 儲存備份的變數

    ; v2 (正確): 備份剪貼簿 (使用 ClipboardAll 變數)
    if (PRESERVE_CLIPBOARD) {
        ClipSaved := ClipboardAll
    }

    ; v2 (正確): 清空剪貼簿 (使用 A_Clipboard 變數)
    A_Clipboard := ""
    Send "^c"

    ; 等待剪貼簿包含資料
    if !ClipWait(0.8, 1) {
        MsgBox "複製文字到剪貼簿失敗 (逾時)。"
        if (PRESERVE_CLIPBOARD)
            ClipboardAll := ClipSaved ; v2 (正確): 還原
        Return -1
    }

    ; v2 (正確): 讀取文字 (使用 A_Clipboard 變數)
    local selectedText := A_Clipboard

    ; --- 1. 文字正規化 (這部分邏輯不變) ---
    selectedText := StrReplace(selectedText, "`r`n", "`n")
    if (InStr(selectedText, "`r")) {
        MsgBox "選取範圍內包含不正確的換行符號 (CR)。請使用正確的方式選取文字。"
        if (PRESERVE_CLIPBOARD)
            ClipboardAll := ClipSaved ; v2 (正確): 還原
        Return -1
    }
    local hadTrimmedRight := false
    if (SubStr(selectedText, -1) == "`n") {
        selectedText := SubStr(selectedText, 1, -1)
        hadTrimmedRight := true
        ;MsgBox("已去除選取文字末尾的多餘換行符號。")
    }
    local txtAry := StrSplit(selectedText, "`n")
    local endLine := txtAry.Length

    ; --- 2. 處理文字 (這部分邏輯不變) ---
    if (StrLen(selectedText) > 0) {
        local finalText := ""
        local isFirstLineEmpty := false
        local startLineNo := 1

        if (RegExMatch(selectedText, "^(\d+)", &existLineNo)) {
            startLineNo := existLineNo[1]
        }

        for index, line in txtAry {
            if (index == 1 && !StrLen(line)) {
                isFirstLineEmpty := true
                ;MsgBox("第一行是空行")
            }
            if (!RegExMatch(line, "^\s*$")) {
                if (RegExMatch(line, "^\s*[-\+\*]*\s*([Vv]arying degree|[Mm]ild).+causing:")) {
                    isSpine := true
                }
                local tmpText := line
                if (!deOrder) {
                    local orderChar := (StrLen(itemChar) > 0 ? itemChar : startLineNo++ . ".")
                    if (isSpine && RegExMatch(line, "^\s*([-\+\*]*|-->)\s*([CcTtLl]\d{1,2}-.+$)", &matchedSpineLevel)) {
                        finalText .= "--> "
                        tmpText := matchedSpineLevel[2]
                    } else {
                        finalText .= orderChar . " "
                    }
                }
                if (StrLen(itemChar) == 0 && discardSeIm) {
                    tmpText := RegExReplace(tmpText, "\s*\(Srs\/Img:[\s,-\/\d;]+\)", "")
                    tmpText := RegExReplace(tmpText, "Mark L\d+:\s*", "")
                }

                ; --- ★ 這裡是您指定的方案 1 ★ ---
                ; v1: "$u7$8"
                ; v2: "$u${7}${8}" (v2 的 $u 指令只作用於緊鄰的下一個變數)
                finalText .= RegExReplace(
                    tmpText,
                    "^(\s*)((\d+\.)|([-\+\*>=])|(\(?\d+\)))?(\s*)(\w?)(.*)",
                    "$u{7}${8}"
                )
                ; --- ★ 修正完畢 ★ ---

                if (index < endLine || hadTrimmedRight) {
                    finalText .= "`r`n"
                }
            } else {
                if (isFirstLineEmpty) {
                    if (index == 1) {
                        finalText .= "`r`n"
                    }
                }
                if (keepEmptyLine) {
                    if (isFirstLineEmpty) {
                        if (!Mod(index, 2)) {
                            finalText .= "`r`n"
                        }
                    } else {
                        if (Mod(index, 2)) {
                            finalText .= "`r`n"
                        }
                    }
                }
            }
        }

        ; --- 3. 輸出 (修正) ---

        ; v2 (正確): 寫入文字 (使用 A_Clipboard 變數)
        A_Clipboard := finalText

        Sleep 100
        Send "^v"

        if (PRESERVE_CLIPBOARD) {
            ; v2 (正確): 還原剪貼簿 (使用 ClipboardAll 變數)
            ClipboardAll := ClipSaved
        }
        Return 0
    } else {
        MsgBox "選取範圍內沒有內容。"
        if (PRESERVE_CLIPBOARD)
            ClipboardAll := ClipSaved ; v2 (正確): 還原
        Return -1
    }
}

Global simReportMap := Map(
  "CHEST PA/AP", Map("CHEST PA/AP+LAT",1),
  "CHEST PA/AP+LAT", Map("CHEST PA/AP",1),
  "KUB", Map("KUB+ABD LAT",1),
  "KUB+L-SPINE LAT(supine)", Map("L-SPINE(AP+LAT)Standing",1),
  "WHOLE  ABDOMEN CT WITH+ WITHOUT CONTRAST", Map("WHOLE  ABDOMEN CT WITHOUT CONTRAST",1),
  "WHOLE  ABDOMEN CT WITHOUT CONTRAST", Map("WHOLE  ABDOMEN CT WITH+ WITHOUT CONTRAST",1),
)

UpdateRisElements()
SetTimer(UpdateRisElements, 5000)

; Mouse Remapping
#HotIf WinActive(RISReportWinTitle)

XButton1::{
  ReorderSelectedText(false, true, "-")
}
XButton2::{
  ReorderSelectedText()
}

#HotIf ; WinActive(RISReportWinTitle)