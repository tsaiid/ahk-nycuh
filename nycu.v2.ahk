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
Global risEle := ""
Global autoNextEle := ""
Global reportSaveEle := ""

#HotIf WinActive(RISReportWinTitle)

;#Include MyScripts\regex-hotstrings.v2.ahk
;#Include MyScripts\others.v2.ahk
;#Include MyScripts\chest-ct.v2.ahk
;#Include MyScripts\abdomen-ct.v2.ahk
;#Include MyScripts\abdomen-mr.v2.ahk
;#Include MyScripts\ct-guide.v2.ahk
;#Include MyScripts\ms-ct.v2.ahk
;#Include MyScripts\ms-mri.v2.ahk
#Include MyScripts\neuro.v2.ahk
;#Include MyScripts\abbreviations.v2.ahk
;#Include MyScripts\mri.v2.ahk
;#Include MyScripts\chest-x-ray.v2.ahk
;#Include MyScripts\kub.v2.ahk
;#Include MyScripts\bone-x-ray.v2.ahk
;#Include MyScripts\other-x-ray.v2.ahk
;#Include MyScripts\sono-guide.v2.ahk
;#Include MyScripts\angio.v2.ahk

#Include MyScripts\lib\ris-common.v2.ahk
^9::{
}

!f::{
  Send "^{Right}"
}

!b::{
  Send "^{Left}"
}

^1::{
  try {
    RisController.PastAllRadio.ControlClick()
  } catch as err {
    MsgBox "操作失敗: " err.Message
  }
}
^2::{
  try {
    RisController.PastModalityRadio.ControlClick()
  } catch as err {
    MsgBox "操作失敗: " err.Message
  }
}
^3::{
  try {
    RisController.PastOnlyMyRadio.ControlClick()
  } catch as err {
    MsgBox "操作失敗: " err.Message
  }
}
;;; Append previous report to FINDINGS and IMPRESSION
AppendPrevReport() {
  try {
    pastImpression := RisController.PastImpressionText.Value
    hImpEdit := RisController.ImpressionEdit.NativeWindowHandle
    Edit_SetSel(hImpEdit, Edit_GetTextLength(hImpEdit))
    Edit_ReplaceSel(hImpEdit, pastImpression)

    pastFinding := RisController.PastFindingText.Value
    hFindingEdit := RisController.FindingEdit.NativeWindowHandle
    Edit_SetSel(hFindingEdit, Edit_GetTextLength(hFindingEdit))
    Edit_ReplaceSel(hFindingEdit, pastFinding)
    Edit_SetSel(hFindingEdit, 0, 0)
    Edit_ScrollCaret(hFindingEdit)
  } catch TargetError as err {
    MsgBox "操作失敗: " err.Message
  }
}
^ESC::{
  AppendPrevReport()
}
;; Ctrl + Y
;; Delete current line
^y::{
  local focusedEle := UIA.GetFocusedElement()
  if (focusedEle.AutomationId = UIA_FindingEdit.AutomationId || focusedEle.AutomationId = UIA_ImpressionEdit.AutomationId) {
    hEdit := focusedEle.NativeWindowHandle
    SelectLogicalLine(hEdit)
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
cacheRequest := UIA.CreateCacheRequest(["AutomationId", "NativeWindowHandle"])

DeletePrevWordUsingCache() {
  global cacheRequest
  ;focusedEle := UIA.GetFocusedElement()
  focusedEle := UIA.GetFocusedElement(cacheRequest)
  ;if (focusedEle.AutomationId = UIA_FindingEdit.AutomationId || focusedEle.AutomationId = UIA_ImpressionEdit.AutomationId) {
  if (focusedEle.CachedAutomationId = UIA_FindingEdit.AutomationId || focusedEle.CachedAutomationId = UIA_ImpressionEdit.AutomationId) {
    hEdit := focusedEle.CachedNativeWindowHandle
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

DeletePrevWord() {
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

BashDeleteWordBackward(hCtrl) {
    try {
        fullText := ControlGetText(hCtrl)
    } catch {
        return
    }

    ; 1. 取得當前游標位置 (0-based)
    caretPosRaw := SendMessage(0x00B0, 0, 0, hCtrl)
    caretPos := caretPosRaw & 0xFFFF ; Low word is starting position

    if (caretPos == 0)
        return ; 已經在最前面，沒東西可刪

    ; 2. 轉成 AHK 1-based 索引
    ; 我們的目標是找到新的起點 (newStart)，範圍是 newStart 到 caretPos
    i := caretPos ; i 代表目前檢查的字元位置 (從游標左邊那個字開始)

    ; 3. 向後搜尋邏輯 (模擬 Bash)

    ; 階段 A: 如果游標緊貼著空白，先吃掉這些空白 (Bash 通常會連同後面的單字一起刪)
    ; 例如 "ls -la  |" -> 按下 Ctrl+W -> "ls |" (刪掉 "-la  ")
    while (i > 0) {
        char := SubStr(fullText, i, 1)
        if (IsSpace(char)) {
            i--
        } else {
            break ; 遇到非空白字元，進入階段 B
        }
    }

    ; 階段 B: 繼續往前吃掉所有「非空白」字元，直到遇到空白或檔頭
    while (i > 0) {
        char := SubStr(fullText, i, 1)
        if (!IsSpace(char)) {
            i--
        } else {
            break ; 遇到空白了，停止 (這個空白要保留)
        }
    }

    ; 此時 i 指向的是「上一個單字前的空白」位置，或者是 0
    ; 因為我們選取範圍是從 i 開始，這會剛好保留那個空白
    selStart := i
    selEnd := caretPos

    ; 4. 執行刪除
    ; 先選取範圍 (保留 Undo 功能的最佳作法)
    SendMessage(0x00B1, selStart, selEnd, hCtrl) ; EM_SETSEL

    ; 發送 Backspace 刪除選取範圍
    ;SendInput "{Backspace}"
    SendMessage(0x303, 0, 0, , hCtrl)
}

; 簡單的空白字元檢查函數
IsSpace(char) {
    return (char == " " || char == "`t" || char == "`r" || char == "`n")
}

^w::{
  ;DeletePrevWord()
  focusedEle := UIA.GetFocusedElement()
  if (focusedEle.AutomationId = UIA_FindingEdit.AutomationId || focusedEle.AutomationId = UIA_ImpressionEdit.AutomationId) {
    hEdit := focusedEle.NativeWindowHandle
    BashDeleteWordBackward(hEdit)
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
    Loop times {
        funcObj()
    }

    ; 4. 記錄結束時間
    DllCall("QueryPerformanceCounter", "Int64*", &end := 0)

    ; 5. 計算耗時 ( (結束-開始) / 頻率 * 1000 轉換為毫秒 )
    return (end - start) / freq * 1000
}


InsertExamname() {
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
  ClickSaveReport()
}

^s::{
  CheckNextAuto(true)
  ClickSaveReport()
}

CheckNextAuto(checked := true){
  try {
    if (checked ^ RisController.AutoNextCheckbox.ToggleState) {
      RisController.AutoNextCheckbox.Toggle()
    }
  } catch as err {
    MsgBox "操作失敗: " err.Message
  }
}

ClickSaveReport(){
  try {
    RisController.ReportSaveButton.ControlClick()
  } catch as err {
    MsgBox "操作失敗: " err.Message
  }
}

!ESC::{
  examName := GetCurrExamName()
  FindSimilarReport(examName)
  ;MsgBox(GetCurrExamName())
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

+Down:: {
  try {
    ; 1. 取得當前作用視窗("A") 中，擁有焦點的控制項名稱 (例如 "Edit1")
    focusedCtl := ControlGetFocus("A")

    ; 2. 將該控制項名稱傳入 Edit 函數
    currentLine := EditGetCurrentLine(focusedCtl, "A")
    lineCount := EditGetLineCount(focusedCtl, "A")

    ; 3. 判斷邏輯
    if (currentLine == lineCount) {
      SendInput "+{End}"
    } else {
      SendInput "+{Down}"
    }
  } catch {
    SendInput "+{Down}"
  }
}

+Up:: {
  try {
    focusedCtl := ControlGetFocus("A")
    currentLine := EditGetCurrentLine(focusedCtl, "A")

    if (currentLine == 1) {
      SendInput "+{Home}"
    } else {
      SendInput "+{Up}"
    }
  } catch {
    SendInput "+{Up}"
  }
}

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

InsertSelectedPrevExamDateCached() {
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
    ;tableEle := winEle.FindFirst(UIA_PastReportTable)
    tableEle := winEle.FindElement(UIA_PastReportTable, , , , , cacheRequest)
    if !IsObject(tableEle)
      throw Error("找不到 Table 物件！`n請檢查您的 UIA_PastReportTable 查詢條件。`n`n目前條件: " . UIA_PastReportTable)

    ;rowElements := tableEle.FindElements({ ControlType: 'Custom' })
    ;rowElements := tableEle.FindElements({ ControlType: 'Custom' }, , , , cacheRequest)
    rowElements := tableEle.FindCachedElements({ ControlType: 'Custom' })
    ;MsgBox(rowElements.Length)
    if (rowElements.Length = 0)
      throw Error("表格找到了，但裡面沒有 'DataItem' (Row)。")

    for i, rowEle in rowElements {
      ;MsgBox(rowEle.CachedControlType)
      if IsObject(rowEle.CachedLegacyIAccessiblePattern) {
      ;if IsObject(rowEle.LegacyIAccessiblePattern) {
        ;Local legacyState := rowEle.LegacyIAccessiblePattern.State
        Local legacyState := rowEle.CachedLegacyIAccessiblePattern.State
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
          ;dateCellEle := rowEle.FindElement({ControlType: "DataItem"}, , 1)
          dateCellEle := rowEle.FindCachedElement({ControlType: "DataItem"}, , 1)
          ;MsgBox(dateCellEle.CachedValue)
          if IsObject(dateCellEle) {
            dateText := dateCellEle.CachedValue
            ;dateText := dateCellEle.Value
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

InsertSelectedPrevExamDate() {
  Local STATE_SYSTEM_SELECTED := 0x2
  try {
    rowElements := RisController.PastReportTable.FindElements({ ControlType: 'Custom' })
    if (rowElements.Length = 0)
      throw Error("表格找到了，但裡面沒有 'DataItem' (Row)。")

    for i, rowEle in rowElements {
      if IsObject(rowEle.LegacyIAccessiblePattern) {
        Local legacyState := rowEle.LegacyIAccessiblePattern.State
        if (legacyState & STATE_SYSTEM_SELECTED) {
          dateText := ""
          dateCellEle := rowEle.FindElement({ControlType: "DataItem"}, , 1)
          if IsObject(dateCellEle) {
            dateText := dateCellEle.Value
          }
          Paste(ConvertRISDate(dateText))
          return
        }
      }
    }
  } catch as err {
    MsgBox("UIA 發生錯誤:`n" . err.Message . "`n`n行: " . err.Line, "UIA Error", 16)
  }
}

InsertSelectedPrevExamDate_old() {
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
  Local STATE_SYSTEM_SELECTED := 0x2
  try {
    rowElements := RisController.PastReportTable.FindAll({ Type: 'Custom' })
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
          Paste(examnameText)
          return
        }
      }
    }
  }
  catch as e {
    MsgBox("UIA 發生錯誤:`n" . e.Message . "`n`n行: " . e.Line, "UIA Error", 16)
  }
}

InsertSelectedPrevExamName_old() {
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
  } Else If (InStr(examname, "US") || InStr(examname, "超音波")) {
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
      UnorderListForFindingsOfCrOrUs()
  }
}

UnorderListForFindingsOfCrOrUs()
{
  If (hEdit := RisController.FindingEdit.NativeWindowHandle) {
    startSel := Edit_FindText(hEdit, "FINDINGS:`r`n|:\s*`r`n\s*`r`n", , , "RegEx", &matchedText)
    If (startSel > -1) {
      startSel += matchedText.Len
      Edit_SetFocus(hEdit)
      Edit_SetSel(hEdit, startSel, -1)
      ;Sleep, 100
      ReorderSelectedText(false, true, "-", false)
    }
  }
}

UnorderListForFindingsOfCtOrMr()
{
  if (hEdit := RisController.FindingEdit.NativeWindowHandle) {
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
      ReorderSelectedText(false, false, "-")
      ;MsgBox % regex_out
    }
  } else {
    MsgBox("FINDING_CONTROL_HWND is invalid!")
  }
}

;;; Formatting FINDINGS
;;;; Reorder Seleted Text And Keep SeIm
SC079::{
  OrderListForFindings()
}
^!,::{
  OrderListForFindings()
}

;;; Formatting IMPRESSION
;;;; Reorder Seleted Text And Discard SeIm
FormatImpressionText() {
  if (hEdit := RisController.ImpressionEdit.NativeWindowHandle) {
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
^!.::{
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
  try {
    reportText := ConvertRISDate(RisController.PathoDateText.Value) . ": " . RisController.PathoDiagnosisText.Value
    if (reportText != "") {
      A_Clipboard := reportText
      MsgBox("病理報告已複製到剪貼簿。")
    } else {
      throw Error("找不到病理報告內容。")
    }
  } catch as err {
    MsgBox("操作失敗: " . err.Message)
  }
}

^+c::{
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
        MouseGetPos ,,, &hCtrl, 2

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
SC029::{
  try {
    If (!WinActive(RISReportWinTitle)) {
      WinActivate(RISReportWinTitle)
      WinWaitActive(RISReportWinTitle)
      focusedEle := UIA.GetFocusedElement()
      if (focusedEle.AutomationId != UIA_FindingEdit.AutomationId && focusedEle.AutomationId != UIA_ImpressionEdit.AutomationId) {
        ControlFocus(RisController.FindingEdit.NativeWindowHandle)
      }
    } Else {
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

class RisController
{
    ; =================================================================
    ; 1. 設定區 (Configuration) - 未來擴充元件都在這裡加
    ; =================================================================
    static WinTitle := RISReportWinTitle

    ; 定義所有子元件的搜尋條件
    static Selectors := Map(
        "AutoNextCheckbox", {AutomationId: "chkAutoNext"},
        "ReportSaveButton", {AutomationId: "btnReportSave"},
        "FindingEdit", {AutomationId: "txtReport"},
        "ImpressionEdit", {AutomationId: "txtImpression"},
        "PastAllRadio", {AutomationId: "rdoPastALL"},
        "PastModalityRadio", {AutomationId: "rdoClassify"},
        "PastOnlyMyRadio", {AutomationId: "rdoPastOnlyMy"},
        "ExamnameText", {AutomationId: "txtExamName"},
        "PastFindingText", {AutomationId: "rtxtPastReport"},
        "PastImpressionText", {AutomationId: "rtxtPastImpression"},
        "PastReportTable", {AutomationId: "dgvPastReport"},
        "PathoDiagnosisText", {AutomationId: "txtDiagnosist"},
        "PathoDateText", {AutomationId: "mtxtRcpDTM"},
    )

    ; =================================================================
    ; 2. 內部狀態 (State)
    ; =================================================================
    static _cache := Map() ; 用來存放所有已經抓到的 UIA Element

    ; =================================================================
    ; 3. 公開屬性 (Public Properties) - 外部呼叫用
    ; =================================================================

    ; 取得最上層 Ris Element (父層)
    static Ris {
        get => this._GetOrUpdateNode("Ris")
    }

    ; 取得 AutoNext Checkbox
    static AutoNextCheckbox {
        get => this._GetOrUpdateNode("AutoNextCheckbox")
    }

    ; 取得 Save Button
    static ReportSaveButton {
        get => this._GetOrUpdateNode("ReportSaveButton")
    }

    static FindingEdit {
        get => this._GetOrUpdateNode("FindingEdit")
    }
    static ImpressionEdit {
        get => this._GetOrUpdateNode("ImpressionEdit")
    }

    static PastAllRadio {
        get => this._GetOrUpdateNode("PastAllRadio")
    }
    static PastModalityRadio {
      get => this._GetOrUpdateNode("PastModalityRadio")
    }
    static PastOnlyMyRadio {
      get => this._GetOrUpdateNode("PastOnlyMyRadio")
    }

    static ExamnameText {
        get => this._GetOrUpdateNode("ExamnameText")
    }

    static PastFindingText {
        get => this._GetOrUpdateNode("PastFindingText")
    }
    static PastImpressionText {
        get => this._GetOrUpdateNode("PastImpressionText")
    }

    static PastReportTable {
        get => this._GetOrUpdateNode("PastReportTable")
    }

    static PathoDiagnosisText {
        get => this._GetOrUpdateNode("PathoDiagnosisText")
    }
    static PathoDateText {
        get => this._GetOrUpdateNode("PathoDateText")
    }

    ; =================================================================
    ; 4. 核心邏輯 (Core Logic) - 包含驗證與自動更新機制
    ; =================================================================
    static _GetOrUpdateNode(nodeName)
    {
        ; --- A. 檢查快取是否有效 (Probe 測試) ---
        if this._cache.Has(nodeName) {
            el := this._cache[nodeName]
            try {
                ; 嘗試讀取一個輕量屬性來確認它還活著
                temp := el.ControlType
                return el ; 沒報錯，直接回傳快取
            } catch {
                ; 報錯代表失效 (Stale)，移除快取準備重新抓取
                this._cache.Delete(nodeName)
            }
        }

        ; --- B. 快取無效，執行重新抓取邏輯 ---

        ; 情境 1: 請求的是 "Ris" (最上層父節點)
        if (nodeName = "Ris") {
            if !WinExist(this.WinTitle)
                throw TargetError("找不到 RIS 視窗: " this.WinTitle)

            try {
                ; 從視窗 Handle 取得新的 Element
                newRoot := UIA.ElementFromHandle(WinExist(this.WinTitle))
                this._cache["Ris"] := newRoot
                return newRoot
            } catch as err {
                throw Error("無法取得 RIS Root Element: " err.Message)
            }
        }

        ; 情境 2: 請求的是子節點 (Checkbox, Button 等)
        else {
            ; 這裡使用了遞迴概念：要找子節點，必須先保證父節點 (this.Ris) 是有效的
            ; 呼叫 this.Ris 會自動觸發上面的 "情境 1" 檢查
            parent := this.Ris

            ; 取得該名稱對應的搜尋條件
            if !this.Selectors.Has(nodeName)
                throw Error("未定義的選取器: " nodeName)
            condition := this.Selectors[nodeName]

            try {
                ; 從父節點往下找
                newChild := parent.FindElement(condition)
                this._cache[nodeName] := newChild
                return newChild
            } catch {
                throw TargetError("在 RIS 介面中找不到元件: " nodeName)
            }
        }
    }
}

UpdateRisElements()
{
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
    if !ClipWait(0.8) {
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
                  /*
                    if (isFirstLineEmpty) {
                        if (!Mod(index, 2)) {
                            finalText .= "`r`n"
                        }
                    } else {
                        if (Mod(index, 2)) {
                            finalText .= "`r`n"
                        }
                    }
                  */
                    finalText .= "`r`n"
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

;UpdateRisElements()
;SetTimer(UpdateRisElements, 60000)

;-----------------------------------------------------------
; Mouse Remapping
#HotIf WinActive(RISReportWinTitle)

XButton1::{
  ReorderSelectedText(false, true, "-")
}
XButton2::{
  ReorderSelectedText()
}

#HotIf ; WinActive(RISReportWinTitle)

;-----------------------------------------------------------
#HotIf WinActive(RISCTMRAbnormalWinTitle)
!1::{
  global ABNORMAL_VALUE_1_CONTROL
  ControlClick(ABNORMAL_VALUE_1_CONTROL)
}
!2::{
  global ABNORMAL_VALUE_2_CONTROL
  ControlClick(ABNORMAL_VALUE_2_CONTROL)
}
!3::{
  global ABNORMAL_VALUE_3_CONTROL
  ControlClick(ABNORMAL_VALUE_3_CONTROL)
}
!4::{
  global ABNORMAL_VALUE_4_CONTROL
  ControlClick(ABNORMAL_VALUE_4_CONTROL)
}

#HotIf ; WinActive(RISCTMRAbnormalWinTitle)