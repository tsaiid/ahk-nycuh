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

Global RISReportWinTitle := "報告作業(frmRISReport)" ; 替換成您的程式標題
Global UIA_PastReportTable := {AutomationId: "dgvPastReport"}
Global UIA_AutoNextCheckbox := {AutomationId: "chkAutoNext"}
Global UIA_ReportSaveButton := {AutomationId: "btnReportSave"}
Global UIA_ExamNameTxt := {AutomationId: "txtExamName"}
Global FINDING_CONTROL := "WindowsForms10.EDIT.app.0.2780b98_r24_ad12"
Global IMPRESSION_CONTROL := "WindowsForms10.EDIT.app.0.2780b98_r24_ad11"

#HotIf WinActive(RISReportWinTitle)
!c::{
  CheckNextAuto(false)
  SaveReport()
}

^s::{
  CheckNextAuto(true)
  SaveReport()
}

CheckNextAuto(checked := true){
  try {
      ; 2. 獲取視窗元素
      winEle := UIA.ElementFromHandle(WinGetID(RISReportWinTitle))
      if !IsObject(winEle)
          throw Error("找不到視窗: " . RISReportWinTitle)

      ; 3. 尋找「下一筆自動勾選框」元素
      nextChkEle := winEle.FindFirst(UIA_AutoNextCheckbox)
      if !IsObject(nextChkEle)
          throw Error("找不到 '下一筆自動勾選框' 物件！`n請檢查您的 UIA_AutoNextCheckbox 查詢條件。`n`n目前條件: " . UIA_AutoNextCheckbox)

      ; 4. 切換勾選框狀態
      if (checked ^ nextChkEle.ToggleState) {
        nextChkEle.Toggle()
      }
      ;MsgBox(nextChkEle.ToggleState)
  }
  catch as e {
      MsgBox("UIA 發生錯誤:`n" . e.Message . "`n`n行: " . e.Line, "UIA Error", 16)
  }
}

SaveReport(){
  try {
      ; 2. 獲取視窗元素
      winEle := UIA.ElementFromHandle(WinGetID(RISReportWinTitle))
      if !IsObject(winEle)
          throw Error("找不到視窗: " . RISReportWinTitle)

      ; 3. 尋找「報告存檔按鈕」元素
      saveBtnEle := winEle.FindFirst(UIA_ReportSaveButton)
      if !IsObject(saveBtnEle)
          throw Error("找不到 '報告存檔按鈕' 物件！`n請檢查您的 UIA_ReportSaveButton 查詢條件。`n`n目前條件: " . UIA_ReportSaveButton)

      ; 4. 點擊按鈕
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
  try {
      ; 2. 獲取視窗元素
      winEle := UIA.ElementFromHandle(WinGetID(RISReportWinTitle))
      if !IsObject(winEle)
          throw Error("找不到視窗: " . RISReportWinTitle)

      ; 3. 尋找「檢查名稱」輸入框元素
      examNameEle := winEle.FindFirst(UIA_ExamNameTxt)
      if !IsObject(examNameEle)
          throw Error("找不到 '檢查名稱' 輸入框！")

      ; 4. 獲取輸入框的值
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

^2::{
  Send "!q"
}

^3::{
  Send "!a"
}

!q::{
  Send "^e"
}

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

;; Insert Selected Prev Exam Date
!d::{
  InsertSelectedPrevExamDate()
}

ConvertRISDate(inputString) {
  ; 1. 從輸入字串 (yyymmddhhmm) 中提取各個部分
  minguoYear := SubStr(inputString, 1, 3)  ; yyy (例如: 114)
  month := SubStr(inputString, 4, 2)       ; mm (例如: 10)
  day := SubStr(inputString, 6, 2)         ; dd (例如: 15)

  ; 2. 將民國年轉換為西元年 (民國年 + 1911 = 西元年)
  ; AHK v1 會自動將字串 "114" 視為數字 114 進行計算
  gregorianYear := minguoYear + 1911

  ; 3. 組合並返回 yyyy-mm-dd 格式的字串
  ; 使用 . 運算子來連接字串
  outputDate := gregorianYear . "-" . month . "-" . day

  Return outputDate
}

InsertSelectedPrevExamDate() {
  Local STATE_SYSTEM_SELECTED := 0x2
  try {
    ; 2. 獲取視窗元素
    winEle := UIA.ElementFromHandle(WinGetID(RISReportWinTitle))
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
          dateText := ""
          dateCellEle := rowEle.FindElement({ControlType: "DataItem"}, , 1)
          if IsObject(dateCellEle) {
            dateText := dateCellEle.Value
          }
          ;MsgBox("找到反白的行！ (透過 Legacy 狀態)`n`n行號 (邏輯): " . i . "`n內容: " . dateText)
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

#HotIf ; 關閉上一個條件

Global simReportMap := Map(
  "CHEST PA/AP", Map("CHEST PA/AP+LAT",1),
  "CHEST PA/AP+LAT", Map("CHEST PA/AP",1),
  "KUB", Map("KUB+ABD LAT",1),
  "KUB+L-SPINE LAT(supine)", Map("L-SPINE(AP+LAT)Standing",1),
)
