; RIS specific functions

SleepThenTab(sleepTime := 400, shiftTab := false)
{
  Sleep sleepTime
  if (shiftTab) {
    Send "+{Tab}"
  } else {
    Send "{Tab}"
  }
  Sleep sleepTime
}

; --- 輔助函式：檢查是否包含任一關鍵字 ---
ContainsKeywords(text, keywords) {
    for kw in keywords {
        if InStr(text, kw)
            return true
    }
    return false
}

; --- 輔助函式：處理 Oxford Comma 格式 ---
FormatList(arr) {
    if (arr.Length == 0)
        return ""

    ; 只有 1 個元素：直接回傳
    if (arr.Length == 1)
        return arr[1]

    ; 只有 2 個元素：A and B (中間不加逗號)
    if (arr.Length == 2)
        return arr[1] " and " arr[2]

    ; 3 個或以上元素：使用 Oxford Comma (A, B, and C)
    str := ""
    loop arr.Length {
        if (A_Index == 1) {
            str .= arr[A_Index]
        } else if (A_Index == arr.Length) {
            str .= ", and " arr[A_Index] ; 最後一項前加逗號和 and
        } else {
            str .= ", " arr[A_Index]
        }
    }
    return str
}