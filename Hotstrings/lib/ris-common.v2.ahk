; RIS specific functions

; ==============================================================================
; [核心函式] 檢查是否包含關鍵字，且「非」否定句 (嚴格限制在單行內判斷)
; 參數:
;   text: 完整的報告內容
;   keywords: 關鍵字陣列 (例如 ["scoliosis", "curve"])
; 回傳:
;   true (只要有任何一行出現該關鍵字且無否定詞，即回傳 true)
; ==============================================================================
HasPositiveFinding(text, keywords) {
    for kw in keywords {
        ; 建構正則搜尋，確保搜尋的是完整單字 (\b)
        pattern := "i)\b" . kw . "\b"

        startPos := 1
        Loop {
            ; 找尋關鍵字的位置
            foundPos := RegExMatch(text, pattern, &match, startPos)
            if (!foundPos)
                break ; 找不到此關鍵字了，換下一個關鍵字詞庫

            ; === 1. 取得關鍵字「之前」的所有文字 ===
            precedingAll := SubStr(text, 1, foundPos - 1)

            ; === 2. 鎖定「當前行」範圍 ===
            ; 尋找最近的一個換行符號 (`n 或 `r)，從後面截斷
            ; 這樣可以確保我們絕對不會讀到「上一行」的 No
            lastLineBreak := Max(InStr(precedingAll, "`n", , -1), InStr(precedingAll, "`r", , -1))

            ; 取得「行首 (或換行後) 到 關鍵字」之間的文字
            currentLinePreceding := SubStr(precedingAll, lastLineBreak + 1)

            ; === 3. 否定句判斷邏輯 (僅針對 currentLinePreceding) ===
            ; 正則說明：
            ; \b(no|without|negative|free|unremarkable|absence)\b  -> 否定關鍵字
            ; [^,.;:]*$  -> 否定詞與病徵之間，不能有「逗號、句號、分號、冒號」
            ;              (注意：這裡不需要放 \n，因為步驟 2 已經把換行濾掉了)

            isNegated := RegExMatch(currentLinePreceding, "i)\b(no|without|negative|neg|free|unremarkable|absence)\b[^,.;:]*$")

            if (!isNegated) {
                ; 找到了一個「非否定」的實例！
                ; 根據邏輯：只要文中有一處提到是陽性，就認定為陽性。
                return true
            }

            ; 如果這次發現的位置被判定為否定 (例如 "No scoliosis")
            ; 我們不回傳 false，而是更新 startPos，繼續往後找下一次出現的位置
            ; 因為下一行可能會寫 "Mild scoliosis is noted."
            startPos := foundPos + StrLen(match[0])
        }
    }
    ; 跑完所有關鍵字的所有出現位置，都沒有發現陽性描述，才回傳 false
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
