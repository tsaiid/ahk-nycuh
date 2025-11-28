#Requires AutoHotkey v2.0

/**
 * 重排選取文字 (無剪貼簿版)
 * @param deOrder 移除序號
 * @param keepEmptyLine 保留空行
 * @param itemChar 項目符號
 * @param discardSeIm 移除 Series/Image 標記
 * @param targetHwnd [選填] 目標 Control 的 Handle。如果有傳入，則完全不使用剪貼簿。
 */
ReorderSelectedText(deOrder := false, keepEmptyLine := false, itemChar := "", discardSeIm := true, targetHwnd := 0) {

    local selectedText := ""

    ; --- 1. 取得文字 (Input) ---
    if (targetHwnd) {
        ; [新方法] 直接從 Control 讀取選取文字，不經剪貼簿
        try {
            selectedText := EditGetSelectedText(targetHwnd)
        } catch {
            MsgBox "無法讀取選取文字 (EditGetSelectedText 失敗)。"
            return -1
        }
    } else {
        ; [舊方法相容] 如果沒傳 Handle，才退回使用剪貼簿 (保留給其他視窗使用)
        local ClipSaved := ""
        ClipSaved := ClipboardAll()

        A_Clipboard := ""
        Send "^c"
        if !ClipWait(0.8) {
            MsgBox "複製文字失敗 (逾時)。"
            A_Clipboard := ClipSaved
            return -1
        }
        selectedText := A_Clipboard
    }

    ; --- 2. 文字處理 (邏輯完全保留) ---
    ; 為了安全起見，先標準化換行
    selectedText := StrReplace(selectedText, "`r`n", "`n")
    if (InStr(selectedText, "`r")) {
        ; 這裡如果直接讀取 Control，通常不會有單獨 `r 的問題，但保留檢查無妨
        MsgBox "選取範圍內包含不正確的換行符號 (CR)。"
        if (!targetHwnd && IsSet(ClipSaved))
            A_Clipboard := ClipSaved
        return -1
    }

    local hadTrimmedRight := false
    if (SubStr(selectedText, -1) == "`n") {
        selectedText := SubStr(selectedText, 1, -1)
        hadTrimmedRight := true
    }

    local txtAry := StrSplit(selectedText, "`n")
    local endLine := txtAry.Length
    local finalText := ""

    if (StrLen(selectedText) > 0) {
        local isSpine := false
        local isFirstLineEmpty := false
        local startLineNo := 1

        if (RegExMatch(selectedText, "^(\d+)", &existLineNo)) {
            startLineNo := existLineNo[1]
        }

        for index, line in txtAry {
            if (index == 1 && !StrLen(line)) {
                isFirstLineEmpty := true
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

                ; 正規表達式替換
                finalText .= RegExReplace(
                    tmpText,
                    "^(\s*)((\d+\.)|([-\+\*>=])|(\(?\d+\)))?(\s*)(\w?)(.*)",
                    "$u{7}${8}"
                )

                if (index < endLine || hadTrimmedRight) {
                    finalText .= "`r`n"
                }
            } else {
                if (isFirstLineEmpty && index == 1) {
                    finalText .= "`r`n"
                }
                if (keepEmptyLine) {
                    finalText .= "`r`n"
                }
            }
        }
    } else {
        ; 無內容
        if (!targetHwnd && IsSet(ClipSaved))
            A_Clipboard := ClipSaved
        return -1
    }

    ; --- 3. 輸出 (Output) ---
    if (targetHwnd) {
        ; [新方法] 直接發送字串替換選取範圍 (不經剪貼簿)
        ; EditPaste 在 v2 中會自動送出 EM_REPLACESEL 訊息
        try {
            EditPaste(finalText, targetHwnd)
        } catch as err {
            MsgBox "寫入失敗: " err.Message
        }
    } else {
        ; [舊方法相容] 剪貼簿貼上
        A_Clipboard := finalText
        Sleep 100
        Send "^v"

        if (IsSet(ClipSaved)) {
            Sleep 200 ; 等待貼上完成再還原
            A_Clipboard := ClipSaved
        }
    }

    return 0
}
