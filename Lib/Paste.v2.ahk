#Requires AutoHotkey v2.0

Paste(text, convertCRLF := true) {
    ; 【閾值設定】
    ; 如果字數少於 50 字，直接用打字的 (毫無延遲)
    ; 如果字數多於 50 字，才用剪貼簿貼上 (避免長文章打字太久)
    if (StrLen(text) < 50) {
        SendText(text)
    }
    else {
        If (convertCRLF) {
            text := RegExReplace(text, "(?<!\r)\n", "`r`n")
        }

        ; 備份舊的剪貼簿內容 (選擇性，如果覺得拖慢速度可註解掉下面這行)
        SavedClip := ClipboardAll()

        A_Clipboard := text

        ; 等待剪貼簿準備好，只要等到有內容就立刻送出，不必死板的 Sleep
        if !ClipWait(0.5) {
            MsgBox("Clipboard failed to set.")
            return
        }

        Send("^v")

        ; 給系統一點時間處理貼上動作，再恢復剪貼簿
        Sleep(100)

        ; 恢復舊剪貼簿 (選擇性)
        A_Clipboard := SavedClip
    }
}