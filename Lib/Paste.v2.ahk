#Requires AutoHotkey v2.0

Paste(text, convertCRLF := true) {
    ; --- 0. 內容前處理 ---
    if (convertCRLF) {
        ; 確保換行符號是 Windows 標準 (CRLF)
        text := StrReplace(text, "`r`n", "`n")
        text := StrReplace(text, "`n", "`r`n")
    }

    ; --- 1. 嘗試取得焦點 Control 的 Hwnd ---
    hCtl := 0
    try {
        hCtl := ControlGetFocus("A")
    }

    ; --- 2. 策略 A: 直接訊息貼上 (優先使用) ---
    if (hCtl) {
        try {
            ; 使用 Hwnd 呼叫 EditPaste，完全繞過剪貼簿
            EditPaste(text, hCtl)
            return ; 成功則直接結束
        } catch {
            ; 失敗 (例如該 Control 不支援 Edit 訊息) 則繼續往下
        }
    }

    ; --- 3. 策略 B: 傳統剪貼簿貼上 (備案) ---

    ; 字數少直接打字 (長度門檻與 FastInput 一致設為 20)
    if (StrLen(text) < 20) {
        SendText(text)
        return
    }

    ; 剪貼簿貼上流程
    SavedClip := ClipboardAll()
    A_Clipboard := ""
    A_Clipboard := text

    if !ClipWait(1) {
        MsgBox "Clipboard failed to set."
        A_Clipboard := SavedClip
        return
    }

    SetKeyDelay 50, 50
    SendEvent "^v"

    Sleep 300
    A_Clipboard := SavedClip
}