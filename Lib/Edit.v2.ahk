Edit_GetText(hEdit, p_Length := -1) {
    ; v2: 靜態變數宣告更簡潔，移除了 v1 中無用的 Dummy 變數
    Static WM_GETTEXT := 0xD
    Static WM_GETTEXTLENGTH := 0xE

    ;-- If needed, determine the length of the text
    ; v2: 'if' 後の ( ) 可省略
    if p_Length < 0 {
        ; v1: SendMessage ... (ErrorLevel 接收返回值)
        ; v2: SendMessage() 函式直接傳回值
        p_Length := SendMessage(WM_GETTEXTLENGTH, 0, 0, , "ahk_id " . hEdit)
    }

    ;-- Add 1 to the length for a trailing null character.
    p_Length += 1

    ;-- Get text
    ; v1: VarSetCapacity(Text, p_Length * (A_IsUnicode ? 2:1))
    ; v2: 使用 Buffer() 取代 VarSetCapacity。v2 永遠是 Unicode (UTF-16)，
    ;     因此大小固定為 (字元數 * 2 bytes)
    TextBuf := Buffer(p_Length * 2)

    ; v1: SendMessage WM_GETTEXT, p_Length, &Text, ...
    ; v2: 將 Buffer 物件 (TextBuf) 直接傳遞給 SendMessage 作為 LParam
    SendMessage(WM_GETTEXT, p_Length, TextBuf, , "ahk_id " . hEdit)

    ; v1: Return Text
    ; v2: 使用 StrGet() 從 Buffer 中讀取字串
    Return StrGet(TextBuf)
}
Edit_SetSel(hEdit, p_StartSelPos := 0, p_EndSelPos := -1) {
    Static EM_SETSEL := 0xB1
    SendMessage(EM_SETSEL, p_StartSelPos, p_EndSelPos, , "ahk_id " . hEdit)
}
Edit_SetText(hEdit, p_Text, p_SetModify := False) {
    Static WM_SETTEXT := 0xC

    ; v1: SendMessage WM_SETTEXT, 0, &p_Text, ... (回傳值在 ErrorLevel)
    ; v2: SendMessage() 直接傳回值，並自動處理 p_Text 字串的指標
    STRC := SendMessage(WM_SETTEXT, 0, StrPtr(p_Text), , "ahk_id " . hEdit)

    ; v1: if STRC:=ErrorLevel
    ; v2: STRC 已被賦值，直接判斷 (WM_SETTEXT 成功時通常傳回 1)
    if STRC {
        if p_SetModify {
            ; 假設 Edit_SetModify 函式也已存在或已轉換為 v2
            Edit_SetModify(hEdit, True)
        }
    }

    Return STRC ;-- Return code from the WM_SETTEXT message
}
Edit_SetModify(hEdit, p_Flag) {
    Static EM_SETMODIFY := 0xB9

    ; v1: SendMessage EM_SETMODIFY, p_Flag, 0, , ahk_id %hEdit%
    ; v2: SendMessage() 函式化, 並串接 ahk_id
    SendMessage(EM_SETMODIFY, p_Flag, 0, , "ahk_id " . hEdit)
}
Edit_GetSel(hEdit, &r_StartSelPos := "", &r_EndSelPos := "") {
    Static EM_GETSEL := 0xB0

    ; v1: VarSetCapacity(s_StartSelPos,4) / VarSetCapacity(s_EndSelPos,4)
    ; v2: 建立 4-byte (32-bit) Buffer 來接收 DWORD (UInt) 值
    local s_StartSelPos := Buffer(4)
    local s_EndSelPos := Buffer(4)

    ; v1: SendMessage EM_GETSEL, &s_StartSelPos, &s_EndSelPos, , ahk_id %hEdit%
    ; v2: WParam 和 LParam 直接傳遞 Buffer 物件
    SendMessage(EM_GETSEL, s_StartSelPos, s_EndSelPos, , "ahk_id " . hEdit)

    ; v1: r_StartSelPos := NumGet(s_StartSelPos, 0, "UInt")
    ; v2: NumGet 語法相同, 從 Buffer 中讀取資料
    r_StartSelPos := NumGet(s_StartSelPos, 0, "UInt")
    r_EndSelPos := NumGet(s_EndSelPos, 0, "UInt")

    Return r_StartSelPos
}
Edit_Clear(hEdit) {
    Static WM_CLEAR := 0x303

    ; v1: SendMessage WM_CLEAR, 0, 0, , ahk_id %hEdit%
    ; v2: SendMessage() 函式化, 並串接 ahk_id
    SendMessage(WM_CLEAR, 0, 0, , "ahk_id " . hEdit)
}
Edit_GetTextRange(hEdit, p_StartPos := 0, p_EndPos := -1) {
    ;-- Parameters
    ; v1: if p_StartPos is not Integer
    ; v2: 使用 IsNumber() 函式 (這比 v1 的 'is Integer' 更靈活, 也更符合 AHK 的精神)
    if (not IsNumber(p_StartPos)) || (p_StartPos < 0) {
        p_StartPos := 0
    }

    ; v1: if p_EndPos is not Integer
    ; v2: 使用 IsNumber()
    if not IsNumber(p_EndPos) {
        p_EndPos := -1
    }

    ;-- Get text range
    ; v1: Return SubStr(Edit_GetText(hEdit,p_EndPos),p_StartPos+1)
    ; v2: 語法完全相同。
    ;     這假設 Edit_GetText() 函式也已成功轉換為 v2。
    Return SubStr(Edit_GetText(hEdit, p_EndPos), p_StartPos + 1)
}
Edit_GetTextLength(hEdit) {
    Static WM_GETTEXTLENGTH := 0xE

    ; v1: SendMessage ... (回傳值在 ErrorLevel)
    ; v2: SendMessage() 函式直接傳回值
    Return SendMessage(WM_GETTEXTLENGTH, 0, 0, , "ahk_id " . hEdit)
}
Edit_ReplaceSel(hEdit, p_Text := "", p_CanUndo := true) {
    Static EM_REPLACESEL := 0xC2

    ; v1: SendMessage EM_REPLACESEL, p_CanUndo, &p_Text, , ahk_id %hEdit%
    ; v2: SendMessage() 函式化, 並自動處理 p_Text 的指標
    SendMessage(EM_REPLACESEL, p_CanUndo, StrPtr(p_Text), , "ahk_id " . hEdit)
}