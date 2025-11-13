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
Edit_FindText(hEdit, p_SearchText, p_Min := 0, p_Max := -1, p_Options := "", &r_RegExOut := "") {
    Static s_Text
    Static WM_GETTEXTLENGTH := 0xE

    ;-- Initialize
    r_RegExOut := "" ; v2: 寫入傳入的變數參考

    ; v1: if InStr(A_Space . p_Options . A_Space," Reset ")
    ; v2: A_Space 變為 " "
    if InStr(" " . p_Options . " ", " Reset ")
        s_Text := ""

    ;-- Bounce and return -1 if there is nothing to search
    if not StrLen(p_SearchText)
        Return -1

    ; v1: SendMessage ... / MaxLen:=ErrorLevel
    ; v2: SendMessage() 函式直接傳回值
    MaxLen := SendMessage(WM_GETTEXTLENGTH, 0, 0, , "ahk_id " . hEdit)
    if (MaxLen = 0)
        Return -1

    ;-- Parameters (v2 偏好用 || 而非 or)
    ; (註: 您的 v1 原始邏輯在這裡有點奇怪, 但我進行了 1:1 轉換)
    if (p_Min < 0 || p_Max > MaxLen)
        p_Min := MaxLen
    if (p_Max < 0 || p_Max > MaxLen)
        p_Max := MaxLen

    ;-- Anything to search? (v2 偏好用 == 進行比較)
    if (p_Min == p_Max)
        Return -1

    ;-- Get text
    if InStr(" " . p_Options . " ", " Static ") {
        if not StrLen(s_Text)
            s_Text := Edit_GetText(hEdit) ; 假設 Edit_GetText 已轉為 v2

        Text := SubStr(s_Text, (p_Max > p_Min) ? p_Min + 1 : p_Max + 1, (p_Max > p_Min) ? p_Max : p_Min)
    } else {
        s_Text := ""
        Text := Edit_GetTextRange(hEdit, (p_Max > p_Min) ? p_Min : p_Max, (p_Max > p_Min) ? p_Max : p_Min) ; 假設 Edit_GetTextRange 已轉為 v2
    }

    ;-- Look for it
    ; v1: ... ? True:False  (v2: InStr() 的傳回值 0 或 >0 本身就是 falsy/truthy)
    MatchCase := InStr(" " . p_Options . " ", " MatchCase ")
    RegEx := InStr(" " . p_Options . " ", " RegEx ")
    WholeWord := InStr(" " . p_Options . " ", " WholeWord ")

    if not (RegEx || WholeWord) {
        ; v1: ... (p_Max>p_Min) ? 1:0)-1  (v1 用 0 代表反向搜尋)
        ; v2: ... (p_Max>p_Min) ? 1:-1)-1 (v2 用 -1 代表反向搜尋)
        FoundPos := InStr(Text, p_SearchText, MatchCase, (p_Max > p_Min) ? 1 : -1) - 1
    }
    else { ;-- RegEx or Whole Word
        if RegEx {
            ; v1: RegExReplace(...,"",1) (v1 的第4參數是 OutputVarCount)
            ; v2: RegExReplace(...,"", 1) (v2 的第4參數是 &OutputVarCount, 第5是 Limit)
            ; 幸運的是, v2 語法會正確將 1 解讀為 Limit=1 (符合 v1 的可能意圖)
            p_SearchText := RegExReplace(p_SearchText, "^P\)?", "",, 1)
        } else { ;-- WholeWord
            p_SearchText := (MatchCase ? "" : "i)") . "\b" . p_SearchText . "\b"
        }

        ;--- v2: RegExMatch 錯誤會拋出例外, 必須使用 Try...Catch ---
        try {
            if (p_Max > p_Min) {
                ;-- Search forward
                ; v1: FoundPos:=RegExMatch(Text,p_SearchText,r_RegExOut,1)-1
                ; v2: &r_RegExOut
                FoundPos := RegExMatch(Text, p_SearchText, &r_RegExOut, 1) - 1
            }
            else {
                ;-- Search backward (***v2 大幅簡化***)
                ; v1: (一整個複雜的二分搜尋迴圈...)
                ; v2: RegExMatch() 原生支援反向搜尋 (StartingPos 設為 -1)
                FoundPos := RegExMatch(Text, p_SearchText, &r_RegExOut, -1) - 1
            }
        } catch as e {
            ; v1: outputdebug, (ltrim join`s ...
            ; v2: A_ThisFunc -> A_ThisFunc.Name, ErrorLevel -> e.Message
            OutputDebug(
                "Function: " . A_ThisFunc.Name . " - RegExMatch error.`n"
                . "Error: " . e.Message
            )
            FoundPos := -1
        }
    }

    ;-- Adjust FoundPos
    if (FoundPos > -1)
        FoundPos += (p_Max > p_Min) ? p_Min : p_Max

    Return FoundPos
}
Edit_ActivateParent(hEdit) {
    ;-- Get the handle to the parent window
    ; v1: DllCall(...) (v1/v2 函式語法相同)
    hParent := DllCall("GetParent", "UPtr", hEdit, "UPtr")

    ;-- Activate if needed
    ; v1: IfWinNotActive ahk_id %hParent%
    ; v2: 使用 !WinActive() 函式, 並串接 ahk_id
    if !WinActive("ahk_id " . hParent) {
        ; v1: WinActivate ahk_id %hParent%
        ; v2: 使用 WinActivate() 函式
        WinActivate("ahk_id " . hParent)

        ;-- Still not active? (rare)
        ; v1: IfWinNotActive ahk_id %hParent%
        if !WinActive("ahk_id " . hParent) {
            ;-- Give the window an additional 250 ms to activate
            ; v1: WinWaitActive ..., if ErrorLevel (timeout)
            ; v2: WinWaitActive() 函式在超時 (timeout) 時傳回 0 (falsy)
            if !WinWaitActive("ahk_id " . hParent, , 0.25)
                Return false
        }
    }

    Return true
}
Edit_HasFocus(hEdit) {
    Static GUITHREADINFO, sizeofGUITHREADINFO

    ;-- Create and initialize GUITHREADINFO structure (one-time init)
    if !IsSet(GUITHREADINFO) {
        sizeofGUITHREADINFO := A_PtrSize = 8 ? 72 : 48
        GUITHREADINFO := Buffer(sizeofGUITHREADINFO, 0)
        NumPut("UInt", sizeofGUITHREADINFO, GUITHREADINFO, 0)
    }

    ;-- Collect GUI Thread Info
    ; v1: DllCall(..., "UPtr", &GUITHREADINFO)
    ; v2: 必須傳遞 .Ptr (位址, 一個 Number), 以滿足 DllCall 的要求
    ; --- 這是修正後的程式碼 ---
    if DllCall("GetGUIThreadInfo", "UInt", 0, "UPtr", GUITHREADINFO.Ptr) {

        ; 奇怪的是, NumGet() 函式 *確實* 接受 Buffer 物件本身
        Return (hEdit = NumGet(GUITHREADINFO, A_PtrSize = 8 ? 16 : 12, "UPtr"))
               ;-- hwndFocus
    }

    ;-- Error
    OutputDebug(
        "Function: " . A_ThisFunc.Name . " - `n"
        . "Call to GetGUIThreadInfo failed. A_LastError: " . A_LastError
    )
    Return false
}
Edit_SetFocus(hEdit, p_ActivateParent := false) {
    ;-- If requested, activate parent
    ; v1: if p_ActivateParent
    ; v2: 相同, 但 v1 的 'False' 變為 v2 的 'false'
    if p_ActivateParent {
        ; 這假設 Edit_ActivateParent 函式也已轉換為 v2
        if not Edit_ActivateParent(hEdit)
            Return false
    }

    ;-- Does the control already have focus?
    ; 這假設 Edit_HasFocus 函式也已轉換為 v2
    ; v1: Return True
    ; v2: v1 的 'True' 變為 v2 的 'true'
    if Edit_HasFocus(hEdit)
        Return true

    ;-- Set focus
    ; v1: ControlFocus,,ahk_id %hEdit%
    ; v1: Return ErrorLevel ? False:True
    ;
    ; v2: ControlFocus() 是一個函式, 它直接傳回 true (成功) 或 false (失敗)
    ;     這完美地取代了 v1 中 "檢查 ErrorLevel" 的邏輯
    Return ControlFocus(hEdit)
}
Edit_GetLineCount(hEdit) {
    Static EM_GETLINECOUNT := 0xBA

    ; 這個訊息傳回的就是邏輯行數 (CRLF count + 1),
    ; 而不是視覺行數 (visual wrapped line count)。
    Return SendMessage(EM_GETLINECOUNT, 0, 0, , "ahk_id " . hEdit)
}
/*
 * Edit_GetLogicalLineCount
 *
 * 終極版:
 * 由於控制項 (可能是 RichEdit 或自訂控制項)
 * 在 EM_GETLINECOUNT 和 EM_LINEFROMCHAR 訊息上
 * 都會「說謊」(傳回視覺行數),
 * 我們被迫採用 100% 可靠的「取得全文並手動計算」方法。
 */
Edit_GetLogicalLineCount(hEdit) {

    ; 1. 取得完整的文字內容。
    ;    我們假設 Edit_GetText(hEdit) 函式 (使用 WM_GETTEXT)
    ;    能正確抓取文字。
    local text := Edit_GetText(hEdit)

    ; 2. 根據 WM_GETTEXT 的標準, 它會將所有換行符 `n`
    ;    轉換為 `r`n (CRLF)。
    ;    因此, 我們使用 `r`n 作為分隔符。

    ; 3. AHK v2 的 StrSplit() 非常高效。
    local lines := StrSplit(text, "`r`n")

    ; 4. .Length 屬性就是我們要的答案。
    ;    - StrSplit("abc")       -> ["abc"]       -> .Length = 1
    ;    - StrSplit("abc`r`nxyz") -> ["abc", "xyz"] -> .Length = 2
    ;    - StrSplit("")          -> [""]          -> .Length = 1 (空控制項也是 1 行)
    Return lines.Length
}
Edit_ScrollCaret(hEdit) {
    Static EM_SCROLLCARET := 0xB7
    SendMessage(EM_SCROLLCARET, 0, 0, , "ahk_id " . hEdit)
}