#Requires AutoHotkey v2.0

/**
 * 通用剪貼簿與安全貼上函式 (Unified Safe Paste)
 * 整合訊息直通 (EditPaste)、短字串打字模擬 (SendText) 與剪貼簿還原貼上。
 *
 * @param {String} text 要輸出的文字
 * @param {Boolean} convertCRLF 是否自動標準化換行符號為 CRLF (預設 true)
 * @param {Integer|Object} targetCtrl 目標控制項 (支援 HWND、控制項名稱或 UIA Element 物件，預設 0 自動尋找焦點控制項)
 * @param {Integer} minLength 觸發剪貼簿貼上的最小字串長度 (預設 20，短於此長度直接打字模擬)
 */
Paste(text, convertCRLF := true, targetCtrl := 0, minLength := 20) {
    if (text == "") {
        return
    }

    ; --- 0. 內容前處理 ---
    if (convertCRLF) {
        ; 確保換行符號是 Windows 標準 (CRLF)
        text := StrReplace(text, "`r`n", "`n")
        text := StrReplace(text, "`n", "`r`n")
    }

    ; --- 1. 嘗試解析目標或焦點 Control 的 Hwnd ---
    hCtl := 0
    isUiaElement := false

    if (IsObject(targetCtrl)) {
        isUiaElement := true
        try {
            if (HasProp(targetCtrl, "NativeWindowHandle")) {
                hCtl := targetCtrl.NativeWindowHandle
            } else if (HasProp(targetCtrl, "Hwnd")) {
                hCtl := targetCtrl.Hwnd
            }
        }
    } else if (targetCtrl) {
        hCtl := targetCtrl
    } else {
        try {
            hCtl := ControlGetFocus("A")
        }
    }

    ; --- 2. 策略 A: 直接訊息貼上 (優先使用，達成 0 延遲且不污染剪貼簿) ---
    canDirectPaste := true
    if (isUiaElement && HasProp(targetCtrl, "FrameworkId")) {
        fw := targetCtrl.FrameworkId
        if (fw != "Win32" && fw != "WinForm") {
            canDirectPaste := false
        }
    }

    if (hCtl && canDirectPaste) {
        try {
            ; 使用 Hwnd 呼叫 EditPaste，完全繞過剪貼簿
            EditPaste(text, hCtl)
            return ; 成功則直接結束
        } catch {
            ; 失敗 (例如該 Control 不支援 Edit 訊息) 則繼續往下進入策略 B
        }
    }

    ; --- 3. 策略 B: 傳統焦點定位與剪貼簿貼上 (備案) ---
    if (isUiaElement && HasMethod(targetCtrl, "SetFocus")) {
        try {
            targetCtrl.SetFocus()
        } catch {
            _PasteNotify("無法聚焦目標欄位")
            return
        }
    }

    ; 字數少直接打字 (長度門檻與 FastInput 一致預設為 20)
    if (StrLen(text) < minLength) {
        SendText(text)
        return
    }

    ; 剪貼簿貼上流程
    SavedClip := ClipboardAll()
    A_Clipboard := ""
    A_Clipboard := text

    if !ClipWait(1) {
        _PasteNotify("剪貼簿設定失敗")
        A_Clipboard := SavedClip
        return
    }

    SetKeyDelay 50, 50
    SendEvent "^v"

    Sleep 300
    A_Clipboard := SavedClip
}

/**
 * 非阻塞式通知輔助函式 (移除 Modal MsgBox，避免干擾臨床操作流程)
 * @param {String} msg 提示訊息
 */
_PasteNotify(msg) {
    try {
        ctrlClass := (%("RisController")%)
        if (HasMethod(ctrlClass, "Notify")) {
            ctrlClass.Notify(msg, 1500)
            return
        }
    }
    try {
        notifyClass := (%("RisNotify")%)
        if (HasMethod(notifyClass, "Show")) {
            notifyClass.Show(msg, 1500)
            return
        }
    }
    ToolTip(msg)
    SetTimer(() => ToolTip(), -1500)
}
