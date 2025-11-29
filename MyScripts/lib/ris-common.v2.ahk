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

class RisController {
    ; =================================================================
    ; 1. 設定區 (Configuration) - 未來擴充元件都在這裡加
    ; =================================================================
    static WinTitle := RISReportWinTitle

    ; 定義所有子元件的搜尋條件
    static Selectors := Map(
        "AutoNextCheckbox", { AutomationId: "chkAutoNext" },
        "ReportSaveButton", { AutomationId: "btnReportSave" },
        "FindingEdit", { AutomationId: "txtReport" },
        "ImpressionEdit", { AutomationId: "txtImpression" },
        "PastAllRadio", { AutomationId: "rdoPastALL" },
        "PastModalityRadio", { AutomationId: "rdoClassify" },
        "PastOnlyMyRadio", { AutomationId: "rdoPastOnlyMy" },
        "ExamnameText", { AutomationId: "txtExamName" },
        "PastFindingText", { AutomationId: "rtxtPastReport" },
        "PastImpressionText", { AutomationId: "rtxtPastImpression" },
        "PastReportTable", { AutomationId: "dgvPastReport" },
        "PathoDiagnosisText", { AutomationId: "txtDiagnosist" },
        "PathoDateText", { AutomationId: "mtxtRcpDTM" },
        "ImpressionLabel", { AutomationId: "label2" },
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
    static _GetOrUpdateNode(nodeName) {
        ; [步驟 1] 取得當前真實視窗的 ID
        ; 如果連視窗都沒開，直接在這裡就會報錯，這也是一種檢查
        currentHwnd := WinExist(this.WinTitle)
        if !currentHwnd
            throw TargetError("找不到 RIS 視窗，請確認程式已開啟。")

        ; [步驟 2] 檢查快取是否存在
        if this._cache.Has(nodeName) {
            el := this._cache[nodeName]

            try {
                ; === 關鍵修正：針對 Root (Ris) 進行嚴格檢查 ===
                if (nodeName = "Ris") {
                    ; 如果快取中的 Element 的視窗 ID 不等於現在的視窗 ID
                    ; 代表視窗已經重開過，這個快取是舊視窗的「屍體」
                    if (el.WindowId != currentHwnd)
                        throw Error("視窗 ID 不匹配 (應用程式可能已重啟)")
                }

                ; 一般的 Probe 測試 (確保元件沒壞)
                temp := el.ControlType
                return el ; 驗證通過，回傳快取
            }
            catch {
                ; 驗證失敗 (可能是視窗重開、元件消失)，清除此快取
                ; DebugLog("快取失效: " nodeName)
                this._cache.Delete(nodeName)

                ; 如果是父層失效，最好連子層快取也清空，避免連鎖反應
                if (nodeName = "Ris")
                    this._cache := Map()
            }
        }

        ; [步驟 3] 快取無效，重新抓取

        ; A. 抓取 Root (Ris)
        if (nodeName = "Ris") {
            try {
                ; 確保使用當前最新的 Hwnd
                newRoot := UIA.ElementFromHandle(currentHwnd)
                this._cache["Ris"] := newRoot
                return newRoot
            } catch as err {
                throw Error("無法取得 RIS Root Element: " err.Message)
            }
        }

        ; B. 抓取子元件
        else {
            ; 遞迴呼叫：這會自動觸發上面的 "Ris" 檢查
            ; 如果 Ris 快取剛被清空，這裡會自動重新抓一個新的 Ris
            parent := this.Ris

            if !this.Selectors.Has(nodeName)
                throw Error("未定義: " nodeName)

            try {
                newChild := parent.FindElement(this.Selectors[nodeName])
                this._cache[nodeName] := newChild
                return newChild
            } catch {
                ; 這裡可以做一個保險：如果找不到子元件，有沒有可能是父元件其實也壞了?
                ; 但因為上面的邏輯已經嚴格檢查過父元件，這裡單純找不到的機率較高。
                throw TargetError("找不到元件: " nodeName)
            }
        }
    }

    ; =================================================================
    ; 通用文字取得方法 (支援所有 Element)
    ; 用法: text := RisController.GetText(RisController.ReportContent)
    ; =================================================================
    static GetText(el) {
        ;MsgBox(el.FrameworkId)
        rawText := ""
        isNativeSuccess := false

        ; [策略 A] 優先嘗試原生 ControlGetText (針對 Win32 Edit Control)
        try {
            ; 1. 取得該 Element 的原生視窗 Handle
            ;    注意：屬性名稱是 NativeWindowHandle
            hwnd := el.NativeWindowHandle

            ; 2. 檢查 Handle 是否有效，且框架是否為 Win32
            ; WinForm 和 Win32 的控制項都有獨立 Handle，適合用 ControlGetText
            if (hwnd && (el.FrameworkId = "Win32" || el.FrameworkId = "WinForm")) {
                rawText := ControlGetText(hwnd)
                isNativeSuccess := true
            }
        }
        catch {
            ; 忽略 ControlGetText 的錯誤，繼續往下嘗試
        }

        ; [策略 B] 如果原生讀取失敗或不適用，使用 UIA 屬性
        if (!isNativeSuccess) {
            try {
                ; -----------------------------------------------------------
                ; 修正點：不要用 IsPatternSupported("Value") 做硬性阻擋
                ; 直接嘗試讀取 .Value，讓函式庫自己去決定是用 ValuePattern 還是 LegacyIAccessible
                ; -----------------------------------------------------------
                rawText := el.Value
            }
            catch {
                ; 如果 .Value 報錯，代表兩者都不支援，這裡什麼都不做，繼續往下試
            }

            ; 如果上面沒拿到值，再試試看 TextPattern (針對 Document)
            if (rawText == "" && el.IsPatternSupported("Text")) {
                try {
                    rawText := el.DocumentRange.GetText()
                }
            }

            ; 最後試試 Name
            if (rawText == "") {
                try {
                    rawText := el.Name
                }
            }
        }

        ; [策略 C] 最終標準化：強制統一換行符號為 Windows 格式 (CRLF)
        if (rawText != "") {
            ; 第一步：把 Windows 標準的 `r`n 轉成 `n
            temp := StrReplace(rawText, "`r`n", "`n")

            ; 第二步：【關鍵修正】把剩餘單獨的 `r 也轉成 `n
            ; 這是為了解決 WinForm/Legacy 有時只回傳 CR 的問題
            temp := StrReplace(temp, "`r", "`n")

            ; 第三步：現在所有換行都統一變成 `n 了，再一次性轉成 `r`n
            return StrReplace(temp, "`n", "`r`n")
        }

        return ""
    }

    ; =================================================================
    ; 2. 專用貼上方法 (外部呼叫用)
    ; =================================================================

    ; 貼上到 Finding 區
    static PasteToFinding(text) {
        this.PasteTo(this.FindingEdit, text)
    }

    ; 貼上到 Impression 區
    static PasteToImpression(text) {
        this.PasteTo(this.ImpressionEdit, text)
    }

    ; =================================================================
    ; 3. 核心貼上邏輯 (整合了 EditPaste 與 Clipboard)
    ; =================================================================
    static PasteTo(targetEl, text)
    {
        if (text == "")
            return

        ; --- 步驟 A: 文字標準化 (確保換行符號為 CRLF) ---
        ; 這是 EditPaste 的硬性要求，否則會擠成一行
        text := StrReplace(text, "`r`n", "`n")
        text := StrReplace(text, "`n", "`r`n")

        ; --- 步驟 B: 嘗試使用原生 EditPaste (最優先、最快) ---
        isNativeSuccess := false
        try {
            hwnd := targetEl.NativeWindowHandle

            ; 確保有 Handle 且框架支援 (Win32/WinForm)
            if (hwnd && (targetEl.FrameworkId == "Win32" || targetEl.FrameworkId == "WinForm")) {
                EditPaste(text, hwnd)
                isNativeSuccess := true
            }
        }
        catch {
            ; 忽略錯誤，準備進入備案
        }

        if (isNativeSuccess)
            return

        ; --- 步驟 C: 備案 - 使用剪貼簿貼上 ---
        ; 如果無法使用 EditPaste，必須先確保該欄位「取得焦點」
        try {
            targetEl.SetFocus()
        } catch {
            MsgBox "無法聚焦目標欄位，貼上失敗。"
            return
        }

        ; 如果字數很少，直接打字 (避開剪貼簿鎖定風險)
        if (StrLen(text) < 50) {
            SendText(text)
            return
        }

        ; 剪貼簿標準流程
        SavedClip := ClipboardAll()
        A_Clipboard := ""
        A_Clipboard := text

        if !ClipWait(1) {
            MsgBox "複製到剪貼簿失敗 (Timeout)。"
            A_Clipboard := SavedClip
            return
        }

        SetKeyDelay 50, 50
        SendEvent "^v"

        Sleep 300 ; 等待目標程式消化
        A_Clipboard := SavedClip
    }

    ; =================================================================
    ; 4. 焦點檢查 helper (供 #HotIf 使用)
    ; =================================================================
    static IsTargetFocused()
    {
        ; 1. 取得目前 Windows 焦點所在的 Hwnd
        try {
            focusedHwnd := ControlGetFocus("A")
        } catch {
            return false
        }

        ; 2. 比對 FindingEdit 的 Handle
        try {
            ; 存取 this.FindingEdit 會觸發 _GetOrUpdateNode
            ; 如果快取還沒建立，這裡會自動建立；如果已建立，讀取非常快
            if (this.FindingEdit.NativeWindowHandle == focusedHwnd)
                return true
        }

        ; 3. 比對 ImpressionEdit 的 Handle
        try {
            if (this.ImpressionEdit.NativeWindowHandle == focusedHwnd)
                return true
        }

        return false
    }

    ; =================================================================
    ; 字體強制模組 (Font Enforcer)
    ; =================================================================
    static _hCustomFont := 0 ; 用來存放 Fira Code 的 Handle

    /**
     * 啟動字體強制功能
     * @param fontName 字體名稱 (預設 "Fira Code")
     * @param fontSize 字體大小 (預設 12)
     */
    static EnableFontEnforcer(fontName := "Cascadia Code", fontSize := 12)
    {
        ; 1. 建立 Font Handle (只做一次，避免記憶體洩漏)
        if (this._hCustomFont == 0) {
            ; 利用一個隱藏的 GUI 來產生合法的 HFONT
            dummyGui := Gui()
            dummyGui.SetFont("s" fontSize, fontName)
            dummyCtrl := dummyGui.Add("Text",, "Dummy")

            ; 發送 WM_GETFONT (0x31) 取得該控制項的 Font Handle
            this._hCustomFont := SendMessage(0x31, 0, 0, dummyCtrl.Hwnd)

            ; 注意：不要 Destroy dummyGui，因為 Font Handle 依附於它
            ; 或是如果要嚴謹一點，應該用 DllCall CreateFont，但這樣寫最簡單
        }

        ; 2. 啟動 Timer，每 1000 ms (1秒) 檢查一次
        ; 使用 ObjBindMethod 將類別內的方法綁定給 Timer
        SetTimer(ObjBindMethod(this, "_EnforceFontTask"), 1000)
    }

    ; =================================================================
    ; 自動排版模組 (Auto Layout)
    ; =================================================================

    ; 設定 Impression 區域的高度 (單位: 像素)
    static _targetImpressionHeight := 95

    ; 綁定的 Timer 任務
    static _EnforceFontTask()
    {
        if !WinActive(this.WinTitle)
            return

        static WM_SETFONT := 0x30

        ; 1. 取得所有必要的 Handles
        try {
            hFind := this.FindingEdit.NativeWindowHandle
            hImp  := this.ImpressionEdit.NativeWindowHandle
        } catch {
            return ; 如果還沒載入完畢，先不動作
        }

        ; 2. 套用字體 (Fira Code)
        if (this._hCustomFont) {
            try SendMessage(WM_SETFONT, this._hCustomFont, 1, , "ahk_id " hFind)
            try SendMessage(WM_SETFONT, this._hCustomFont, 1, , "ahk_id " hImp)
        }

        ; 3. 執行排版運算 (Layout Calculation)
        this._ApplyLayout(hFind, hImp)
    }

    static _ApplyLayout(hFind, hImp)
    {
        ; 取得目前控制項的位置與大小
        ControlGetPos(&fX, &fY, &fW, &fH, hFind)
        ControlGetPos(&iX, &iY, &iW, &iH, hImp)

        ; --- 計算錨點 (Anchor) ---
        ; 我們假設目前的 Impression 底部是正確的邊界 (畫面的最下方)
        currentBottom := iY + iH

        ; --- 設定目標參數 ---
        targetImpH := this._targetImpressionHeight ; 您希望的 Impression 高度 (5行)
        gap := 30 ; Finding 和 Impression 之間的間距 (留給標籤用)

        ; 計算 Impression 的新 Y 軸 (底部錨點 - 目標高度)
        targetImpY := currentBottom - targetImpH

        ; 計算 Finding 的新高度 (填滿上方剩餘空間)
        ; 新高度 = (Impression的新Y位置 - 間距) - Finding原本的Y位置
        targetFindH := (targetImpY - gap) - fY

        ; --- 檢查是否需要移動 (避免重複執行導致閃爍) ---
        ; 容許 5px 的誤差
        if (Abs(iH - targetImpH) < 5 && Abs(iY - targetImpY) < 5 && Abs(fH - targetFindH) < 5)
            return

        ; --- 開始移動 ---

        ; 1. 移動 Impression Edit (至底部，變矮)
        ControlMove(,, iW, targetImpH, hImp) ; 只改 Y 和 H
        ControlMove(, targetImpY,,, hImp)    ; 分開寫比較穩定

        ; 2. 移動 Finding Edit (變高)
        ControlMove(,,, targetFindH, hFind)

        ; 3. 移動 Impression 標籤 (Label)
        ; 嘗試透過 UIA 抓取 Label Handle
        try {
            ; 嘗試從 Cache 抓，沒有就抓新的
            elLabel := this._GetOrUpdateNode("ImpressionLabel")
            hLabel := elLabel.NativeWindowHandle

            if (hLabel) {
                ; 標籤應該放在 Impression Edit 的正上方
                ; 假設標籤高度約 20px
                labelNewY := targetImpY - 25
                ControlMove(, labelNewY,,, hLabel)
            }
        } catch {
            ; 如果抓不到標籤，不做動作 (但 Edit 已經移好了)
        }
    }

    /**
     * 刪除游標所在的整行 (包含換行符號)
     * @returns {Boolean} true 代表已執行刪除; false 代表不在目標區，未執行
     */
    static DeleteCurrentLine() {
        ; 1. 檢查焦點是否在 Finding 或 Impression
        if !this.IsTargetFocused()
            return false

        ; 2. 取得焦點的 Handle
        try {
            hFocus := ControlGetFocus("A")
            if !hFocus
                return false
        } catch {
            return false
        }

        ; 3. 執行選取並刪除
        ; 使用新的 Win32 算法，不需讀取整段文字
        this._SelectLineWin32(hFocus)

        ; 4. 發送清除指令 (WM_CLEAR = 0x0303)
        SendMessage(0x0303, 0, 0, hFocus)

        return true
    }

    /**
     * 刪除前一個 Word (Bash Style: 以前方空白為界)
     * @returns {Boolean} true: 已執行刪除; false: 不在目標區，未執行
     */
    static DeleteWordBackward() {
        ; 1. 檢查焦點
        if !this.IsTargetFocused()
            return false

        try {
            hCtrl := ControlGetFocus("A")
            if !hCtrl
                return false
        } catch {
            return false
        }

        ; 2. 執行核心演算法
        this._BashDeleteAlgo(hCtrl)
        return true
    }

    ; --- 私有演算法實作 ---
    static _BashDeleteAlgo(hCtrl) {
        try {
            fullText := ControlGetText(hCtrl)
        } catch {
            return
        }

        ; 取得當前游標位置 (0-based)
        ; 0x00B0 = EM_GETSEL
        caretPosRaw := SendMessage(0x00B0, 0, 0, hCtrl)
        caretPos := caretPosRaw & 0xFFFF

        if (caretPos == 0)
            return ; 已經在最前面

        ; 轉成 AHK 1-based 索引
        i := caretPos

        ; --- 階段 A: 吃掉緊貼游標的空白 ---
        ; "Hello   |" -> "Hello|"
        while (i > 0) {
            char := SubStr(fullText, i, 1)
            if (this._IsSpace(char)) {
                i--
            } else {
                break
            }
        }

        ; --- 階段 B: 吃掉單字，直到遇到下一個空白 ---
        ; "Hello|" -> "|" (遇到 Hello 前面的空白停止)
        while (i > 0) {
            char := SubStr(fullText, i, 1)
            if (!this._IsSpace(char)) {
                i--
            } else {
                break
            }
        }

        ; 計算選取範圍
        selStart := i
        selEnd := caretPos

        ; 執行刪除
        ; 0x00B1 = EM_SETSEL
        SendMessage(0x00B1, selStart, selEnd, hCtrl)

        ; 0x0303 = WM_CLEAR
        SendMessage(0x0303, 0, 0, hCtrl)
    }

    ; --- 私有輔助函數 ---
    static _IsSpace(char) {
        return (char == " " || char == "`t" || char == "`r" || char == "`n")
    }

    ; =================================================================
    ; 5. [優化] Win32 輔助方法
    ; =================================================================

    /**
     * 使用 Win32 API 快速選取整行 (比字串分析快且穩)
     * 邏輯：選取「本行開頭」到「下一行開頭」之間的所有內容
     */
    static _SelectLineWin32(hCtrl) {
        static EM_LINEFROMCHAR := 0x00C9
        static EM_LINEINDEX    := 0x00BB
        static EM_SETSEL       := 0x00B1

        ; 1. 取得目前游標在哪一行 (0-based)
        ; -1 代表使用目前游標位置
        lineIdx := SendMessage(EM_LINEFROMCHAR, -1, 0, hCtrl)

        ; 2. 取得這一行的第一個字元索引 (Start)
        lineStart := SendMessage(EM_LINEINDEX, lineIdx, 0, hCtrl)

        ; 3. 取得「下一行」的第一個字元索引 (End)
        ; 如果下一行存在，這就會包含本行的換行符號 (`r`n)
        nextLineStart := SendMessage(EM_LINEINDEX, lineIdx + 1, 0, hCtrl)

        if (nextLineStart == -1) {
            ; 狀況 A: 沒有下一行了 (這是最後一行)
            ; 我們就選到文字的最末端
            ; 這裡偷懶用一個很大的數字代表「到底」，EM_SETSEL 支援這種寫法
            ; 或者您也可以用 GetTextLength 來算，但 -1 通常有效
            selStart := lineStart
            selEnd   := -1  ; 選到最後
        } else {
            ; 狀況 B: 有下一行
            ; 選取範圍就是 [本行開頭, 下一行開頭)
            selStart := lineStart
            selEnd   := nextLineStart
        }

        ; 4. 執行選取
        SendMessage(EM_SETSEL, selStart, selEnd, hCtrl)
    }
}
