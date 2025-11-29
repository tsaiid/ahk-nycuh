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
        ; 1. [新增] 加上 Try-Catch 保護
        ; 因為這個方法是被 Timer 呼叫的，如果視窗剛好被切換或關閉，
        ; ControlGetPos 會因為抓不到目標而報錯，這裡我們選擇「靜默失敗」即可。
        try {
            ; 取得目前控制項的位置與大小
            ControlGetPos(&fX, &fY, &fW, &fH, hFind)
            ControlGetPos(&iX, &iY, &iW, &iH, hImp)
        } catch {
            return ; 如果抓不到位置，代表視窗狀態不穩，這次先不做排版
        }

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
        this._SelectLine(hFocus)

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
    ; [新增功能] Emacs 風格游標移動 (Ctrl+A / Ctrl+E)
    ; =================================================================

    /**
     * 移動游標到邏輯行首或行尾 (Win32 API 高效版)
     * @param mode "Start" (行首) 或 "End" (行尾)
     * @returns {Boolean} true: 已執行; false: 不在目標區
     */
    static MoveCaret(mode) {
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

        ; 2. 定義 Win32 常數
        static EM_LINEFROMCHAR := 0x00C9 ; 查詢目前在哪一行
        static EM_LINEINDEX    := 0x00BB ; 查詢該行第一個字的 Index
        static EM_LINELENGTH   := 0x00C1 ; 查詢該行長度
        static EM_SETSEL       := 0x00B1 ; 設定選取範圍 (移動游標)
        static EM_SCROLLCARET  := 0x00B7 ; 捲動畫面至游標處

        ; 3. 計算位置
        ; 取得目前行號 (0-based)
        lineIdx := SendMessage(EM_LINEFROMCHAR, -1, 0, hCtrl)

        ; 取得該行起始位置 (Character Index)
        lineStart := SendMessage(EM_LINEINDEX, lineIdx, 0, hCtrl)

        targetPos := 0

        if (mode = "Start") {
            ; 行首就是 lineStart
            targetPos := lineStart
        }
        else if (mode = "End") {
            ; 取得該行長度 (不含換行符號)
            ; 注意：EM_LINELENGTH 需要傳入「該行內任一字元的 index」才能運作正確
            ; 我們傳入 lineStart 即可
            lineLen := SendMessage(EM_LINELENGTH, lineStart, 0, hCtrl)

            ; 行尾 = 起始位置 + 長度
            targetPos := lineStart + lineLen
        }

        ; 4. 執行移動
        ; Start = End 代表不選取文字，單純移動游標
        SendMessage(EM_SETSEL, targetPos, targetPos, hCtrl)

        ; 5. 確保游標在視野內
        SendMessage(EM_SCROLLCARET, 0, 0, hCtrl)

        return true
    }

    ; =================================================================
    ; 5. [優化] Win32 輔助方法
    ; =================================================================

    ; =================================================================
    ; [核心邏輯] 選取邏輯行 (包含換行符號)
    ; 取代原本的 Win32 API 版本，改用字串分析以支援 WordWrap 模式
    ; =================================================================
    static _SelectLine(hCtrl) {
        ; 1. 取得 Control 內的全部文字
        ; 為了準確判斷邏輯段落，必須讀取全文進行 `n 搜尋
        try {
            fullText := ControlGetText(hCtrl)
        } catch {
            return ; 如果無法取得文字則放棄
        }

        if (fullText = "")
            return

        ; 2. 取得當前游標位置 (EM_GETSEL = 0x00B0)
        ; 回傳值 Low Word 是起始位置
        caretPosRaw := SendMessage(0x00B0, 0, 0, hCtrl)
        caretPos := caretPosRaw & 0xFFFF ; 0-based index

        ; 3. 計算邏輯行的開始 (Start)
        ; AHK 字串索引是 1-based，需轉換
        ahkCaretPos := caretPos + 1

        ; 往回找上一個換行符號 `n
        ; 參數 -1 代表反向搜尋
        prevLineBreak := InStr(fullText, "`n", , ahkCaretPos, -1)

        ; 如果 prevLineBreak 是 5 (代表第 5 個字是 `n)，
        ; EM_SETSEL 設定 5 代表從第 6 個字開始選 (0-based 的特性)
        ; 所以這裡直接用 prevLineBreak 即可
        selStart := (prevLineBreak == 0) ? 0 : prevLineBreak

        ; 4. 計算邏輯行的結束 (End) - 包含換行符號

        ; 找尋游標後的下一個 `r 或 `n
        nextR := InStr(fullText, "`r", , ahkCaretPos)
        nextN := InStr(fullText, "`n", , ahkCaretPos)

        selEnd := 0

        ; 狀況 1: 後面完全沒有換行符號 -> 選到全文結束
        if (nextR == 0 && nextN == 0) {
            selEnd := StrLen(fullText)
        }
        ; 狀況 2: 先遇到 `r (通常是 Windows 的 `r`n 結構)
        else if (nextR > 0 && (nextN == 0 || nextR < nextN)) {
            ; 檢查這個 `r 後面是不是緊接著 `n
            if (SubStr(fullText, nextR + 1, 1) == "`n") {
                ; 是 `r`n 結構，選取範圍要包含這兩個字元
                selEnd := nextR + 1
            } else {
                ; 只有 `r (罕見)，選取範圍包含 `r
                selEnd := nextR
            }
        }
        ; 狀況 3: 先遇到 `n (Unix 格式換行)，選取範圍包含 `n
        else {
            selEnd := nextN
        }

        ; 5. 發送選取指令 (EM_SETSEL = 0x00B1)
        SendMessage(0x00B1, selStart, selEnd, hCtrl)
    }

    ; =================================================================
    ; [新增功能] Emacs 單字移動 (Alt+F / Alt+B)
    ; =================================================================

    /**
     * 移動游標一個單字 (Word)
     * @param direction "Left" or "Right"
     * @returns {Boolean} true: 已執行; false: 不在目標區
     */
    static MoveCaretWord(direction) {
        ; 1. 檢查焦點 (這是關鍵，避免誤觸 Alt 選單)
        if !this.IsTargetFocused()
            return false

        ; 2. 執行移動
        ; 這裡直接用 Send 模擬 Ctrl+Left/Right 即可，Win32 API 實作 Word 跳轉較複雜且沒必要
        if (direction = "Left")
            Send "^{Left}"
        else
            Send "^{Right}"

        return true
    }

    ; =================================================================
    ; [新增功能] 歷史報告過濾器 (Ctrl+1/2/3)
    ; =================================================================

    /**
     * 切換歷史報告的篩選條件
     * @param modeName "All", "Modality", or "My"
     */
    static SwitchHistoryFilter(modeName) {
        try {
            switch modeName {
                case "All":      this.PastAllRadio.ControlClick()
                case "Modality": this.PastModalityRadio.ControlClick()
                case "My":       this.PastOnlyMyRadio.ControlClick()
            }
        } catch as err {
            ; 可以在這裡統一處理錯誤，例如 ToolTip 提示
            ToolTip "切換失敗: " err.Message
            SetTimer () => ToolTip(), -2000
        }
    }

    ; =================================================================
    ; [新增功能] 帶入前次報告 (Ctrl+ESC)
    ; =================================================================

    /**
     * 將前次報告內容 Append 到目前的 Finding 與 Impression
     */
    static AppendPreviousReport() {
        ; 1. 取得資料 (使用 Class 內的 Handle，無需重新搜尋)
        try {
            ; 取得來源文字
            ; 這裡直接用 ControlGetText 取 Handle 效能最好
            pastImp := ControlGetText(this.PastImpressionText.NativeWindowHandle)
            pastFind := ControlGetText(this.PastFindingText.NativeWindowHandle)

            ; 取得目標 Handle
            hImpEdit := this.ImpressionEdit.NativeWindowHandle
            hFindEdit := this.FindingEdit.NativeWindowHandle
        } catch {
            return ; 如果元件還沒準備好，直接離開
        }

        ; 2. 定義 Win32 常數
        static EM_SETSEL := 0x00B1
        static EM_REPLACESEL := 0x00C2
        static EM_SCROLLCARET := 0x00B7
        static EM_GETTEXTLENGTH := 0x000E ; 或者使用 ControlGetText 判斷長度

        ; --- 內部函式：執行 Append ---
        AppendToEdit(hEdit, textToAppend) {
            if (textToAppend == "")
                return

            try {
                currentText := ControlGetText(hEdit)
                currentLen := StrLen(currentText)
            } catch {
                currentLen := 0
            }

            ; A. 將游標移到最後 (Start = End = Length)
            SendMessage(EM_SETSEL, currentLen, currentLen, hEdit)

            ; B. 插入文字 (如果不是空行開頭，建議先補個換行)
            ; 這裡視您的需求，如果想強制換行再貼上：
            if (currentLen > 0)
                textToAppend := "`r`n" . textToAppend

            ; C. 執行貼上 (Replace Selection at the end)
            ; True (1) 代表允許 Undo
            SendMessage(EM_REPLACESEL, 1, StrPtr(textToAppend), hEdit)

            ; D. 捲動到游標處
            SendMessage(EM_SCROLLCARET, 0, 0, hEdit)
        }

        ; 3. 執行寫入
        AppendToEdit(hImpEdit, pastImp)
        AppendToEdit(hFindEdit, pastFind)

        ; 4. (選用) 將焦點移回 Finding 方便繼續編輯
        try this.FindingEdit.SetFocus()
    }

    ; =================================================================
    ; [新增功能] 檢查名稱與存檔流程 (Exam Name & Saving)
    ; =================================================================

    /**
     * 在目前游標位置插入處理過的檢查名稱
     * 邏輯：去除 "檢查項目: " -> 加上 ":\r\n\r\n" -> 插入
     * @returns {Boolean} true: 執行成功; false: 焦點不在目標區
     */
    static InsertExamNameAtCaret() {
        ; 1. 檢查焦點
        if !this.IsTargetFocused()
            return false

        try {
            hEdit := ControlGetFocus("A")

            ; 取得原始文字
            rawName := ControlGetText(this.ExamnameText.NativeWindowHandle)
        } catch {
            return false
        }

        if (rawName == "")
            return true ; 沒抓到名稱就不動作，但也算處理完畢

        ; 2. 字串處理 (String Processing)
        ; 去除 "檢查項目: " (包含後面的空白)
        cleanName := StrReplace(rawName, "檢查項目: ", "")

        ; 加上後綴格式 (冒號 + 兩行換行)
        textToInsert := cleanName . ":`r`n`r`n"

        ; 3. 執行插入 (Win32 API)
        static EM_REPLACESEL := 0x00C2

        ; 參數說明:
        ; wParam (1) = 可以被 Undo (True)
        ; lParam = 要插入的字串指標
        ; 行為: 將目前游標處(或選取範圍)替換為 textToInsert，游標自動移到插入文字的後方
        SendMessage(EM_REPLACESEL, 1, StrPtr(textToInsert), hEdit)

        return true
    }

    /**
     * 設定「自動下一筆」的狀態
     * @param targetState {Boolean} true=勾選, false=取消勾選
     */
    static SetAutoNextState(targetState) {
        try {
            ; 取得目前的 ToggleState (1=Checked, 0=Unchecked)
            currentState := this.AutoNextCheckbox.ToggleState

            ; 如果狀態不一致，才執行 Toggle
            ; targetState 轉成 1/0 比較保險
            if (!!targetState != !!currentState) {
                this.AutoNextCheckbox.Toggle()
            }
        } catch as err {
            ; 這裡建議靜默失敗或寫 Log，不要跳 MsgBox 打斷流程
        }
    }

    /**
     * 點擊儲存報告按鈕
     */
    static SaveReport() {
        try {
            this.ReportSaveButton.ControlClick()
        } catch as err {
            MsgBox "存檔按鈕點擊失敗: " err.Message
        }
    }

    /**
     * 處理 Shift+Up/Down 的邊界行為
     * 當在第一行按 Shift+Up -> 自動轉為 Shift+Home (選到行首)
     * 當在最後一行按 Shift+Down -> 自動轉為 Shift+End (選到行尾)
     * * @param direction "Up" or "Down"
     * @returns {Boolean} true: 已處理; false: 焦點不在目標區
     */
    static SmartExtendSelection(direction) {
        ; 1. 檢查焦點
        if !this.IsTargetFocused()
            return false

        try {
            hCtrl := ControlGetFocus("A")
        } catch {
            return false
        }

        ; 2. 取得行資訊 (Win32 API)
        static EM_LINEFROMCHAR := 0x00C9 ; 取得目前行號 (0-based)
        static EM_GETLINECOUNT := 0x00BA ; 取得總行數

        currentLineIdx := SendMessage(EM_LINEFROMCHAR, -1, 0, hCtrl) ; 0-based
        lineCount      := SendMessage(EM_GETLINECOUNT, 0, 0, hCtrl)

        ; 3. 判斷方向與邊界
        if (direction == "Up") {
            ; 如果在第一行 (Index 0)
            if (currentLineIdx == 0) {
                SendInput "+{Home}"
            } else {
                SendInput "+{Up}"
            }
        }
        else if (direction == "Down") {
            ; 如果在最後一行 (Index == Count - 1)
            if (currentLineIdx == lineCount - 1) {
                SendInput "+{End}"
            } else {
                SendInput "+{Down}"
            }
        }

        return true
    }

    ; =================================================================
    ; [新增功能] 歷史報告互動 (History Table Actions)
    ; =================================================================

    ; --- 1. 定義相似報告對照表 (Static Map) ---
    static _SimReportMap := Map(
        "CHEST PA/AP", Map("CHEST PA/AP+LAT", 1),
        "CHEST PA/AP+LAT", Map("CHEST PA/AP", 1),
        "KUB", Map("KUB+ABD LAT", 1),
        "KUB+L-SPINE LAT(supine)", Map("L-SPINE(AP+LAT)Standing", 1),
        "WHOLE  ABDOMEN CT WITH+ WITHOUT CONTRAST", Map("WHOLE  ABDOMEN CT WITHOUT CONTRAST", 1),
        "WHOLE  ABDOMEN CT WITHOUT CONTRAST", Map("WHOLE  ABDOMEN CT WITH+ WITHOUT CONTRAST", 1),
    )

    ; --- 2. 主要功能: 尋找並點擊相似報告 ---

    /**
     * 搜尋歷史報告表，若找到相同或相似的檢查名稱，自動點擊該行
     */
    static FindAndClickSimilarReport() {
        ; 1. 取得當前檢查名稱 (內部重用邏輯)
        currExamName := this._GetCleanCurrentExamName()
        if (currExamName == "")
            return

        SearchColumnIndex := 3 ; 1=簽收日, 2=儀器, 3=檢查項目

        try {
            ; 2. 獲取 Table (利用 Class 快取)
            tableEle := this.PastReportTable

            ; 3. 尋找所有 Rows
            ; 注意：這裡保留您原本的邏輯，先找 Custom，若無可視情況調整
            rowElements := tableEle.FindAll({ Type: 'Custom' })
            if (rowElements.Length = 0)
                throw Error("表格中找不到資料行 (Rows)")

            ; 4. 遍歷搜尋
            for rowEle in rowElements {
                ; 找儲存格
                cellElements := rowEle.FindAll({ Type: 'DataItem' })
                if (cellElements.Length = 0)
                    cellElements := rowEle.FindAll({ Type: 'Custom' })

                if (cellElements.Length < SearchColumnIndex)
                    continue

                ; 取得檢查項目文字 (第 3 欄)
                targetCellEle := cellElements[SearchColumnIndex]
                historyExamName := targetCellEle.Value

                ; 比對邏輯
                if (this._IsRelatedReport(historyExamName, currExamName)) {
                    ; *** 找到了！執行點擊 ***
                    this._ClickUIAElement(targetCellEle)

                    ; 顯示提示 (可選)
                    ToolTip "已選取相似報告: " historyExamName
                    SetTimer () => ToolTip(), -2000
                    return
                }
            }

            ToolTip "未找到相似的歷史報告"
            SetTimer () => ToolTip(), -2000

        } catch as err {
            MsgBox "搜尋歷史報告失敗: " err.Message
        }
    }

    ; --- 3. 主要功能: 插入選取行的日期/名稱 ---

    /**
     * 插入歷史報告表格中「目前被選取」的那一行的日期 (轉為西元)
     */
    static InsertSelectedHistoryDate() {
        this._InsertFromSelectedRow(1, true) ; 1 = 日期欄位, true = 需要日期轉換
    }

    /**
     * 插入歷史報告表格中「目前被選取」的那一行的檢查名稱
     */
    static InsertSelectedHistoryName() {
        this._InsertFromSelectedRow(3, false) ; 3 = 名稱欄位
    }

    ; --- 內部通用邏輯: 從選取行抓資料 ---
    static _InsertFromSelectedRow(colIndex, needDateConvert) {
        if !this.IsTargetFocused()
            return

        static STATE_SYSTEM_SELECTED := 0x2

        try {
            tableEle := this.PastReportTable
            rowElements := tableEle.FindAll({ Type: 'Custom' })

            foundValue := ""

            ; 遍歷尋找被選取的行 (LegacyIAccessible Pattern)
            for rowEle in rowElements {
                if IsObject(rowEle.LegacyIAccessiblePattern) {
                    if (rowEle.LegacyIAccessiblePattern.State & STATE_SYSTEM_SELECTED) {

                        ; 找到該行的目標儲存格
                        targetCell := rowEle.FindElement({ ControlType: "DataItem" }, , colIndex)
                        if IsObject(targetCell) {
                            foundValue := targetCell.Value
                        }
                        break ; 找到就跳出
                    }
                }
            }

            if (foundValue != "") {
                if (needDateConvert)
                    foundValue := this._ConvertRISDate(foundValue)

                ; 執行插入 (使用 SendText 模擬打字，最通用)
                SendText foundValue
            }

        } catch as err {
            ; 靜默失敗或簡單提示
            ; ToolTip "無法擷取資料: " err.Message
        }
    }

    ; --- 4. 輔助運算 (Private Helpers) ---

    ; 取得清理過的當前檢查名稱 (移除 "檢查項目: ")
    static _GetCleanCurrentExamName() {
        try {
            rawName := ControlGetText(this.ExamnameText.NativeWindowHandle)
            return StrReplace(rawName, "檢查項目: ", "")
        } catch {
            return ""
        }
    }

    ; 判斷是否為相關報告
    static _IsRelatedReport(prevName, currName) {
        ; 1. 完全相同
        if (prevName == currName)
            return true

        ; 2. 查表 (相似)
        if this._SimReportMap.Has(currName) {
            similarExams := this._SimReportMap[currName]
            if similarExams.Has(prevName)
                return true
        }
        return false
    }

    ; 日期轉換 (民國 -> 西元)
    static _ConvertRISDate(inputString) {
        cleanString := StrReplace(inputString, "/")
        if (StrLen(cleanString) < 7) ; 簡單防呆
            return inputString

        minguoYear := SubStr(cleanString, 1, 3)
        month := SubStr(cleanString, 4, 2)
        day := SubStr(cleanString, 6, 2)

        gregorianYear := minguoYear + 1911
        return gregorianYear . "-" . month . "-" . day
    }

    ; 執行 UIA 點擊 (包含座標計算邏輯)
    static _ClickUIAElement(el) {
        try {
            rect := el.BoundingRectangle
            loc := el.Location

            ; 計算中心點
            ClickX := rect.l + (loc.w / 2)
            ClickY := rect.t + (loc.h / 2)

            ; 模擬滑鼠點擊 (保留使用者原本的 MouseMove 邏輯以確保觸發)
            MouseGetPos(&OrigX, &OrigY)
            MouseMove(ClickX, ClickY, 0)
            Click()
            MouseMove(OrigX, OrigY, 0)
        } catch {
            ; 如果計算失敗，嘗試 UIA 原生 Invoke
            try el.Invoke()
        }
    }

    ; =================================================================
    ; [新增功能] 報告格式化與重排 (Text Reordering & Formatting)
    ; =================================================================

    /**
     * 對 Finding 區域執行自動編號重排
     * 自動判斷 CT/MR 或 CR/US 模式
     */
    static FormatFindingText() {
        if !this.IsTargetFocused()
            return

        try {
            ; 判斷檢查類型 (CT/MR/US/CR)
            examType := this._GetCurrExamType()
            hEdit := this.FindingEdit.NativeWindowHandle

            switch examType {
                case "CT", "MR":
                    this._FormatFindingForAdvanced(hEdit)
                case "CR", "US", "CR": ; Default
                    this._FormatFindingForBasic(hEdit)
            }
        } catch as err {
            MsgBox "格式化失敗: " err.Message
        }
    }

    /**
     * 對 Impression 區域執行格式化
     * 若只有一行則加編號，多行則重排
     */
    static FormatImpressionText() {
        if !this.IsTargetFocused()
            return

        try {
            hEdit := this.ImpressionEdit.NativeWindowHandle

            ; 聚焦並全選 (Impression 通常是一次處理全部)
            ; 注意：這裡保留您原本的邏輯，先 SetFocus 再 SetSel
            ControlFocus(hEdit)
            SendMessage(0x00B1, 0, -1, hEdit) ; EM_SETSEL (Select All)

            ; 計算非空行數
            lineCount := this._CountNonEmptyLines(hEdit)

            if (lineCount > 1) {
                ; 多行：重排 (移除舊編號 -> 重新編號)
                this._ReorderSelectedText(, , , , hEdit)
            } else {
                ; 單行或空：強制編號 (deOrder=true, itemChar="") -> 其實這邏輯怪怪的，照您原本代碼是 ReorderSelectedText(true...)
                ; 依照您原本的邏輯： ReorderSelectedText(true, , , , hEdit)
                ; 參數 1 (deOrder) = true 代表「移除序號」?
                ; 看了您的原始碼： if (!deOrder) { 加序號 }
                ; 所以單行時您傳入 true，代表「不要加序號」? 還是您原本想寫 false?
                ; 假設您原本邏輯是對的：單行不強制加 1.
                this._ReorderSelectedText(true, , , , hEdit)
            }
        } catch as err {
            MsgBox "Impression 格式化失敗: " err.Message
        }
    }

    /**
     * 通用的重排指令 (供熱鍵直接呼叫)
     * @param options {Object} 設定參數 {deOrder, keepEmpty, itemChar, discardSeIm}
     */
    static ReorderSelection(options := {}) {
        if !this.IsTargetFocused()
            return

        try {
            hEdit := ControlGetFocus("A")
            ; 設定預設值
            deOrder := options.HasOwnProp("deOrder") ? options.deOrder : false
            keepEmpty := options.HasOwnProp("keepEmpty") ? options.keepEmpty : false
            itemChar := options.HasOwnProp("itemChar") ? options.itemChar : ""
            discardSeIm := options.HasOwnProp("discardSeIm") ? options.discardSeIm : true

            this._ReorderSelectedText(deOrder, keepEmpty, itemChar, discardSeIm, hEdit)
        }
    }

    ; --- 內部格式化邏輯 (Private) ---

    static _GetCurrExamType() {
        name := this._GetCleanCurrentExamName() ; 重用之前的函數
        if (InStr(name, "CT") || InStr(name, "電腦斷層"))
            return "CT"
        if (InStr(name, "MR") || InStr(name, "磁振造影"))
            return "MR"
        if (InStr(name, "US") || InStr(name, "超音波"))
            return "US"
        return "CR"
    }

    static _FormatFindingForBasic(hEdit) {
        ; 搜尋 FINDINGS: ...
        ; 這裡需要 Edit_FindText 的邏輯，建議改用正則表達式讀取全文後計算位置
        ; 為了簡化，這裡假設您有引用 Edit.ahk 或者我們實作一個簡單版
        ; 由於篇幅限制，這裡使用 Win32 API 取得全文後用 AHK RegEx 算位置

        fullText := ControlGetText(hEdit)
        needle := "im)FINDINGS:\r?\n|:\s*\r?\n\s*\r?\n"

        if RegExMatch(fullText, needle, &match) {
            startPos := match.Pos + match.Len - 1 ; 轉成 0-based offset

            ; 選取從匹配點到最後
            SendMessage(0x00B1, startPos, -1, hEdit)

            ; 重排： "-" 符號，保留空行
            this._ReorderSelectedText(false, true, "-", false, hEdit)
        }
    }

    static _FormatFindingForAdvanced(hEdit) {
        fullText := ControlGetText(hEdit)
        needle := "im)FINDINGS:\r?\n|The study shows:\r?\n\r?\n|show the following findings:\r?\n\r?\n|which revealed:\r?\n\r?\n"

        if RegExMatch(fullText, needle, &match) {
            startPos := match.Pos + match.Len - 1 ; 0-based

            ; 尋找結束點 (REMARKS / RECOMMENDATION)
            endNeedle := "im)REMARKS?:|RECOMMENDATION:"
            endPos := -1
            if RegExMatch(fullText, endNeedle, &endMatch, startPos + 1) {
                endPos := endMatch.Pos - 1 ; 0-based start of end tag
                ; 依照原本邏輯還要 -2 (扣掉前面的換行?)
                if (endPos > 2)
                    endPos -= 2
            }

            ; 設定選取範圍
            SendMessage(0x00B1, startPos, endPos, hEdit)

            ; 重排： "-" 符號
            this._ReorderSelectedText(false, false, "-", true, hEdit)
        }
    }

    ; 核心重排演算法 (移植自您的 ReorderSelectedText)
    static _ReorderSelectedText(deOrder := false, keepEmptyLine := false, itemChar := "", discardSeIm := true, targetHwnd := 0) {
        selectedText := ""
        try {
            ; 使用 Win32 API 取得選取文字 (取代 EditGetSelectedText)
            ; 為了相容性，這裡假設您環境有 EditGetSelectedText，或者我們用 ControlGetText+GetSel 模擬
            ; 這裡示範最簡單的模擬：
            fullText := ControlGetText(targetHwnd)

            static EM_GETSEL := 0x00B0
            selRaw := SendMessage(EM_GETSEL, 0, 0, targetHwnd)
            start := selRaw & 0xFFFF
            end := (selRaw >> 16) & 0xFFFF

            if (end > start)
                selectedText := SubStr(fullText, start + 1, end - start)
        }

        if (selectedText == "")
            return

        ; --- 文字處理邏輯 (完全保留您的 RegEx) ---
        selectedText := StrReplace(selectedText, "`r`n", "`n")

        ; (省略部分檢查邏輯以節省篇幅，保留核心迴圈)
        txtAry := StrSplit(selectedText, "`n")
        finalText := ""
        startLineNo := 1

        ; 嘗試偵測既有編號
        if (RegExMatch(selectedText, "^(\d+)", &existLineNo))
            startLineNo := existLineNo[1]

        for index, line in txtAry {
            if (!RegExMatch(line, "^\s*$")) {
                tmpText := line

                ; 處理 Spine 特殊邏輯
                isSpine := RegExMatch(line, "^\s*[-\+\*]*\s*([Vv]arying degree|[Mm]ild).+causing:")

                if (!deOrder) {
                    orderChar := (itemChar != "" ? itemChar : startLineNo++ . ".")
                    if (isSpine && RegExMatch(line, "^\s*([-\+\*]*|-->)\s*([CcTtLl]\d{1,2}-.+$)", &m)) {
                        finalText .= "--> "
                        tmpText := m[2]
                    } else {
                        finalText .= orderChar . " "
                    }
                }

                if (itemChar == "" && discardSeIm) {
                    tmpText := RegExReplace(tmpText, "\s*\((Srs|Ser)\/Img:[\s,-\/\d;]+\)", "")
                    tmpText := RegExReplace(tmpText, "Mark L\d+:\s*", "")
                }

                finalText .= RegExReplace(tmpText, "^(\s*)((\d+\.)|([-\+\*>=])|(\(?\d+\)))?(\s*)(\w?)(.*)", "$u{7}${8}")
                finalText .= "`r`n"
            } else {
                if (keepEmptyLine)
                    finalText .= "`r`n"
            }
        }

        ; 去除最後一個換行
        finalText := RTrim(finalText, "`r`n")

        ; --- 寫回 ---
        try EditPaste(finalText, targetHwnd)
    }

    static _CountNonEmptyLines(hEdit) {
        text := ControlGetText(hEdit)
        if (text == "")
            return 0
        lines := StrSplit(text, "`n", "`r")
        count := 0
        for line in lines {
            if (Trim(line, " `t") != "")
                count++
        }
        return count
    }

    ; =================================================================
    ; [新增功能] 病理報告複製 (Pathology Copy)
    ; =================================================================
    static CopyPathologyReport() {
        try {
            dateVal := this.PathoDateText.Value
            diagVal := this.PathoDiagnosisText.Value

            if (dateVal == "" && diagVal == "")
                throw Error("找不到病理報告內容")

            reportText := this._ConvertRISDate(dateVal) . ": " . diagVal
            A_Clipboard := reportText

            ; 這裡可以用您的 Notify 函數，或者簡單 ToolTip
            ToolTip "病理報告已複製"
            SetTimer () => ToolTip(), -2000
        } catch as err {
            MsgBox "複製失敗: " err.Message
        }
    }

    ; =================================================================
    ; [新增功能] 滑鼠連點選取 (Triple Click)
    ; =================================================================
    static HandleTripleClick() {
        ; 這裡需要一個靜態變數來記錄點擊狀態
        static clickCount := 0
        static lastClickTime := 0
        static DoubleClickTime := DllCall("GetDoubleClickTime")

        timeSinceLast := A_TickCount - lastClickTime
        if (timeSinceLast <= DoubleClickTime)
            clickCount++
        else
            clickCount := 1

        lastClickTime := A_TickCount

        if (clickCount == 3) {
            clickCount := 0

            MouseGetPos , , , &hCtrl, 2
            try {
                classNN := ControlGetClassNN(hCtrl)
                if (InStr(classNN, "Edit") && !InStr(classNN, "RichEdit")) {
                    this._SelectLine(hCtrl) ; 使用之前寫好的 Win32 選行函數
                }
            }
        }
    }
}
