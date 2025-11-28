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
}
