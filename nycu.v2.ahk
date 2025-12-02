#Requires AutoHotkey v2.0
#SingleInstance Force
ProcessSetPriority "High"  ; 提高優先級

; 先設定選項：SendInput (SI) 和 忽略終止符 (O)
#Hotstring SI O
; 再設定終止字元：只有 Tab
#Hotstring EndChars `t

; === 解決 DPI 座標偏移問題 ===
; 1. 宣告 DPI 感知 (讓 BoundingRectangle 回報實體座標)
DllCall("SetThreadDpiAwarenessContext", "ptr", -3, "ptr")
; 2. 確保 AHK 的 Click 使用螢幕座標
CoordMode "Mouse", "Screen"
; ================================
SetMouseDelay -1
SetKeyDelay -1

global PRESERVE_CLIPBOARD := 0

#Include <RisController.v2>
;#Include <UIA.v2>
;#Include <Paste.v2>
;#Include <Edit.v2>

;#Include MyScripts\hotkey\reorder-selected-text.v2.ahk
#Include Hotkeys\remapping-original-hotkeys-infinitt.v2.ahk

; 腳本啟動時執行，設定為 Fira Code, 大小 14 (根據您的螢幕解析度調整)
;RisController.EnableFontEnforcer("Fira Code", 12)
;RisController.EnableFontEnforcer("Cascadia Code", 12)
;RisController.EnableFontEnforcer("Sarasa Mono TC", 12)
RisController.EnableFontEnforcer("Maple Mono CN", 11)

#HotIf WinActive(RisController.WinTitle)

    #Include Hotstrings\regex-hotstrings.v2.ahk
    #Include Hotstrings\others.v2.ahk
    #Include Hotstrings\chest-ct.v2.ahk
    #Include Hotstrings\abdomen-ct.v2.ahk
    #Include Hotstrings\abdomen-mr.v2.ahk
    #Include Hotstrings\ct-guide.v2.ahk
    #Include Hotstrings\ms-ct.v2.ahk
    #Include Hotstrings\ms-mri.v2.ahk
    #Include Hotstrings\neuro.v2.ahk
    #Include Hotstrings\abbreviations.v2.ahk
    #Include Hotstrings\mri.v2.ahk
    #Include Hotstrings\chest-x-ray.v2.ahk
    #Include Hotstrings\kub.v2.ahk
    #Include Hotstrings\bone-x-ray.v2.ahk
    #Include Hotstrings\other-x-ray.v2.ahk
    #Include Hotstrings\sono-guide.v2.ahk
    #Include Hotstrings\angio.v2.ahk
    #Include Hotstrings\sono.v2.ahk
    #Include Hotstrings\comparisons.v2.ahk

    ^9:: {
    }

    ; --- Emacs Word Movement ---
    ; 加上防呆機制：只有在打字時，Alt+F 才是向右移動
    ; 否則它會保留 Windows 預設行為 (開啟檔案選單)
    !f:: {
        if !RisController.MoveCaretWord("Right")
            Send "!f" ; 透傳
    }

    !b:: {
        if !RisController.MoveCaretWord("Left")
            Send "!b" ; 透傳
    }

    ; --- History Filter Switching ---
    ; 這裡不需要透傳，因為 Ctrl+數字鍵通常沒有其他重要功能
    ^1:: RisController.SwitchHistoryFilter("All")
    ^2:: RisController.SwitchHistoryFilter("Modality")
    ^3:: RisController.SwitchHistoryFilter("My")

    ; --- Business Logic ---
    ^Esc:: RisController.AppendPreviousReport()

    ; Ctrl+A: Emacs 行首 (若不在目標框則為全選)
    ^a:: {
        ; 嘗試移動到行首
        didMove := RisController.MoveCaret("Start")

        ; 如果不在 Finding/Impression 框框內，則送出原本的 Ctrl+A (全選)
        if (!didMove) {
            Send "^a"
        }
    }

    ; Ctrl+E: Emacs 行尾
    ^e:: {
        didMove := RisController.MoveCaret("End")

        ; 如果不在目標框，送出原生的 Ctrl+E (有些系統可能是置中對齊或其他功能)
        if (!didMove) {
            Send "^e"
        }
    }

    ^d:: {
        Send "{Del}"
    }

    ; Ctrl+Y 刪除整行
    ^y:: {
        ; 呼叫 Controller 執行刪除
        isDeleted := RisController.DeleteCurrentLine()

        ; 如果 Controller 回傳 false (代表沒在目標框內，或執行失敗)
        ; 則將 Ctrl+D 原封不動送回給系統
        if (!isDeleted) {
            Send("^d")
        }
    }

    ; Ctrl+W: 刪除前一個字 (Bash Style)
    ^w:: {
        ; 嘗試呼叫 Controller 執行刪除
        didDelete := RisController.DeleteWordBackward()

        ; 如果 Controller 回傳 false (代表焦點不在 Finding/Impression 框框內)
        ; 則透傳 Ctrl+W 給系統 (例如：醫師在瀏覽器想關分頁，或在其他地方想用原生的 Ctrl+W)
        if (!didDelete) {
            Send "^w"
        }
    }

    ; Alt+E: 在游標處插入檢查名稱
    !e:: {
        ; 如果焦點在編輯框內，執行插入
        ; 如果焦點在其他地方，送出 Alt+E (開啟系統選單)
        if !RisController.InsertExamNameAtCaret() {
            Send "!e"
        }
    }

    ; Alt+C: 取消 AutoNext 並存檔 (Save Only)
    !c:: {
        ; 1. 設定 AutoNext 為 False (取消勾選)
        RisController.SetAutoNextState(false)
        ; 2. 存檔
        RisController.SaveReport()
    }

    ; Ctrl+S: 勾選 AutoNext 並存檔 (Save & Next)
    ^s:: {
        ; 1. 設定 AutoNext 為 True (勾選)
        RisController.SetAutoNextState(true)
        ; 2. 存檔
        RisController.SaveReport()
    }

    !q:: {
        Send "^e"
    }

    ^k:: {
        Send "+{End}"
        Send "{Del}"
    }

    #a:: {
        Send "^a"
    }

    #d:: {
        Send "^a"
        Sleep 100
        Send "{Del}"
    }

    ; Shift + Up: 智慧向上選取
    +Up:: {
        ; 如果沒執行智慧選取 (不在編輯框內)，則送出原生的 Shift+Up
        if !RisController.SmartExtendSelection("Up") {
            SendInput "+{Up}"
        }
    }

    ; Shift + Down: 智慧向下選取
    +Down:: {
        if !RisController.SmartExtendSelection("Down") {
            SendInput "+{Down}"
        }
    }

    ; Alt+Esc: 根據目前的檢查名稱，自動搜尋並選取歷史報告中的相似項目
    !Esc:: RisController.FindAndClickSimilarReport()

    ; Alt+Shift+D: 插入目前選取的歷史報告日期 (自動轉西元)
    !+d:: {
        if !RisController.IsTargetFocused()
            Send "!d" ; 透傳
        else
            RisController.InsertSelectedHistoryDate()
    }

    ; Alt+D: 插入以複製的歷史報告日期 (自動轉西元)
    !d:: {
        if !RisController.IsTargetFocused()
            Send "!d" ; 透傳
        else
            RisController.InsertCopiedReportDate()
    }

    ; Ctrl+Alt+E: 插入目前選取的歷史報告名稱
    ^!e:: {
        if !RisController.IsTargetFocused()
            Send "^!e" ; 透傳
        else
            RisController.InsertSelectedHistoryName()
    }

    ; --- Findings Formatting ---
    SC079:: RisController.FormatFindingText() ; 日文鍵盤的轉換鍵?
    ^!,::   RisController.FormatFindingText()

    ; --- Impression Formatting ---
    SC070:: RisController.FormatImpressionText() ; 日文鍵盤的無變換鍵?
    ^!.::   RisController.FormatImpressionText()


    ; --- Selection Reordering (Manual) ---
    ; 重排選取文字 (預設：自動編號 1. 2. 3.)
    ^!o:: RisController.ReorderSelection()

    ; 重排選取文字 (保留 Series/Image 標記)
    ^!+o:: RisController.ReorderSelection({discardSeIm: false})

    ; 移除編號，改用 "*"
    ^+*:: RisController.ReorderSelection({deOrder: false, keepEmpty: true, itemChar: "*"})

    ; 移除編號，改用 "-"
    ^+-:: RisController.ReorderSelection({deOrder: false, keepEmpty: true, itemChar: "-"})

    ; 移除編號，改用 ">"
    ^+>:: RisController.ReorderSelection({deOrder: false, keepEmpty: true, itemChar: ">"})

    ; --- Pathology Copy ---
    ^!c:: RisController.CopyPathologyReport()

    ;-----------------------------------------------------------
    ; Mouse Remapping
    XButton1:: RisController.ReorderSelection({deOrder: false, keepEmpty: true, itemChar: "-"})

    XButton2:: RisController.ReorderSelection()

    ; --- Triple Click Handler ---
    ; 讓滑鼠左鍵通過，同時觸發連點檢查
    ~LButton:: RisController.HandleTripleClick()

#HotIf ; WinActive(RisController.WinTitle)

;; for JIS keyboard
SC029:: RisController.ActivateOrToggleFocus() ; SC029 通常是 `~ 鍵


; 危急值視窗熱鍵區
#HotIf WinActive(RisController.AbnormalWinTitle)

    !1:: RisController.ClickAbnormalButton(1)
    !2:: RisController.ClickAbnormalButton(2)
    !3:: RisController.ClickAbnormalButton(3)
    !4:: RisController.ClickAbnormalButton(4)

    ; 如果想加存檔熱鍵也很容易：
    ; ^s:: RisController.ClickAbnormalButton("Save")

#HotIf ; WinActive(RisController.AbnormalWinTitle)


;
; Global Remap
;
^!r:: Reload

#^p:: {
    ProcessClose("G3PACS.exe")
}

SC07B::LButton

; ==============================================================================
; 3. Benchmark 工具函數 (高精確度)
; ==============================================================================
; 參數:
; funcObj: 要測試的函數物件 (使用 CallbackCreate 或 Func.Bind)
; times: 執行的次數
; 回傳: 執行總耗時 (毫秒, ms)
; ==============================================================================
Benchmark(funcObj, times := 1) {
    ; 1. 獲取計時器頻率 (每秒多少 ticks)
    DllCall("QueryPerformanceFrequency", "Int64*", &freq := 0)

    ; 2. 記錄開始時間
    DllCall("QueryPerformanceCounter", "Int64*", &start := 0)

    ; 3. 執行迴圈
    loop times {
        funcObj()
    }

    ; 4. 記錄結束時間
    DllCall("QueryPerformanceCounter", "Int64*", &end := 0)

    ; 5. 計算耗時 ( (結束-開始) / 頻率 * 1000 轉換為毫秒 )
    return (end - start) / freq * 1000
}