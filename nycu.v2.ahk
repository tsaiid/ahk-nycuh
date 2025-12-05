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
    #Include Hotstrings\others.v2.ahk

    ^9:: {
    }

    ; --- History Filter Switching ---
    ; 這裡不需要透傳，因為 Ctrl+數字鍵通常沒有其他重要功能
    ^1:: RisController.SwitchHistoryFilter("All")
    ^2:: RisController.SwitchHistoryFilter("Modality")
    ^3:: RisController.SwitchHistoryFilter("My")

    ; --- Business Logic ---
    ^Esc:: RisController.AppendPreviousReport()

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

    ; --- Pathology Copy ---
    ^!c:: RisController.CopyPathologyReport()

    ; Alt+Esc: 根據目前的檢查名稱，自動搜尋並選取歷史報告中的相似項目
    !Esc:: RisController.FindAndClickSimilarReport()

    ; --- Findings Formatting ---
    SC079:: RisController.FormatFindingText() ; 日文鍵盤的轉換鍵?
    ^!,::   RisController.FormatFindingText()

    ; --- Impression Formatting ---
    SC070:: RisController.FormatImpressionText() ; 日文鍵盤的無變換鍵?
    ^!.::   RisController.FormatImpressionText()

#HotIf ; WinActive(RisController.WinTitle)

; 只有在「RIS 視窗作用中」且「焦點在輸入框內」時，這個熱鍵才存在
#HotIf WinActive(RisController.WinTitle) && RisController.IsTargetFocused()

    ; --- Emacs Word Movement ---
    ; Ctrl+A: Emacs 行首 (若不在目標框則為全選)
    ^a:: RisController.MoveCaret("Start")

    ; Ctrl+E: Emacs 行尾
    ^e:: RisController.MoveCaret("End")

    !f:: RisController.MoveCaretWord("Right")

    !b:: RisController.MoveCaretWord("Left")

    #d:: RisController.ClearCurrentEdit()

    ; Ctrl+Y 刪除整行
    ^y:: RisController.DeleteCurrentLine()

    ; Ctrl+W: 刪除前一個字 (Bash Style)
    ^w:: RisController.DeleteWordBackward()

    ; Emacs style Kill Line
    ^k::RisController.KillLine()

    ; Alt+E: 在游標處插入檢查名稱
    !e:: RisController.InsertExamNameAtCaret()

    ; Shift + Up: 智慧向上選取
    +Up:: RisController.SmartExtendSelection("Up")

    ; Shift + Down: 智慧向下選取
    +Down:: RisController.SmartExtendSelection("Down")

    ; Alt+Shift+D: 插入目前選取的歷史報告日期 (自動轉西元)
    !+d:: RisController.InsertSelectedHistoryDate()

    ; Alt+D: 插入以複製的歷史報告日期 (自動轉西元)
    !d:: RisController.InsertCopiedReportDate()

    ; Ctrl+Alt+E: 插入目前選取的歷史報告名稱
    ^!e:: RisController.InsertSelectedHistoryName()

    ; --- Selection Reordering (Manual) ---
    ; 重排選取文字 (預設：自動編號 1. 2. 3.)
    ^!o:: RisController.ReorderSelection()

    ; 重排選取文字 (保留 Series/Image 標記)
    ^!+o:: RisController.ReorderSelection({discardSeIm: false})

    ; 移除編號，改用 "*"
    ^+*:: {
        Send("{Blind}{vkE8}") ; 發送一個無效按鍵，阻斷 Windows 的語言切換偵測
        RisController.ReorderSelection({deOrder: false, keepEmpty: true, itemChar: "*"})
    }

    ; 移除編號，改用 "-"
    ^+-:: {
        Send("{Blind}{vkE8}") ; 發送一個無效按鍵，阻斷 Windows 的語言切換偵測
        RisController.ReorderSelection({deOrder: false, keepEmpty: true, itemChar: "-"})
    }

    ; 移除編號，改用 ">"
    ^+>:: {
        Send("{Blind}{vkE8}") ; 發送一個無效按鍵，阻斷 Windows 的語言切換偵測
        RisController.ReorderSelection({deOrder: false, keepEmpty: true, itemChar: ">"})
    }

    ^d:: {
        Send "{Del}"
    }

    #a:: {
        Send "^a"
    }

    ;-----------------------------------------------------------
    ; Mouse Remapping
    XButton1:: RisController.ReorderSelection({deOrder: false, keepEmpty: true, itemChar: "-"})

    XButton2:: RisController.ReorderSelection()

    ; --- Triple Click Handler ---
    ; 讓滑鼠左鍵通過，同時觸發連點檢查
    ~LButton:: RisController.HandleTripleClick()

#HotIf  ; WinActive(RisController.WinTitle) && RisController.IsTargetFocused()

;; for JIS keyboard
SC029:: RisController.ActivateOrToggleFocus() ; SC029 通常是 `~ 鍵


; 危急值視窗熱鍵區
#HotIf WinActive(RisController.AbnormalWinTitle)

    !1:: RisController.ClickAbnormalButton(1)
    !2:: RisController.ClickAbnormalButton(2)
    !3:: RisController.ClickAbnormalButton(3)
    !4:: RisController.ClickAbnormalButton(4)

    ; 如果想加存檔熱鍵也很容易：
    !s:: RisController.ClickAbnormalButton("Save")
    ESC:: RisController.ClickAbnormalButton("Cancel")

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