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

; [新增] 設定工作列 (Tray) 上的 Icon
TraySetIcon(A_ScriptDir "\assets\nycu_icon.png")

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

; 啟動自動更新排程
;RisController.EnableAutoWorklistUpdate()

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
    #Include Hotstrings\special.v2.ahk
    #Include Hotstrings\breast-mr.v2.ahk

    ^9:: {
    }

    ; --- History Filter Switching ---
    ; 這裡不需要透傳，因為 Ctrl+數字鍵通常沒有其他重要功能
    ^1:: RisController.SwitchHistoryFilter("All")
    ^2:: RisController.SwitchHistoryFilter("Modality")
    ^3:: RisController.SwitchHistoryFilter("My")

    ; --- Right List Tab Switching (Tab: tbRightList) ---
    !1:: RisController.SelectRightTab(1)
    !2:: RisController.SelectRightTab(2)
    !3:: RisController.SelectRightTab(3)
    !4:: RisController.SelectRightTab(4)

    ; --- Clinical Data Tab Switching (Tab: tbcClinicalData) ---
    ^!2:: RisController.SelectClinicalTab(2)
    ^!4:: RisController.SelectClinicalTab(4)

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

    ; Alt+E: 在游標處插入檢查名稱
    !e:: RisController.InsertExamNameAtCaret()

    ; Alt+Shift+D: 插入目前選取的歷史報告日期 (自動轉西元)
    !+d:: {
        Send("{Blind}{vkE8}") ; 發送一個無效按鍵，阻斷 Windows 的語言切換偵測
        RisController.InsertSelectedHistoryDate()
        Send("{Blind}{vkE8}") ; 發送一個無效按鍵，阻斷 Windows 的語言切換偵測
    }

    ; Alt+D: 插入以複製的歷史報告日期 (自動轉西元)
    !d:: RisController.InsertCopiedReportDate()

    ; Ctrl+Alt+E: 插入目前選取的歷史報告名稱
    ^!e:: RisController.InsertSelectedHistoryName()

    ; --- Ph Exam or Pathology Copy ---
    ^+c:: {
        Send("{Blind}{vkE8}") ; 發送一個無效按鍵，阻斷 Windows 的語言切換偵測
        RisController.CopyOtherReport()
    }

    ; Alt+Esc: 根據目前的檢查名稱，自動搜尋並選取歷史報告中的相似項目
    !Esc:: RisController.FindAndClickSimilarReport()

    ; --- Findings Formatting ---
    SC079:: RisController.FormatFindingText() ; 日文鍵盤的轉換鍵?
    ^!,::   RisController.FormatFindingText()

    ; --- Impression Formatting ---
    SC070:: RisController.FormatImpressionText() ; 日文鍵盤的無變換鍵?
    ^!.::   RisController.FormatImpressionText()

    ; --- Copy Finding to Impression ---
    #v:: RisController.CopyFindingToImpression()

    ; Ctrl+W: 刪除前一個字 (Bash Style)
    ^w:: {
        Critical
        RisController.DeleteWordBackward()
    }

    !Up::RisController.MoveCurrentLine("Up")
    !Down::RisController.MoveCurrentLine("Down")

    ; --- Triple Click Handler ---
    ; 讓滑鼠左鍵通過，同時觸發連點檢查
    ~LButton:: RisController.HandleTripleClick()

#HotIf ; WinActive(RisController.WinTitle)

; 只有在「RIS 視窗作用中」且「焦點在輸入框內」時，這個熱鍵才存在
#HotIf WinActive(RisController.WinTitle) && RisController.IsTargetFocused()
    ; Insert Indication by AI
    !i:: RisController.GenerateAndInsertIndication()

    ; Insert Impression by AI (Summary of Findings)
    !s:: RisController.GenerateAndInsertImpression()
    !+s:: {
        Send("{Blind}{vkE8}")
        RisController.GenerateAndInsertImpression(true)
        Send("{Blind}{vkE8}")
    }

    ; --- Emacs Word Movement ---
    ; Ctrl+A: Emacs 行首 (若不在目標框則為全選)
    ^a:: RisController.MoveCaret("Start")

    ; Ctrl+E: Emacs 行尾
    ^e:: RisController.MoveCaret("End")

    !f:: RisController.MoveCaretWord("Right")

    !b:: RisController.MoveCaretWord("Left")

    #d:: RisController.ClearCurrentEdit()

    ; Ctrl+C 複製選取文字或整行
    ^c:: RisController.CopyLineOrSelection()

    ; Ctrl+X 剪下選取文字或整行
    ^x:: RisController.CutLineOrSelection()

    ; Ctrl+Y 刪除整行
    ^y:: RisController.DeleteCurrentLine()

    ; Emacs style Kill Line
    ^k::RisController.KillLine()

    ; Shift + Up: 智慧向上選取
    +Up:: RisController.SmartExtendSelection("Up")

    ; Shift + Down: 智慧向下選取
    +Down:: RisController.SmartExtendSelection("Down")

    ; --- Selection Reordering (Manual) ---
    ; 重排選取文字 (預設：自動編號 1. 2. 3.)
    ^!o:: RisController.ReorderSelection()

    ; 重排選取文字 (保留 Series/Image 標記)
    ^!+o:: {
        Send("{Blind}{vkE8}")
        RisController.ReorderSelection({discardSeIm: false})
        Send("{Blind}{vkE8}")
    }

    ; 移除編號，改用 "*"
    ^+*:: {
        Send("{Blind}{vkE8}") ; 發送一個無效按鍵，阻斷 Windows 的語言切換偵測
        RisController.ReorderSelection({keepEmpty: true, itemChar: "*"})
    }

    ; 移除編號，改用 "-"
    ^+-:: {
        Send("{Blind}{vkE8}") ; 發送一個無效按鍵，阻斷 Windows 的語言切換偵測
        RisController.ReorderSelection({keepEmpty: true, itemChar: "-"})
    }

    ; 移除編號，自動偵測
    ^!+-:: {
        Send("{Blind}{vkE8}") ; 發送一個無效按鍵，阻斷 Windows 的語言切換偵測
        RisController.ReorderSelection({autoDetectItemChar: true, keepEmpty: true, itemChar: "-"})
        Send("{Blind}{vkE8}") ; 發送一個無效按鍵，阻斷 Windows 的語言切換偵測
    }

    ; 移除編號，改用 ">"
    ^+>:: {
        Send("{Blind}{vkE8}") ; 發送一個無效按鍵，阻斷 Windows 的語言切換偵測
        RisController.ReorderSelection({keepEmpty: true, itemChar: ">"})
    }

    ^d:: {
        Send "{Del}"
    }

    #a:: {
        Send "^a"
    }

    ^Up::RisController.SmartPageMove("Up")
    ^Down::RisController.SmartPageMove("Down")

    ^Enter::RisController.InsertNewLine("Below") ; Ctrl+Enter: 下方插入
    +Enter::RisController.InsertNewLine("Above") ; Shift+Enter: 上方插入

    ;-----------------------------------------------------------
    ; Mouse Remapping
    XButton1:: RisController.ReorderSelection({autoDetectItemChar: true, keepEmpty: true})

    XButton2:: RisController.ReorderSelection()

#HotIf  ; WinActive(RisController.WinTitle) && RisController.IsTargetFocused()

#w:: {
    RisController.GetWorklistJson()
}

;; for JIS keyboard
; 取得目前活動視窗的輸入法語言 ID
GetKeyboardHKL() {
    try {
        hWnd := WinActive("A")
        if !hWnd
            return 0
        ThreadID := DllCall("GetWindowThreadProcessId", "Ptr", hWnd, "Ptr", 0)

        ; 【關鍵修正】加上 & 0xFFFFFFFF
        ; 這會確保無論系統回傳的是 64 bit 還是有號整數，我們都只取低位的 32 bit ID
        return DllCall("GetKeyboardLayout", "UInt", ThreadID, "Ptr") & 0xFFFFFFFF
    } catch {
        return 0
    }
}
F2:: {
    CurrentID := GetKeyboardHKL()
    MsgBox("偵測到的 ID: " . Format("0x{:X}", CurrentID) . "`n" . "目標 ID: 0x04110409`n" . "是否相等? " . (CurrentID == 0x04110409 ? "YES" : "NO"))
}

; === 定義熱鍵 ===

; 【情境 A：US 鍵盤】ID 通常為 0x04090409
; 當偵測到是 US 鍵盤時，將 Right Alt (RAlt) 設為觸發鍵
#HotIf (GetKeyboardHKL() == 0x04090409)
    RAlt::RisController.ActivateOrToggleFocus()
#HotIf

; 【情境 B：日文鍵盤】ID 為你查到的 0x04110409
; 當偵測到是日文鍵盤時，將 SC029 (全形半形鍵) 設為觸發鍵
#HotIf (GetKeyboardHKL() == 0x04110409)
    SC029::RisController.ActivateOrToggleFocus()
#HotIf


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

; =================================================================
; 會診視窗專屬熱鍵
; =================================================================
#HotIf WinActive(RisController.ConsultationWinTitle)

    #Include Hotstrings\consultation.v2.ahk

^t::
{
    RisController.AddConsultationTime(20)
}

#HotIf ; 關閉條件判斷區

;
; Global Remap
;
^!r:: Reload

#^p:: {
    ProcessClose("G3PACS.exe")
}

SC07B::LButton