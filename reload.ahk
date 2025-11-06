#Persistent
DetectHiddenWindows, On ; <--- 關鍵點 1: 必須開啟

; 我們的自訂訊息 ID
global MY_MSG := 0x5555

; 1. 註冊訊息監聽器
OnMessage(MY_MSG, "V1_ReceiveAction")

MsgBox, "V1 腳本已啟動並開始偵聽 " . MY_MSG ; 啟動提示
return

; ===================================
; 2. 觸發熱鍵
^F1::
    ; 顯示這個 MsgBox, 證明熱鍵有被觸發
    MsgBox, "步驟 A: V1 熱鍵 ^F1 已觸發。"

    ; 3. 廣播訊息
    ;    我們使用 %MY_MSG% (強制表達式) 確保傳遞的是數字
    PostMessage, %MY_MSG%, 0, 0, , ahk_class AutoHotkey

    MsgBox, "步驟 B: V1 PostMessage 已發送。"
return

; ===================================
; 4. 接收訊息的函式
V1_ReceiveAction(wParam, lParam, msg, hwnd)
{
    ; 能夠執行到這裡, 才算成功
    MsgBox, "步驟 C: V1 OnMessage 成功接收！"
}
return