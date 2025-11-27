#Requires AutoHotkey v2.0
#SingleInstance force
SetWorkingDir A_ScriptDir
SetControlDelay -1
CoordMode "Mouse", "Screen"

#Include <AHKHID.v2>

; Initialize global variables
global shuttlepro_speed_saved := 0
global shuttlepro_shuttle_start := True
global shuttlepro_old_4 := 255
global shuttlepro_old_5 := 255
global shuttlepro_shuttle_saved := 0
global timer_active_hwnd := -1
global hlbxInput := 0
global a := "" ; Global debug string accumulator

; Create GUI
global MainGui := Gui("+LastFound -Resize -MaximizeBox -MinimizeBox", "ShuttlePro Debug")
MainGui.SetFont("w700 s8", "Courier New")
global lbxInput := MainGui.Add("ListBox", "h300 w600 vlbxInput") ; 加寬一點以容納詳細資訊
hlbxInput := lbxInput.Hwnd

GuiHandle := MainGui.Hwnd

; Register ShuttlePro Page 12, Usage 1
AHKHID.Register(12, 1, GuiHandle, AHKHID.RIDEV_INPUTSINK)

; Intercept WM_INPUT
OnMessage(0x00FF, ShuttleProIntercept)

MainGui.Show()

; Functions
ShuttleProIntercept(wParam, lParam, msg, hwnd) {
    Critical
    global shuttlepro_speed_saved, shuttlepro_shuttle_start, shuttlepro_old_4, shuttlepro_old_5,
        shuttlepro_shuttle_saved
    global lbxInput, hlbxInput, MainGui
    global a := "" ; Reset debug string for this event

    devicetype := AHKHID.GetInputInfo(lParam, AHKHID.II_DEVTYPE)

    if (devicetype = AHKHID.RIM_TYPEHID) {
        hid_handle := AHKHID.GetInputInfo(lParam, AHKHID.II_DEVHANDLE)

        vendor_id := AHKHID.GetDevInfo(hid_handle, AHKHID.DI_HID_VENDORID, True)
        product_id := AHKHID.GetDevInfo(hid_handle, AHKHID.DI_HID_PRODUCTID, True)

        if (vendor_id = 2867 && product_id = 48) {
            ; ShuttlePro v2

            ; Get raw data buffer
            uData := AHKHID.GetInputData(lParam)
            if (uData.Size = 0)
                return

            ; Read bytes (Offsets corrected: +1 compared to 0-based index to match v1 logic)
            byte1 := NumGet(uData, 1, "UChar") ; Offset 1: Outer Ring (Speed)
            byte2 := NumGet(uData, 2, "UChar") ; Offset 2: Inner Ring (Jog)
            byte4 := NumGet(uData, 4, "UChar") ; Offset 4: Buttons 1-8
            byte5 := NumGet(uData, 5, "UChar") ; Offset 5: Buttons 9-15

            ; Parse Logic
            byte4_new := byte4 & shuttlepro_old_4
            byte5_new := byte5 & shuttlepro_old_5
            shuttlepro_old_4 := ~byte4
            shuttlepro_old_5 := ~byte5

            ; Byte 1: Outer Wheel (Speed)
            if (byte1 != shuttlepro_speed_saved) {
                stop_all_speed_timers()
                execute_shuttlepro_speed(byte1, 1)
            }
            shuttlepro_speed_saved := byte1

            ; Raw Data Display
            ; 這裡將 V1 的 raw data 顯示邏輯整合進來
            rawInfo := Format("B1:{} B2:{} B4:{} B5:{}", byte1, byte2, byte4, byte5)
            a .= " " . rawInfo

            ; Byte 2: Inner Wheel (Shuttle)
            if (shuttlepro_shuttle_start) {
                shuttlepro_shuttle_saved := byte2
                shuttlepro_shuttle_start := False
            } else {
                if (byte2 != shuttlepro_shuttle_saved) {
                    execute_shuttlepro_shuttle(byte2, 1)
                }
            }
            shuttlepro_shuttle_saved := byte2

            ; Buttons: Byte 4 (Key 1-8)
            loop 8 {
                if (byte4_new & 1)
                    execute_shuttlepro(A_Index, 1)
                byte4_new >>= 1
            }

            ; Buttons: Byte 5 (Key 9-15)
            i := 9
            loop 7 {
                if (byte5_new & 1)
                    execute_shuttlepro(i, 1)
                byte5_new >>= 1
                i++
            }

            ; Update GUI (Append Mode)
            try {
                ; Add timestamp for better debugging
                finalMsg := Format("{:02}:{:02}:{:02} {}", A_Hour, A_Min, A_Sec, a)

                lbxInput.Add([finalMsg])
                count := SendMessage(0x018B, 0, 0, lbxInput.Hwnd)
                if (count > 0)
                    SendMessage(0x0186, count - 1, 0, lbxInput.Hwnd)
            }
        }
    }
}

execute_shuttlepro(key, layer) {
    global a

    ; Re-enabled debug info string
    a .= " key: " . key . " in layer " . layer

    if WinActive("ahk_exe WEBVIE~1.EXE") {
        if (key = 1)
            Send "q"
        else if (key = 2)
            Send "{F4}"
        else if (key = 3)
            Send "{F3}"
        else if (key = 4)
            Send "{F1}"
        else if (key = 5)
            Send "!t"
        else if (key = 8)
            Send "{F5}"
        else if (key = 9)
            Send "{F2}"
        else if (key = 10)
            Send "^{F12}"
        else if (key = 11)
            Send "f"
        else if (key = 13)
            Send "w"
        else if (key = 14)
            Send "!p"
        else if (key = 15)
            Send "^!w"

    } else if (WinActive("ahk_exe G3PACS.exe") || WinActive("172.17.12.174 - 遠端桌面連線")) {
        if (key = 1)
            Send "q"
        else if (key = 2)
            Send "7"    ; skull window
        else if (key = 3)
            Send "5"    ; bone window
        else if (key = 4)
            Send "6"    ; lung window
        else if (key = 5)
            Send "d"
        else if (key = 6)
            Send "8"    ; neck window
        else if (key = 7)
            Send "1"    ; brain window
        else if (key = 8)
            Send "9"    ; liver window
        else if (key = 9)
            Send "4"    ; abdomen window
        else if (key = 10)
            Send "z"    ; zoom in
        else if (key = 11) {
            if WinActive("172.17.12.174 - 遠端桌面連線")
                Send "f"
            else
                ToggleDiffExamSync()
        } else if (key = 12) {
            Send "+z"   ; zoom out
        } else if (key = 13) {
            if WinActive("172.17.12.174 - 遠端桌面連線")
                Send "w"
            else
                ToggleSync()
        } else if (key = 14)
            Send "o"
        else if (key = 15)
            Send "{Down}"

    } else if WinActive("ahk_exe Report.exe") {
        if (key = 4) { ; Close prev exam list window
            if WinActive("查詢歷史報告") {
                try {
                    isPopHisReport := ControlGetVisible("TBitBtn6", "查詢歷史報告")
                    if (isPopHisReport)
                        ControlClick "TBitBtn6", "查詢歷史報告"
                    else {
                        ControlClick "TBitBtn1", "查詢歷史報告"
                        Sleep 100
                        ControlFocus "TMemo6", "ahk_exe Report.exe"
                    }
                }
            }
        }
    }
}

execute_shuttlepro_speed(speed, layer) {
    global shuttlepro_speed_saved, timer_active_hwnd, a

    ; Convert unsigned byte to signed logic (-7 to 7)
    if (speed > 200)
        speed := speed - 256

    corrected_speed_saved := shuttlepro_speed_saved
    if (corrected_speed_saved > 200)
        corrected_speed_saved := corrected_speed_saved - 256

    timer_active_hwnd := WinExist("A")

    ; Restored V1 debug strings
    if (WinActive("172.17.12.174 - 遠端桌面連線")) {
        a .= "Remote Desktop: "
        set_scroll_speed(corrected_speed_saved, speed, 800, 600, 333, 200, 100, 50, 20)
    } else if (WinActive("ahk_exe G3PACS.exe")) {
        a .= "GEUV: "
        set_scroll_speed(corrected_speed_saved, speed, 800, 600, 333, 200, 100, 50, 20)
    } else if (WinActive("ahk_exe WEBVIE~1.EXE")) {
        a .= "EBM: "
        set_scroll_speed(corrected_speed_saved, speed, 800, 600, 333, 200, 100, 50, 20)
    } else if (WinActive("ahk_exe XWinGEAWE32.exe")) {
        a .= "AWS: "
        set_scroll_speed(corrected_speed_saved, speed, 800, 600, 333, 200, 100, 50, 20)
    }

    a .= "corrected_speed_saved: " . corrected_speed_saved . ", speed: " . speed
}

execute_shuttlepro_shuttle(shuttle_value, layer) {
    global shuttlepro_shuttle_saved, a

    ; Restored V1 debug strings
    if WinActive("172.17.12.174 - 遠端桌面連線") {
        a .= "Remote Desktop "
        if is_shuttle_clockwise(shuttle_value, shuttlepro_shuttle_saved)
            Click "WheelDown"
        else
            Click "WheelUp"

    } else if WinActive("ahk_exe G3PACS.exe") {
        a .= "GEUV "
        if is_shuttle_clockwise(shuttle_value, shuttlepro_shuttle_saved)
            Click "WheelDown"
        else
            Click "WheelUp"

    } else if WinActive("ahk_exe XWinGEAWE32.exe") {
        a .= "AWS "
        if is_shuttle_clockwise(shuttle_value, shuttlepro_shuttle_saved)
            Click "WheelDown"
        else
            Click "WheelUp"

    } else {
        try {
            MouseGetPos , , &id
            title := WinGetTitle(id)
            a .= '"' . title . '" '

            if is_shuttle_clockwise(shuttle_value, shuttlepro_shuttle_saved)
                Click "WheelDown"
            else
                Click "WheelUp"
        }
    }

    a .= "shuttle: " . shuttle_value . "(prev: " . shuttlepro_shuttle_saved . ") in layer " . layer
}

is_shuttle_clockwise(shuttle_value, shuttlepro_shuttle_saved) {
    if (shuttle_value = 0)
        return (shuttlepro_shuttle_saved > 127)
    else if (shuttlepro_shuttle_saved = 0)
        return (shuttle_value < 128)
    else
        return (shuttle_value > shuttlepro_shuttle_saved)
}

stop_all_speed_timers() {
    SetTimer UpScroll, 0
    SetTimer DownScroll, 0
}

UpKey() {
    Send "{Up}"
}

DownKey() {
    Send "{Down}"
}

set_scroll_speed(corrected_speed_saved, speed, s1, s2, s3, s4, s5, s6, s7) {
    global scroll_direction

    if (corrected_speed_saved < speed && speed > 0) {
        Click "WheelDown"
    } else if (corrected_speed_saved > speed && speed < 0) {
        Click "WheelUp"
    }

    if (speed < 0) {
        scroll_direction := UpScroll
    } else {
        scroll_direction := DownScroll
    }

    abs_speed := Abs(speed)

    period := 0
    switch abs_speed {
        case 1: period := s1
        case 2: period := s2
        case 3: period := s3
        case 4: period := s4
        case 5: period := s5
        case 6: period := s6
        case 7: period := s7
    }

    if (period > 0)
        SetTimer scroll_direction, period
}

UpScroll() {
    global timer_active_hwnd
    if (WinExist("A") != timer_active_hwnd) {
        SetTimer UpScroll, 0
    } else {
        Click "WheelUp"
    }
}

DownScroll() {
    global timer_active_hwnd
    if (WinExist("A") != timer_active_hwnd) {
        SetTimer DownScroll, 0
    } else {
        Click "WheelDown"
    }
}

ToggleSync() {
    try {
        focusedHwnd := ControlGetFocus("A")
        FocusedClassNN := ""
        if (focusedHwnd)
            FocusedClassNN := ControlGetClassNN(focusedHwnd)
        OutputVar := WinGetTitle("A")

        if (OutputVar = "INFINITT PACS" && SubStr(FocusedClassNN, 1, 3) == "Afx") {
            ;MsgBox("Toggling Sync Button")
            DiffSyncBtns := ["Button1", "Button85"]
            for idx, btn in DiffSyncBtns {
                try {
                    t := ControlGetText(btn, "A")
                    if (t = "自動同步") {
                        ControlClick btn, "A"
                        break
                    }
                }
            }
        }
    }
}

ToggleDiffExamSync() {
    try {
        focusedHwnd := ControlGetFocus("A")
        FocusedClassNN := ""
        if (focusedHwnd)
            FocusedClassNN := ControlGetClassNN(focusedHwnd)
        OutputVar := WinGetTitle("A")

        if (OutputVar = "INFINITT PACS" && SubStr(FocusedClassNN, 1, 3) == "Afx") {
            DiffSyncBtns := ["Button2", "Button86", "Button91"]
            for idx, btn in DiffSyncBtns {
                try {
                    t := ControlGetText(btn, "A")
                    if (t = "不同檢查同步 ") {
                        ControlClick btn, "A"
                        break
                    }
                }
            }
        }
    }
}

MainGui.OnEvent("Close", (*) => ExitApp())