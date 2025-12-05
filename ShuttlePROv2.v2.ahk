#Requires AutoHotkey v2.0
#SingleInstance force
SetWorkingDir A_ScriptDir
SetControlDelay -1
CoordMode "Mouse", "Screen"

#Include <AHKHID.v2>

global App := ShuttleProController()

class ShuttleProController {
    LastSpeed := 0
    LastShuttle := 0
    LastBtn1_8 := 255
    LastBtn9_15 := 255

    ; 統一使用 IsFirstRun 作為旗標
    IsFirstRun := True

    ; 預設滾動速度 (毫秒)
    DefaultSpeeds := [800, 600, 333, 200, 100, 50, 20]

    ScrollDirection := 0
    CurrentTimerObj := unset
    GuiObj := unset
    LbxLog := unset
    AppList := []

    __New() {
        this.CurrentTimerObj := ObjBindMethod(this, "AutoScroll")
        this.InitGui()
        this.InitConfig()
        this.RegisterHID()
    }

    ; ==========================================================================
    ; [核心設定區]
    ; ==========================================================================
    InitConfig() {
        ; 1. EBM Web Viewer
        this.RegisterApp("WEBVIEWER", "ahk_exe WEBVIE~1.EXE", Map(
            1, "q",
            2, "{F4}",
            3, "{F3}",
            4, "{F1}",
            5, "!t",
            8, "{F5}",
            9, "{F2}",
            10, "^{F12}",
            11, "f",
            13, "w",
            14, "!p",
            15, "^!w"
        ))

        ; 2. Remote Desktop (RDP)
        rdpSpeeds := [1000, 800, 500, 300, 200, 100, 50]
        this.RegisterApp("RDP", "172.17.12.174 - 遠端桌面連線", Map(
            1, "q",
            2, "7",
            3, "5",
            4, "6",
            5, "d",
            6, "8",
            7, "1",
            8, "9",
            9, "4",
            10, "z",
            11, "f",
            12, "+z",
            13, "w",
            14, "o",
            15, "{Down}"
        ), rdpSpeeds)

        ; 3. G3PACS (Infinitt)
        pacsSpeeds := [800, 600, 333, 200, 100, 50, 20]
        this.RegisterApp("G3PACS", "ahk_exe G3PACS.exe", Map(
            1, "q",
            2, "7",
            3, "5",
            4, "6",
            5, "d",
            6, "8",
            7, "1",
            8, "9",
            9, "4",
            10, "z",
            11, (*) => this.ToggleDiffExamSync(),
            12, "+z",
            13, (*) => this.ToggleSync(),
            14, "o", 15, "{Down}"
        ), pacsSpeeds)

        ; 4. AWS (GE)
        this.RegisterApp("AWS", "ahk_exe XWinGEAWE32.exe", Map(
            1, "q",
            2, "7",
            3, "5",
            4, "6",
            5, "d",
            6, "8",
            7, "1",
            8, "9",
            9, "4",
            10, "z",
            11, "f",
            12, "+z",
            13, "w",
            14, "o",
            15, "{Down}"
        ))

        ; 5. Report System
        this.RegisterApp("REPORT", "ahk_exe Report.exe", Map(
            4, (*) => this.HandleReportHistory()
        ))
    }

    RegisterApp(name, winTitle, keyMap, speeds := []) {
        if (speeds.Length == 0)
            useSpeeds := this.DefaultSpeeds
        else
            useSpeeds := speeds

        this.AppList.Push({
            Name: name,
            WinTitle: winTitle,
            KeyMap: keyMap,
            Speeds: useSpeeds
        })
    }

    GetActiveContext() {
        for app in this.AppList {
            if WinActive(app.WinTitle)
                return app
        }
        return ""
    }

    ; ==========================================================================
    ; HID 處理邏輯
    ; ==========================================================================
    RegisterHID() {
        GuiHandle := this.GuiObj.Hwnd
        AHKHID.Register(12, 1, GuiHandle, AHKHID.RIDEV_INPUTSINK)
        OnMessage(0x00FF, (wParam, lParam, msg, hwnd) => this.OnInput(wParam, lParam, msg, hwnd))
    }

    OnInput(wParam, lParam, msg, hwnd) {
        Critical

        devicetype := AHKHID.GetInputInfo(lParam, AHKHID.II_DEVTYPE)
        if (devicetype != AHKHID.RIM_TYPEHID)
            return

        hid_handle := AHKHID.GetInputInfo(lParam, AHKHID.II_DEVHANDLE)
        vendor_id := AHKHID.GetDevInfo(hid_handle, AHKHID.DI_HID_VENDORID, True)
        product_id := AHKHID.GetDevInfo(hid_handle, AHKHID.DI_HID_PRODUCTID, True)

        if (vendor_id != 2867 || product_id != 48)
            return

        uData := AHKHID.GetInputData(lParam)
        if (uData.Size = 0)
            return

        byte1 := NumGet(uData, 1, "UChar")
        byte2 := NumGet(uData, 2, "UChar")
        byte4 := NumGet(uData, 4, "UChar")
        byte5 := NumGet(uData, 5, "UChar")

        debugStr := Format("B1:{} B2:{} B4:{} B5:{}", byte1, byte2, byte4, byte5)

        currentApp := this.GetActiveContext()
        appName := (currentApp != "") ? currentApp.Name : "Unknown"

        ; [核心修改] 第一次執行時，只同步狀態，不執行動作
        if (this.IsFirstRun) {
            this.LastSpeed := byte1
            this.LastShuttle := byte2
            this.LastBtn1_8 := byte4
            this.LastBtn9_15 := byte5
            this.IsFirstRun := False
            this.Log("Initialized: " . debugStr)
            return ; 直接返回，不處理後續邏輯
        }

        ; --- 1. 處理 Outer Ring (Speed) ---
        if (byte1 != this.LastSpeed) {
            this.HandleOuterRing(byte1, currentApp)
        }
        this.LastSpeed := byte1

        ; --- 2. 處理 Inner Ring (Jog) ---
        if (byte2 != this.LastShuttle) {
            this.HandleInnerJog(byte2)
            debugStr .= " Jog"
        }
        this.LastShuttle := byte2

        ; --- 3. 處理 Buttons 1-8 ---
        diff4 := (this.LastBtn1_8 ^ byte4) & this.LastBtn1_8
        if (diff4 > 0) {
            loop 8 {
                if (diff4 & 1)
                    this.ExecuteAction(A_Index, currentApp)
                diff4 >>= 1
            }
        }
        this.LastBtn1_8 := byte4

        ; --- 4. 處理 Buttons 9-15 ---
        diff5 := (this.LastBtn9_15 ^ byte5) & this.LastBtn9_15
        if (diff5 > 0) {
            i := 9
            loop 7 {
                if (diff5 & 1)
                    this.ExecuteAction(i, currentApp)
                diff5 >>= 1
                i++
            }
        }
        this.LastBtn9_15 := byte5

        this.Log(debugStr . " [" . appName . "]")
    }

    ExecuteAction(keyID, appCtx) {
        if (appCtx == "")
            return

        if (appCtx.KeyMap.Has(keyID)) {
            action := appCtx.KeyMap[keyID]
            if (Type(action) = "Func" || Type(action) = "Closure") {
                action.Call()
            } else {
                Send(action)
            }
            this.Log("Action Key: " . keyID)
        }
    }

    ; ==========================================================================
    ; 滾輪與轉盤邏輯 (修正版)
    ; ==========================================================================
    HandleOuterRing(rawSpeed, appCtx) {
        ; 1. 轉換為有號整數 (-7 到 7)
        newSpeed := (rawSpeed > 200) ? (rawSpeed - 256) : rawSpeed
        oldSpeed := (this.LastSpeed > 200) ? (this.LastSpeed - 256) : this.LastSpeed

        ; 2. 停止當前的 Timer (重置節奏)
        SetTimer this.CurrentTimerObj, 0

        ; 3. [核心還原]: 模擬舊版的 Bounce 機制
        ; 只有在「加速」的時候才觸發立即滾動，減速時不觸發
        if (oldSpeed < newSpeed && newSpeed > 0) {
            ; 正向加速 (例如 0->1, 1->2)
            Click "WheelDown"
        } else if (oldSpeed > newSpeed && newSpeed < 0) {
            ; 反向加速 (例如 0->-1, -1->-2)
            Click "WheelUp"
        }

        ; 4. 設定後續的自動滾動 Timer
        if (newSpeed = 0)
            return

        ; 設定滾動方向 (供 Timer 使用)
        this.ScrollDirection := (newSpeed > 0) ? 1 : -1

        ; 取得對應 App 的速度設定
        speedSettings := this.DefaultSpeeds
        if (appCtx != "") {
            speedSettings := appCtx.Speeds
        }

        ; 啟動 Timer
        absSpeed := Abs(newSpeed)
        if (absSpeed >= 1 && absSpeed <= speedSettings.Length) {
            period := speedSettings[absSpeed]
            SetTimer this.CurrentTimerObj, period
        }
    }

    AutoScroll() {
        if (this.ScrollDirection = 1)
            Click "WheelDown"
        else
            Click "WheelUp"
    }

    HandleInnerJog(newVal) {
        ; 1. 計算當前數值與上一次數值的差值
        diff := newVal - this.LastShuttle

        ; 2. 處理 0 <-> 255 的邊界跨越問題 (Wrap-around Correction)
        if (diff > 128)
            diff := diff - 256
        else if (diff < -128)
            diff := diff + 256

        ; 3. 根據修正後的 diff 執行動作
        if (diff > 0) {
            Click "WheelDown"
        } else if (diff < 0) {
            Click "WheelUp"
        }
    }

    ; ==========================================================================
    ; 特殊功能函數
    ; ==========================================================================
    HandleReportHistory() {
        if WinActive("查詢歷史報告") {
            try {
                if ControlGetVisible("TBitBtn6", "查詢歷史報告")
                    ControlClick "TBitBtn6", "查詢歷史報告"
                else {
                    ControlClick "TBitBtn1", "查詢歷史報告"
                    Sleep 100
                    ControlFocus "TMemo6", "ahk_exe Report.exe"
                }
            }
        }
    }

    ToggleSync() {
        this.ClickButtonIfTextMatches("INFINITT PACS", ["Button1", "Button85"], "自動同步")
    }

    ToggleDiffExamSync() {
        this.ClickButtonIfTextMatches("INFINITT PACS", ["Button2", "Button86", "Button91"], "不同檢查同步 ")
    }

    ClickButtonIfTextMatches(winTitle, btnList, targetText) {
        if !WinActive(winTitle)
            return

        focusedHwnd := ControlGetFocus("A")
        if (!focusedHwnd || SubStr(ControlGetClassNN(focusedHwnd), 1, 3) != "Afx")
            return

        for btn in btnList {
            try {
                if (ControlGetText(btn, "A") = targetText) {
                    ControlClick btn, "A"
                    this.Log("Clicked " . btn)
                    return
                }
            }
        }
    }

    ; ==========================================================================
    ; Debug GUI
    ; ==========================================================================
    InitGui() {
        this.GuiObj := Gui("+LastFound -Resize -MaximizeBox -MinimizeBox", "ShuttlePro V2 Configurable")
        this.GuiObj.SetFont("w700 s8", "Courier New")
        this.LbxLog := this.GuiObj.Add("ListBox", "h300 w600")
        this.GuiObj.OnEvent("Close", (*) => ExitApp())
        this.GuiObj.Show()
    }

    Log(msg) {
        timestamp := Format("{:02}:{:02}:{:02}", A_Hour, A_Min, A_Sec)
        try {
            this.LbxLog.Add([timestamp . " " . msg])
            SendMessage(0x018B, 0, 0, this.LbxLog.Hwnd)
            count := SendMessage(0x018B, 0, 0, this.LbxLog.Hwnd)
            if (count > 0)
                SendMessage(0x0186, count - 1, 0, this.LbxLog.Hwnd)
        }
    }
}