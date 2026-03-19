#Requires AutoHotkey v2.0
#SingleInstance force
SetWorkingDir A_ScriptDir
SetControlDelay -1
CoordMode "Mouse", "Screen"

; [新增] 設定工作列 (Tray) 上的 Icon
TraySetIcon(A_ScriptDir "\assets\ShuttlePROv2_icon.png")

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

    ; [新增] 用於處理過渡時間的屬性
    TargetPeriod := 0      ; 記錄目標循環時間
    IsTransitioning := False ; 標記是否正處於加速過渡期

    ; [新增] 啟動緩衝相關屬性
    StartupTimerObj := unset
    IsStartupPending := False ; 是否正處於「剛起步觀察期」
    StartupDelay := 80        ; 觀察期毫秒數 (建議 50~80 之間)

    GuiObj := unset
    LbxLog := unset
    AppList := []

    __New() {
        this.CurrentTimerObj := ObjBindMethod(this, "AutoScroll")
        ; [新增] 初始化啟動緩衝 Timer
        this.StartupTimerObj := ObjBindMethod(this, "ExecuteStartup")

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
            1, "8",
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
            14, "o",
            15, "0"
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
    ; 滾輪與轉盤邏輯 (雙向平滑版 + 零延遲啟動)
    ; ==========================================================================
    HandleOuterRing(rawSpeed, appCtx) {
        ; 1. 轉換數值 (-7 到 7)
        newSpeed := (rawSpeed > 200) ? (rawSpeed - 256) : rawSpeed
        oldSpeed := (this.LastSpeed > 200) ? (this.LastSpeed - 256) : this.LastSpeed

        ; 2. 判斷是否歸零 (停車)
        if (newSpeed = 0) {
            SetTimer this.CurrentTimerObj, 0  ; 停止循環 Timer
            SetTimer this.StartupTimerObj, 0  ; 停止啟動觀察 Timer
            this.IsTransitioning := False
            this.IsStartupPending := False
            return
        }

        ; 3. 設定方向與取得速度表
        this.ScrollDirection := (newSpeed > 0) ? 1 : -1
        speedSettings := (appCtx != "") ? appCtx.Speeds : this.DefaultSpeeds

        ; 4. 計算新週期 (Period)
        absNew := Abs(newSpeed)
        newPeriod := 0
        if (absNew >= 1 && absNew <= speedSettings.Length)
            newPeriod := speedSettings[absNew]

        if (newPeriod == 0)
            return

        ; ======================================================================
        ; [核心修改] 啟動緩衝邏輯
        ; ======================================================================

        ; 情況 A: 正在觀察期內 (例如 0->1 剛發生，尚未觸發，馬上又變成 2)
        if (this.IsStartupPending) {
            ; 這裡什麼都不用做 (Do Nothing)
            ; 因為我們只是更新了 this.LastSpeed (在外部 OnInput 會更新)，
            ; 並且更新了 newPeriod。
            ; 等到 Startup Timer 時間到時，它會自動讀取最新的速度去執行。
            ; 這樣就實現了「忽略中間檔位」的效果。
            return
        }

        ; 情況 B: 從靜止啟動 (0 -> X)
        if (oldSpeed == 0) {
            this.IsStartupPending := True
            ; 設定一個單次執行的 Timer，延遲 80ms
            SetTimer this.StartupTimerObj, -this.StartupDelay
            return
        }

        ; ======================================================================
        ; 以下為「已經在轉動中」的變速邏輯 (1 -> 2 或 2 -> 1)
        ; ======================================================================

        ; 必須先停止原本的 Timer，準備變速
        SetTimer this.CurrentTimerObj, 0
        this.IsTransitioning := False

        absOld := Abs(oldSpeed)
        oldPeriod := 0
        if (absOld >= 1 && absOld <= speedSettings.Length)
            oldPeriod := speedSettings[absOld]

        isAccelerating := (absNew > absOld)
        isDecelerating := (absNew < absOld)
        waitDelay := 0

        if (isAccelerating) {
            ; --- 加速邏輯 ---
            ; 等待時間 = 差值的一半 (避免太快暴衝)
            waitDelay := Floor(Abs(oldPeriod - newPeriod) / 2)
        } else if (isDecelerating) {
            ; --- 減速邏輯 (新增) ---
            ; 等待時間 = 兩者平均值 (填補時間空隙，模擬慣性)
            ; 例如：50ms -> 200ms，中間插入一個 125ms 的等待
            waitDelay := Floor((oldPeriod + newPeriod) / 2)
        } else {
            ; --- 穩定狀態或從靜止啟動 ---
            ; 直接使用新週期
            SetTimer this.CurrentTimerObj, newPeriod
            return
        }

        ; 6. 執行過渡 Timer
        ; 人類感知閾值 (約 40ms)，如果延遲太短直接執行以免無感
        if (waitDelay < 40) {
            this.AutoScroll()
            SetTimer this.CurrentTimerObj, newPeriod
        } else {
            this.IsTransitioning := True
            this.TargetPeriod := newPeriod
            SetTimer this.CurrentTimerObj, -waitDelay
        }
    }

    AutoScroll() {
        if (this.ScrollDirection = 1)
            Click "WheelDown"
        else
            Click "WheelUp"

        ; [新增] 如果剛剛是執行「過渡的一次性 Timer」
        ; 執行完這次動作後，立刻將 Timer 設回目標的穩定循環週期
        if (this.IsTransitioning) {
            SetTimer this.CurrentTimerObj, this.TargetPeriod
            this.IsTransitioning := False
        }
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

    ExecuteStartup() {
        ; 1. 緩衝期結束，標記解除
        this.IsStartupPending := False

        ; 2. 為了安全，重新獲取一次當前的目標週期
        ; 因為在等待的 60ms 內，User 可能已經換了 App 或是速度變了
        currentApp := this.GetActiveContext()
        speedSettings := (currentApp != "") ? currentApp.Speeds : this.DefaultSpeeds

        ; 獲取最新的速度 (利用 LastSpeed，因為它在 OnInput 已被更新到最新)
        currSpeed := (this.LastSpeed > 200) ? (this.LastSpeed - 256) : this.LastSpeed
        absSpeed := Abs(currSpeed)

        if (currSpeed == 0) ; 如果等待期間使用者又停下來了
            return

        finalPeriod := 0
        if (absSpeed >= 1 && absSpeed <= speedSettings.Length)
            finalPeriod := speedSettings[absSpeed]

        if (finalPeriod == 0)
            return

        ; 3. 立即執行第一槍 (達成無延遲感的啟動)
        ; 這時候執行的就是「最新的檔位」，中間的檔位被跳過了
        this.AutoScroll()

        ; 4. 設定循環 Timer 進入穩定狀態
        SetTimer this.CurrentTimerObj, finalPeriod
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
        this.ClickButtonIfTextMatches("INFINITT PACS", ["Button1", "Button85", "Button90", "Button102"], [" Auto sync", "自動同步"])
    }

    ToggleDiffExamSync() {
        this.ClickButtonIfTextMatches("INFINITT PACS", ["Button2", "Button86", "Button91", "Button103"], [" Sync with other exams", "不同檢查同步 "])
    }

    ClickButtonIfTextMatches(winTitle, btnList, targetText) {
        if !WinActive(winTitle)
            return

        focusedHwnd := ControlGetFocus("A")
        if (!focusedHwnd || SubStr(ControlGetClassNN(focusedHwnd), 1, 3) != "Afx")
            return

        for btn in btnList {
            try {
                curText := ControlGetText(btn, "A")
                isMatch := false

                if (Type(targetText) = "Array") {
                    for t in targetText {
                        if (curText = t) {
                            isMatch := true
                            break
                        }
                    }
                } else {
                    if (curText = targetText)
                        isMatch := true
                }

                if (isMatch) {
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
        ; 移除 -MinimizeBox 以允許視窗最小化
        this.GuiObj := Gui("+LastFound -Resize -MaximizeBox", "ShuttlePro V2 Configurable")
        this.GuiObj.SetFont("w700 s8", "Courier New")
        this.LbxLog := this.GuiObj.Add("ListBox", "h300 w600")

        ; 在右下方新增 Reload 按鈕 (X 座標設為 510 以對齊 ListBox 右邊界)
        btnReload := this.GuiObj.Add("Button", "w100 x510", "Reload Script")
        btnReload.OnEvent("Click", (*) => Reload())

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