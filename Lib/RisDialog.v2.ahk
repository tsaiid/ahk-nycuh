#Requires AutoHotkey v2.0

#Include .\ShowGUIatCurrScreenCenter.v2.ahk

/**
 * 專案標準對話框基底模組
 * 封裝 GUI 主題風格、螢幕置中計算、DWM 渲染與生命週期管理
 */
class RisDialog {
    static BG_COLOR := "F4F5F7"
    static DEFAULT_FONT_NAME := "Microsoft JhengHei UI"
    static DEFAULT_FONT_SIZE := "s11"
    static DEFAULT_MARGIN_X := 16
    static DEFAULT_MARGIN_Y := 14

    /**
     * 一鍵建立已配置標準主題的 Gui 物件
     * @param {String} title 視窗標題
     * @param {String} options Gui 建立選項 (預設 "+AlwaysOnTop +ToolWindow")
     * @param {Object} styleOpts 主題覆蓋選項 { BackColor, MarginX, MarginY, FontName, FontSize, FontOptions }
     * @returns {Gui}
     */
    static Create(title := "", options := "+AlwaysOnTop +ToolWindow", styleOpts := 0) {
        guiObj := Gui(options, title)
        this.ApplyTheme(guiObj, styleOpts)
        return guiObj
    }

    /**
     * 為既有 Gui 物件套用標準主題風格
     * @param {Gui} guiObj 目標 Gui 物件
     * @param {Object} styleOpts 主題覆蓋選項
     * @returns {Gui}
     */
    static ApplyTheme(guiObj, styleOpts := 0) {
        bgColor := (IsObject(styleOpts) && styleOpts.HasOwnProp("BackColor")) ? styleOpts.BackColor : this.BG_COLOR
        marginX := (IsObject(styleOpts) && styleOpts.HasOwnProp("MarginX")) ? styleOpts.MarginX : this.DEFAULT_MARGIN_X
        marginY := (IsObject(styleOpts) && styleOpts.HasOwnProp("MarginY")) ? styleOpts.MarginY : this.DEFAULT_MARGIN_Y
        fontName := (IsObject(styleOpts) && styleOpts.HasOwnProp("FontName")) ? styleOpts.FontName : this.DEFAULT_FONT_NAME
        fontSize := (IsObject(styleOpts) && styleOpts.HasOwnProp("FontSize")) ? styleOpts.FontSize : this.DEFAULT_FONT_SIZE
        fontOptions := (IsObject(styleOpts) && styleOpts.HasOwnProp("FontOptions")) ? styleOpts.FontOptions : ""

        guiObj.BackColor := bgColor
        guiObj.MarginX := marginX
        guiObj.MarginY := marginY

        fontSpec := Trim(fontSize . " " . fontOptions)
        guiObj.SetFont(fontSpec, fontName)

        return guiObj
    }

    /**
     * 將 Gui 顯示於當前滑鼠所在螢幕中央，並套用 DWM 視窗樣式
     * @param {Gui} guiObj 目標 Gui 物件
     * @param {String} extraShowOpts 額外 Show 參數 (例如 "AutoSize")
     */
    static ShowAtCurrentMonitorCenter(guiObj, extraShowOpts := "") {
        currentMonitorIndex := GetCurrentMonitorIndex()
        guiObj.Show("Hide")
        guiWidth := 0, guiHeight := 0
        GetClientSize(guiObj.Hwnd, &guiWidth, &guiHeight)
        guiX := CoordXCenterScreen(guiWidth, currentMonitorIndex)
        guiY := CoordYCenterScreen(guiHeight, currentMonitorIndex)
        showCmd := "x" . guiX . " y" . guiY . (extraShowOpts != "" ? " " . extraShowOpts : "")
        guiObj.Show(showCmd)
        this.ApplyWindowStyle(guiObj.Hwnd)
    }

    /**
     * 將 Gui 顯示於主螢幕中央，並套用 DWM 視窗樣式
     * @param {Gui} guiObj 目標 Gui 物件
     * @param {String} extraShowOpts 額外 Show 參數 (例如 "AutoSize")
     */
    static ShowCenter(guiObj, extraShowOpts := "") {
        showCmd := "Center" . (extraShowOpts != "" ? " " . extraShowOpts : "")
        guiObj.Show(showCmd)
        this.ApplyWindowStyle(guiObj.Hwnd)
    }

    /**
     * 銷毀 Gui 並將焦點還原至原視窗
     * @param {Gui} guiObj 目標 Gui 物件
     * @param {Integer|String} parentWnd 父視窗 HWND
     */
    static CloseAndRestoreFocus(guiObj, parentWnd := 0) {
        try {
            guiObj.Destroy()
        }
        if (parentWnd && WinExist("ahk_id " . parentWnd)) {
            try {
                WinActivate("ahk_id " . parentWnd)
            }
        }
    }

    /**
     * 套用視窗樣式（Win10 移除 DWM 陰影與邊框，Win11 保留陰影但消除邊框）
     * @param {Integer} hwnd 視窗控制代碼
     */
    static ApplyWindowStyle(hwnd) {
        try {
            if (this.GetWindowsBuildNumber() < 22000) {
                renderingPolicy := 1 ; DWMNCRP_DISABLED: remove Win10 DWM shadow/edge.
                DllCall("Dwmapi\DwmSetWindowAttribute", "Ptr", hwnd, "UInt", 2, "Int*", renderingPolicy, "UInt", 4)
                return
            }

            borderColor := 0xFFFFFFFE ; DWMWA_COLOR_NONE: remove Win11 DWM outline while keeping shadow.
            DllCall("Dwmapi\DwmSetWindowAttribute", "Ptr", hwnd, "UInt", 34, "UInt*", borderColor, "UInt", 4)
        } catch {
        }
    }

    /**
     * 取得 Windows 系統 Build 編號
     * @returns {Integer}
     */
    static GetWindowsBuildNumber() {
        return RegExMatch(A_OSVersion, "^\d+\.\d+\.(\d+)", &match) ? Integer(match[1]) : 0
    }
}
