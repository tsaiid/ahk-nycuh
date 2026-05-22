#Requires AutoHotkey v2.0

class G3PacsNotify {
    static _gui := 0
    static _text := 0
    static _hideTimer := 0
    static _width := 360
    static _scale := 1.0

    static Show(message, duration := 1800) {
        message := Trim(message)
        if (message = "")
            return

        ; 在建立 GUI 之前，先獲取滑鼠所在點的座標
        MouseGetPos(&mouseX, &mouseY)

        if this._gui {
            this._gui.Destroy()
            this._gui := 0
        }
        this._EnsureGui(message, mouseX, mouseY)
        SetTimer(this._hideTimer, 0)
        this._gui.Show("AutoSize Hide")

        WinGetPos(, , &width, &height, this._gui.Hwnd)
        workArea := this._GetMonitorWorkAreaAtPoint(mouseX, mouseY)
        x := workArea.left + Floor(((workArea.right - workArea.left) - width) / 2)
        y := workArea.top + Floor(((workArea.bottom - workArea.top) - height) / 2)

        this._gui.Show(Format("NoActivate x{1} y{2}", Max(workArea.left, x), Max(workArea.top, y)))
        this._ApplyVisualStyle()

        if (duration > 0)
            SetTimer(this._hideTimer, -duration)
    }

    static _EnsureGui(message := "", mouseX := 0, mouseY := 0) {
        if this._gui
            return

        ; 取得滑鼠所在螢幕的 DPI 並計算縮放比例
        dpi := this._GetDpiAtPoint(mouseX, mouseY)
        this._scale := dpi / 96

        ; 使用 -DPIScale 停用 AHK 預設的主螢幕 DPI 縮放，完全由我們手動計算以適配多螢幕 DPI
        g := Gui("+AlwaysOnTop -Caption +ToolWindow +E0x08000000 -DPIScale")
        g.BackColor := "F8FAFC"
        
        ; 根據 DPI 比例縮放邊距
        g.MarginX := Round(16 * this._scale)
        g.MarginY := Round(8 * this._scale)
        
        ; 根據 DPI 比例縮放字型大小
        fontSize := Round(12 * this._scale)
        g.SetFont("s" fontSize " c111827 bold", "Microsoft JhengHei UI")

        this._gui := g
        
        ; 根據 DPI 比例縮放控制項寬度
        ctrlWidth := Round(this._width * this._scale)
        this._text := g.Add("Text", "w" ctrlWidth " Center", message)
        this._hideTimer := ObjBindMethod(this, "_Hide")
    }

    static _Hide() {
        if this._gui
            this._gui.Hide()
    }

    static _GetDpiAtPoint(x, y) {
        hMonitor := 0
        if (A_PtrSize == 8) {
            hMonitor := DllCall("MonitorFromPoint", "Int64", (x & 0xFFFFFFFF) | (y << 32), "UInt", 2, "Ptr")
        } else {
            hMonitor := DllCall("MonitorFromPoint", "Int", x, "Int", y, "UInt", 2, "Ptr")
        }
        if (hMonitor) {
            dpiX := 0, dpiY := 0
            try {
                ; MDT_EFFECTIVE_DPI = 0
                DllCall("Shcore\GetDpiForMonitor", "Ptr", hMonitor, "Int", 0, "UInt*", &dpiX, "UInt*", &dpiY)
                if (dpiX > 0)
                    return dpiX
            }
        }
        return 96 ; 預設 96 DPI (100% 縮放)
    }

    static _GetMonitorWorkAreaAtPoint(x, y) {
        monitorCount := MonitorGetCount()
        loop monitorCount {
            MonitorGet(A_Index, &left, &top, &right, &bottom)
            if (x >= left && x < right && y >= top && y < bottom) {
                MonitorGetWorkArea(A_Index, &workLeft, &workTop, &workRight, &workBottom)
                return {left: workLeft, top: workTop, right: workRight, bottom: workBottom}
            }
        }

        MonitorGetWorkArea(, &workLeft, &workTop, &workRight, &workBottom)
        return {left: workLeft, top: workTop, right: workRight, bottom: workBottom}
    }

    static _ApplyVisualStyle() {
        if !this._gui
            return

        hwnd := this._gui.Hwnd
        WinGetPos(, , &width, &height, hwnd)
        try {
            ; 圓角半徑也根據 DPI 比例進行縮放
            r := Round(10 * this._scale)
            WinSetRegion("0-0 w" width " h" height " r" r "-" r, hwnd)

            style := DllCall("GetClassLongPtr", "Ptr", hwnd, "Int", -26, "Ptr")
            DllCall("SetClassLongPtr", "Ptr", hwnd, "Int", -26, "Ptr", style | 0x00020000)

            WinSetTransparent(245, hwnd)
        }
    }
}
