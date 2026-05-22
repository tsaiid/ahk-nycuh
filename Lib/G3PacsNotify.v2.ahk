#Requires AutoHotkey v2.0

class G3PacsNotify {
    static _gui := 0
    static _text := 0
    static _hideTimer := 0
    static _width := 680

    static Show(message, duration := 1800) {
        message := Trim(message)
        if (message = "")
            return

        if this._gui {
            this._gui.Destroy()
            this._gui := 0
        }
        this._EnsureGui(message)
        SetTimer(this._hideTimer, 0)
        this._gui.Show("AutoSize Hide")

        WinGetPos(, , &width, &height, this._gui.Hwnd)
        MouseGetPos(&mouseX, &mouseY)
        workArea := this._GetMonitorWorkAreaAtPoint(mouseX, mouseY)
        x := workArea.left + Floor(((workArea.right - workArea.left) - width) / 2)
        y := workArea.top + Floor(((workArea.bottom - workArea.top) - height) / 2)

        this._gui.Show(Format("NoActivate x{1} y{2}", Max(workArea.left, x), Max(workArea.top, y)))
        this._ApplyVisualStyle()

        if (duration > 0)
            SetTimer(this._hideTimer, -duration)
    }

    static _EnsureGui(message := "") {
        if this._gui
            return

        g := Gui("+AlwaysOnTop -Caption +ToolWindow +E0x08000000")
        g.BackColor := "F8FAFC"
        g.MarginX := 32
        g.MarginY := 20
        g.SetFont("s18 c111827 bold", "Microsoft JhengHei UI")

        this._gui := g
        this._text := g.Add("Text", "w" this._width " Center", message)
        this._hideTimer := ObjBindMethod(this, "_Hide")
    }

    static _Hide() {
        if this._gui
            this._gui.Hide()
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
            WinSetRegion("0-0 w" width " h" height " r14-14", hwnd)

            style := DllCall("GetClassLongPtr", "Ptr", hwnd, "Int", -26, "Ptr")
            DllCall("SetClassLongPtr", "Ptr", hwnd, "Int", -26, "Ptr", style | 0x00020000)

            WinSetTransparent(245, hwnd)
        }
    }
}
