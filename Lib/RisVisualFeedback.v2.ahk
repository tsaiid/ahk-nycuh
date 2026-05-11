#Requires AutoHotkey v2.0

class RisVisualFeedback {
    static _workingCursorCache := {AppStarting: "", CursorBaseSize: "", Path: ""}

    ; 優先依照使用者目前的 cursor scheme 與 CursorBaseSize 選擇 working 游標檔。
    ; 注意：SetSystemCursor 只會套用靜態游標內容，無法保留 .ani 動畫。
    static ShowWaitCursor() {
        try {
            OCR_NORMAL := 32512
            OCR_IBEAM := 32513
            cursorPath := this.ResolveWorkingCursorPath()

            if (cursorPath = "")
                return

            ; SetSystemCursor 會接手並銷毀 handle，因此需各自載入一份複本。
            hCopyNormal := DllCall("LoadCursorFromFile", "Str", cursorPath, "Ptr")
            hCopyIBeam  := DllCall("LoadCursorFromFile", "Str", cursorPath, "Ptr")

            if (!hCopyNormal || !hCopyIBeam)
                return

            DllCall("SetSystemCursor", "Ptr", hCopyNormal, "Int", OCR_NORMAL)
            DllCall("SetSystemCursor", "Ptr", hCopyIBeam, "Int", OCR_IBEAM)
        }
    }

    static ResolveWorkingCursorPath() {
        try {
            appStarting := RegRead("HKCU\Control Panel\Cursors", "AppStarting", "")
        } catch {
            appStarting := ""
        }

        try {
            cursorBaseSizeRaw := RegRead("HKCU\Control Panel\Cursors", "CursorBaseSize", "32")
        } catch {
            cursorBaseSizeRaw := "32"
        }

        cache := this._workingCursorCache
        if (cache.AppStarting = appStarting
            && cache.CursorBaseSize = cursorBaseSizeRaw
            && cache.Path != ""
            && FileExist(cache.Path)) {
            return cache.Path
        }

        cursorBaseSize := Integer(cursorBaseSizeRaw)
        cursorPath := appStarting

        if (cursorPath = "" || !FileExist(cursorPath)) {
            cursorPath := A_WinDir . "\Cursors\aero_working.ani"
        }

        if (cursorBaseSize >= 64) {
            sizedPath := RegExReplace(cursorPath, "i)(\.[^\\.]*)$", "_xl$1")
            if (sizedPath != cursorPath && FileExist(sizedPath)) {
                this._workingCursorCache := {AppStarting: appStarting, CursorBaseSize: cursorBaseSizeRaw, Path: sizedPath}
                return sizedPath
            }
        }

        if (cursorBaseSize >= 48) {
            sizedPath := RegExReplace(cursorPath, "i)(\.[^\\.]*)$", "_l$1")
            if (sizedPath != cursorPath && FileExist(sizedPath)) {
                this._workingCursorCache := {AppStarting: appStarting, CursorBaseSize: cursorBaseSizeRaw, Path: sizedPath}
                return sizedPath
            }
        }

        resolvedPath := FileExist(cursorPath) ? cursorPath : ""
        this._workingCursorCache := {AppStarting: appStarting, CursorBaseSize: cursorBaseSizeRaw, Path: resolvedPath}
        return resolvedPath
    }

    static RestoreCursor() {
        ; SPI_SETCURSORS = 0x0057, 重置系統所有游標回預設值
        DllCall("SystemParametersInfo", "UInt", 0x0057, "UInt", 0, "Ptr", 0, "UInt", 0)
    }

    ; 紅色 + 半透明 + 圓形，先裁切再顯示以避免閃爍。
    static HighlightCaret(hTargetCtrl := 0) {
        try {
            CoordMode "Caret", "Screen"
            CoordMode "Mouse", "Screen"

            x := 0, y := 0
            isFound := false

            if CaretGetPos(&cx, &cy) {
                x := cx
                y := cy
                isFound := true
            } else if (hTargetCtrl) {
                try {
                    WinGetPos(&wx, &wy, &ww, &wh, "ahk_id " hTargetCtrl)
                    x := wx + (ww / 2) - 20
                    y := wy + (wh / 2) - 20
                    isFound := true
                }
            }

            if (!isFound)
                return

            if (x == cx) {
                finalX := x - 20
                finalY := y - 10
            } else {
                finalX := x
                finalY := y
            }

            g := Gui("+AlwaysOnTop -Caption +ToolWindow +E0x20 +E0x08000000 -DPIScale")
            g.BackColor := "Red"

            try {
                WinSetRegion("0-0 w40 h40 E", g.Hwnd)
                WinSetTransparent(100, g.Hwnd)
            }

            g.Show("NA x" finalX " y" finalY " w40 h40")
            SetTimer () => (IsObject(g) ? g.Destroy() : ""), -400
        } catch {
        }
    }
}
