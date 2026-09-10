#Requires AutoHotkey v2.0

class G3PacsProbe {
    static ClassPrefix := "Afx:00400000:b:00000000:00000013:00000000"
    static Configs := [
        [3, 10, 29, 101],
        [3, 10, 31, 101],
        [3, 10, 33, 101],
        [3, 10, 35, 101],
        [3, 10, 37, 101],
        [3, 10, 39, 101],
        [3, 10, 41, 101],
        [3, 10, 43, 101],
        [3, 10, 45, 101],
        [3, 10, 47, 101],
        [3, 10, 49, 101],
        [3, 10, 51, 101],
        [3, 10, 53, 101],
        [3, 10, 55, 101],
        [5, 57, 29, 185],
        [5, 57, 31, 185],
        [5, 57, 33, 185],
        [5, 57, 35, 185],
        [5, 57, 37, 185],
        [5, 57, 39, 185],
        [5, 57, 41, 185],
        [5, 57, 43, 185],
        [5, 57, 45, 185],
        [5, 57, 47, 185],
        [5, 57, 49, 185],
        [5, 57, 51, 185],
        [5, 57, 53, 185],
    ]

    static GetPatternList() {
        patterns := []
        for cfg in this.Configs {
            descOffsets := (cfg[4] == 101 || cfg[4] == 185) ? [0, 1] : [0]

            for descInc in descOffsets {
                pMap := Map()
                descBase := cfg[4] + descInc
                pName := "Pattern_" . cfg[1] . "_" . cfg[2] . "_" . cfg[3] . "_" . descBase
                isOffset := (descInc > 0)

                Loop 8 {
                    offset := A_Index - 1
                    focusClassNN := this.ClassPrefix . (cfg[1] + offset)
                    imgClassNN := "ComboBox" . (cfg[2] + (offset * 6))
                    imgGroupClassNN := "ComboBox" . (cfg[2] + 3 + (offset * 6))
                    srsClassNN := "AfxWnd140u" . (cfg[3] + (offset * 3))
                    descClassNN := "Button" . (descBase + (offset * 5))
                    pMap[focusClassNN] := {
                        img: imgClassNN,
                        imgGroup: imgGroupClassNN,
                        srs: srsClassNN,
                        desc: descClassNN,
                        type: pName
                    }
                }
                patterns.Push({name: pName, map: pMap, isOffset: isOffset})
            }
        }
        return patterns
    }

    static GetSeriesControlsForFocusClassNN(focusClassNN, hwnd, patterns := "") {
        match := this.GetSeriesMatchForFocusClassNN(focusClassNN, hwnd, patterns)
        return match ? match.candidate : false
    }

    static GetSrsControlForFocusClassNN(focusClassNN, hwnd, patterns := "") {
        controls := this.GetSeriesControlsForFocusClassNN(focusClassNN, hwnd, patterns)
        return controls ? controls.srs : ""
    }

    static GetSeriesMatchForFocusClassNN(focusClassNN, hwnd, patterns := "") {
        if (patterns == "")
            patterns := this.GetPatternList()

        for patternData in patterns {
            pMap := patternData.map
            if !pMap.Has(focusClassNN)
                continue

            candidate := pMap[focusClassNN]
            descText := ""
            if this.IsSeriesPatternMatch(candidate, focusClassNN, hwnd, &descText)
                return {candidate: candidate, name: patternData.name, desc: descText, isOffset: patternData.HasOwnProp("isOffset") ? patternData.isOffset : false}
        }
        return false
    }

    static IsSeriesPatternMatch(candidate, focusClassNN, hwnd, &descText := "") {
        try {
            srsText := ControlGetText(candidate.srs, hwnd)
            if !InStr(srsText, "VMTool")
                return false

            descText := ControlGetText(candidate.desc, hwnd)
            if !RegExMatch(descText, "i)^\([a-z0-9]+\)\s")
                return false
        } catch {
            return false
        }

        return this.IsSpatialControlMatch(candidate.srs, focusClassNN, hwnd)
            && this.IsSpatialControlMatch(candidate.img, focusClassNN, hwnd)
            && this.IsSpatialControlMatch(candidate.desc, focusClassNN, hwnd)
    }

    static IsSpatialControlMatch(childClassNN, focusClassNN, hwnd) {
        try {
            ControlGetPos(&childX, &childY, &childW, &childH, childClassNN, hwnd)
            ControlGetPos(&focusX, &focusY, &focusW, &focusH, focusClassNN, hwnd)
            if (childW <= 0 || childH <= 0 || focusW <= 0 || focusH <= 0)
                return false

            childBottom := childY + childH
            if (Abs(focusY - childBottom) > 50)
                return false

            childCenterX := childX + (childW / 2)
            return childCenterX >= (focusX - 10) && childCenterX <= (focusX + focusW + 10)
        }
        return false
    }

    static FocusSeriesUnderMouse(&hwnd := 0, &controlHwnd := 0) {
        MouseGetPos(&mouseX, &mouseY, &hwnd, &controlHwnd, 2)
        if !hwnd {
            return false
        }

        if !WinActive("ahk_id " hwnd) {
            try {
                if (WinGetProcessName("ahk_id " hwnd) = "G3PACS.exe") {
                    WinActivate("ahk_id " hwnd)
                }
            }
        }

        if !this.IsActiveSeriesUnderMouse(hwnd, controlHwnd)
            && !this.IsRecentLeftClick(mouseX, mouseY) {
            if this.TryControlClickSrsUnderMouse(hwnd, controlHwnd) {
                ; ControlClick the Srs row to focus without moving or dragging the image.
            } else {
                Click()
            }
            this.RecordLeftClick(mouseX, mouseY)
        }
        return true
    }

    static ClickUnderMouseAndSendKey(keyName, syncMpr := false) {
        this.FocusSeriesUnderMouse(&hwnd, &controlHwnd)
        Send("{" keyName "}")
        if (syncMpr) {
            Sleep(10)
            this.SendMprNavigationKey(hwnd, controlHwnd)
        }
    }

    static IsActiveSeriesUnderMouse(hwnd, controlHwnd) {
        if !hwnd || !controlHwnd {
            return false
        }

        try {
            controlClassNN := ControlGetClassNN(controlHwnd)
        } catch {
            return false
        }

        srsClassNN := this.GetSrsControlForFocusClassNN(controlClassNN, hwnd)
        return srsClassNN != "" && this.GetSrsControlFocusState(srsClassNN, hwnd) = "active"
    }

    static TryControlClickSrsUnderMouse(hwnd, controlHwnd) {
        if !hwnd || !controlHwnd {
            return false
        }

        try {
            controlClassNN := ControlGetClassNN(controlHwnd)
        } catch {
            return false
        }

        srsClassNN := this.GetSrsControlForFocusClassNN(controlClassNN, hwnd)
        if (srsClassNN = "") {
            return false
        }

        try {
            ControlClick(srsClassNN, "ahk_id " hwnd,, "Left", 1, "NA")
            return true
        }
        return false
    }

    static GetSrsControlFocusState(srsClassNN, hwnd) {
        try {
            ControlGetPos(&x, &y, &w, &h, srsClassNN, hwnd)
            if (w <= 8 || h <= 8) {
                return "unknown"
            }

            pt := Buffer(8, 0)
            NumPut("int", x, pt, 0)
            NumPut("int", y, pt, 4)
            DllCall("ClientToScreen", "ptr", hwnd, "ptr", pt)
            screenX := NumGet(pt, 0, "int")
            screenY := NumGet(pt, 4, "int")

            return this.GetSrsColorFocusState(screenX, screenY, w, h)
        }

        return "unknown"
    }

    static GetSrsColorFocusState(screenX, screenY, width, height) {
        sample := this.GetScreenPixelColor(
            screenX + 3,
            screenY + 3
        )
        if !sample.ok {
            return "unknown"
        }
        if this.IsColorNear(sample, 0x1B, 0x1D, 0x20, 18) {
            return "active"
        }
        if this.IsColorNear(sample, 0x4B, 0x4D, 0x5D, 18) {
            return "inactive"
        }
        return "unknown"
    }

    static IsColorNear(sample, red, green, blue, tolerance := 30) {
        return Abs(sample.red - red) + Abs(sample.green - green) + Abs(sample.blue - blue) <= tolerance
    }

    static GetScreenPixelColor(x, y) {
        hdc := DllCall("GetDC", "ptr", 0, "ptr")
        if !hdc {
            return {ok: false, hex: "GetDC failed", brightness: 255}
        }

        try {
            color := DllCall("GetPixel", "ptr", hdc, "int", x, "int", y, "uint")
            if (color = 0xFFFFFFFF) {
                return {ok: false, hex: "GetPixel failed", brightness: 255}
            }

            red := color & 0xFF
            green := (color >> 8) & 0xFF
            blue := (color >> 16) & 0xFF
            return {
                ok: true,
                hex: Format("#{1:02X}{2:02X}{3:02X}", red, green, blue),
                brightness: Round((red + green + blue) / 3, 1),
                red: red,
                green: green,
                blue: blue,
            }
        } finally {
            DllCall("ReleaseDC", "ptr", 0, "ptr", hdc)
        }
    }

    static RecordLeftClick(mouseX, mouseY) {
        state := this.GetLastLeftClick()
        state.time := A_TickCount
        state.x := mouseX
        state.y := mouseY
    }

    static IsRecentLeftClick(mouseX, mouseY) {
        state := this.GetLastLeftClick()
        return state.time
            && this.IsWithinDoubleClick(mouseX, mouseY, state.x, state.y, A_TickCount - state.time)
    }

    static GetLastLeftClick() {
        static state := {time: 0, x: 0, y: 0}
        return state
    }

    static IsWithinDoubleClick(mouseX, mouseY, lastClickX, lastClickY, elapsedMs) {
        static doubleClickTime := DllCall("GetDoubleClickTime", "UInt")
        static doubleClickWidth := DllCall("GetSystemMetrics", "Int", 36, "Int") ; SM_CXDOUBLECLK
        static doubleClickHeight := DllCall("GetSystemMetrics", "Int", 37, "Int") ; SM_CYDOUBLECLK

        return elapsedMs <= doubleClickTime
            && Abs(mouseX - lastClickX) <= doubleClickWidth
            && Abs(mouseY - lastClickY) <= doubleClickHeight
    }

    static SendMprNavigationKey(hwnd, controlHwnd) {
        if !hwnd || !controlHwnd {
            return
        }

        try {
            controlClassNN := ControlGetClassNN(controlHwnd)
        } catch {
            return
        }

        controls := this.GetSeriesControlsForFocusClassNN(controlClassNN, hwnd)
        if !controls {
            return
        }

        try {
            descVal := ControlGetText(controls.desc, hwnd)
        } catch {
            return
        }

        if RegExMatch(descVal, "i)MPR|MIP|COR|SAG")
            && !RegExMatch(descVal, "i)t1|t2|dwi|adc|dual|stir|fl2d|pd") {
            Send("y")
        }
    }
}
