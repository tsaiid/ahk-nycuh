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
                patterns.Push({name: pName, map: pMap, isOffset: isOffset, hits: 0})
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

    static GetDpiScale(hwnd := 0) {
        dpi := 0
        if (hwnd) {
            try {
                dpi := DllCall("User32\GetDpiForWindow", "Ptr", hwnd, "UInt")
            }
        }
        if (!dpi) {
            try {
                dpi := A_ScreenDPI
            }
        }
        return (dpi > 0) ? (dpi / 96) : 1.0
    }

    static HasConfig(cfg) {
        if !(cfg is Array) || (cfg.Length < 4) {
            return false
        }
        for c in this.Configs {
            descMatches := (c[4] == cfg[4]) || ((c[4] == 101 && cfg[4] == 102) || (c[4] == 185 && cfg[4] == 186))
            if (c[1] == cfg[1] && c[2] == cfg[2] && c[3] == cfg[3] && descMatches) {
                return true
            }
        }
        return false
    }

    static IsSpatialControlMatch(childClassNN, focusClassNN, hwnd, &diag := unset) {
        try {
            ControlGetPos(&childX, &childY, &childW, &childH, childClassNN, hwnd)
            ControlGetPos(&focusX, &focusY, &focusW, &focusH, focusClassNN, hwnd)
            if (childW <= 0 || childH <= 0 || focusW <= 0 || focusH <= 0) {
                if (IsSet(diag)) {
                    diag := Format("尺寸無效: child={1}x{2}, focus={3}x{4}", childW, childH, focusW, focusH)
                }
                return false
            }

            dpiScale := this.GetDpiScale(hwnd)
            vTolerance := Round(50 * dpiScale)
            hMargin := Round(10 * dpiScale)

            childBottom := childY + childH
            vDiff := Abs(focusY - childBottom)
            if (vDiff > vTolerance) {
                if (IsSet(diag)) {
                    diag := Format("垂直差距 {1}px 超過公差 {2}px (DPI {3}%)", vDiff, vTolerance, Round(dpiScale * 100))
                }
                return false
            }

            childCenterX := childX + (childW / 2)
            if (childCenterX < (focusX - hMargin) || childCenterX > (focusX + focusW + hMargin)) {
                if (IsSet(diag)) {
                    diag := Format("水平中心 {1}px 超出範圍 [{2}, {3}] (DPI {4}%)", Round(childCenterX), focusX - hMargin, focusX + focusW + hMargin, Round(dpiScale * 100))
                }
                return false
            }

            if (IsSet(diag)) {
                diag := Format("幾何相符: 垂直差距 {1}px <= {2}px, 水平中心 {3}px (DPI {4}%)", vDiff, vTolerance, Round(childCenterX), Round(dpiScale * 100))
            }
            return true
        } catch Error as e {
            if (IsSet(diag)) {
                diag := "取得控制項座標失敗: " . e.Message
            }
            return false
        }
    }

    static GetMatchingDiagnostics(focusClassNN, hwnd, patterns := "") {
        if (patterns == "") {
            patterns := this.GetPatternList()
        }

        matchedCandidates := []
        for patternData in patterns {
            pMap := patternData.map
            if (!pMap.Has(focusClassNN)) {
                continue
            }
            matchedCandidates.Push({name: patternData.name, candidate: pMap[focusClassNN], isOffset: patternData.HasOwnProp("isOffset") ? patternData.isOffset : false})
        }

        if (matchedCandidates.Length == 0) {
            return "  ⚠️ 控制項 " . focusClassNN . " 未包含在任何已配置 Pattern 的焦點清單中。`n"
        }

        dpiScale := this.GetDpiScale(hwnd)
        vTolerance := Round(50 * dpiScale)
        hMargin := Round(10 * dpiScale)

        report := "  🔍 焦點控制項 " . focusClassNN . " 在配置中找到 " . matchedCandidates.Length . " 組候選規則，詳細比對結果：`n"
        report .= "  (螢幕 DPI 縮放: " . Round(dpiScale * 100) . "% | 垂直門檻: " . vTolerance . "px | 水平容差: " . hMargin . "px)`n"

        for item in matchedCandidates {
            cand := item.candidate
            report .= "  --------------------------------------------------`n"
            report .= "  📌 候選模式: " . item.name . (item.isOffset ? " (+1 容錯)" : "") . "`n"

            ; 1. Srs 文字
            srsText := ""
            srsOk := false
            try {
                srsText := ControlGetText(cand.srs, hwnd)
                srsOk := InStr(srsText, "VMTool") > 0
            } catch Error as e {
                srsText := "讀取失敗: " . e.Message
            }
            report .= "    • Srs 文字 (" . cand.srs . "): [" . Trim(srsText) . "] -> " . (srsOk ? "✅ 符合 (含 VMTool)" : "❌ 不符 (未含 VMTool)") . "`n"

            ; 2. Desc 文字
            descText := ""
            descOk := false
            try {
                descText := ControlGetText(cand.desc, hwnd)
                descOk := RegExMatch(descText, "i)^\([a-z0-9]+\)\s") > 0
            } catch Error as e {
                descText := "讀取失敗: " . e.Message
            }
            report .= "    • Desc 文字 (" . cand.desc . "): [" . Trim(descText) . "] -> " . (descOk ? "✅ 符合序號格式" : "❌ 不符序號格式") . "`n"

            ; 3. Srs 空間幾何
            diagSrs := ""
            srsSpatial := this.IsSpatialControlMatch(cand.srs, focusClassNN, hwnd, &diagSrs)
            report .= "    • Srs 幾何 (" . cand.srs . "): " . (srsSpatial ? "✅ " : "❌ ") . diagSrs . "`n"

            ; 4. Img 空間幾何
            diagImg := ""
            imgSpatial := this.IsSpatialControlMatch(cand.img, focusClassNN, hwnd, &diagImg)
            report .= "    • Img 幾何 (" . cand.img . "): " . (imgSpatial ? "✅ " : "❌ ") . diagImg . "`n"

            ; 5. Desc 空間幾何
            diagDesc := ""
            descSpatial := this.IsSpatialControlMatch(cand.desc, focusClassNN, hwnd, &diagDesc)
            report .= "    • Desc 幾何 (" . cand.desc . "): " . (descSpatial ? "✅ " : "❌ ") . diagDesc . "`n"
        }
        return report
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

    ; ==========================================================================
    ; PACS Toolbar 同步與輔助按鈕巡覽
    ; ==========================================================================
    static SyncButtons := ["Button1", "Button85", "Button90", "Button102"]
    static SyncTexts := [" Auto sync", "自動同步"]

    static DiffExamSyncButtons := ["Button2", "Button86", "Button91", "Button103"]
    static DiffExamSyncTexts := [" Sync with other exams", "不同檢查同步 "]

    static ScoutLineButtons := ["Button4", "Button78"]
    static ScoutLineTexts := [" Scout lines"]

    static IsPacsImageViewActive(winTitle := "INFINITT PACS") {
        if !WinActive(winTitle) {
            return false
        }

        try {
            focusedHwnd := ControlGetFocus("A")
            focusedClassNN := ControlGetClassNN(focusedHwnd)
            return SubStr(focusedClassNN, 1, 3) == "Afx"
        }
        return false
    }

    static IsTextMatch(text, targetText) {
        if (targetText is Array) {
            for t in targetText {
                if (text == t || Trim(text) == Trim(t)) {
                    return true
                }
            }
            return false
        }
        return (text == targetText || Trim(text) == Trim(targetText))
    }

    static ToggleSync(winTitle := "INFINITT PACS") {
        return this.ClickButtonIfTextMatches(this.SyncButtons, this.SyncTexts, winTitle)
    }

    static ToggleDiffExamSync(winTitle := "INFINITT PACS") {
        return this.ClickButtonIfTextMatches(this.DiffExamSyncButtons, this.DiffExamSyncTexts, winTitle)
    }

    static ToggleScoutLines(winTitle := "INFINITT PACS") {
        return this.ClickButtonIfTextMatches(this.ScoutLineButtons, this.ScoutLineTexts, winTitle)
    }

    static ClickButtonIfTextMatches(btnList, targetText, winTitle := "INFINITT PACS") {
        if !this.IsPacsImageViewActive(winTitle) {
            return ""
        }

        for btn in btnList {
            try {
                curText := ControlGetText(btn, "A")
                if this.IsTextMatch(curText, targetText) {
                    ControlClick(btn, "A")
                    return btn
                }
            }
        }
        return ""
    }
}
