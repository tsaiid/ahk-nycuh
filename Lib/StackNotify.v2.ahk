#Requires AutoHotkey v2.0

class StackNotify {
    static Show(text, duration := unset) {
        startedAt := A_TickCount
        text := Trim(text)
        if (text = "")
            return

        if !IsSet(duration)
            duration := this.DefaultDuration

        ref := this._ResolveReferencePoint()
        dpi := this._GetDpiAtPoint(ref.x, ref.y)
        currentScale := dpi / 96

        if (this._gui && this._scale != currentScale) {
            this._gui.Destroy()
            this._gui := 0
        }

        this._scale := currentScale
        this._refX := ref.x
        this._refY := ref.y
        this._PruneExpired()

        existingId := this._FindRecentId(text)
        if (existingId) {
            this._RefreshDuration(existingId, duration)
            return
        }

        this._nextId += 1
        expiresAt := (duration > 0) ? (A_TickCount + duration) : 0
        this._queue.Push({
            id: this._nextId,
            text: text,
            displayText: this._BuildDisplayText(text, startedAt, 0),
            createdAt: A_TickCount,
            duration: duration,
            expiresAt: expiresAt
        })

        if (this._queue.Length > this.MaxVisible)
            this._queue.RemoveAt(1, this._queue.Length - this.MaxVisible)

        this._EnsureGui()
        this._Render(startedAt, true)
        this._UpdateSweepTimer()
    }

    static _BuildDisplayText(text, startedAt := 0, themeMs := 0) {
        if !(this.HasOwnProp("DebugBenchmark") && this.DebugBenchmark)
            return text

        elapsedMs := startedAt ? A_TickCount - startedAt : 0
        return Format("{1}`nbench: p{2} t{3} ms", text, elapsedMs, themeMs)
    }

    static _ResolveReferencePoint() {
        MouseGetPos(&x, &y)
        return {x: x, y: y}
    }

    static _EnsureGui() {
        if this._gui
            return this._gui

        g := Gui("+AlwaysOnTop -Caption +ToolWindow +E0x08000000 -DPIScale")
        g.BackColor := this._GetBackColor(this.Theme)
        g.MarginX := Round(this.PaddingX * this._scale)
        g.MarginY := Round(this.PaddingY * this._scale)
        g.SetFont(this._GetFontOptions(this.Theme), this.FontName)

        this._gui := g
        this._slots := []
        this._slotItemIds := []

        loop this.MaxVisible {
            px := Round(this.PaddingX * this._scale)
            py := Round(this.PaddingY * this._scale)
            w := Round(this.Width * this._scale)
            sh := Round(this.SlotHeight * this._scale)

            slot := g.Add("Text", Format("x{1} y{2} w{3} h{4} Center Hidden", px, py, w, sh), "")
            if this.ClickToDismiss
                slot.OnEvent("Click", ObjBindMethod(this, "_HandleSlotClick", A_Index))
            this._slots.Push([slot])
            this._slotItemIds.Push(0)
        }

        this._sweepTimer := ObjBindMethod(this, "_Sweep")
        this._ApplyTheme(this.Theme)
        return g
    }

    static _Render(startedAt := 0, updateTheme := false) {
        g := this._EnsureGui()
        visibleCount := this._queue.Length
        innerWidth := this._GetContentWidth()

        paddingY := Round(this.PaddingY * this._scale)
        paddingX := Round(this.PaddingX * this._scale)
        slotGap := Round(this.SlotGap * this._scale)
        minSlotHeight := Round(this.SlotHeight * this._scale)
        textPaddingY := Round(this.TextPaddingY * this._scale)

        height := paddingY * 2
        y := paddingY

        for index, slotLines in this._slots {
            if (index <= visibleCount) {
                item := this._queue[index]
                displayLines := this._WrapLines(item.displayText, innerWidth)
                textHeight := this._MeasureLinesHeight(displayLines)
                slotHeight := Max(minSlotHeight, textHeight + textPaddingY * 2)
                textY := y + Floor((slotHeight - textHeight) / 2)
                lineHeight := this._MeasureLineHeight()
                lineAdvance := Ceil(lineHeight * this.LineSpacingScale)

                for lineIndex, lineText in displayLines {
                    lineSlot := this._GetLineSlot(index, lineIndex)
                    lineSlot.Text := lineText
                    lineSlot.Move(paddingX, textY + (lineIndex - 1) * lineAdvance, innerWidth, lineHeight)
                    lineSlot.Opt("-Hidden")
                }

                this._HideLineSlots(index, displayLines.Length + 1)
                this._slotItemIds[index] := item.id
                height := y + slotHeight + paddingY
                y += slotHeight + slotGap
            } else {
                this._HideLineSlots(index)
                this._slotItemIds[index] := 0
            }
        }

        if (visibleCount = 0) {
            try g.Hide()
            this._visible := false
            return
        }

        totalWidth := innerWidth + paddingX * 2
        position := this._GetWindowPosition(totalWidth, height, this._refX, this._refY)

        themeMs := 0
        theme := this._theme
        if (updateTheme && !this._visible) {
            themeStart := A_TickCount
            theme := this._ChooseThemeForRegion(position.x, position.y, totalWidth, height)
            themeMs := A_TickCount - themeStart
        }
        if (this.HasOwnProp("DebugBenchmark") && this.DebugBenchmark && visibleCount > 0) {
            item := this._queue[visibleCount]
            item.displayText := this._BuildDisplayText(item.text, startedAt, themeMs)
            this._queue[visibleCount] := item
        }
        this._ApplyTheme(theme)

        g.Show(Format("NoActivate x{1} y{2} w{3} h{4}", position.x, position.y, totalWidth, height))
        this._visible := true
        this._ApplyVisualStyle()
    }

    static _GetWindowPosition(width, height, refX, refY) {
        workArea := this._GetMonitorWorkAreaAtPoint(refX, refY)
        workWidth := workArea.right - workArea.left
        workHeight := workArea.bottom - workArea.top
        x := workArea.left + Floor((workWidth - width) / 2)
        y := workArea.top + Floor((workHeight * this.VerticalPositionRatio) - (height / 2))

        return {
            x: Max(workArea.left, x),
            y: Max(workArea.top, y)
        }
    }

    static _GetContentWidth() {
        minW := Round(this.MinWidth * this._scale)
        maxW := Round(this.MaxWidth * this._scale)
        width := minW

        for item in this._queue
            width := Max(width, this._MeasureNaturalWidth(item.displayText))

        return Min(maxW, width)
    }

    static _MeasureNaturalWidth(text) {
        hdcState := this._BeginTextMeasure()
        if !hdcState.hdc
            return Round(this.Width * this._scale)

        width := 0
        sizeBuffer := Buffer(8, 0)

        for line in StrSplit(text, "`n", "`r") {
            lineText := (line = "") ? " " : line
            if DllCall("GetTextExtentPoint32", "Ptr", hdcState.hdc, "Str", lineText, "Int", StrLen(lineText), "Ptr", sizeBuffer.Ptr, "Int")
                width := Max(width, NumGet(sizeBuffer, 0, "Int"))
        }

        this._EndTextMeasure(hdcState)
        return Max(Round(this.MinWidth * this._scale), width + Round(20 * this._scale))
    }

    static _MeasureLinesHeight(lines) {
        lineCount := lines.Length
        if (lineCount = 0)
            return this._MeasureLineHeight()

        lineHeight := this._MeasureLineHeight()
        lineAdvance := Ceil(lineHeight * this.LineSpacingScale)
        return lineHeight + (lineCount - 1) * lineAdvance
    }

    static _MeasureLineHeight() {
        hdcState := this._BeginTextMeasure()
        if !hdcState.hdc
            return Round(this.SlotHeight * this._scale) - Round(this.TextPaddingY * this._scale) * 2

        lineHeight := 0
        sizeBuffer := Buffer(8, 0)
        sampleText := "Ag"
        if DllCall("GetTextExtentPoint32", "Ptr", hdcState.hdc, "Str", sampleText, "Int", StrLen(sampleText), "Ptr", sizeBuffer.Ptr, "Int")
            lineHeight := NumGet(sizeBuffer, 4, "Int")
        this._EndTextMeasure(hdcState)
        return Max(1, lineHeight)
    }

    static _WrapLines(text, maxWidth) {
        wrappedLines := []
        for rawLine in StrSplit(text, "`n", "`r")
            this._AppendWrappedLine(wrappedLines, rawLine, maxWidth)

        return wrappedLines.Length ? wrappedLines : [text]
    }

    static _GetLineSlot(slotIndex, lineIndex) {
        while (this._slots[slotIndex].Length < lineIndex) {
            px := Round(this.PaddingX * this._scale)
            py := Round(this.PaddingY * this._scale)
            w := Round(this.Width * this._scale)
            sh := Round(this.SlotHeight * this._scale)

            lineSlot := this._gui.Add("Text", Format("x{1} y{2} w{3} h{4} Center Hidden", px, py, w, sh), "")
            if this.ClickToDismiss
                lineSlot.OnEvent("Click", ObjBindMethod(this, "_HandleSlotClick", slotIndex))
            this._slots[slotIndex].Push(lineSlot)
        }

        return this._slots[slotIndex][lineIndex]
    }

    static _HideLineSlots(slotIndex, startIndex := 1) {
        for lineIndex, lineSlot in this._slots[slotIndex] {
            if (lineIndex >= startIndex) {
                lineSlot.Text := ""
                lineSlot.Opt("Hidden")
            }
        }
    }

    static _AppendWrappedLine(lines, text, maxWidth) {
        if (text = "") {
            lines.Push("")
            return
        }

        current := ""
        tokenPattern := "(\s+|[^\s]+)"
        startPos := 1
        tokenFound := false

        while RegExMatch(text, tokenPattern, &match, startPos) {
            tokenFound := true
            token := match[1]
            candidate := current . token

            if (current = "" || this._MeasureLineWidth(candidate) <= maxWidth) {
                current := candidate
            } else {
                trimmed := Trim(current, " ")
                if (trimmed != "")
                    lines.Push(trimmed)
                this._AppendWrappedLongToken(lines, token, maxWidth, &current)
            }

            startPos := match.Pos + match.Len
        }

        if !tokenFound {
            this._AppendWrappedLongToken(lines, text, maxWidth, &current)
            tokenFound := true
        }

        finalLine := Trim(current, " ")
        if (tokenFound && finalLine != "")
            lines.Push(finalLine)
    }

    static _AppendWrappedLongToken(lines, token, maxWidth, &current) {
        token := Trim(token, " ")
        if (token = "") {
            current := ""
            return
        }

        if (this._MeasureLineWidth(token) <= maxWidth) {
            current := token
            return
        }

        current := ""
        chunk := ""
        loop parse token {
            char := A_LoopField
            candidate := chunk . char
            if (chunk = "" || this._MeasureLineWidth(candidate) <= maxWidth) {
                chunk := candidate
            } else {
                lines.Push(chunk)
                chunk := char
            }
        }
        current := chunk
    }

    static _MeasureLineWidth(text) {
        if (text = "")
            return 0

        hdcState := this._BeginTextMeasure()
        if !hdcState.hdc
            return Round(this.Width * this._scale)

        sizeBuffer := Buffer(8, 0)
        width := 0
        if DllCall("GetTextExtentPoint32", "Ptr", hdcState.hdc, "Str", text, "Int", StrLen(text), "Ptr", sizeBuffer.Ptr, "Int")
            width := NumGet(sizeBuffer, 0, "Int")

        this._EndTextMeasure(hdcState)
        return width
    }

    static _BeginTextMeasure() {
        this._EnsureGui()

        if !this._slots.Length
            return {hdc: 0, hwnd: 0, oldFont: 0}

        hwnd := this._slots[1][1].Hwnd
        hdc := DllCall("GetDC", "Ptr", hwnd, "Ptr")
        if !hdc
            return {hdc: 0, hwnd: hwnd, oldFont: 0}

        hFont := SendMessage(0x0031, 0, 0, hwnd)
        oldFont := hFont ? DllCall("SelectObject", "Ptr", hdc, "Ptr", hFont, "Ptr") : 0
        return {hdc: hdc, hwnd: hwnd, oldFont: oldFont}
    }

    static _EndTextMeasure(hdcState) {
        if !hdcState.hdc
            return

        if hdcState.oldFont
            DllCall("SelectObject", "Ptr", hdcState.hdc, "Ptr", hdcState.oldFont, "Ptr")

        DllCall("ReleaseDC", "Ptr", hdcState.hwnd, "Ptr", hdcState.hdc)
    }

    static _ChooseThemeForRegion(x, y, width, height) {
        if !this.AutoTheme
            return this.Theme

        brightness := this._GetAverageScreenBrightness(x, y, width, height)
        if (brightness = "")
            return this._theme

        return (brightness >= 128) ? "dark" : "light"
    }

    static _ApplyTheme(theme) {
        if !this._gui
            return

        this._theme := theme
        this._gui.BackColor := this._GetBackColor(theme)
        for slotLines in this._slots {
            for lineControl in slotLines
                lineControl.SetFont(this._GetFontOptions(theme), this.FontName)
        }
    }

    static _GetBackColor(theme) {
        return (theme = "dark") ? this.DarkBackColor : this.LightBackColor
    }

    static _GetFontOptions(theme) {
        fontSize := Round(this.FontSize * this._scale)
        fontColor := (theme = "dark") ? this.DarkTextColor : this.LightTextColor
        return "s" fontSize " c" fontColor " bold"
    }

    static _GetAverageScreenBrightness(x, y, width, height) {
        columns := 11
        rows := 7
        hdcScreen := DllCall("GetDC", "Ptr", 0, "Ptr")
        if !hdcScreen
            return ""

        hdcMem := DllCall("CreateCompatibleDC", "Ptr", hdcScreen, "Ptr")
        if !hdcMem {
            DllCall("ReleaseDC", "Ptr", 0, "Ptr", hdcScreen)
            return ""
        }

        hbm := DllCall("CreateCompatibleBitmap", "Ptr", hdcScreen, "Int", columns, "Int", rows, "Ptr")
        if !hbm {
            DllCall("DeleteDC", "Ptr", hdcMem)
            DllCall("ReleaseDC", "Ptr", 0, "Ptr", hdcScreen)
            return ""
        }

        oldBitmap := DllCall("SelectObject", "Ptr", hdcMem, "Ptr", hbm, "Ptr")
        DllCall("SetStretchBltMode", "Ptr", hdcMem, "Int", 3)
        copied := DllCall(
            "StretchBlt",
            "Ptr", hdcMem, "Int", 0, "Int", 0, "Int", columns, "Int", rows,
            "Ptr", hdcScreen, "Int", x, "Int", y, "Int", width, "Int", height,
            "UInt", 0x00CC0020,
            "Int"
        )

        brightness := ""
        if copied {
            bmi := Buffer(40, 0)
            NumPut("UInt", 40, bmi, 0)
            NumPut("Int", columns, bmi, 4)
            NumPut("Int", -rows, bmi, 8)
            NumPut("UShort", 1, bmi, 12)
            NumPut("UShort", 32, bmi, 14)
            pixels := Buffer(columns * rows * 4, 0)

            if DllCall("GetDIBits", "Ptr", hdcMem, "Ptr", hbm, "UInt", 0, "UInt", rows, "Ptr", pixels.Ptr, "Ptr", bmi.Ptr, "UInt", 0, "Int") {
                total := 0
                count := columns * rows

                loop count {
                    offset := (A_Index - 1) * 4
                    blue := NumGet(pixels, offset, "UChar")
                    green := NumGet(pixels, offset + 1, "UChar")
                    red := NumGet(pixels, offset + 2, "UChar")
                    total += (red * 299 + green * 587 + blue * 114) / 1000
                }

                brightness := total / count
            }
        }

        if oldBitmap
            DllCall("SelectObject", "Ptr", hdcMem, "Ptr", oldBitmap, "Ptr")
        DllCall("DeleteObject", "Ptr", hbm)
        DllCall("DeleteDC", "Ptr", hdcMem)
        DllCall("ReleaseDC", "Ptr", 0, "Ptr", hdcScreen)
        return brightness
    }

    static _ApplyVisualStyle() {
        if !this._gui
            return

        hwnd := this._gui.Hwnd
        WinGetPos(, , &w, &h, hwnd)

        try {
            r := Round(this.CornerRadius * this._scale)
            WinSetRegion("0-0 w" w " h" h " r" r "-" r, hwnd)

            style := DllCall("GetClassLongPtr", "Ptr", hwnd, "Int", -26, "Ptr")
            DllCall("SetClassLongPtr", "Ptr", hwnd, "Int", -26, "Ptr", style | 0x00020000)

            WinSetTransparent(this.Transparent, hwnd)
        }
    }

    static _HandleSlotClick(slotIndex, ctrl, *) {
        if (slotIndex < 1 || slotIndex > this._slotItemIds.Length)
            return

        itemId := this._slotItemIds[slotIndex]
        if !itemId
            return

        this._RemoveById(itemId)
    }

    static _RemoveById(itemId) {
        for index, item in this._queue {
            if (item.id = itemId) {
                this._queue.RemoveAt(index)
                this._Render(0, true)
                this._UpdateSweepTimer()
                return
            }
        }
    }

    static _FindRecentId(text) {
        if (this.DedupeWindow <= 0)
            return 0

        for index, item in this._queue {
            if (item.text == text && A_TickCount - item.createdAt <= this.DedupeWindow)
                return item.id
        }
        return 0
    }

    static _RefreshDuration(itemId, duration) {
        for index, item in this._queue {
            if (item.id != itemId)
                continue

            item.createdAt := A_TickCount
            item.duration := duration
            item.expiresAt := (duration > 0) ? (A_TickCount + duration) : 0
            item.displayText := this._BuildDisplayText(item.text)
            this._queue[index] := item
            this._Render(0, true)
            this._UpdateSweepTimer()
            return
        }
    }

    static _PruneExpired() {
        removedAny := false
        loop {
            removed := false
            for index, item in this._queue {
                if (item.expiresAt > 0 && A_TickCount >= item.expiresAt) {
                    this._queue.RemoveAt(index)
                    removed := true
                    removedAny := true
                    break
                }
            }
        } until !removed
        return removedAny
    }

    static _Sweep() {
        if this._PruneExpired()
            this._Render()
        this._UpdateSweepTimer()
    }

    static _UpdateSweepTimer() {
        if !this._sweepTimer
            return

        SetTimer(this._sweepTimer, 0)

        for item in this._queue {
            if (item.expiresAt > 0) {
                SetTimer(this._sweepTimer, -150)
                return
            }
        }
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
                DllCall("Shcore\GetDpiForMonitor", "Ptr", hMonitor, "Int", 0, "UInt*", &dpiX, "UInt*", &dpiY)
                if (dpiX > 0)
                    return dpiX
            }
        }
        return 96
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
}
