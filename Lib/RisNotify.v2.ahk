#Requires AutoHotkey v2.0

class RisNotify {
    static _gui := 0
    static _queue := []
    static _slots := []
    static _slotItemIds := []
    static _sweepTimer := 0
    static _nextId := 0
    static _maxVisible := 5
    static _dedupeWindow := 800
    static _minWidth := 320
    static _width := 420
    static _maxWidth := 720
    static _paddingX := 24
    static _paddingY := 14
    static _slotGap := 8
    static _slotHeight := 36
    static _textPaddingY := 5
    static _lineSpacingScale := 1.2

    static Show(text, duration := 1500) {
        text := Trim(text)
        if (text == "")
            return

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
            createdAt: A_TickCount,
            duration: duration,
            expiresAt: expiresAt
        })

        if (this._queue.Length > this._maxVisible)
            this._queue.RemoveAt(1, this._queue.Length - this._maxVisible)

        this._EnsureGui()
        this._Render()
        this._UpdateSweepTimer()
    }

    static _EnsureGui() {
        if this._gui
            return this._gui

        g := Gui("+AlwaysOnTop -Caption +ToolWindow +E0x08000000")
        g.BackColor := "202020"
        g.SetFont("s12 cWhite bold", "Microsoft JhengHei UI")
        g.MarginX := this._paddingX
        g.MarginY := this._paddingY

        this._gui := g
        this._slots := []
        this._slotItemIds := []

        loop this._maxVisible {
            slot := g.Add("Text", Format("x{1} y{2} w{3} h{4} Center Hidden", this._paddingX, this._paddingY, this._width, this._slotHeight), "")
            slot.OnEvent("Click", ObjBindMethod(this, "_HandleSlotClick", A_Index))
            this._slots.Push([slot])
            this._slotItemIds.Push(0)
        }

        this._sweepTimer := ObjBindMethod(this, "_Sweep")
        return g
    }

    static _Render() {
        g := this._EnsureGui()
        visibleCount := this._queue.Length
        innerWidth := this._GetContentWidth()
        height := this._paddingY * 2
        y := this._paddingY

        for index, slotLines in this._slots {
            if (index <= visibleCount) {
                item := this._queue[index]
                displayLines := this._WrapLines(item.text, innerWidth)
                textHeight := this._MeasureLinesHeight(displayLines)
                slotHeight := Max(this._slotHeight, textHeight + this._textPaddingY * 2)
                textY := y + Floor((slotHeight - textHeight) / 2)
                lineHeight := this._MeasureLineHeight()
                lineAdvance := Ceil(lineHeight * this._lineSpacingScale)

                for lineIndex, lineText in displayLines {
                    lineSlot := this._GetLineSlot(index, lineIndex)
                    lineSlot.Text := lineText
                    lineSlot.Move(this._paddingX, textY + (lineIndex - 1) * lineAdvance, innerWidth, lineHeight)
                    lineSlot.Opt("-Hidden")
                }

                this._HideLineSlots(index, displayLines.Length + 1)
                this._slotItemIds[index] := item.id
                height := y + slotHeight + this._paddingY
                y += slotHeight + this._slotGap
            } else {
                this._HideLineSlots(index)
                this._slotItemIds[index] := 0
            }
        }

        if (visibleCount = 0) {
            try g.Hide()
            return
        }

        totalWidth := innerWidth + this._paddingX * 2
        position := this._GetWindowPosition(totalWidth, height)
        g.Show(Format("NoActivate x{1} y{2} w{3} h{4}", position.x, position.y, totalWidth, height))
        this._ApplyVisualStyle()
    }

    static _GetWindowPosition(width, height) {
        MonitorGetWorkArea(, &left, &top, &right, &bottom)
        workWidth := right - left
        workHeight := bottom - top
        x := left + Floor((workWidth - width) / 2)
        y := top + Floor((workHeight * 2 / 5) - (height / 2))

        return {
            x: Max(left, x),
            y: Max(top, y)
        }
    }

    static _GetContentWidth() {
        width := this._minWidth

        for item in this._queue
            width := Max(width, this._MeasureNaturalWidth(item.text))

        return Min(this._maxWidth, width)
    }

    static _MeasureNaturalWidth(text) {
        hdcState := this._BeginTextMeasure()
        if !hdcState.hdc
            return this._width

        width := 0
        sizeBuffer := Buffer(8, 0)

        for line in StrSplit(text, "`n", "`r") {
            lineText := (line = "") ? " " : line
            if DllCall("GetTextExtentPoint32", "Ptr", hdcState.hdc, "Str", lineText, "Int", StrLen(lineText), "Ptr", sizeBuffer.Ptr, "Int")
                width := Max(width, NumGet(sizeBuffer, 0, "Int"))
        }

        this._EndTextMeasure(hdcState)
        return Max(this._minWidth, width + 20)
    }

    static _MeasureLinesHeight(lines) {
        lineCount := lines.Length
        if (lineCount = 0)
            return this._MeasureLineHeight()

        lineHeight := this._MeasureLineHeight()
        lineAdvance := Ceil(lineHeight * this._lineSpacingScale)
        return lineHeight + (lineCount - 1) * lineAdvance
    }

    static _MeasureLineHeight() {
        hdcState := this._BeginTextMeasure()
        if !hdcState.hdc
            return this._slotHeight - this._textPaddingY * 2

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
            lineSlot := this._gui.Add("Text", Format("x{1} y{2} w{3} h{4} Center Hidden", this._paddingX, this._paddingY, this._width, this._slotHeight), "")
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
            return this._width

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

    static _ApplyVisualStyle() {
        if !this._gui
            return

        hwnd := this._gui.Hwnd
        WinGetPos(&x, &y, &w, &h, hwnd)

        try {
            WinSetRegion("0-0 w" w " h" h " r12-12", hwnd)

            style := DllCall("GetClassLongPtr", "Ptr", hwnd, "Int", -26, "Ptr")
            DllCall("SetClassLongPtr", "Ptr", hwnd, "Int", -26, "Ptr", style | 0x00020000)

            WinSetTransparent(235, hwnd)
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
                this._Render()
                this._UpdateSweepTimer()
                return
            }
        }
    }

    static _FindRecentId(text) {
        for index, item in this._queue {
            if (item.text == text && A_TickCount - item.createdAt <= this._dedupeWindow)
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
            this._queue[index] := item
            this._Render()
            this._UpdateSweepTimer()
            return
        }
    }

    static _PruneExpired() {
        loop {
            removed := false
            for index, item in this._queue {
                if (item.expiresAt > 0 && A_TickCount >= item.expiresAt) {
                    this._queue.RemoveAt(index)
                    removed := true
                    break
                }
            }
        } until !removed
    }

    static _Sweep() {
        this._PruneExpired()
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
}
