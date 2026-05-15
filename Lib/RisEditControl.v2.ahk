#Requires AutoHotkey v2.0

/**
 * Win32 Edit control 的低階 selection / replace helper。
 */
class RisEditControl {
    static MSG := {
        GETSEL:      0x00B0,
        SETSEL:      0x00B1,
        LINESCROLL:  0x00B6,
        GETLINECOUNT: 0x00BA,
        REPLACESEL:  0x00C2,
        LINEFROMCHAR: 0x00C9,
        SCROLLCARET: 0x00B7,
        GETFIRSTVISIBLELINE: 0x00CE,
        CUT: 0x0300,
        COPY: 0x0301,
        CLEAR: 0x0303
    }

    static SetSel(hCtrl, startPos, endPos) {
        SendMessage(this.MSG.SETSEL, startPos, endPos, hCtrl)
    }

    static ReplaceSel(hCtrl, text) {
        SendMessage(this.MSG.REPLACESEL, 1, StrPtr(text), hCtrl)
    }

    static ScrollCaret(hCtrl) {
        SendMessage(this.MSG.SCROLLCARET, 0, 0, hCtrl)
    }

    static GetFirstVisibleLine(hCtrl) {
        return SendMessage(this.MSG.GETFIRSTVISIBLELINE, 0, 0, hCtrl)
    }

    static LineScroll(hCtrl, lineCount, columnCount := 0) {
        SendMessage(this.MSG.LINESCROLL, columnCount, lineCount, hCtrl)
    }

    static LineFromChar(hCtrl, charPos := -1) {
        return SendMessage(this.MSG.LINEFROMCHAR, charPos, 0, hCtrl)
    }

    static GetLineCount(hCtrl) {
        return SendMessage(this.MSG.GETLINECOUNT, 0, 0, hCtrl)
    }

    static ReplaceSelectionAndScroll(hCtrl, text) {
        this.ReplaceSel(hCtrl, text)
        this.ScrollCaret(hCtrl)
    }

    static ReplaceSelectionPreserveFirstVisibleLine(hCtrl, text) {
        firstVisibleLineBefore := this.GetFirstVisibleLine(hCtrl)
        this.ReplaceSel(hCtrl, text)
        firstVisibleLineAfter := this.GetFirstVisibleLine(hCtrl)

        linesToScroll := firstVisibleLineBefore - firstVisibleLineAfter
        if (linesToScroll != 0) {
            this.LineScroll(hCtrl, linesToScroll)
        }
    }

    static GetSel(hCtrl) {
        startBuf := Buffer(4, 0), endBuf := Buffer(4, 0)
        SendMessage(this.MSG.GETSEL, startBuf.Ptr, endBuf.Ptr, hCtrl)
        return {Start: NumGet(startBuf, "UInt"), End: NumGet(endBuf, "UInt")}
    }

    static GetSelectedText(hCtrl) {
        sel := this.GetSel(hCtrl)
        if (sel.End <= sel.Start) {
            return ""
        }

        fullText := ControlGetText(hCtrl)
        return SubStr(fullText, sel.Start + 1, sel.End - sel.Start)
    }

    static CountNonEmptyLines(hCtrl) {
        try {
            text := ControlGetText(hCtrl)
        } catch {
            return 0
        }

        if (text == "") {
            return 0
        }

        lines := StrSplit(text, "`n", "`r")
        count := 0

        for line in lines {
            if (Trim(line, " `t") != "") {
                count++
            }
        }

        return count
    }

    static GetLogicalLineBoundaries(hCtrl, specificPos := -1) {
        try {
            fullText := ControlGetText(hCtrl)
        } catch {
            return {Start: 0, ContentEnd: 0, FullEnd: 0}
        }

        if (specificPos != -1) {
            caretPos := specificPos
        } else {
            caretPos := this.GetSel(hCtrl).Start
        }

        prevLineBreak := InStr(fullText, "`n", , caretPos + 1, -1)
        lineStart := (prevLineBreak == 0) ? 0 : prevLineBreak

        rPos := InStr(fullText, "`r", , lineStart + 1)
        nPos := InStr(fullText, "`n", , lineStart + 1)

        if (rPos == 0 && nPos == 0) {
            nextLineBreak := 0
        } else if (rPos == 0) {
            nextLineBreak := nPos
        } else if (nPos == 0) {
            nextLineBreak := rPos
        } else {
            nextLineBreak := Min(rPos, nPos)
        }

        if (nextLineBreak == 0) {
            contentEnd := StrLen(fullText)
            fullEnd := contentEnd
        } else {
            contentEnd := nextLineBreak - 1

            char := SubStr(fullText, nextLineBreak, 1)
            if (char == "`r" && SubStr(fullText, nextLineBreak + 1, 1) == "`n") {
                fullEnd := nextLineBreak + 1
            } else {
                fullEnd := nextLineBreak
            }
        }

        return {Start: lineStart, ContentEnd: contentEnd, FullEnd: fullEnd}
    }

    static SelectLine(hCtrl) {
        bounds := this.GetLogicalLineBoundaries(hCtrl)
        if (bounds.FullEnd > bounds.Start) {
            this.SetSel(hCtrl, bounds.Start, bounds.FullEnd)
        }
    }

    static SelectLineForRemoval(hCtrl) {
        bounds := this.GetLogicalLineBoundaries(hCtrl)
        isLastLine := (bounds.FullEnd == bounds.ContentEnd)

        if (!isLastLine) {
            this.SetSel(hCtrl, bounds.Start, bounds.FullEnd)
        } else if (bounds.Start == 0) {
            this.SetSel(hCtrl, 0, bounds.FullEnd)
        } else {
            this.SetSel(hCtrl, bounds.Start - 2, bounds.FullEnd)
        }
    }

    static InsertNewLine(hCtrl, mode := "Below") {
        lineInfo := this._GetCurrentLineInfo(hCtrl)
        bounds := lineInfo ? lineInfo.Bounds : this.GetLogicalLineBoundaries(hCtrl)
        smartPrefix := lineInfo ? this._GetSmartListInsertPrefix(lineInfo, mode) : ""

        if (mode = "Above") {
            this.SetSel(hCtrl, bounds.Start, bounds.Start)
            this.ReplaceSel(hCtrl, smartPrefix . "`r`n")
            this.SetSel(hCtrl, bounds.Start + StrLen(smartPrefix), bounds.Start + StrLen(smartPrefix))
        } else {
            this.SetSel(hCtrl, bounds.ContentEnd, bounds.ContentEnd)
            this.ReplaceSel(hCtrl, "`r`n" . smartPrefix)
        }

        this.ScrollCaret(hCtrl)
    }

    static SmartListEnter(hCtrl) {
        lineInfo := this._GetCurrentLineInfo(hCtrl)
        if (!lineInfo || lineInfo.Sel.Start != lineInfo.Sel.End) {
            return false
        }

        if (lineInfo.Sel.Start != lineInfo.Bounds.ContentEnd) {
            return false
        }

        if (this._CanRemoveEmptySmartPrefix(lineInfo)) {
            this._RemoveLineContent(hCtrl, lineInfo)
            return true
        }

        nextPrefix := this._GetSmartListNextPrefix(lineInfo)
        if (nextPrefix == "") {
            return false
        }

        this.ReplaceSel(hCtrl, "`r`n" . nextPrefix)
        this.ScrollCaret(hCtrl)
        return true
    }

    static ShouldSmartListEnter(hCtrl) {
        lineInfo := this._GetCurrentLineInfo(hCtrl)
        if (!lineInfo || lineInfo.Sel.Start != lineInfo.Sel.End) {
            return false
        }

        if (lineInfo.Sel.Start != lineInfo.Bounds.ContentEnd) {
            return false
        }

        return this._CanRemoveEmptySmartPrefix(lineInfo) || this._GetSmartListNextPrefix(lineInfo) != ""
    }

    static SmartListBackspace(hCtrl) {
        lineInfo := this._GetCurrentLineInfo(hCtrl)
        if (!lineInfo || lineInfo.Sel.Start != lineInfo.Sel.End) {
            return false
        }

        if (lineInfo.Sel.Start != lineInfo.Bounds.ContentEnd) {
            return false
        }

        if !this._CanRemoveEmptySmartPrefix(lineInfo) {
            return false
        }

        this._RemoveLineContent(hCtrl, lineInfo)
        return true
    }

    static ShouldSmartListBackspace(hCtrl) {
        lineInfo := this._GetCurrentLineInfo(hCtrl)
        if (!lineInfo || lineInfo.Sel.Start != lineInfo.Sel.End) {
            return false
        }

        if (lineInfo.Sel.Start != lineInfo.Bounds.ContentEnd) {
            return false
        }

        return this._CanRemoveEmptySmartPrefix(lineInfo)
    }

    static _GetCurrentLineInfo(hCtrl) {
        try {
            fullText := ControlGetText(hCtrl)
            sel := this.GetSel(hCtrl)
            bounds := this.GetLogicalLineBoundaries(hCtrl, sel.Start)
        } catch {
            return false
        }

        lineText := SubStr(fullText, bounds.Start + 1, bounds.ContentEnd - bounds.Start)
        return {Sel: sel, Bounds: bounds, Text: lineText}
    }

    static _GetSmartListNextPrefix(lineInfo) {
        parsedLine := this._ParseSmartLine(lineInfo.Text)
        if (!parsedLine) {
            return ""
        }

        if (Trim(parsedLine.Body, " `t") == "") {
            return ""
        }

        nextMarker := ""
        if (parsedLine.HasMarker) {
            nextMarker := this._GetNextListMarker(parsedLine.Marker)
        }

        nextSpineSegment := this._GetNextSpineSegmentPrefix(parsedLine.Body)
        if (nextSpineSegment != "") {
            return parsedLine.Indent . nextMarker . nextSpineSegment
        }

        if (parsedLine.HasMarker) {
            return parsedLine.Indent . nextMarker
        }

        return ""
    }

    static _GetSmartListInsertPrefix(lineInfo, mode) {
        if (!lineInfo || lineInfo.Sel.Start != lineInfo.Sel.End) {
            return ""
        }

        return (mode = "Above")
            ? this._GetSmartListPreviousPrefix(lineInfo)
            : this._GetSmartListNextInsertPrefix(lineInfo)
    }

    static _GetSmartListNextInsertPrefix(lineInfo) {
        parsedLine := this._ParseSmartLine(lineInfo.Text)
        if (!parsedLine) {
            return ""
        }

        nextMarker := ""
        if (parsedLine.HasMarker) {
            nextMarker := this._GetNextListMarker(parsedLine.Marker)
        }

        nextSpineSegment := this._GetNextSpineSegmentPrefix(parsedLine.Body, false)
        if (nextSpineSegment != "") {
            return parsedLine.Indent . nextMarker . nextSpineSegment
        }

        if (parsedLine.HasMarker) {
            return parsedLine.Indent . nextMarker
        }

        return ""
    }

    static _GetSmartListPreviousPrefix(lineInfo) {
        parsedLine := this._ParseSmartLine(lineInfo.Text)
        if (!parsedLine) {
            return ""
        }

        previousSpineSegment := this._GetPreviousSpineSegmentPrefix(parsedLine.Body)
        if (previousSpineSegment != "") {
            return parsedLine.Indent . (parsedLine.HasMarker ? parsedLine.Marker . " " : "") . previousSpineSegment
        }

        if (parsedLine.HasMarker) {
            return parsedLine.Indent . parsedLine.Marker . " "
        }

        return ""
    }

    static _CanRemoveEmptySmartPrefix(lineInfo) {
        parsedLine := this._ParseSmartLine(lineInfo.Text)
        if (!parsedLine) {
            return false
        }

        body := Trim(parsedLine.Body, " `t")
        if (body != "") {
            if !this._IsEmptySpineSegmentPrefix(body) {
                return false
            }
        } else if (!parsedLine.HasMarker) {
            return false
        }

        return true
    }

    static _RemoveLineContent(hCtrl, lineInfo) {
        this.SetSel(hCtrl, lineInfo.Bounds.Start, lineInfo.Bounds.ContentEnd)
        this.ReplaceSel(hCtrl, "")
        this.ScrollCaret(hCtrl)
    }

    static _ParseSmartLine(lineText) {
        if !RegExMatch(lineText, "^([ \t]*)(?:(-->|[-*+>]|\d+[.)])[ \t]+)?(.*)$", &match) {
            return false
        }

        return {
            Indent: match[1],
            HasMarker: match[2] != "",
            Marker: match[2],
            Body: match[3]
        }
    }

    static _GetNextListMarker(marker) {
        if RegExMatch(marker, "^(\d+)([.)])$", &numberMatch) {
            return (Integer(numberMatch[1]) + 1) . numberMatch[2] . " "
        }

        return marker . " "
    }

    static _GetNextSpineSegmentPrefix(text, requireTrailingText := true) {
        segmentInfo := this._ParseSpineSegmentPrefix(text)
        if (!segmentInfo || (requireTrailingText && Trim(segmentInfo.TrailingText, " `t") == "")) {
            return ""
        }

        nextStart := segmentInfo.LastEnd
        nextEnd := this._GetNextVertebra(nextStart.Region, nextStart.Number)
        if (!nextEnd) {
            return ""
        }

        return nextStart.Region . nextStart.Number . "-" . nextEnd.Region . nextEnd.Number . ": "
    }

    static _GetPreviousSpineSegmentPrefix(text) {
        segmentInfo := this._ParseSpineSegmentPrefix(text)
        if (!segmentInfo) {
            return ""
        }

        previousStart := this._GetPreviousVertebra(segmentInfo.FirstStart.Region, segmentInfo.FirstStart.Number)
        if (!previousStart) {
            return ""
        }

        currentStart := segmentInfo.FirstStart
        return previousStart.Region . previousStart.Number . "-" . currentStart.Region . currentStart.Number . ": "
    }

    static _IsEmptySpineSegmentPrefix(text) {
        segmentInfo := this._ParseSpineSegmentPrefix(text)
        return segmentInfo && Trim(segmentInfo.TrailingText, " `t") == ""
    }

    static _ParseSpineSegmentPrefix(text) {
        if !RegExMatch(text, "^(.+):([ \t]*.*)$", &prefixMatch) {
            return false
        }

        segmentItems := StrSplit(prefixMatch[1], ",")
        firstStart := false
        previousEnd := false
        lastEnd := false

        for segmentText in segmentItems {
            segmentText := Trim(segmentText, " `t")
            if !RegExMatch(segmentText, "i)^([CTLS])(\d+)-([CTLS])(\d+)$", &segmentMatch) {
                return false
            }

            start := {Region: StrUpper(segmentMatch[1]), Number: Integer(segmentMatch[2])}
            end := {Region: StrUpper(segmentMatch[3]), Number: Integer(segmentMatch[4])}
            if (!firstStart) {
                firstStart := start
            }
            expectedEnd := this._GetNextVertebra(start.Region, start.Number)

            if (!expectedEnd || expectedEnd.Region != end.Region || expectedEnd.Number != end.Number) {
                return false
            }

            if (previousEnd && (previousEnd.Region != start.Region || previousEnd.Number != start.Number)) {
                return false
            }

            previousEnd := end
            lastEnd := end
        }

        return {FirstStart: firstStart, LastEnd: lastEnd, TrailingText: prefixMatch[2]}
    }

    static _GetNextVertebra(region, number) {
        switch StrUpper(region) {
            case "C":
                return (number < 7)
                    ? {Region: "C", Number: number + 1}
                    : (number == 7 ? {Region: "T", Number: 1} : false)
            case "T":
                return (number < 12)
                    ? {Region: "T", Number: number + 1}
                    : (number == 12 ? {Region: "L", Number: 1} : false)
            case "L":
                return (number < 5)
                    ? {Region: "L", Number: number + 1}
                    : (number == 5 ? {Region: "S", Number: 1} : false)
            case "S":
                return false
        }

        return false
    }

    static _GetPreviousVertebra(region, number) {
        switch StrUpper(region) {
            case "C":
                return (number > 1) ? {Region: "C", Number: number - 1} : false
            case "T":
                return (number > 1)
                    ? {Region: "T", Number: number - 1}
                    : {Region: "C", Number: 7}
            case "L":
                return (number > 1)
                    ? {Region: "L", Number: number - 1}
                    : {Region: "T", Number: 12}
            case "S":
                return (number > 1)
                    ? {Region: "S", Number: number - 1}
                    : {Region: "L", Number: 5}
        }

        return false
    }

    static KillLine(hCtrl) {
        sel := this.GetSel(hCtrl)
        currentPos := sel.Start
        text := ControlGetText(hCtrl)

        foundPos := InStr(text, "`r", , currentPos + 1)
        targetPos := (foundPos == 0) ? StrLen(text) : foundPos - 1

        if (targetPos > currentPos) {
            this.SetSel(hCtrl, currentPos, targetPos)
            this.ReplaceSel(hCtrl, "")
        }
    }

    static DeleteCurrentLine(hCtrl) {
        this.SelectLineForRemoval(hCtrl)
        SendMessage(this.MSG.CLEAR, 0, 0, hCtrl)
        this.ScrollCaret(hCtrl)
    }

    static CutLineOrSelection(hCtrl) {
        sel := this.GetSel(hCtrl)
        if (sel.Start == sel.End) {
            this.SelectLineForRemoval(hCtrl)
        }

        SendMessage(this.MSG.CUT, 0, 0, hCtrl)
        this.ScrollCaret(hCtrl)
    }

    static CopyLineOrSelection(hCtrl) {
        sel := this.GetSel(hCtrl)
        didAutoSelect := false

        if (sel.Start == sel.End) {
            this.SelectLine(hCtrl)
            didAutoSelect := true
        }

        SendMessage(this.MSG.COPY, 0, 0, hCtrl)

        if (didAutoSelect) {
            this.SetSel(hCtrl, sel.Start, sel.Start)
        }
    }

    static MoveCaret(hCtrl, mode) {
        bounds := this.GetLogicalLineBoundaries(hCtrl)

        targetPos := 0
        if (mode = "Start") {
            targetPos := bounds.Start
        } else if (mode = "End") {
            targetPos := bounds.ContentEnd
        }

        this.SetSel(hCtrl, targetPos, targetPos)
        this.ScrollCaret(hCtrl)
    }

    static DeleteWordBackward(hCtrl) {
        try {
            fullText := ControlGetText(hCtrl)
        } catch {
            return
        }

        caretPos := this.GetSel(hCtrl).Start
        if (caretPos == 0) {
            return
        }

        i := caretPos

        GetCharType(char) {
            if (RisEditControl.IsSpace(char))
                return 1
            if (IsAlnum(char) || char == "_")
                return 2
            return 3
        }

        while (i > 0 && this.IsSpace(SubStr(fullText, i, 1))) {
            i--
        }

        if (i > 0) {
            targetType := GetCharType(SubStr(fullText, i, 1))

            while (i > 0) {
                currentChar := SubStr(fullText, i, 1)
                if (GetCharType(currentChar) != targetType)
                    break
                i--
            }
        }

        this.SetSel(hCtrl, i, caretPos)
        this.ReplaceSel(hCtrl, "")
    }

    static IsSpace(char) {
        return char == " " || char == "`t" || char == "`r" || char == "`n"
    }

    static MoveCurrentLine(hCtrl, direction) {
        sel := this.GetSel(hCtrl)
        if (sel.Start != sel.End) {
            return
        }

        fullText := ControlGetText(hCtrl)
        currLine := this.GetLogicalLineBoundaries(hCtrl, sel.Start)

        if (direction == "Up") {
            if (currLine.Start == 0) {
                return
            }

            searchPos := currLine.Start - 1
            if (searchPos > 0 && SubStr(fullText, searchPos, 1) == "`r") {
                searchPos -= 1
            }

            targetLine := this.GetLogicalLineBoundaries(hCtrl, searchPos)
            topLine := targetLine
            btmLine := currLine
        } else {
            if (currLine.FullEnd == StrLen(fullText)) {
                return
            }

            targetLine := this.GetLogicalLineBoundaries(hCtrl, currLine.FullEnd)
            topLine := currLine
            btmLine := targetLine
        }

        txtTop := SubStr(fullText, topLine.Start + 1, topLine.FullEnd - topLine.Start)
        txtBtm := SubStr(fullText, btmLine.Start + 1, btmLine.FullEnd - btmLine.Start)

        if (SubStr(txtTop, -1) == "`n" && SubStr(txtBtm, -1) != "`n") {
            txtTop := SubStr(txtTop, 1, StrLen(txtTop) - 2)
            txtBtm .= "`r`n"
        }

        this.SetSel(hCtrl, topLine.Start, btmLine.FullEnd)
        this.ReplaceSel(hCtrl, txtBtm . txtTop)

        offset := sel.Start - currLine.Start
        if (direction == "Up") {
            newPos := topLine.Start + offset
        } else {
            newPos := topLine.Start + StrLen(txtBtm) + offset
        }

        this.SetSel(hCtrl, newPos, newPos)
        this.ScrollCaret(hCtrl)
    }

    static SmartExtendSelection(hCtrl, direction) {
        lineIdx := this.LineFromChar(hCtrl)
        lineCount := this.GetLineCount(hCtrl)

        if (direction == "Up") {
            SendInput (lineIdx == 0) ? "+{Home}" : "+{Up}"
        } else if (direction == "Down") {
            SendInput (lineIdx == lineCount - 1) ? "+{End}" : "+{Down}"
        } else {
            return false
        }

        return true
    }

    static SmartPageMove(hCtrl, direction, extend := false) {
        modifierKey := extend ? "+" : ""
        prevFirstLine := this.GetFirstVisibleLine(hCtrl)

        if (direction == "Up") {
            SendInput(modifierKey "{PgUp}")
            Sleep 10
            currFirstLine := this.GetFirstVisibleLine(hCtrl)

            if (prevFirstLine == 0 && currFirstLine == 0) {
                if (extend) {
                    SendInput("+^{Home}")
                } else {
                    this.SetSel(hCtrl, 0, 0)
                }
                this.ScrollCaret(hCtrl)
            }
        } else if (direction == "Down") {
            SendInput(modifierKey "{PgDn}")
            Sleep 10
            currFirstLine := this.GetFirstVisibleLine(hCtrl)

            if (prevFirstLine == currFirstLine) {
                if (extend) {
                    SendInput("+^{End}")
                } else {
                    fullText := ControlGetText(hCtrl)
                    this.SetSel(hCtrl, StrLen(fullText), StrLen(fullText))
                }
                this.ScrollCaret(hCtrl)
            }
        } else {
            return false
        }

        return true
    }
}
