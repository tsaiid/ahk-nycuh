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
        GETFIRSTVISIBLELINE: 0x00CE
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

    static GetSel(hCtrl) {
        startBuf := Buffer(4, 0), endBuf := Buffer(4, 0)
        SendMessage(this.MSG.GETSEL, startBuf.Ptr, endBuf.Ptr, hCtrl)
        return {Start: NumGet(startBuf, "UInt"), End: NumGet(endBuf, "UInt")}
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
        bounds := this.GetLogicalLineBoundaries(hCtrl)

        if (mode = "Above") {
            this.SetSel(hCtrl, bounds.Start, bounds.Start)
            this.ReplaceSel(hCtrl, "`r`n")
            this.SetSel(hCtrl, bounds.Start, bounds.Start)
        } else {
            this.SetSel(hCtrl, bounds.ContentEnd, bounds.ContentEnd)
            this.ReplaceSel(hCtrl, "`r`n")
        }

        this.ScrollCaret(hCtrl)
    }
}
