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
}
