#Requires AutoHotkey v2.0
#Include .\StackNotify.v2.ahk

class RisNotify extends StackNotify {
    static TargetTitles := []
    static _gui := 0
    static _queue := []
    static _slots := []
    static _slotItemIds := []
    static _sweepTimer := 0
    static _nextId := 0
    static _scale := 1.0
    static _refX := 0
    static _refY := 0
    static _theme := "dark"
    static _visible := false

    static DefaultDuration := 1500
    static MaxVisible := 5
    static DedupeWindow := 800
    static MinWidth := 320
    static Width := 420
    static MaxWidth := 720
    static PaddingX := 24
    static PaddingY := 14
    static SlotGap := 8
    static SlotHeight := 36
    static TextPaddingY := 5
    static LineSpacingScale := 1.2
    static VerticalPositionRatio := 0.4
    static CornerRadius := 12
    static Transparent := 235
    static FontName := "Microsoft JhengHei UI"
    static FontSize := 12
    static Theme := "dark"
    static AutoTheme := false
    static LightBackColor := "F8FAFC"
    static LightTextColor := "111827"
    static DarkBackColor := "202020"
    static DarkTextColor := "White"
    static ClickToDismiss := true

    static _ResolveReferencePoint() {
        hwnd := 0

        for title in this.TargetTitles {
            if (h := WinExist(title)) {
                try {
                    if (WinGetMinMax(h) != -1) {
                        hwnd := h
                        break
                    }
                }
            }
        }

        if (!hwnd) {
            if (h := WinActive("A")) {
                try {
                    if (WinGetMinMax(h) != -1)
                        hwnd := h
                }
            }
        }

        if (hwnd) {
            try {
                WinGetPos(&wx, &wy, &ww, &wh, hwnd)
                return {x: wx + Floor(ww / 2), y: wy + Floor(wh / 2)}
            }
        }

        MouseGetPos(&refX, &refY)
        return {x: refX, y: refY}
    }
}
