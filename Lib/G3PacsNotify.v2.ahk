#Requires AutoHotkey v2.0
#Include .\StackNotify.v2.ahk

class G3PacsNotify extends StackNotify {
    static _gui := 0
    static _queue := []
    static _slots := []
    static _slotItemIds := []
    static _sweepTimer := 0
    static _nextId := 0
    static _scale := 1.0
    static _refX := 0
    static _refY := 0
    static _theme := "light"
    static _visible := false

    static DefaultDuration := 1800
    static MaxVisible := 5
    static DedupeWindow := 800
    static MinWidth := 320
    static Width := 360
    static MaxWidth := 720
    static PaddingX := 16
    static PaddingY := 8
    static SlotGap := 8
    static SlotHeight := 28
    static TextPaddingY := 4
    static LineSpacingScale := 1.2
    static VerticalPositionRatio := 0.5
    static CornerRadius := 10
    static Transparent := 245
    static FontName := "Microsoft JhengHei UI"
    static FontSize := 12
    static Theme := "light"
    static AutoTheme := true
    static LightBackColor := "F8FAFC"
    static LightTextColor := "111827"
    static DarkBackColor := "202020"
    static DarkTextColor := "White"
    static ClickToDismiss := true
    static DebugBenchmark := false
}
