#Requires AutoHotkey v2.0
A_MaxHotkeysPerInterval := 200

;; for INFINITT PACS
#HotIf MouseIsOverG3Pacs()
$LButton::HandleG3PacsLeftClick()
$WheelUp::FocusG3PacsUnderMouseAndScroll("WheelUp")
$WheelDown::FocusG3PacsUnderMouseAndScroll("WheelDown")
#HotIf

#HotIf IsG3PacsHotkeyContext()
^s::SelectG3PacsSortBySliceLocationDesc()
^!d::CopyG3PacsActiveSeriesDebugInfo()
$Up::ClickG3PacsUnderMouseAndSendKey("Up")
$Down::ClickG3PacsUnderMouseAndSendKey("Down")
#HotIf

#HotIf IsG3PacsCalciumScoreContext()
!c::AnalyzeG3PacsCalciumScoreFromClipboard()
#HotIf

MouseIsOverG3Pacs() {
    MouseGetPos(,, &hwnd)
    try return WinGetProcessName("ahk_id " hwnd) = "G3PACS.exe"
    return false
}

IsG3PacsHotkeyContext() {
    return WinActive("ahk_exe G3PACS.exe") && MouseIsOverG3Pacs()
}

IsG3PacsCalciumScoreContext() {
    return IsG3PacsHotkeyContext() && FileExist(A_ScriptDir "\config\private.ini")
}

FocusG3PacsUnderMouseAndScroll(direction) {
    MouseGetPos(,, &hwnd)
    if (hwnd && !WinActive("ahk_id " hwnd)) {
        WinActivate("ahk_id " hwnd)
    }
    Click(direction)
}

ClickG3PacsUnderMouseAndSendKey(keyName) {
    MouseGetPos(&mouseX, &mouseY, &hwnd, &controlHwnd, 2)
    if !IsG3PacsActiveSeriesUnderMouse(hwnd, controlHwnd)
        && !IsRecentG3PacsLeftClick(mouseX, mouseY) {
        Click()
        RecordG3PacsLeftClick(mouseX, mouseY)
    }
    Send("{" keyName "}")
}

IsG3PacsActiveSeriesUnderMouse(hwnd, controlHwnd) {
    if !hwnd || !controlHwnd
        return false

    try controlClassNN := ControlGetClassNN(controlHwnd)
    catch {
        return false
    }

    srsClassNN := GetG3PacsSrsControlForFocusClassNN(controlClassNN, hwnd)
    return srsClassNN != "" && GetG3PacsSrsControlFocusState(srsClassNN, hwnd) = "active"
}

GetG3PacsSrsControlForFocusClassNN(focusClassNN, hwnd) {
    static classPrefix := "Afx:00400000:b:00000000:00000013:00000000"
    static configs := [
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
        [5, 57, 37, 185],
        [5, 57, 39, 185],
        [5, 57, 41, 185],
        [5, 57, 43, 185],
        [5, 57, 45, 185],
        [5, 57, 47, 185],
        [5, 57, 49, 185],
        [5, 57, 51, 185],
    ]

    for cfg in configs {
        Loop 8 {
            offset := A_Index - 1
            expectedFocusNN := classPrefix . (cfg[1] + offset)
            if (focusClassNN != expectedFocusNN)
                continue

            imgClassNN := "ComboBox" . (cfg[2] + (offset * 6))
            srsClassNN := "AfxWnd140u" . (cfg[3] + (offset * 3))
            descClassNN := "Button" . (cfg[4] + (offset * 5))
            if IsG3PacsSeriesPatternMatch(srsClassNN, imgClassNN, descClassNN, focusClassNN, hwnd)
                return srsClassNN
        }
    }

    return ""
}

IsG3PacsSeriesPatternMatch(srsClassNN, imgClassNN, descClassNN, focusClassNN, hwnd) {
    try {
        srsText := ControlGetText(srsClassNN, hwnd)
        if !InStr(srsText, "VMTool")
            return false

        descText := ControlGetText(descClassNN, hwnd)
        if !RegExMatch(descText, "^\(\d+\)\s")
            return false
    } catch {
        return false
    }

    return IsG3PacsSpatialControlMatch(srsClassNN, focusClassNN, hwnd)
        && IsG3PacsSpatialControlMatch(imgClassNN, focusClassNN, hwnd)
        && IsG3PacsSpatialControlMatch(descClassNN, focusClassNN, hwnd)
}

IsG3PacsSpatialControlMatch(childClassNN, focusClassNN, hwnd) {
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

GetG3PacsSrsControlFocusState(srsClassNN, hwnd) {
    try {
        ControlGetPos(&x, &y, &w, &h, srsClassNN, hwnd)
        if (w <= 8 || h <= 8)
            return "unknown"

        pt := Buffer(8, 0)
        NumPut("int", x, pt, 0)
        NumPut("int", y, pt, 4)
        DllCall("ClientToScreen", "ptr", hwnd, "ptr", pt)
        screenX := NumGet(pt, 0, "int")
        screenY := NumGet(pt, 4, "int")

        return GetG3PacsSrsColorFocusState(screenX, screenY, w, h)
    }

    return "unknown"
}

GetG3PacsSrsColorFocusState(screenX, screenY, width, height) {
    sample := GetG3PacsScreenPixelColor(
        screenX + 3,
        screenY + 3
    )
    if !sample.ok
        return "unknown"
    if IsG3PacsColorNear(sample, 0x1B, 0x1D, 0x20, 18)
        return "active"
    if IsG3PacsColorNear(sample, 0x4B, 0x4D, 0x5D, 18)
        return "inactive"
    return "unknown"
}

IsG3PacsColorNear(sample, red, green, blue, tolerance := 30) {
    return Abs(sample.red - red) + Abs(sample.green - green) + Abs(sample.blue - blue) <= tolerance
}

GetG3PacsScreenPixelColor(x, y) {
    hdc := DllCall("GetDC", "ptr", 0, "ptr")
    if !hdc
        return {ok: false, hex: "GetDC failed", brightness: 255}

    try {
        color := DllCall("GetPixel", "ptr", hdc, "int", x, "int", y, "uint")
        if (color = 0xFFFFFFFF)
            return {ok: false, hex: "GetPixel failed", brightness: 255}

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

CopyG3PacsActiveSeriesDebugInfo() {
    debugInfo := BuildG3PacsActiveSeriesDebugInfo()
    A_Clipboard := debugInfo
    ToolTip("G3PACS focus debug 已複製到剪貼簿")
    SetTimer(() => ToolTip(), -2500)
}

BuildG3PacsActiveSeriesDebugInfo() {
    MouseGetPos(&mouseX, &mouseY, &hwnd, &controlHwnd, 2)

    lines := "G3PACS active series debug`r`n"
    lines .= "Time: " FormatTime(A_Now, "yyyy-MM-dd HH:mm:ss") "`r`n"
    lines .= "Mouse: " mouseX ", " mouseY "`r`n"
    lines .= "Window hwnd: " hwnd "`r`n"
    try lines .= "Window title: " WinGetTitle("ahk_id " hwnd) "`r`n"
    catch as err
        lines .= "Window title: error - " err.Message "`r`n"

    lines .= "Mouse control hwnd: " controlHwnd "`r`n"
    controlClassNN := ""
    try controlClassNN := ControlGetClassNN(controlHwnd)
    catch as err
        lines .= "Mouse control ClassNN: error - " err.Message "`r`n"
    if (controlClassNN != "")
        lines .= "Mouse control ClassNN: " controlClassNN "`r`n"

    focusHwnd := 0
    try focusHwnd := ControlGetFocus("ahk_id " hwnd)
    catch as err
        lines .= "Win32 focused hwnd: error - " err.Message "`r`n"
    if focusHwnd {
        lines .= "Win32 focused hwnd: " focusHwnd "`r`n"
        try lines .= "Win32 focused ClassNN: " ControlGetClassNN(focusHwnd) "`r`n"
    } else {
        lines .= "Win32 focused hwnd: none`r`n"
    }

    srsClassNN := controlClassNN != "" ? GetG3PacsSrsControlForFocusClassNN(controlClassNN, hwnd) : ""
    lines .= "Mapped Srs ClassNN: " (srsClassNN != "" ? srsClassNN : "not found") "`r`n"
    srsState := srsClassNN != "" ? GetG3PacsSrsControlFocusState(srsClassNN, hwnd) : "unknown"
    lines .= "Active decision: " srsState " (Srs target active #1B1D20, inactive #4B4D5D)`r`n"

    if (srsClassNN = "")
        return lines

    try lines .= "Srs text: " Trim(ControlGetText(srsClassNN, hwnd)) "`r`n"
    catch as err
        lines .= "Srs text: error - " err.Message "`r`n"

    try {
        ControlGetPos(&x, &y, &w, &h, srsClassNN, hwnd)
        lines .= "Srs client rect: x=" x " y=" y " w=" w " h=" h "`r`n"

        pt := Buffer(8, 0)
        NumPut("int", x, pt, 0)
        NumPut("int", y, pt, 4)
        DllCall("ClientToScreen", "ptr", hwnd, "ptr", pt)
        screenX := NumGet(pt, 0, "int")
        screenY := NumGet(pt, 4, "int")
        lines .= "Srs screen rect: x=" screenX " y=" screenY " w=" w " h=" h "`r`n"
        lines .= GetG3PacsSrsTargetColorScanDebug(screenX, screenY, w, h)
        lines .= GetG3PacsColorScanDebug("Srs horizontal scan y=25%", screenX, screenY + Integer(h * 0.25), w, 17)
        lines .= GetG3PacsColorScanDebug("Srs horizontal scan y=50%", screenX, screenY + Integer(h * 0.50), w, 17)
        lines .= GetG3PacsColorScanDebug("Srs horizontal scan y=75%", screenX, screenY + Integer(h * 0.75), w, 17)
        lines .= "Pixel samples (screen x,y RGB brightness):`r`n"

        samplePoints := [
            [0.15, 0.25], [0.50, 0.25], [0.85, 0.25],
            [0.15, 0.50], [0.50, 0.50], [0.85, 0.50],
            [0.15, 0.75], [0.50, 0.75], [0.85, 0.75],
        ]
        for point in samplePoints {
            px := screenX + Integer(w * point[1])
            py := screenY + Integer(h * point[2])
            sample := GetG3PacsScreenPixelColor(px, py)
            lines .= "  " px "," py " " sample.hex " brightness=" sample.brightness "`r`n"
        }
    } catch as err {
        lines .= "Srs rect/samples: error - " err.Message "`r`n"
    }

    return lines
}

GetG3PacsSrsTargetColorScanDebug(screenX, screenY, width, height) {
    activeCount := 0
    inactiveCount := 0
    activeHits := ""
    inactiveHits := ""
    xStep := 16
    lines := "Srs target color scan step=" xStep " (active #1B1D20, inactive #4B4D5D):`r`n"

    for yRatio in [0.25, 0.50, 0.75] {
        py := screenY + Integer(height * yRatio)
        x := screenX + 2
        while (x <= screenX + width - 3) {
            px := x
            sample := GetG3PacsScreenPixelColor(px, py)
            if sample.ok && IsG3PacsColorNear(sample, 0x1B, 0x1D, 0x20, 18) {
                activeCount += 1
                if (activeCount <= 12)
                    activeHits .= "  active " px "," py " " sample.hex " b=" sample.brightness "`r`n"
            } else if sample.ok && IsG3PacsColorNear(sample, 0x4B, 0x4D, 0x5D, 18) {
                inactiveCount += 1
                if (inactiveCount <= 12)
                    inactiveHits .= "  inactive " px "," py " " sample.hex " b=" sample.brightness "`r`n"
            }
            x += xStep
        }
    }

    lines .= "  active-like count: " activeCount "`r`n"
    lines .= "  inactive-like count: " inactiveCount "`r`n"
    lines .= activeHits
    lines .= inactiveHits
    return lines
}

GetG3PacsColorScanDebug(label, startX, startY, length, sampleCount, vertical := false) {
    if (length <= 0 || sampleCount <= 0)
        return label ": invalid scan`r`n"

    step := sampleCount > 1 ? length / (sampleCount - 1) : 0
    colors := Map()
    minBrightness := 255
    maxBrightness := 0
    sumBrightness := 0
    lines := label " (" (vertical ? "vertical" : "horizontal") ", samples=" sampleCount "):`r`n"

    Loop sampleCount {
        offset := Integer((A_Index - 1) * step)
        px := vertical ? startX : startX + offset
        py := vertical ? startY + offset : startY
        sample := GetG3PacsScreenPixelColor(px, py)
        colors[sample.hex] := colors.Has(sample.hex) ? colors[sample.hex] + 1 : 1
        minBrightness := Min(minBrightness, sample.brightness)
        maxBrightness := Max(maxBrightness, sample.brightness)
        sumBrightness += sample.brightness
        lines .= "  " px "," py " " sample.hex " b=" sample.brightness "`r`n"
    }

    lines .= "  brightness min/avg/max=" minBrightness "/" Round(sumBrightness / sampleCount, 1) "/" maxBrightness "`r`n"
    lines .= "  unique colors: " FormatG3PacsColorCounts(colors) "`r`n"
    return lines
}

FormatG3PacsColorCounts(colors) {
    output := ""
    for color, count in colors {
        output .= (output != "" ? ", " : "") color "=" count
    }
    return output
}

HandleG3PacsLeftClick() {
    static lastClickTime := 0
    static lastClickX := 0
    static lastClickY := 0

    MouseGetPos(&mouseX, &mouseY,, &controlClassNN)
    RecordG3PacsLeftClick(mouseX, mouseY)

    if lastClickTime && IsWithinG3PacsDoubleClick(mouseX, mouseY, lastClickX, lastClickY, A_TickCount - lastClickTime) {
        lastClickTime := 0
        if IsG3PacsImageClassNN(controlClassNN) {
            KeyWait("LButton")
            Send("{Space}")
            return
        }
    }

    lastClickTime := A_TickCount
    lastClickX := mouseX
    lastClickY := mouseY
    Click("Down")
    KeyWait("LButton")
    Click("Up")
}

IsWithinG3PacsDoubleClick(mouseX, mouseY, lastClickX, lastClickY, elapsedMs) {
    static doubleClickTime := DllCall("GetDoubleClickTime", "UInt")
    static doubleClickWidth := DllCall("GetSystemMetrics", "Int", 36, "Int") ; SM_CXDOUBLECLK
    static doubleClickHeight := DllCall("GetSystemMetrics", "Int", 37, "Int") ; SM_CYDOUBLECLK

    return elapsedMs <= doubleClickTime
        && Abs(mouseX - lastClickX) <= doubleClickWidth
        && Abs(mouseY - lastClickY) <= doubleClickHeight
}

RecordG3PacsLeftClick(mouseX, mouseY) {
    state := GetG3PacsLastLeftClick()
    state.time := A_TickCount
    state.x := mouseX
    state.y := mouseY
}

IsRecentG3PacsLeftClick(mouseX, mouseY) {
    state := GetG3PacsLastLeftClick()
    return state.time
        && IsWithinG3PacsDoubleClick(mouseX, mouseY, state.x, state.y, A_TickCount - state.time)
}

GetG3PacsLastLeftClick() {
    static state := {time: 0, x: 0, y: 0}
    return state
}

IsG3PacsImageClassNN(controlClassNN) {
    return RegExMatch(controlClassNN, "^Afx:[0-9A-Fa-f]{8}:[0-9A-Fa-f]+:[0-9A-Fa-f]{8}:[0-9A-Fa-f]{8}:[0-9A-Fa-f]+$")
}

SelectG3PacsSortBySliceLocationDesc() {
    Click("Right")
    Sleep(250)
    popupHwnd := WaitForG3PacsPopupMenu("排序")
    if !popupHwnd {
        NotifyG3PacsHotkeyError("找不到 G3PACS 右鍵選單")
        return
    }

    hMenu := SendMessage(0x01E1, 0, 0,, "ahk_id " popupHwnd) ; MN_GETHMENU
    if !hMenu {
        NotifyG3PacsHotkeyError("無法取得右鍵選單 HMENU")
        return
    }

    sortItem := FindMenuItemByText(hMenu, "排序")
    if !sortItem || !sortItem.submenu {
        NotifyG3PacsHotkeyError("找不到選單項目: 排序", hMenu)
        return
    }

    HoverMenuItem(popupHwnd, hMenu, sortItem.index)

    subMenuHwnd := WaitForG3PacsPopupMenu("切面位置 (遞減)", popupHwnd)
    if !subMenuHwnd {
        NotifyG3PacsHotkeyError("找不到排序子選單")
        return
    }

    subMenu := SendMessage(0x01E1, 0, 0,, "ahk_id " subMenuHwnd) ; MN_GETHMENU
    if !subMenu {
        NotifyG3PacsHotkeyError("無法取得排序子選單 HMENU")
        return
    }

    sliceDescItem := FindMenuItemByText(subMenu, "切面位置 (遞減)")
    if !sliceDescItem {
        NotifyG3PacsHotkeyError("找不到選單項目: 切面位置 (遞減)", subMenu)
        return
    }

    ClickMenuItem(subMenuHwnd, subMenu, sliceDescItem.index)
}

WaitForG3PacsPopupMenu(requiredItemText := "", excludeHwnd := 0, timeoutMs := 1500) {
    deadline := A_TickCount + timeoutMs
    while A_TickCount < deadline {
        for hwnd in WinGetList("ahk_class #32768 ahk_exe G3PACS.exe") {
            if hwnd = excludeHwnd
                continue
            if IsG3PacsPopupMenu(hwnd) && PopupMenuHasItem(hwnd, requiredItemText)
                return hwnd
        }

        Sleep(20)
    }
    return 0
}

IsG3PacsPopupMenu(hwnd) {
    if !hwnd
        return false

    try return WinGetClass("ahk_id " hwnd) = "#32768"
        && WinGetProcessName("ahk_id " hwnd) = "G3PACS.exe"
    return false
}

PopupMenuHasItem(hwnd, itemText) {
    if itemText = ""
        return true

    hMenu := SendMessage(0x01E1, 0, 0,, "ahk_id " hwnd) ; MN_GETHMENU
    return hMenu && FindMenuItemByText(hMenu, itemText)
}

FindMenuItemByText(hMenu, targetText) {
    itemCount := DllCall("GetMenuItemCount", "Ptr", hMenu, "Int")
    selectableOffset := 0
    normalizedTargetText := NormalizeMenuItemText(targetText)

    loop itemCount {
        index := A_Index - 1
        itemText := GetMenuItemText(hMenu, index)
        if !IsSelectableMenuItem(hMenu, index, itemText)
            continue

        subMenu := DllCall("GetSubMenu", "Ptr", hMenu, "Int", index, "Ptr")
        if NormalizeMenuItemText(itemText) = normalizedTargetText
            return {index: index, offset: selectableOffset, submenu: subMenu}

        selectableOffset += 1
    }

    return false
}

IsSelectableMenuItem(hMenu, index, itemText) {
    return itemText != ""
}

NormalizeMenuItemText(itemText) {
    itemText := StrReplace(itemText, "&")
    itemText := StrReplace(itemText, "‧‧", "..")
    itemText := RegExReplace(itemText, "\t.*$")
    return Trim(itemText)
}

GetMenuItemText(hMenu, index) {
    length := DllCall("GetMenuString", "Ptr", hMenu, "UInt", index, "Ptr", 0, "Int", 0, "UInt", 0x0400, "Int") ; MF_BYPOSITION
    if length < 1
        return ""

    textBuffer := Buffer((length + 1) * 2, 0)
    DllCall("GetMenuString", "Ptr", hMenu, "UInt", index, "Ptr", textBuffer, "Int", length + 1, "UInt", 0x0400, "Int")
    return StrGet(textBuffer, "UTF-16")
}

HoverMenuItem(hwnd, hMenu, index) {
    point := GetMenuItemCenter(hwnd, hMenu, index)
    if !point {
        Send("{Home}")
        return
    }

    MouseMove(point.x, point.y, 0)
    Sleep(250)
}

ClickMenuItem(hwnd, hMenu, index) {
    point := GetMenuItemCenter(hwnd, hMenu, index)
    if !point {
        Send("{Enter}")
        return
    }

    Click(point.x, point.y)
}

GetMenuItemCenter(hwnd, hMenu, index) {
    rect := Buffer(16, 0)
    if !DllCall("GetMenuItemRect", "Ptr", hwnd, "Ptr", hMenu, "UInt", index, "Ptr", rect, "Int")
        return false

    left := NumGet(rect, 0, "Int")
    top := NumGet(rect, 4, "Int")
    right := NumGet(rect, 8, "Int")
    bottom := NumGet(rect, 12, "Int")
    return {x: (left + right) // 2, y: (top + bottom) // 2}
}

DumpMenuItems(hMenu) {
    itemCount := DllCall("GetMenuItemCount", "Ptr", hMenu, "Int")
    if itemCount < 1
        return "(no items)"

    lines := ""
    loop itemCount {
        index := A_Index - 1
        itemText := GetMenuItemText(hMenu, index)
        subMenu := DllCall("GetSubMenu", "Ptr", hMenu, "Int", index, "Ptr")
        lines .= Format("{:02}", A_Index) ". submenu: " (subMenu ? subMenu : "no") " text: " (itemText != "" ? itemText : "(blank)") "`n"
    }
    return RTrim(lines, "`n")
}

NotifyG3PacsHotkeyError(message, hMenu := 0) {
    if hMenu
        A_Clipboard := message "`n`nMenu items:`n" DumpMenuItems(hMenu)
    ToolTip(message)
    SetTimer(() => ToolTip(), -3000)
}

AnalyzeG3PacsCalciumScoreFromClipboard() {
    static isRunning := false

    if isRunning {
        NotifyG3PacsAIStatus("Calcium Score AI 仍在處理中", 2500)
        return
    }

    isRunning := true
    startedAt := A_TickCount
    pngPath := ""
    showDebugWindow := RisConfig.AI.CalciumScoreImage.HasOwnProp("ShowDebugWindow")
        ? RisConfig.AI.CalciumScoreImage.ShowDebugWindow
        : false

    try {
        NotifyG3PacsAIStatus("G3PACS: 複製影像到 clipboard...", 2500)
        ClearClipboardAllFormats()
        KeyWait("Alt")
        Sleep(80)
        SendEvent("{Ctrl down}c{Ctrl up}")

        if !WaitForG3PacsClipboardImage(5000) {
            throw Error("未偵測到 clipboard 影像。請確認目前 G3PACS 畫面可用 Ctrl+C 複製影像。")
        }

        copyTime := A_TickCount - startedAt
        saveResult := SaveClipboardImageToPng()
        pngPath := saveResult.Path
        imageInfo := showDebugWindow ? GetImageFileInfo(pngPath) : ""
        imageBytes := FileGetSize(pngPath)
        imageBase64 := Base64EncodeFile(pngPath)
        encodeTime := A_TickCount - startedAt - copyTime

        NotifyG3PacsAIStatus("Gemini 分析 Calcium Score 影像中...", 4000)
        promptText := RisConfig.AI.CalciumScoreImage.Prompt
        apiStart := A_TickCount
        response := RisAIService.CallGoogleWithInlineImage(
            promptText,
            { MimeType: "image/png", Base64: imageBase64 },
            RisConfig.AI.CalciumScoreImage,
            { Notify: NotifyG3PacsAIStatus }
        )
        apiTime := A_TickCount - apiStart

        resultText := RisAIOrchestration.NormalizeResult(response.Result)
        tempFileStatus := DeleteTempImageFile(pngPath)
        debugInfo := {
            APIKeyName: response.APIKeyName,
            Model: response.Model,
            Provider: response.Provider,
            CopyTime: copyTime,
            EncodeTime: encodeTime,
            ApiTime: apiTime,
            TotalTime: A_TickCount - startedAt,
            PromptChars: StrLen(promptText),
            ImagePath: pngPath,
            ImageBytes: imageBytes,
            ImageBase64Chars: StrLen(imageBase64),
            ImageInfo: imageInfo,
            SaveMethod: saveResult.Method,
            TempFileStatus: tempFileStatus,
            ClipboardFormats: showDebugWindow ? GetClipboardFormatDebug() : ""
        }

        A_Clipboard := resultText
        if showDebugWindow {
            ShowG3PacsCalciumScoreResult(resultText, debugInfo, true)
            NotifyG3PacsAIStatus("Calcium Score AI 完成，結果已複製並顯示 debug", 2500)
        } else {
            NotifyG3PacsAIStatus("Calcium Score AI 完成，結果已複製到剪貼簿", 2500)
        }
    } catch as err {
        errMsg := "G3PACS Calcium Score AI 失敗：" . err.Message
        if (pngPath != "" && FileExist(pngPath)) {
            errMsg .= "`r`n`r`nTemp image retained:`r`n" . pngPath
        }
        RisAIDebugGui.ShowDebugError(errMsg, { Notify: NotifyG3PacsAIStatus })
    } finally {
        isRunning := false
    }
}

WaitForG3PacsClipboardImage(timeoutMs := 2500) {
    deadline := A_TickCount + timeoutMs
    while A_TickCount < deadline {
        if HasClipboardImageFormat()
            return true
        Sleep(50)
    }
    return false
}

HasClipboardImageFormat() {
    return DllCall("IsClipboardFormatAvailable", "UInt", 2, "Int") ; CF_BITMAP
        || DllCall("IsClipboardFormatAvailable", "UInt", 8, "Int") ; CF_DIB
        || DllCall("IsClipboardFormatAvailable", "UInt", 17, "Int") ; CF_DIBV5
}

ClearClipboardAllFormats() {
    if !DllCall("OpenClipboard", "Ptr", A_ScriptHwnd, "Int")
        return false

    DllCall("EmptyClipboard", "Int")
    DllCall("CloseClipboard", "Int")
    return true
}

SaveClipboardImageToPng() {
    pngPath := A_Temp "\G3PacsCalciumScore_" A_Now "_" A_TickCount ".png"
    return { Path: SaveClipboardImageToPngWithPowerShell(pngPath), Method: "PowerShell System.Windows.Forms" }
}

DeleteTempImageFile(filePath) {
    if !FileExist(filePath) {
        return "not found: " . filePath
    }

    try {
        FileDelete(filePath)
        return "deleted: " . filePath
    } catch as err {
        return "delete failed: " . err.Message . " | " . filePath
    }
}

SaveClipboardImageToPngWithPowerShell(pngPath) {
    psPath := A_Temp "\G3PacsClipboardImageToPng.ps1"
    psScript := "
    (
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
$targetPath = $args[0]
$image = [System.Windows.Forms.Clipboard]::GetImage()
if ($null -eq $image) { exit 2 }
try {
    $image.Save($targetPath, [System.Drawing.Imaging.ImageFormat]::Png)
} finally {
    $image.Dispose()
}
    )"

    if FileExist(psPath)
        FileDelete(psPath)
    FileAppend(psScript, psPath, "UTF-8")

    exitCode := RunWait("powershell.exe -NoProfile -WindowStyle Hidden -STA -ExecutionPolicy Bypass -File " QuoteCmdArg(psPath) " " QuoteCmdArg(pngPath), , "Hide")
    if (exitCode != 0 || !FileExist(pngPath)) {
        throw Error("clipboard 影像轉 PNG 失敗 (PowerShell exit code: " . exitCode . ")")
    }

    return pngPath
}

GetImageFileInfo(filePath) {
    psPath := A_Temp "\G3PacsImageInfo.ps1"
    outputPath := A_Temp "\G3PacsImageInfo_" A_TickCount ".txt"
    psScript := "
    (
Add-Type -AssemblyName System.Drawing
$image = [System.Drawing.Image]::FromFile($args[0])
try {
    $lines = @('Width=' + $image.Width, 'Height=' + $image.Height, 'PixelFormat=' + $image.PixelFormat)
    Set-Content -LiteralPath $args[1] -Value $lines -Encoding UTF8
} finally {
    $image.Dispose()
}
    )"

    if FileExist(psPath)
        FileDelete(psPath)
    FileAppend(psScript, psPath, "UTF-8")

    exitCode := RunWait("powershell.exe -NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File " QuoteCmdArg(psPath) " " QuoteCmdArg(filePath) " " QuoteCmdArg(outputPath), , "Hide")
    if (exitCode != 0 || !FileExist(outputPath)) {
        return "Image info unavailable (PowerShell exit code: " . exitCode . ")"
    }

    try {
        return Trim(FileRead(outputPath, "UTF-8"), "`r`n `t")
    } finally {
        try FileDelete(outputPath)
    }
}

GetClipboardFormatDebug() {
    formatNames := Map(
        1, "CF_TEXT",
        2, "CF_BITMAP",
        3, "CF_METAFILEPICT",
        8, "CF_DIB",
        13, "CF_UNICODETEXT",
        14, "CF_ENHMETAFILE",
        17, "CF_DIBV5"
    )

    if !DllCall("OpenClipboard", "Ptr", A_ScriptHwnd, "Int")
        return "OpenClipboard failed"

    try {
        lines := ""
        currentFormat := 0
        while (currentFormat := DllCall("EnumClipboardFormats", "UInt", currentFormat, "UInt")) {
            formatName := formatNames.Has(currentFormat) ? formatNames[currentFormat] : GetRegisteredClipboardFormatName(currentFormat)
            lines .= currentFormat . ": " . formatName . "`r`n"
        }
        return RTrim(lines, "`r`n")
    } finally {
        DllCall("CloseClipboard", "Int")
    }
}

GetRegisteredClipboardFormatName(formatId) {
    nameBuffer := Buffer(256 * 2, 0)
    length := DllCall("GetClipboardFormatNameW", "UInt", formatId, "Ptr", nameBuffer.Ptr, "Int", 256, "Int")
    if (length > 0)
        return StrGet(nameBuffer, length, "UTF-16")
    return "(registered/unknown)"
}

Base64EncodeFile(filePath) {
    imageFile := FileOpen(filePath, "r")
    if !imageFile {
        throw Error("無法讀取影像檔: " . filePath)
    }

    fileSize := imageFile.Length
    if (fileSize < 1) {
        imageFile.Close()
        throw Error("影像檔為空: " . filePath)
    }

    bytes := Buffer(fileSize, 0)
    bytesRead := imageFile.RawRead(bytes, fileSize)
    imageFile.Close()

    chars := 0
    flags := 0x40000001 ; CRYPT_STRING_BASE64 | CRYPT_STRING_NOCRLF
    if !DllCall("Crypt32\CryptBinaryToStringW", "Ptr", bytes.Ptr, "UInt", bytesRead, "UInt", flags, "Ptr", 0, "UInt*", &chars) {
        throw Error("計算 base64 長度失敗")
    }

    output := Buffer(chars * 2, 0)
    if !DllCall("Crypt32\CryptBinaryToStringW", "Ptr", bytes.Ptr, "UInt", bytesRead, "UInt", flags, "Ptr", output.Ptr, "UInt*", &chars) {
        throw Error("影像 base64 編碼失敗")
    }

    return StrGet(output, chars, "UTF-16")
}

QuoteCmdArg(value) {
    return '"' StrReplace(value, '"', '\"') '"'
}

NotifyG3PacsAIStatus(message, duration := 1800) {
    G3PacsNotify.Show(message, duration)
}

ShowG3PacsCalciumScoreResult(resultText, debugInfo, showDebugWindow := false) {
    resultGui := Gui("+AlwaysOnTop +Resize", "G3PACS Calcium Score AI")
    resultGui.MarginX := 14
    resultGui.MarginY := 12
    resultGui.SetFont("s10", "Microsoft JhengHei UI")

    resultGui.Add("Text", "w760", "Result:")
    resultEditHeight := showDebugWindow ? 110 : 180
    resultEdit := resultGui.Add("Edit", "w760 h" . resultEditHeight . " ReadOnly Multi -Wrap", resultText)

    if showDebugWindow {
        debugText := Format(
            "Provider: {1}`r`nModel: {2}`r`nAPI Key: {3}`r`nCopy: {4} ms | Encode: {5} ms | API: {6} | Total: {7} ms`r`nPrompt: {8} chars | Image: {9} bytes | Base64: {10}`r`nSave method: {11}`r`nTemp file: {12}`r`n{13}`r`n`r`nClipboard formats:`r`n{14}",
            debugInfo.Provider,
            debugInfo.Model,
            debugInfo.APIKeyName,
            debugInfo.CopyTime,
            debugInfo.EncodeTime,
            debugInfo.ApiTime,
            debugInfo.TotalTime,
            debugInfo.PromptChars,
            debugInfo.ImageBytes,
            debugInfo.ImageBase64Chars,
            debugInfo.SaveMethod,
            debugInfo.TempFileStatus,
            debugInfo.ImageInfo,
            debugInfo.ClipboardFormats
        )

        resultGui.Add("Text", "xm y+12 w760", "Debug:")
        debugEdit := resultGui.Add("Edit", "w760 h220 ReadOnly Multi -Wrap", debugText)
    }

    btnCopyResult := resultGui.Add("Button", "Default w140 xm y+12", "複製結果")
    btnCopyResult.OnEvent("Click", (*) => (
        A_Clipboard := resultText,
        NotifyG3PacsAIStatus("已複製結果", 1500)
    ))

    if showDebugWindow {
        btnCopyDebug := resultGui.Add("Button", "w140 x+10 yp", "複製 debug")
        btnCopyDebug.OnEvent("Click", (*) => (
            A_Clipboard := debugText,
            NotifyG3PacsAIStatus("已複製 debug", 1500)
        ))
    }

    btnClose := resultGui.Add("Button", "w100 x+10 yp", "關閉")
    btnClose.OnEvent("Click", (*) => resultGui.Destroy())
    resultGui.OnEvent("Escape", (*) => resultGui.Destroy())

    resultGui.Show("AutoSize Center")
    SendMessage(0x00B1, 0, 0, resultEdit.Hwnd)
    if showDebugWindow {
        SendMessage(0x00B1, 0, 0, debugEdit.Hwnd)
    }
    btnCopyResult.Focus()
}

#HotIf WinActive("ahk_exe G3PACS.exe")
w::
{
    try {
        ; v2 的 ControlGetFocus 回傳 HWND，需轉為 ClassNN 才能做字串比對
        hCtl := ControlGetFocus("A")
        FocusedControl := ControlGetClassNN(hCtl)
    } catch {
        FocusedControl := ""
    }

    OutputVar := WinGetTitle("A")
    ;MsgBox(OutputVar)

    if (OutputVar = "INFINITT PACS" && SubStr(FocusedControl, 1, 3) = "Afx") {
        DiffSyncBtns := ["Button1", "Button85", "Button90", "Button102"]
        for idx, btn in DiffSyncBtns {
            try {
                t := ControlGetText(btn)
                if (t == " Auto sync" || t == "自動同步") {
                    ControlClick(btn)
                    break
                }
            }
        }
    } else {
        Send("w")
    }
}

f::
{
    try {
        hCtl := ControlGetFocus("A")
        FocusedControl := ControlGetClassNN(hCtl)
    } catch {
        FocusedControl := ""
    }

    OutputVar := WinGetTitle("A")
    ;MsgBox(OutputVar)

    if (OutputVar = "INFINITT PACS" && SubStr(FocusedControl, 1, 3) = "Afx") {
        DiffSyncBtns := ["Button2", "Button86", "Button91", "Button103"]
        for idx, btn in DiffSyncBtns {
            try {
                t := ControlGetText(btn)
                if (t == " Sync with other exams" || t == "不同檢查同步 ") {
                    ControlClick(btn)
                    break
                }
            }
        }
    } else {
        Send("f")
    }
}

e::
{
    try {
        hCtl := ControlGetFocus("A")
        FocusedControl := ControlGetClassNN(hCtl)
    } catch {
        FocusedControl := ""
    }
    ;MsgBox(FocusedControl)

    OutputVar := WinGetTitle("A")
    if (OutputVar = "INFINITT PACS" && SubStr(FocusedControl, 1, 3) = "Afx") {
        DiffSyncBtns := ["Button4", "Button78"]
        for idx, btn in DiffSyncBtns {
            try {
                t := ControlGetText(btn)
                if (t = " Scout lines") { ; 注意這裡原代碼有一個前導空白
                    ControlClick(btn)
                    break
                }
            }
        }
    } else {
        Send("e")
    }
}

/*
;; activate RIS and insert exam name
!e::
{
    global FINDING_CONTROL ; 引用外部全域變數
    if (!WinActive("報告作業(frmRISReport)")) {
        ; 檢查視窗是否存在，避免報錯
        if (WinExist("報告作業(frmRISReport)")) {
            WinActivate("報告作業(frmRISReport)")
            WinWaitActive("報告作業(frmRISReport)")

            ; 確保 FINDING_CONTROL 已定義且控制項存在
            if IsSet(FINDING_CONTROL) {
                try ControlFocus(FINDING_CONTROL)
            }

            ; 假設 InsertExamname() 是一個已定義的函數
            try InsertExamname()
        }
    }
}
*/
#HotIf ; end for INFINITT PACS
