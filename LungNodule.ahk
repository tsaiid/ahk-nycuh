#Requires AutoHotkey v2.0
#Include <Acc.v2>
#Include <OCR.v2>

; ==============================================================================
; ★ 系統初始化
; ==============================================================================
DllCall("SetThreadDpiAwarenessContext", "ptr", -4, "ptr")
CoordMode("Mouse", "Screen")
CoordMode("ToolTip", "Screen")

global NoduleData := Map("RUL", [], "RML", [], "RLL", [], "LUL", [], "LLL", [])
global GuiX := 100, GuiY := 150, MyGui := ""
global COL_LEFT_X := 10, COL_RIGHT_X := 170, COL_WIDTH := 140

; ★ 定義兩個映射表
global MapPatternA := Map() ; 原始模式
global MapPatternB := Map() ; 新發現模式 (高效能 & 高機率)

GenerateMaps()
UpdateGUI()

; ==============================================================================
; ★ 映射表生成 (雙模式)
; ==============================================================================
GenerateMaps() {
    classPrefix := "Afx:00400000:b:00000000:00000013:00000000"

    ; --- 模式 A (低機率) ---
    ; Focus: 3... | Combo: 10... | Afx: 47...
    Loop 8 {
        i := A_Index - 1
        f := classPrefix . (3 + i)
        c := "ComboBox" . (10 + (i * 6))
        a := "AfxWnd140u" . (47 + (i * 3))
        MapPatternA[f] := {img: c, srs: a, type: "Pattern A"}
    }

    ; --- 模式 B (高機率 - 優先) ---
    ; Focus: 5... | Combo: 57... | Afx: 47...
    Loop 8 {
        i := A_Index - 1
        f := classPrefix . (5 + i)
        c := "ComboBox" . (57 + (i * 6))
        a := "AfxWnd140u" . (47 + (i * 3))
        MapPatternB[f] := {img: c, srs: a, type: "Pattern B"}
    }
}

; ==============================================================================
; ★ 快捷鍵區域
; ==============================================================================
#HotIf WinActive("ahk_exe G3PACS.exe")

; F11: 效能測試
F11::RunSmartBenchmark()

; F12: 探針工具 (Debug用)
F12::ProbeControl()

!q::CaptureNodule("RUL")
!a::CaptureNodule("RML")
!z::CaptureNodule("RLL")
!w::CaptureNodule("LUL")
!s::CaptureNodule("LLL")

+!q::DirectCopy("RUL")
+!a::DirectCopy("RML")
+!z::DirectCopy("RLL")
+!w::DirectCopy("LUL")
+!s::DirectCopy("LLL")
+!c::SimpleDirectCopy()

#HotIf

; ==============================================================================
; ★ 核心邏輯：智慧抓取 (Smart Capture) - 優先順序已調整
; ==============================================================================

GetSmartInfo() {
    ; 1. 嘗試極速方法 (ClassNN)
    fastInfo := GetInfo_ByProbe()
    if (fastInfo.valid) {
        return fastInfo
    }

    ; 2. 失敗則回退到慢速方法 (Acc)
    return GetNoduleInfoFromFocus()
}

/**
 * 探針函數：嘗試用 ClassNN 抓取
 * 優化：優先檢查 Pattern B
 */
GetInfo_ByProbe() {
    try {
        focusHwnd := ControlGetFocus("A")
        if (!focusHwnd) {
            return {srs: "", img: "", valid: false}
        }

        focusNN := ControlGetClassNN(focusHwnd)
        hwnd := WinActive("A")
        target := ""

        ; === 策略優化：優先檢查模式 B (高機率) ===
        if (MapPatternB.Has(focusNN)) {
            candidate := MapPatternB[focusNN]
            if (ProbeComboBox(candidate.img, hwnd)) {
                target := candidate
            }
        }

        ; === 若 B 失敗，檢查模式 A (低機率) ===
        if (target == "" && MapPatternA.Has(focusNN)) {
            candidate := MapPatternA[focusNN]
            if (ProbeComboBox(candidate.img, hwnd)) {
                target := candidate
            }
        }

        ; === 兩者皆空，放棄 ===
        if (target == "") {
            return {srs: "", img: "", valid: false}
        }

        ; === 命中目標，開始取值 ===

        ; 1. 取 Img
        imgVal := ControlGetText(target.img, hwnd)

        ; 2. 取 Srs (OCR)
        srsVal := ""
        try {
            ControlGetPos(&cX, &cY, &cW, &cH, target.srs, hwnd)

            ; 座標轉換 (Client -> Screen)
            pt := Buffer(8)
            NumPut("int", cX, pt, 0)
            NumPut("int", cY, pt, 4)
            DllCall("ClientToScreen", "ptr", hwnd, "ptr", pt)
            screenX := NumGet(pt, 0, "int")
            screenY := NumGet(pt, 4, "int")

            if (cW > 0 && cH > 0) {
                ocrResult := OCR.FromRect(screenX, screenY, cW, cH, {scale: 2})
                srsVal := ParseSrs(ocrResult.Text)
            }
        }

        return {srs: srsVal, img: imgVal, valid: true, method: target.type}

    } catch {
        return {srs: "", img: "", valid: false}
    }
}

ProbeComboBox(controlName, hwnd) {
    try {
        txt := ControlGetText(controlName, hwnd)
        return IsNumber(Trim(txt))
    } catch {
        return false
    }
}

ParseSrs(text) {
    splitText := StrSplit(text, ",")
    if (splitText.Length > 0) {
        if (RegExMatch(splitText[1], "(\d+)", &match)) {
            return match[1]
        }
    }
    return ""
}

; ==============================================================================
; ★ 應用層函數
; ==============================================================================

CaptureNodule(location) {
    try {
        info := GetSmartInfo()

        if (info.img == "" || info.srs == "") {
            ShowTip("⚠️ 抓取失敗 (請確認焦點)", 1500)
            return
        }

        ; 檢查重複
        For existingItem in NoduleData[location] {
            if (existingItem.srs == info.srs && existingItem.img == info.img) {
                ShowTip("⚠️ 已存在 (忽略)", 1000)
                return
            }
        }

        NoduleData[location].Push(info)
        UpdateGUI()

        methodTag := (info.HasOwnProp("method")) ? "[" info.method "] " : "[Acc] "
        ShowTip("✅ " methodTag location ": " info.srs "/" info.img, 1000)

    } catch Error as e {
        ShowTip("❌ " e.Message, 2000)
    }
}

DirectCopy(location) {
    try {
        info := GetSmartInfo()
        if (info.img == "" || info.srs == "") {
            ShowTip("⚠️ 抓取失敗", 2000)
            return
        }
        reportStr := location . " of lung (Srs/Img: " . info.srs . "/" . info.img . ")"
        A_Clipboard := reportStr
        ShowTip("📋 Copied:`n" reportStr, 2000)
    } catch Error as e {
        ShowTip("❌ " e.Message, 2000)
    }
}

SimpleDirectCopy() {
    try {
        info := GetSmartInfo()
        if (info.img == "" || info.srs == "") {
            ShowTip("⚠️ 抓取失敗", 2000)
            return
        }
        reportStr := "(Srs/Img: " . info.srs . "/" . info.img . ")"
        A_Clipboard := reportStr
        ShowTip("📋 Copied:`n" reportStr, 2000)
    } catch Error as e {
        ShowTip("❌ " e.Message, 2000)
    }
}

ShowTip(msg, duration) {
    ToolTip(msg)
    SetTimer(() => ToolTip(), -duration)
}

; ==============================================================================
; ★ 工具：F12 探針 (保留給您除錯用)
; ==============================================================================
ProbeControl() {
    MouseGetPos(,, &hwnd, &ctrlClassNN)

    try {
        txt := ControlGetText(ctrlClassNN, hwnd)
    } catch {
        txt := "無文字或無法讀取"
    }

    msg := "【探針資訊】`n"
    msg .= "ClassNN: " ctrlClassNN "`n"
    msg .= "Text內容: " txt "`n`n"
    msg .= "對應檢查：`n"
    msg .= "Pattern A (Focus->Combo): +1 -> +6`n"
    msg .= "Pattern B (Focus->Combo): +5 -> +57"

    MsgBox(msg)
}

; ==============================================================================
; ★ 效能測試 (F11)
; ==============================================================================
RunSmartBenchmark() {
    resultText := "★ 極速模式測試結果 (Priority: B) ★`n`n"

    ; 1. 測試新方法 (探針)
    start := A_TickCount
    newInfo := GetInfo_ByProbe()
    newTime := A_TickCount - start

    if (newInfo.valid) {
        resultText .= "🚀 新方法成功 (" newInfo.method ")`n"
        resultText .= "耗時: " newTime " ms`n"
        resultText .= "數值: Srs " newInfo.srs " / Img " newInfo.img "`n`n"
    } else {
        resultText .= "❌ 新方法失敗 (未命中任何模式)`n"
        resultText .= "耗時: " newTime " ms`n`n"
    }

    ; 2. 測試舊方法 (Acc) 對照
    start := A_TickCount
    oldInfo := GetNoduleInfoFromFocus()
    oldTime := A_TickCount - start

    resultText .= "🐢 舊方法 (Acc)`n"
    resultText .= "耗時: " oldTime " ms`n"
    resultText .= "數值: Srs " oldInfo.srs " / Img " oldInfo.img "`n`n"

    ; 結論
    if (newInfo.valid) {
        speedup := Round(oldTime / (newTime > 0 ? newTime : 1), 1)
        resultText .= "🏆 速度提升: " speedup " 倍!"
    }

    MsgBox(resultText)
}

; ==============================================================================
; ★ 舊方法保留 (Acc)
; ==============================================================================
GetNoduleInfoFromFocus() {
    try {
        focusedEl := Acc.ElementFromPoint()
        if (!focusedEl) {
            return {srs: "", img: ""}
        }

        fullPath := GetFullPath(focusedEl)
        pathParts := StrSplit(fullPath, ",")
        if (pathParts.Length < 2) {
            return {srs: "", img: ""}
        }

        targetIdx := pathParts.Length - 1

        ; Img
        pathParts[targetIdx] := Integer(pathParts[targetIdx]) + 1
        basePath := ""
        Loop targetIdx {
            basePath .= pathParts[A_Index] ","
        }
        imgPath := basePath . pathParts[pathParts.Length] . ",1,4,2,4"
        try {
            imgEl := Acc.GetRootElement()[imgPath]
            imgVal := Trim(imgEl.Value)
        } catch {
            imgVal := ""
        }

        ; Srs
        pathParts[targetIdx] := Integer(pathParts[targetIdx]) + 1
        basePath := ""
        Loop targetIdx {
            basePath .= pathParts[A_Index] ","
        }
        srsPath := basePath . pathParts[pathParts.Length]
        srsVal := ""
        try {
            srsEl := Acc.GetRootElement()[srsPath]
            loc := srsEl.Location
            if (loc.w > 0 && loc.h > 0) {
                ocrResult := OCR.FromRect(loc.x, loc.y, loc.w, loc.h, {scale: 2})
                srsVal := ParseSrs(ocrResult.Text)
            }
        }

        return {srs: srsVal, img: imgVal}
    } catch {
        return {srs: "", img: ""}
    }
}

GetFullPath(oEl) {
    path := ""
    curr := oEl
    try {
        while (curr.Parent) {
            p := curr.Parent
            for idx, child in p {
                if (child.IsEqual(curr)) {
                    path := idx "," path
                    break
                }
            }
            curr := p
        }
    }
    return Trim(path, ",")
}

; ==============================================================================
; ★ GUI 介面繪製
; ==============================================================================
UpdateGUI() {
    if (MyGui && WinExist("ahk_id " MyGui.Hwnd)) {
        WinGetPos(&currentX, &currentY,,, "ahk_id " MyGui.Hwnd)
        global GuiX := currentX
        global GuiY := currentY
        MyGui.Destroy()
    }
    global MyGui := Gui("+AlwaysOnTop +ToolWindow +Caption +Border", "Nodule Tracker")
    MyGui.SetFont("s10", "Segoe UI")
    MyGui.BackColor := "FFFFE0"
    MyGui.OnEvent("Close", WindowClosed)

    MyGui.SetFont("s11 Bold", "Segoe UI")
    MyGui.Add("Text", "w320 Center", "Nodule Tracker")

    MyGui.SetFont("s9 Norm", "Segoe UI")
    btnX := (320 - 220) / 2
    btnCopy := MyGui.Add("Button", "x" btnX " w100 h30", "Copy Report")
    btnCopy.OnEvent("Click", CopyReport)
    btnClear := MyGui.Add("Button", "x+20 w100 h30", "Clear All")
    btnClear.OnEvent("Click", ClearAll)
    MyGui.Add("Text", "x10 y+10 w300 h1 0x10")

    For key, arr in NoduleData {
        SortNoduleData(arr)
    }

    MyGui.SetFont("s11 Bold", "Segoe UI")
    MyGui.Add("Text", "Section x" COL_LEFT_X " y+10 w" COL_WIDTH " Center cBlue", "Right Lung")
    RenderSection("RUL", COL_LEFT_X)
    RenderSection("RML", COL_LEFT_X)
    RenderSection("RLL", COL_LEFT_X)

    MyGui.SetFont("s11 Bold", "Segoe UI")
    MyGui.Add("Text", "ys x" COL_RIGHT_X " w" COL_WIDTH " Center cBlue", "Left Lung")
    RenderSection("LUL", COL_RIGHT_X)

    if (NoduleData["RML"].Length > 0) {
        MyGui.Add("Text", "x" COL_RIGHT_X " y+28", "")
    } else {
        MyGui.Add("Text", "x" COL_RIGHT_X " y+5", "")
    }
    RenderSection("LLL", COL_RIGHT_X)

    MyGui.Show("x" GuiX " y" GuiY " NoActivate AutoSize")
}

RenderSection(label, xPos) {
    global MyGui
    MyGui.SetFont("s10 Bold", "Segoe UI")
    MyGui.Add("Text", "x" xPos " y+5 w" COL_WIDTH " Center c003366", label)
    items := NoduleData[label]
    if (items.Length == 0) {
        MyGui.SetFont("s9 Norm cGray", "Segoe UI")
        MyGui.Add("Text", "xp y+2 w" COL_WIDTH " Center", "-")
    } else {
        MyGui.SetFont("s10 Norm cDefault", "Segoe UI")
        For index, item in items {
            displayText := item.srs . "/" . item.img
            textX := xPos + (COL_WIDTH / 2) - 35
            MyGui.Add("Text", "x" textX " y+5 w50 Right", displayText)
            btnDel := MyGui.Add("Button", "x+5 yp-3 w20 h20", "x")
            btnDel.OnEvent("Click", DeleteItem.Bind(label, index))
        }
    }
}

DeleteItem(location, index, *) {
    NoduleData[location].RemoveAt(index)
    UpdateGUI()
}

ClearAll(*) {
    For key, arr in NoduleData {
        NoduleData[key] := []
    }
    UpdateGUI()
}

CopyReport(*) {
    reportParts := []
    global LobeOrder := ["RUL", "RML", "RLL", "LUL", "LLL"]
    For label in LobeOrder {
        items := NoduleData[label]
        if (items.Length > 0) {
            SortNoduleData(items)
            seriesMap := Map()
            For item in items {
                s := item.srs, i := item.img
                if (!seriesMap.Has(s)) {
                    seriesMap[s] := []
                }
                seriesMap[s].Push(i)
            }
            srsKeys := []
            For k, v in seriesMap {
                srsKeys.Push(k)
            }
            if (srsKeys.Length > 1) {
                Loop srsKeys.Length {
                    idx := A_Index
                    Loop srsKeys.Length - idx {
                        j := A_Index
                        if (Integer(srsKeys[j]) > Integer(srsKeys[j+1])) {
                            t := srsKeys[j], srsKeys[j] := srsKeys[j+1], srsKeys[j+1] := t
                        }
                    }
                }
            }
            lobeStr := ""
            For sKey in srsKeys {
                imgArr := seriesMap[sKey]
                imgStr := ""
                For val in imgArr {
                    imgStr .= val . ","
                }
                imgStr := Trim(imgStr, ",")
                lobeStr .= sKey . "/" . imgStr . "; "
            }
            lobeStr := Trim(lobeStr, "; ")
            reportParts.Push(label . " (Srs/Img: " . lobeStr . ")")
        }
    }
    if (reportParts.Length == 0) {
        ShowTip("! 無資料可複製", 2000)
        return
    }
    finalStr := ""
    if (reportParts.Length == 1) {
        finalStr := reportParts[1]
    } else if (reportParts.Length == 2) {
        finalStr := reportParts[1] . " and " . reportParts[2]
    } else {
        Loop reportParts.Length - 1 {
            finalStr .= reportParts[A_Index] . ", "
        }
        finalStr .= "and " . reportParts[reportParts.Length]
    }
    A_Clipboard := finalStr
    ShowTip("Copied:`n" finalStr, 3000)
}

SortNoduleData(arr) {
    if (arr.Length < 2) {
        return
    }
    Loop arr.Length {
        i := A_Index
        Loop arr.Length - i {
            j := A_Index
            s1 := Integer(arr[j].srs), s2 := Integer(arr[j+1].srs)
            i1 := Integer(arr[j].img), i2 := Integer(arr[j+1].img)
            if (s1 > s2) || (s1 == s2 && i1 > i2) {
                t := arr[j], arr[j] := arr[j+1], arr[j+1] := t
            }
        }
    }
}

WindowClosed(*) {
    try {
        if (MyGui && WinExist("ahk_id " MyGui.Hwnd)) {
            WinGetPos(&currentX, &currentY,,, "ahk_id " MyGui.Hwnd)
            global GuiX := currentX
            global GuiY := currentY
        }
    }
    For key, arr in NoduleData {
        NoduleData[key] := []
    }
}

; 視窗專用快捷鍵
#HotIf WinActive("Nodule Tracker ahk_class AutoHotkeyGUI")
^c::CopyReport()
Esc::ClearAll()
#HotIf