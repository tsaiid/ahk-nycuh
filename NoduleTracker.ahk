; ==============================================================================
; ★ Nodule Tracker (PACS Workflow Optimizer)
; ==============================================================================
; Description : 自動抓取 PACS 影像中的 Series 與 Image 編號並分類肺葉。
;               支援智慧探針 (ClassNN) 與 UI 自動化 (Acc) 雙模切換。
; Author      : Tsai, I-Ta (放射科醫師)
; GitHub      : tsaiid
; License     : MIT License
; Version     : 2026.01.30
;
; ------------------------------------------------------------------------------
; 【使用說明】
; 1. 環境需求：需安裝 AutoHotkey v2.0+，並包含 Acc.v2 與 OCR.v2 函式庫。
; 2. 基本操作：
;    - Alt + Q/A/Z : 抓取並記錄至 RUL / RML / RLL (右肺)
;    - Alt + W/S   : 抓取並記錄至 LUL / LLL (左肺)
; 3. 直接複製 (不存入 GUI)：
;    - Shift + Alt + Q/A/Z/W/S : 抓取單筆資訊並格式化複製 (含肺葉名稱)
;    - Shift + Alt + C         : 僅複製 (Srs/Img) 格式字串
; 4. 輔助工具：
;    - F11         : 執行效能基準測試 (Benchmark)，驗證 Pattern B 命中率。
;    - F12         : 開啟探針工具，查看當前視窗控制項 ClassNN (除錯用)。
; 5. GUI 功能：
;    - Copy Report : 整合所有記錄並按肺葉順序自動排序複製到剪貼簿。
;    - Clear All   : 清空當前所有結節紀錄。
; ------------------------------------------------------------------------------

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
; ★ 映射表生成 (擴充至 Pattern D)
; ==============================================================================
GenerateMaps() {
    classPrefix := "Afx:00400000:b:00000000:00000013:00000000"

    ; 初始化所有 Map
    global MapPatternA := Map(), MapPatternB := Map()
    global MapPatternC := Map(), MapPatternD := Map()

    Loop 8 {
        i := A_Index - 1

        ; --- Pattern A ---
        fA := classPrefix . (3 + i)
        cA := "ComboBox" . (10 + (i * 6))
        aA := "AfxWnd140u" . (47 + (i * 3))
        MapPatternA[fA] := {img: cA, srs: aA, type: "Pattern A"}

        ; --- Pattern B ---
        fB := classPrefix . (5 + i)
        cB := "ComboBox" . (57 + (i * 6))
        aB := "AfxWnd140u" . (47 + (i * 3))
        MapPatternB[fB] := {img: cB, srs: aB, type: "Pattern B"}

        ; --- Pattern C (新) ---
        ; Focus: ...3 | Combo: 10... | Afx: 31...
        fC := classPrefix . (3 + i)
        cC := "ComboBox" . (10 + (i * 6))
        aC := "AfxWnd140u" . (31 + (i * 3))
        MapPatternC[fC] := {img: cC, srs: aC, type: "Pattern C"}

        ; --- Pattern D (新) ---
        ; Focus: ...3 | Combo: 10... | Afx: 37...
        fD := classPrefix . (3 + i)
        cD := "ComboBox" . (10 + (i * 6))
        aD := "AfxWnd140u" . (37 + (i * 3))
        MapPatternD[fD] := {img: cD, srs: aD, type: "Pattern D"}
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

; 使用大括號語法 (AHK v2)，並加入 vkE8 阻斷語言切換偵測

+!q:: {
    Send("{Blind}{vkE8}")
    DirectCopy("RUL")
}

+!a:: {
    Send("{Blind}{vkE8}")
    DirectCopy("RML")
}

+!z:: {
    Send("{Blind}{vkE8}")
    DirectCopy("RLL")
}

+!w:: {
    Send("{Blind}{vkE8}")
    DirectCopy("LUL")
}

+!s:: {
    Send("{Blind}{vkE8}")
    DirectCopy("LLL")
}

+!c:: {
    Send("{Blind}{vkE8}")
    SimpleDirectCopy()
}

#HotIf

; ==============================================================================
; ★ 修改核心邏輯：加入詳細錯誤診斷 (Diagnostic Capture)
; ==============================================================================

GetSmartInfo() {
    ; 1. 嘗試極速方法 (ClassNN)
    fastInfo := GetInfo_ByProbe()
    if (fastInfo.valid) {
        return fastInfo
    }

    ; 保存 Probe 的錯誤原因，如果 Acc 也失敗，可以參考
    lastError := fastInfo.HasOwnProp("error") ? fastInfo.error : "Unknown Probe Error"

    ; 2. 失敗則回退到慢速方法 (Acc)
    accInfo := GetNoduleInfoFromFocus()
    if (accInfo.srs != "" && accInfo.img != "") {
        accInfo.valid := true
        accInfo.method := "Acc"
        return accInfo
    }

    ; 3. 全部失敗，回傳最後的錯誤原因
    return {srs: "", img: "", valid: false, error: lastError}
}

; ==============================================================================
; ★ 核心邏輯：精確 Pattern 判定 (三重驗證：ComboBox -> VMTool 標籤 -> OCR)
; ==============================================================================
GetInfo_ByProbe() {
    try {
        focusHwnd := ControlGetFocus("A")
        if (!focusHwnd) {
            return {srs: "", img: "", valid: false, error: "無法取得焦點控制項"}
        }

        focusNN := ControlGetClassNN(focusHwnd)
        hwnd := WinActive("A")

        ; 依序測試各個模式映射表
        patternList := [MapPatternA, MapPatternB, MapPatternC, MapPatternD]

        for pMap in patternList {
            if (pMap.Has(focusNN)) {
                candidate := pMap[focusNN]

                ; --- 第一重驗證：ComboBox 必須有數字 (Image No.) ---
                imgVal := ""
                try {
                    imgVal := ControlGetText(candidate.img, hwnd)
                }
                if (!IsNumber(Trim(imgVal))) {
                    continue
                }

                ; --- 第二重驗證：Series 控制項的 Text 必須包含 "VMTool" ---
                ; 這是極速過濾，避免對非 Series 區域執行 OCR
                try {
                    ctrlText := ControlGetText(candidate.srs, hwnd)
                    if (!InStr(ctrlText, "VMTool")) {
                        continue ; Text 不符，表示這不是我們要找的 Series 區域
                    }
                } catch {
                    continue ; 無法取得 Text 亦跳過
                }

                ; --- 第三重驗證：OCR 實際解析 Series 編號 ---
                srsVal := ""
                try {
                    ControlGetPos(&cX, &cY, &cW, &cH, candidate.srs, hwnd)

                    ; 座標轉換 (Client -> Screen)
                    pt := Buffer(8), NumPut("int", cX, pt, 0), NumPut("int", cY, pt, 4)
                    DllCall("ClientToScreen", "ptr", hwnd, "ptr", pt)
                    screenX := NumGet(pt, 0, "int"), screenY := NumGet(pt, 4, "int")

                    if (cW > 0 && cH > 0) {
                        ocrResult := OCR.FromRect(screenX, screenY, cW, cH, {scale: 2})
                        srsVal := ParseSrs(ocrResult.Text)
                    }
                }

                ; 只有當前兩重文字驗證與最終 OCR 都通過時，才視為命中
                if (srsVal != "") {
                    return {srs: srsVal, img: imgVal, valid: true, method: candidate.type}
                }
            }
        }

        return {srs: "", img: "", valid: false, error: "所有模式驗證失敗 (標籤不符或 OCR 無效)"}

    } catch Error as e {
        return {srs: "", img: "", valid: false, error: "Probe Runtime Error: " . e.Message}
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
; ★ 修改應用層函數：顯示具體錯誤
; ==============================================================================

CaptureNodule(location) {
    try {
        info := GetSmartInfo()

        if (!info.valid) {
            errLog := info.HasOwnProp("error") ? info.error : "未知錯誤"
            ShowTip("⚠️ " . errLog, 2500)
            return
        }

        ; 檢查重複 (保留原有邏輯)
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
        ShowTip("❌ Critical: " e.Message, 3000)
    }
}

; 修改 DirectCopy 以支援錯誤詳解
DirectCopy(location) {
    info := GetSmartInfo()
    if (!info.valid) {
        ShowTip("⚠️ 複製失敗: " . (info.HasOwnProp("error") ? info.error : "無法抓取"), 2500)
        return
    }
    reportStr := location . " of lung (Srs/Img: " . info.srs . "/" . info.img . ")"
    A_Clipboard := reportStr
    ShowTip("📋 Copied:`n" reportStr, 2000)
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
; ★ 工具：F12 探針 (更新顯示資訊)
; ==============================================================================
ProbeControl() {
    MouseGetPos(,, &hwnd, &ctrlClassNN)
    ; ... (前面邏輯不變)
    msg := "【探針資訊】`nClassNN: " ctrlClassNN "`n`n模式檢查：`n"
    msg .= "Pattern B: " . (MapPatternB.Has(ctrlClassNN) ? "✅" : "❌") . "`n"
    msg .= "Pattern C: " . (MapPatternC.Has(ctrlClassNN) ? "✅" : "❌") . "`n"
    msg .= "Pattern A: " . (MapPatternA.Has(ctrlClassNN) ? "✅" : "❌")
    MsgBox(msg)
}

; ==============================================================================
; ★ 效能測試 (F11) - 支援 A/B/C/D 全模式診斷
; ==============================================================================
RunSmartBenchmark() {
    resultText := "★ 智慧抓取效能測試 (Pattern A-D) ★`n`n"
    hwnd := WinActive("A")
    focusHwnd := ControlGetFocus("A")

    if (!focusHwnd) {
        MsgBox("❌ 測試失敗：無法取得視窗焦點。")
        return
    }

    focusNN := ControlGetClassNN(focusHwnd)
    resultText .= "當前焦點: " . focusNN . "`n" . "----------------------------------`n"

    ; 1. 測試極速模式 (Probe)
    startProbe := A_TickCount
    info := GetInfo_ByProbe() ; 這裡會跑完 A->B->C->D 邏輯
    probeTime := A_TickCount - startProbe

    if (info.valid) {
        resultText .= "🚀 極速模式成功！`n"
        resultText .= "命中模式: " . info.method . "`n"
        resultText .= "耗時: " . probeTime . " ms`n"
        resultText .= "數值: Srs " . info.srs . " / Img " . info.img . "`n`n"
    } else {
        errReason := info.HasOwnProp("error") ? info.error : "未知原因"
        resultText .= "❌ 極速模式失敗`n"
        resultText .= "原因: " . errReason . "`n`n"
    }

    ; 2. 測試舊方法 (Acc) 作為基準
    startAcc := A_TickCount
    accInfo := GetNoduleInfoFromFocus()
    accTime := A_TickCount - startAcc

    resultText .= "🐢 基準方法 (Acc)`n"
    resultText .= "耗時: " . accTime . " ms`n"
    resultText .= "數值: Srs " . (accInfo.srs ? accInfo.srs : "N/A") . " / Img " . (accInfo.img ? accInfo.img : "N/A") . "`n"
    resultText .= "----------------------------------`n"

    ; 3. 效能總結
    if (info.valid) {
        speedup := Round(accTime / (probeTime > 0 ? probeTime : 1), 1)
        resultText .= "🏆 速度提升: " . speedup . " 倍"
    } else {
        resultText .= "💡 建議：請檢查 F12 探針資訊以修正 Pattern 映射表。"
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