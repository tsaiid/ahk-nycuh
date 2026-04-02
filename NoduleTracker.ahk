; ==============================================================================
; ★ Nodule Tracker (PACS Workflow Optimizer)
; ==============================================================================
; Description : 自動抓取 PACS 影像中的 Series 與 Image 編號並分類肺葉。
;               支援智慧探針 (ClassNN) 與 UI 自動化 (Acc) 雙模切換。
;               具備 OCR 驗證、命中率動態排序及 Cor/Sag 影像自動補償。
; Author      : Tsai, I-Ta (放射科醫師)
; GitHub      : tsaiid
; License     : MIT License
; Version     : 2026.04.01
;
; ------------------------------------------------------------------------------
; 【使用說明】
; 1. 環境需求：需安裝 AutoHotkey v2.0+，並包含 Acc.v2 與 OCR.v2 函式庫。
; 2. 基本操作 (存入 GUI)：
;    - Alt + Q/A/Z : 抓取並記錄至 RUL / RML / RLL (右肺)
;    - Alt + W/S   : 抓取並記錄至 LUL / LLL (左肺)
; 3. 直接複製 (不存入 GUI)：
;    - Shift + Alt + Q/A/Z/W/S : 抓取單筆並帶入肺葉名稱 (e.g., RUL of lung...)
;    - m                       : 僅抓取並格式化為 (Srs/Img: 3/69)
; 4. 輔助工具：
;    - Ctrl + G    : 快速跳至指定影像編號 (透過定位 PACS 控制項寫入)
;    - F11         : 執行效能基準測試 (Benchmark)，驗證探針模式命中率。
;    - F12         : 開啟探針工具，深度除錯當前控制項 ClassNN 與 Acc 路徑。
; 5. GUI 功能：
;    - Full        : 整合所有記錄並按肺葉順序自動排序複製 (標準報告格式)。
;    - Loc:Img     : 依肺葉分類複製影像編號 (e.g., RUL:15;22 LUL:10)。
;    - Img No      : 提取所有 Image Number，排序並以分號分隔複製。
; 6. 特殊邏輯：
;    - 自動補償：偵測到 Coronal 或 Sagittal 序列時，影像編號自動 +1 (適配 PACS)。
;    - 動態排序：依據歷史命中率自動調整探針順序，提升抓取速度。
; ------------------------------------------------------------------------------

#Requires AutoHotkey v2.0
#Include <Acc.v2>
#Include <OCR.v2>

; ==============================================================================
; ★ 系統初始化
; ==============================================================================
; [新增] 設定工作列 (Tray) 上的 Icon
TraySetIcon(A_ScriptDir "\assets\NoduleTracker_icon.png")

DllCall("SetThreadDpiAwarenessContext", "ptr", -4, "ptr")
CoordMode("Mouse", "Screen")
CoordMode("ToolTip", "Screen")

global NoduleData := Map("RUL", [], "RML", [], "RLL", [], "LUL", [], "LLL", [])
global LobeOrder := ["RUL", "RML", "RLL", "LUL", "LLL"]
global GuiX := 100, GuiY := 150, MyGui := ""
global COL_LEFT_X := 10, COL_RIGHT_X := 170, COL_WIDTH := 140

; ★ 新增：Acc 慢速模式的總開關 (強烈建議設為 false，徹底杜絕卡死)
global EnableAccFallback := true

; ★ 新增：OCR 視覺化與文字除錯開關 (發布或實戰時請改為 false)
global DebugOCR := false

; ★ 新增：Ctrl+G 除錯視窗開關
global ShowDebugQuickSet := true

global GuiStatusMsg := "Ready"
global txtStatus := ""

; ★ 將分散的 Map 改為統一的 PatternList 陣列統一管理
global PatternList := []
global StatsFile := A_ScriptDir "\PatternStats.ini" ; ★ 新增：統計數據儲存路徑

GenerateMaps()
OptimizePatternOrder() ; ★ 新增：初始化時依照歷史命中率重新排序 PatternList
UpdateGUI()

; ==============================================================================
; ★ 映射表生成 (極簡陣列配置)
; ==============================================================================
GenerateMaps() {
    classPrefix := "Afx:00400000:b:00000000:00000013:00000000"
    global PatternList := []

    ; ★ 極簡配置：只需依序填入 [Focus起始, Combo起始, Afx起始, Desc起始]
    configs := [
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
        [5, 57, 29, 185],
        [5, 57, 31, 101],
        [5, 57, 39, 101],
        [5, 57, 41, 101],
        [5, 57, 45, 101],
        [5, 57, 47, 101],
    ]

    for idx, cfg in configs {
        pMap := Map()
        ; ★ 核心修正：使用配置參數組成唯一 ID (例如 Pattern_5_57_31_101)
        pName := "Pattern_" . cfg[1] . "_" . cfg[2] . "_" . cfg[3] . "_" . cfg[4]

        Loop 8 {
            i := A_Index - 1
            f := classPrefix . (cfg[1] + i)
            c := "ComboBox" . (cfg[2] + (i * 6))
            a := "AfxWnd140u" . (cfg[3] + (i * 3))
            d := "Button" . (cfg[4] + (i * 5))
            pMap[f] := {img: c, srs: a, desc: d, type: pName}
        }

        ; 將生成好的 Map 及其自動生成的名稱推入全域列表
        PatternList.Push({name: pName, map: pMap})
    }
}

; ==============================================================================
; ★ 快捷鍵區域
; ==============================================================================
#HotIf WinActive("ahk_exe G3PACS.exe") && WinActive("INFINITT PACS")

; F11: 效能測試
F11::RunSmartBenchmark()

; F12: 探針工具 (Debug用)
F12::ProbeControl()

!q::CaptureNodule("RUL")
!a::CaptureNodule("RML")
!z::CaptureNodule("RLL")
!w::CaptureNodule("LUL")
!s::CaptureNodule("LLL")

; ★ 新增：Ctrl+G 快速跳至指定影像編號 (Image Number)
^g::QuickSetImage()

; 使用大括號語法 (AHK v2)，並加入 vkE8 阻斷語言切換偵測

+!q:: {
    Send("{Blind}{vkE8}")
    DirectCopy("RUL")
    Send("{Blind}{vkE8}")
}

+!a:: {
    Send("{Blind}{vkE8}")
    DirectCopy("RML")
    Send("{Blind}{vkE8}")
}

+!z:: {
    Send("{Blind}{vkE8}")
    DirectCopy("RLL")
    Send("{Blind}{vkE8}")
}

+!w:: {
    Send("{Blind}{vkE8}")
    DirectCopy("LUL")
    Send("{Blind}{vkE8}")
}

+!s:: {
    Send("{Blind}{vkE8}")
    DirectCopy("LLL")
    Send("{Blind}{vkE8}")
}

m:: {
    ;Send("{Blind}{vkE8}")
    SimpleDirectCopy()
}

#HotIf

; ==============================================================================
; ★ 修改核心邏輯：移除過渡提示，僅回傳最終診斷結果
; ==============================================================================
GetSmartInfo() {
    fastInfo := GetInfo_ByProbe()
    if (fastInfo.valid) {
        return fastInfo
    }

    lastError := fastInfo.HasOwnProp("error") ? fastInfo.error : "Unknown Probe Error"

    if (!EnableAccFallback) {
        return {srs: "", img: "", valid: false, error: "Probe 失敗 (" . lastError . ")，已停用 Acc"}
    }

    ; 直接執行 Acc 備用方案 (不顯示過渡狀態)
    accInfo := GetNoduleInfoFromFocus()
    if (accInfo.srs != "" && accInfo.img != "") {
        accInfo.valid := true
        accInfo.method := "Acc"
        return accInfo
    }

    ; 若連 Acc 也失敗，將兩者的錯誤訊息合併回傳
    accError := accInfo.HasOwnProp("error") ? accInfo.error : "未知 Acc 錯誤"
    return {srs: "", img: "", valid: false, error: "Probe 失敗 (" . lastError . ")，Acc 也失敗 (" . accError . ")"}
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

        for patternData in PatternList {
            pMap := patternData.map

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
                try {
                    ctrlText := ControlGetText(candidate.srs, hwnd)
                    if (!InStr(ctrlText, "VMTool")) {
                        continue
                    }

                    if (DebugOCR) {
                        ControlGetPos(&cX, &cY, &cW, &cH, candidate.srs, hwnd)
                        ptC := Buffer(8), NumPut("int", cX, ptC, 0), NumPut("int", cY, ptC, 4)
                        DllCall("ClientToScreen", "ptr", hwnd, "ptr", ptC)
                        srsX := NumGet(ptC, 0, "int"), srsY := NumGet(ptC, 4, "int")

                        ControlGetPos(&fX, &fY, &fW, &fH, focusNN, hwnd)
                        ptF := Buffer(8), NumPut("int", fX, ptF, 0), NumPut("int", fY, ptF, 4)
                        DllCall("ClientToScreen", "ptr", hwnd, "ptr", ptF)
                        focX := NumGet(ptF, 0, "int"), focY := NumGet(ptF, 4, "int")

                        ShowDebugRects([
                            {x: srsX, y: srsY, w: cW, h: cH, color: "Red"},
                            {x: focX, y: focY, w: fW, h: fH, color: "Green"}
                        ])
                    }

                    if (!VerifySpatialMatch(candidate.srs, focusNN, hwnd)) {
                        continue
                    }
                } catch {
                    continue
                }

                ; --- 第三重驗證：OCR 實際解析 Series 編號 ---
                srsVal := ""
                try {
                    ControlGetPos(&cX, &cY, &cW, &cH, candidate.srs, hwnd)
                    pt := Buffer(8), NumPut("int", cX, pt, 0), NumPut("int", cY, pt, 4)
                    DllCall("ClientToScreen", "ptr", hwnd, "ptr", pt)
                    screenX := NumGet(pt, 0, "int"), screenY := NumGet(pt, 4, "int")

                    if (cW > 0 && cH > 0) {
                        ; 限制掃描最大寬度為 150 px
                        scanW := (cW > 150) ? 150 : cW

                        ; ★ 改用帶有半透明紅框濾鏡的 OCR 函數
                        ocrResult := CaptureOcrWithFilter(screenX, screenY, scanW, cH, 2)
                        srsVal := ParseSrs(ocrResult.Text)

                        if (DebugOCR) {
                            ShowDebugWindow("【OCR 除錯資訊】`n`nScreen X: " screenX "`nScreen Y: " screenY "`nWidth: " scanW "`nHeight: " cH "`n`n[原始 OCR 抓取文字]:`n" ocrResult.Text "`n`n[ParseSrs 解析結果]: " srsVal "`n`n[幾何驗證]: 通過", "Debug OCR", 5)
                        }
                    }
                }

                ; ★ 修改：只有當前兩重文字驗證與最終 OCR 都通過時，才視為命中並記錄數據
                if (srsVal != "") {
                    ; 嘗試抓取 Probe 模式下的 Series Description (新增)
                    descVal := ""
                    try {
                        descVal := ControlGetText(candidate.desc, hwnd)
                    }

                    ; ★ 新增：如果 desc 包含 cor 或 sag (不分大小寫)，且非 MRI 關鍵字 (t1, t2, dwi, adc)，則 img + 1
                    if (descVal != "" && RegExMatch(descVal, "i)cor|sag") && !RegExMatch(descVal, "i)t1|t2|dwi|adc|dual|stir|fl2d|pd")) {
                        if (IsNumber(imgVal)) {
                            imgVal := String(Integer(imgVal) + 1)
                        }
                    }

                    RecordPatternHit(patternData.name) ; ★ 新增：寫入命中統計
                    return {srs: srsVal, img: imgVal, desc: descVal, valid: true, method: candidate.type}
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
; ★ 修改應用層函數：精準分流狀態列 (GUI) 與浮動提示 (ToolTip)
; ==============================================================================
CaptureNodule(location) {
    try {
        info := GetSmartInfo()

        ; 狀態 1：Probe 與 Acc 雙雙失敗 -> 顯示於 GUI 狀態列
        if (!info.valid) {
            errLog := info.HasOwnProp("error") ? info.error : "未知錯誤"
            global GuiStatusMsg := "❌ " . errLog
            UpdateGUI()
            return
        }

        ; 檢查重複：恢復使用 ToolTip 顯示，不干擾 GUI
        For existingItem in NoduleData[location] {
            if (existingItem.srs == info.srs && existingItem.img == info.img) {
                ShowTip("⚠️ 已存在 (忽略)", 1000)
                return
            }
        }

        ; 加入資料陣列
        NoduleData[location].Push(info)

        ; 狀態分流邏輯
        if (info.HasOwnProp("method") && info.method == "Acc") {
            ; 狀態 2：Probe 失敗但 Acc 成功 -> 顯示於 GUI 狀態列 (常駐警告)
            global GuiStatusMsg := "⚠️ Probe 失敗，Acc 抓取成功 (建議按 F12 擴充)"
            UpdateGUI()
        } else {
            ; 狀態 3：Probe 完美成功 -> 清空 GUI 狀態列，並透過 ToolTip 顯示
            global GuiStatusMsg := ""
            UpdateGUI()

            methodTag := (info.HasOwnProp("method")) ? "[" info.method "] " : ""
            ShowTip("✅ " methodTag location ": " info.srs "/" info.img, 1000)
        }

    } catch Error as e {
        global GuiStatusMsg := "❌ Critical: " e.Message
        UpdateGUI()
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
        if (!info.valid) {
            ShowTip("⚠️ 抓取失敗: " . (info.HasOwnProp("error") ? info.error : "無法抓取"), 2000)
            return
        }

        clipStr := A_Clipboard
        entries := []

        ; 解析現有內容 (支援 3/69; 3/77 或 3/69,77 格式)
        if (RegExMatch(clipStr, "i)^\s*\(Srs/Img:\s*(.+)\)\s*$", &match)) {
            existingEntries := match[1]
            Loop Parse, existingEntries, ";", " " {
                if (RegExMatch(A_LoopField, "(\d+)/([\d,]+)", &m)) {
                    srs := m[1]
                    imgs := StrSplit(m[2], ",")
                    for img in imgs {
                        ; 檢查重複
                        isDup := false
                        for e in entries {
                            if (e.srs == srs && e.img == img) {
                                isDup := true
                                break
                            }
                        }
                        if (!isDup)
                            entries.Push({srs: srs, img: img})
                    }
                }
            }
        }

        ; 加入本次新抓取的資訊 (避免重複)
        isNewDup := false
        for e in entries {
            if (e.srs == info.srs && e.img == info.img) {
                isNewDup := true
                break
            }
        }
        if (!isNewDup)
            entries.Push({srs: info.srs, img: info.img})

        ; 排序 (先 Srs 後 Img)
        if (entries.Length > 1) {
            Loop entries.Length {
                i := A_Index
                Loop entries.Length - i {
                    j := A_Index
                    s1 := Integer(entries[j].srs), s2 := Integer(entries[j+1].srs)
                    i1 := Integer(entries[j].img), i2 := Integer(entries[j+1].img)
                    if (s1 > s2) || (s1 == s2 && i1 > i2) {
                        t := entries[j], entries[j] := entries[j+1], entries[j+1] := t
                    }
                }
            }
        }

        ; 分組重組字串
        groupedStr := ""
        currentSrs := ""
        for e in entries {
            if (e.srs != currentSrs) {
                if (groupedStr != "")
                    groupedStr .= "; "
                groupedStr .= e.srs . "/" . e.img
                currentSrs := e.srs
            } else {
                groupedStr .= "," . e.img
            }
        }

        reportStr := "(Srs/Img: " . groupedStr . ")"
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
; ★ 工具：F12 探針 (動態顯示資訊 - 新增 Acc 抓取文字與 OCR 原始結果除錯)
; ==============================================================================
ProbeControl() {
    MouseGetPos(,, &hwnd, &ctrlClassNN)
    focusHwnd := ControlGetFocus(hwnd)
    if (!focusHwnd) {
        focusHwnd := hwnd
    }

    msg := "【Pattern 探針資訊】`n當前指向 ClassNN: " ctrlClassNN "`n`n"
    matchFound := false

    ; ★ 動態驗證所有 Pattern
    for patternData in PatternList {
        pMap := patternData.map

        if (pMap.Has(ctrlClassNN)) {
            candidate := pMap[ctrlClassNN]

            imgVal := ""
            try {
                imgVal := ControlGetText(candidate.img, hwnd)
            }

            srsText := ""
            try {
                srsText := ControlGetText(candidate.srs, hwnd)
            }

            if (IsNumber(Trim(imgVal)) && InStr(srsText, "VMTool")) {

                if (DebugOCR) {
                    ControlGetPos(&cX, &cY, &cW, &cH, candidate.srs, hwnd)
                    ptC := Buffer(8), NumPut("int", cX, ptC, 0), NumPut("int", cY, ptC, 4)
                    DllCall("ClientToScreen", "ptr", hwnd, "ptr", ptC)
                    srsX := NumGet(ptC, 0, "int"), srsY := NumGet(ptC, 4, "int")

                    ControlGetPos(&fX, &fY, &fW, &fH, ctrlClassNN, hwnd)
                    ptF := Buffer(8), NumPut("int", fX, ptF, 0), NumPut("int", fY, ptF, 4)
                    DllCall("ClientToScreen", "ptr", hwnd, "ptr", ptF)
                    focX := NumGet(ptF, 0, "int"), focY := NumGet(ptF, 4, "int")

                    ShowDebugRects([
                        {x: srsX, y: srsY, w: cW, h: cH, color: "Red"},
                        {x: focX, y: focY, w: fW, h: fH, color: "Green"}
                    ])
                }

                spatialPass := VerifySpatialMatch(candidate.srs, ctrlClassNN, hwnd)

                if (DebugOCR) {
                     msg .= "【幾何驗證結果】: " (spatialPass ? "✅ 通過" : "❌ 失敗") "`n------------------`n"
                }

                if (!spatialPass) {
                    continue
                }

                msg .= "🎯 實際命中: " patternData.name "`n"
                msg .= "  - Img 控制項: " candidate.img " (數值: " Trim(imgVal) ")`n"
                msg .= "  - Srs 控制項: " candidate.srs " (標籤吻合，且通過空間驗證)`n"

                ; 顯示 Probe 抓到的 Desc (新增)
                try {
                    dText := ControlGetText(candidate.desc, hwnd)
                    msg .= "  - Desc 控制項: " candidate.desc " (數值: " Trim(dText) ")`n"
                } catch {
                    msg .= "  - Desc 控制項: " candidate.desc " (無法讀取)`n"
                }

                matchFound := true
                break
            }
        }
    }

    if (!matchFound) {
        msg .= "❌ 狀態：未命中任何完整 Pattern 規則。`n"
    }

    ; ==============================================================================
    ; ★ Acc 備用方案深度除錯 (新增文字解析顯示)
    ; ==============================================================================
    msg .= "`n====================`n【Acc 備用方案除錯】`n"

    try {
        pacsRoot := Acc.ElementFromHandle(hwnd)
        ControlGetPos(,, &cW, &cH, focusHwnd, "ahk_id " hwnd)

        pt := Buffer(8, 0)
        DllCall("ClientToScreen", "ptr", focusHwnd, "ptr", pt)
        targetX := NumGet(pt, 0, "int") + (cW // 2)
        targetY := NumGet(pt, 4, "int") + (cH // 2)

        msg .= "- 中心座標: X" targetX ", Y" targetY "`n"

        focusedEl := Acc.ElementFromPoint(targetX, targetY)
        if (!focusedEl) {
            msg .= "❌ 無法從該座標取得 Acc 節點`n"
        } else {
            fullPath := GetRelativePath(focusedEl, pacsRoot)

            if (fullPath == "") {
                msg .= "❌ 無法解析相對路徑 (越界或超時)`n"
            } else {
                msg .= "- 原始相依路徑: " fullPath "`n"

                pathParts := StrSplit(fullPath, ",")
                if (pathParts.Length >= 2) {
                    targetIdx := pathParts.Length - 1

                    ; --- 計算 Img 路徑 (Index + 1) ---
                    pathParts[targetIdx] := Integer(pathParts[targetIdx]) + 1
                    basePathImg := ""
                    Loop targetIdx {
                        basePathImg .= pathParts[A_Index] ","
                    }
                    imgPath := basePathImg . pathParts[pathParts.Length] . ",1,4,2,4"

                    ; --- 計算 Srs 路徑 (Index 再 + 1) ---
                    pathParts[targetIdx] := Integer(pathParts[targetIdx]) + 1
                    basePathSrs := ""
                    Loop targetIdx {
                        basePathSrs .= pathParts[A_Index] ","
                    }
                    srsPath := basePathSrs . pathParts[pathParts.Length]

                    try {
                        imgEl := pacsRoot[imgPath]
                        msg .= "   ✅ 抓取 Img 值: [" Trim(imgEl.Value) "]`n"
                    } catch {
                        msg .= "   ❌ 預測的 Img 路徑無效`n"
                    }

                    try {
                        srsEl := pacsRoot[srsPath]
                        loc := srsEl.Location
                        msg .= "   ✅ Srs 座標框: W" loc.w " H" loc.h "`n"

                        ; ★ 新增：顯示 Series Description (descVal)
                        try {
                            descPath := srsPath . ",2,4"
                            descEl := pacsRoot[descPath]
                            rawDesc := descEl.Value ? descEl.Value : descEl.Name
                            if (rawDesc != "")
                                msg .= "   📝 Series Desc: [" rawDesc "]`n"
                        }

                        ; ★ 新增：顯示底層到底是抓到什麼字！
                        rawText := ""
                        try {
                            rawText := srsEl.Name
                        }
                        if (rawText == "") {
                            try {
                                rawText := srsEl.Value
                            }
                        }

                        if (rawText != "") {
                            msg .= "   🔍 Acc 屬性文字: [" rawText "]`n"
                            msg .= "   🎯 解析結果: [" ParseSrs(rawText) "]`n"
                        } else {
                            msg .= "   ⚠️ Acc 無內建文字，啟動 OCR...`n"
                            if (loc.w > 0 && loc.h > 0) {
                                scanW := (loc.w > 150) ? 150 : loc.w
                                ocrResult := OCR.FromRect(loc.x, loc.y, scanW, loc.h, {scale: 2})
                                ; 避免文字太多換行破壞版面，替換為空格
                                safeText := StrReplace(ocrResult.Text, "`n", " ")
                                msg .= "   🔍 OCR 原始文字: [" safeText "]`n"
                                msg .= "   🎯 解析結果: [" ParseSrs(ocrResult.Text) "]`n"
                            }
                        }

                    } catch {
                        msg .= "   ❌ 預測的 Srs 路徑無效`n"
                    }
                } else {
                    msg .= "❌ 路徑太短，無法計算偏移`n"
                }
            }
        }
    } catch Error as e {
        msg .= "❌ Acc 發生執行期錯誤: " e.Message "`n"
    }

    ShowDebugWindow(Trim(msg, "`n"), "Pattern 探針資訊")
}

; ==============================================================================
; ★ 效能測試 (F11) - 支援 A/B/C/D 全模式診斷
; ==============================================================================
RunSmartBenchmark() {
    resultText := "★ 智慧抓取效能測試 (Pattern A-D) ★`n`n"
    hwnd := WinActive("A")
    focusHwnd := ControlGetFocus("A")

    if (!focusHwnd) {
        ShowDebugWindow("❌ 測試失敗：無法取得視窗焦點。", "效能測試錯誤")
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

    ShowDebugWindow(resultText, "智慧抓取效能測試")
}

; ==============================================================================
; ★ 方案 B 核心：限制 Acc 掃描範圍於當前視窗 (防止 VS Code/系統卡死)
; ★ 修正：限制 Srs OCR 掃描寬度，避免讀取到全螢幕雜訊
; ==============================================================================
GetNoduleInfoFromFocus() {
    try {
        ; 1. 取得當前活動視窗 (PACS) 的 Handle 與 Acc 根節點
        pacsHwnd := WinActive("A")
        if (!pacsHwnd) {
            return {srs: "", img: "", error: "Acc: 找不到活動視窗"}
        }
        pacsRoot := Acc.ElementFromHandle(pacsHwnd)

        ; 2. 取得當前焦點控制項
        focusHwnd := ControlGetFocus("ahk_id " pacsHwnd)
        if (!focusHwnd) {
            return {srs: "", img: "", error: "Acc: 無法取得焦點控制項"}
        }

        ; 3. 精準取得控制項在螢幕上的絕對中心點座標 (Screen 座標)
        pt := Buffer(8), NumPut("int", 0, pt, 0), NumPut("int", 0, pt, 4)
        DllCall("ClientToScreen", "ptr", focusHwnd, "ptr", pt)
        screenX := NumGet(pt, 0, "int")
        screenY := NumGet(pt, 4, "int")
        ControlGetPos(,, &cW, &cH, focusHwnd, "ahk_id " pacsHwnd)

        targetX := screenX + (cW // 2)
        targetY := screenY + (cH // 2)

        ; 4. 從精確座標取得目標節點
        focusedEl := Acc.ElementFromPoint(targetX, targetY)
        if (!focusedEl) {
            return {srs: "", img: "", error: "Acc: 無法從座標取得節點"}
        }

        ; 5. 取得「相對於 PACS 視窗」的路徑
        fullPath := GetRelativePath(focusedEl, pacsRoot)
        if (fullPath == "") {
            return {srs: "", img: "", error: "Acc: 路徑不在當前視窗內或解析超時"}
        }

        pathParts := StrSplit(fullPath, ",")
        if (pathParts.Length < 2) {
            return {srs: "", img: "", error: "Acc: 路徑層級過淺"}
        }

        targetIdx := pathParts.Length - 1

        ; 解析 Img
        pathParts[targetIdx] := Integer(pathParts[targetIdx]) + 1
        basePath := ""
        Loop targetIdx {
            basePath .= pathParts[A_Index] ","
        }
        imgPath := basePath . pathParts[pathParts.Length] . ",1,4,2,4"
        try {
            imgEl := pacsRoot[imgPath]
            imgVal := Trim(imgEl.Value)
        } catch {
            imgVal := ""
        }

        ; 解析 Srs
        pathParts[targetIdx] := Integer(pathParts[targetIdx]) + 1
        basePath := ""
        Loop targetIdx {
            basePath .= pathParts[A_Index] ","
        }
        srsPath := basePath . pathParts[pathParts.Length]
        srsVal := ""
        descVal := ""

        ; 解析 Series Description (新增)
        try {
            descPath := srsPath . ",2,4"
            descEl := pacsRoot[descPath]
            descVal := Trim(descEl.Value)
            if (descVal == "")
                descVal := Trim(descEl.Name)
        }

        try {
            srsEl := pacsRoot[srsPath]

            ; ★ 優化 1：先嘗試直接從 Acc 屬性讀取文字，省去 OCR 負擔與錯誤率
            rawText := ""
            try {
                rawText := srsEl.Name
            }
            if (rawText == "") {
                try {
                    rawText := srsEl.Value
                }
            }

            if (rawText != "") {
                srsVal := ParseSrs(rawText)
            }

            ; ★ 優化 2：如果 Acc 屬性沒有文字，才使用 OCR
            if (srsVal == "") {
                loc := srsEl.Location
                if (loc.w > 0 && loc.h > 0) {
                    scanW := (loc.w > 150) ? 150 : loc.w
                    ; ★ 改用帶有半透明紅框濾鏡的 OCR 函數
                    ocrResult := CaptureOcrWithFilter(loc.x, loc.y, scanW, loc.h, 2)
                    srsVal := ParseSrs(ocrResult.Text)
                }
            }
        }

        ; ★ 新增：如果 desc 包含 cor 或 sag (不分大小寫)，且非 MRI 關鍵字 (t1, t2, dwi, adc)，則 img + 1
        if (descVal != "" && RegExMatch(descVal, "i)cor|sag") && !RegExMatch(descVal, "i)t1|t2|dwi|adc|dual|stir|fl2d|pd")) {
            if (IsNumber(imgVal)) {
                imgVal := String(Integer(imgVal) + 1)
            }
        }

        return {srs: srsVal, img: imgVal, desc: descVal}
    } catch Error as e {
        return {srs: "", img: "", error: "Acc Error: " . e.Message}
    }
}

; ==============================================================================
; ★ 搭配 Helper：只在指定根節點內尋找路徑 (防死鎖關鍵)
; ==============================================================================
GetRelativePath(targetEl, rootEl) {
    path := ""
    curr := targetEl
    startTime := A_TickCount
    maxDepth := 50
    timeoutMs := 2500

    try {
        Loop maxDepth {
            ; ★ 如果已經往上爬到我們指定的根節點 (PACS 視窗)，就停止
            if (curr.IsEqual(rootEl)) {
                break
            }

            ; 如果爬到沒有父節點 (撞到桌面)，代表發生異常越界
            if (!curr.Parent) {
                return ""
            }

            ; 防死鎖超時判定
            if (A_TickCount - startTime > timeoutMs) {
                throw Error("Acc 相對路徑解析超時")
            }

            p := curr.Parent
            matchFound := false
            for idx, child in p {
                if (child.IsEqual(curr)) {
                    path := idx "," path
                    matchFound := true
                    break
                }
            }

            if (!matchFound) {
                break
            }
            curr := p
        }
    } catch {
        return ""
    }

    return Trim(path, ",")
}

; ==============================================================================
; ★ 修改 Helper：多重區域視覺化 (除錯用，支援不同顏色與常駐)
; ==============================================================================
ShowDebugRects(rectList) {
    global debugGuis ; 改用陣列儲存多個 GUI

    ; 每次呼叫前，先清空舊的框框
    if (IsSet(debugGuis) && debugGuis) {
        For dGui in debugGuis {
            try {
                dGui.Destroy()
            }
        }
    }
    global debugGuis := []

    For r in rectList {
        ; 建立一個穿透點擊、無邊框、永遠在上的純色 GUI
        dGui := Gui("+AlwaysOnTop -Caption +ToolWindow -DPIScale +E0x20")
        dGui.BackColor := r.color

        ; 設定透明度 (0-255，100 為半透明)
        dGui.Opt("+LastFound")
        WinSetTransparent(100)

        dGui.Show("x" r.x " y" r.y " w" r.w " h" r.h " NoActivate")
        debugGuis.Push(dGui) ; 存入陣列防銷毀
    }
}

; ==============================================================================
; ★ 新增 Helper：帶有物理濾鏡的 OCR 抓取 (解決對比度與中文字體幻覺)
; ==============================================================================
CaptureOcrWithFilter(x, y, w, h, scale := 2) {
    ; 1. 建立一個穿透點擊、無邊框的半透明紅色 GUI 作為濾鏡
    filterGui := Gui("+AlwaysOnTop -Caption +ToolWindow -DPIScale +E0x20")
    filterGui.BackColor := "Red"
    filterGui.Opt("+LastFound")
    WinSetTransparent(100)

    ; 2. 顯示濾鏡覆蓋在目標區域
    filterGui.Show("x" x " y" y " w" w " h" h " NoActivate")

    ; 3. 給予極短的延遲，讓 Windows DWM 渲染這個紅框
    Sleep(50)

    ; 4. 執行 OCR 截圖 (此時畫面已加上紅底濾鏡)
    try {
        ocrResult := OCR.FromRect(x, y, w, h, {scale: scale})
    } catch {
        ocrResult := {Text: ""}
    }

    ; 5. 截圖完畢，立刻銷毀濾鏡 (視覺上只會閃爍一下)
    filterGui.Destroy()

    return ocrResult
}

; ==============================================================================
; ★ 修改 Helper：邊緣距離防錯驗證 (嚴格比對 Focus 上緣與 Srs 下緣)
; ==============================================================================
VerifySpatialMatch(childNN, focusNN, hwnd) {
    try {
        ControlGetPos(&cX, &cY, &cW, &cH, childNN, hwnd)
        ControlGetPos(&fX, &fY, &fW, &fH, focusNN, hwnd)

        focusTop := fY
        srsBottom := cY + cH

        ; 1. 垂直驗證：Focus 的上緣 (fY) 離 Srs 的下緣 (cY + cH)，距離不能超過 50 px
        ; 使用 Abs() 確保無論 Srs 是在 Focus 框外正上方，還是 Focus 框內的最上緣，都能被涵蓋
        if (Abs(focusTop - srsBottom) > 50) {
            return false
        }

        ; 2. 水平驗證：確保 Srs 沒有跑到左右隔壁的格子
        ; 容許 10 pixel 誤差，Srs 中心點必須落在 Focus 的 X 座標範圍內
        centerX := cX + (cW / 2)
        hMargin := 10
        if !(centerX >= (fX - hMargin) && centerX <= (fX + fW + hMargin)) {
            return false
        }

        return true
    } catch {
        return false
    }
}

; ==============================================================================
; ★ 命中率統計與動態排序 Helper
; ==============================================================================
RecordPatternHit(patternName) {
    try {
        ; 1. 更新該 Pattern 的累計命中數
        hits := Integer(IniRead(StatsFile, "Hits", patternName, 0))
        IniWrite(hits + 1, StatsFile, "Hits", patternName)

        ; 2. 檢查是否需要執行跨日重新排序 (每天第一次命中時觸發)
        lastSort := IniRead(StatsFile, "Settings", "LastSortDate", "")
        currentDate := FormatTime(A_Now, "yyyyMMdd")

        if (lastSort != currentDate) {
            IniWrite(currentDate, StatsFile, "Settings", "LastSortDate")
            OptimizePatternOrder()
        }
    }
}

OptimizePatternOrder() {
    global PatternList
    try {
        ; 讀取每個 Pattern 的歷史命中數，存入物件屬性中
        for patternData in PatternList {
            patternData.hits := Integer(IniRead(StatsFile, "Hits", patternData.name, 0))
        }

        ; 氣泡排序：依命中數 (hits) 降冪排序 (高頻命中排前面)
        Loop PatternList.Length {
            i := A_Index
            Loop PatternList.Length - i {
                j := A_Index
                if (PatternList[j].hits < PatternList[j+1].hits) {
                    temp := PatternList[j]
                    PatternList[j] := PatternList[j+1]
                    PatternList[j+1] := temp
                }
            }
        }
    }
}

; ==============================================================================
; ★ UI 狀態列 Helper
; ==============================================================================
SetGuiStatus(msg, color := "cRed") {
    global GuiStatusMsg := msg
    if (IsSet(txtStatus) && Type(txtStatus) == "Gui.Text") {
        try {
            txtStatus.Opt(color)
            txtStatus.Value := msg
        }
    }
}

; ==============================================================================
; ★ GUI 介面繪製 (佈局優化與狀態列置底)
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

    ; --- 標題區 ---
    MyGui.SetFont("s11 Bold", "Segoe UI")
    MyGui.Add("Text", "x10 w300 Center", "Nodule Tracker")

    ; --- 按鈕區 ---
    MyGui.SetFont("s9 Norm", "Segoe UI")
    btnX := 15
    btnCopy := MyGui.Add("Button", "x" btnX " y+10 w85 h30", "Full")
    btnCopy.OnEvent("Click", CopyReport)
    btnCopyLocImg := MyGui.Add("Button", "x+10 w90 h30", "Loc:Img")
    btnCopyLocImg.OnEvent("Click", CopyLocImg)
    btnCopyImg := MyGui.Add("Button", "x+10 w85 h30", "Img No")
    btnCopyImg.OnEvent("Click", CopyImgNo)

    ; --- 上分隔線 ---
    MyGui.Add("Text", "x10 y+15 w300 h1 0x10")

    ; --- 列表區 ---
    For key, arr in NoduleData {
        SortNoduleData(arr)
    }

    MyGui.SetFont("s11 Bold", "Segoe UI")
    MyGui.Add("Text", "Section x" COL_LEFT_X " y+10 w" COL_WIDTH " Center cBlue", "Right Lung")
    RenderSection("RUL", COL_LEFT_X)
    RenderSection("RML", COL_LEFT_X)
    RenderSection("RLL", COL_LEFT_X)

    ; ★ 取得左欄最底部的 Y 座標
    dummyLeft := MyGui.Add("Text", "x" COL_LEFT_X " y+0 w0 h0", "")
    dummyLeft.GetPos(, &leftY,, &leftH)
    maxLeftY := leftY + leftH

    MyGui.SetFont("s11 Bold", "Segoe UI")
    MyGui.Add("Text", "ys x" COL_RIGHT_X " w" COL_WIDTH " Center cBlue", "Left Lung")
    RenderSection("LUL", COL_RIGHT_X)

    if (NoduleData["RML"].Length > 0) {
        MyGui.Add("Text", "x" COL_RIGHT_X " y+28", "")
    } else {
        MyGui.Add("Text", "x" COL_RIGHT_X " y+5", "")
    }
    RenderSection("LLL", COL_RIGHT_X)

    ; ★ 取得右欄最底部的 Y 座標
    dummyRight := MyGui.Add("Text", "x" COL_RIGHT_X " y+0 w0 h0", "")
    dummyRight.GetPos(, &rightY,, &rightH)
    maxRightY := rightY + rightH

    ; ★ 計算兩欄中的最大 Y 座標，確保下方的控制項絕對不會重疊
    bottomY := (maxLeftY > maxRightY) ? maxLeftY : maxRightY

    ; --- 狀態列區 (置底) ---
    ; 加入下分隔線 (使用絕對座標取代相對座標)
    MyGui.Add("Text", "x10 y" (bottomY + 15) " w300 h1 0x10")

    MyGui.SetFont("s9 Bold", "Segoe UI")
    statusColor := InStr(GuiStatusMsg, "✅") ? "cBlue" : "cRed"
    displayMsg := (GuiStatusMsg == "") ? " " : GuiStatusMsg

    global txtStatus := MyGui.Add("Text", "x10 y+5 w300 h30 Center " . statusColor, displayMsg)

    ; --- 選項區 (修正語法錯誤) ---
    MyGui.SetFont("s8 Norm", "Segoe UI")
    chkDebug := MyGui.Add("Checkbox", "x10 y+2 Checked" (ShowDebugQuickSet ? "1" : "0"), "顯示 Ctrl+G 除錯資訊")

    ToggleDebug(ctrl, *) {
        global ShowDebugQuickSet := ctrl.Value
    }
    chkDebug.OnEvent("Click", ToggleDebug)

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

CopyImgNo(*) {
    imgMap := Map()
    for label, items in NoduleData {
        for item in items {
            if IsNumber(item.img) {
                imgMap[Integer(item.img)] := true
            }
        }
    }

    imgList := []
    for val, _ in imgMap {
        imgList.Push(val)
    }

    if (imgList.Length == 0) {
        ShowTip("! 無資料可複製", 2000)
        return
    }

    ; 數值排序 (Bubble Sort)
    Loop imgList.Length {
        i := A_Index
        Loop imgList.Length - i {
            j := A_Index
            if (imgList[j] > imgList[j+1]) {
                temp := imgList[j]
                imgList[j] := imgList[j+1]
                imgList[j+1] := temp
            }
        }
    }

    finalStr := ""
    for val in imgList {
        finalStr .= val . ";"
    }
    finalStr := Trim(finalStr, ";")

    A_Clipboard := finalStr
    ShowTip("📋 Copied Img No:`n" finalStr, 3000)
}

CopyLocImg(*) {
    lobeParts := []
    global LobeOrder
    For label in LobeOrder {
        items := NoduleData[label]
        if (items.Length > 0) {
            imgMap := Map()
            for item in items {
                if IsNumber(item.img) {
                    imgMap[Integer(item.img)] := true
                }
            }
            imgList := []
            for val, _ in imgMap {
                imgList.Push(val)
            }
            ; 數值排序
            if (imgList.Length > 1) {
                Loop imgList.Length {
                    i := A_Index
                    Loop imgList.Length - i {
                        j := A_Index
                        if (imgList[j] > imgList[j+1]) {
                            temp := imgList[j], imgList[j] := imgList[j+1], imgList[j+1] := temp
                        }
                    }
                }
            }
            imgStr := ""
            for val in imgList {
                imgStr .= val . ";"
            }
            imgStr := Trim(imgStr, ";")
            lobeParts.Push(label . ":" . imgStr)
        }
    }
    if (lobeParts.Length == 0) {
        ShowTip("! 無資料可複製", 2000)
        return
    }
    finalStr := ""
    for part in lobeParts {
        finalStr .= part . " "
    }
    finalStr := Trim(finalStr)
    A_Clipboard := finalStr
    ShowTip("📋 Copied Loc:Img:`n" finalStr, 3000)
}

ClearAll(*) {
    For key, arr in NoduleData {
        NoduleData[key] := []
    }
    global GuiStatusMsg := "Ready" ; ★ 清空時重置狀態
    UpdateGUI()

    ; ★ 關閉所有除錯紅框
    global debugGuis
    if (IsSet(debugGuis) && debugGuis) {
        For dGui in debugGuis {
            try {
                dGui.Destroy()
            }
        }
        global debugGuis := []
    }
}

CopyReport(*) {
    reportParts := []
    global LobeOrder
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

; ==============================================================================
; ★ 新增：Ctrl+G 實作 (快速跳至指定影像編號)
; ==============================================================================
QuickSetImage() {
    ; 0. 在彈出 GUI 前，先抓取目標視窗與當前滑鼠位置/焦點
    targetHwnd := WinActive("A")
    MouseGetPos(&mX, &mY, &mHwnd, &mCtrlNN)

    ; 優先取得當前視窗的焦點控制項，若無則使用滑鼠下的控制項
    targetFocusHwnd := ControlGetFocus("ahk_id " targetHwnd)
    if (!targetFocusHwnd && mHwnd == targetHwnd) {
        try targetFocusHwnd := ControlGetHwnd(mCtrlNN, "ahk_id " targetHwnd)
    }

    if (!targetFocusHwnd) {
        ShowTip("❌ 無法鎖定目標控制項", 1500)
        return
    }

    ; 1. 建立一個極簡輸入框 GUI
    inputGui := Gui("+AlwaysOnTop -Caption +Border", "Jump to Image")
    inputGui.SetFont("s12 Bold", "Segoe UI")
    inputGui.BackColor := "White"

    editNum := inputGui.Add("Edit", "w60 h30 Center Number Limit3")

    ; 設定 Enter 鍵觸發執行
    btnOk := inputGui.Add("Button", "Default w0 h0", "OK")
    btnOk.OnEvent("Click", (*) => ProcessJump(editNum.Value))

    ; 處理 Esc 鍵取消
    inputGui.OnEvent("Escape", (*) => inputGui.Destroy())

    ; 顯示並激活，確保可以直接輸入
    inputGui.Show("x" (mX - 30) " y" (mY - 15))
    editNum.Focus()

    ; ★ 穩定方案：使用計時器偵測是否失去活動狀態 (預存 HWND 防止銷毀後報錯)
    guiHwnd := inputGui.Hwnd
    SetTimer(WatchFocus, 100)

    WatchFocus() {
        try {
            if (!WinExist("ahk_id " guiHwnd)) {
                SetTimer(WatchFocus, 0)
                return
            }
            if (!WinActive("ahk_id " guiHwnd)) {
                SetTimer(WatchFocus, 0)
                inputGui.Destroy()
            }
        } catch {
            SetTimer(WatchFocus, 0)
        }
    }

    ProcessJump(val) {
        if (val == "" || !IsNumber(val)) {
            inputGui.Destroy()
            return
        }

        targetNum := Integer(val)
        inputGui.Destroy()

        msg := "【Ctrl+G 執行除錯】`n目標影像編號: " targetNum "`n"

        try {
            ; 使用預先鎖定的控制項
            focusNN := ControlGetClassNN(targetFocusHwnd)
            msg .= "- 目標焦點: " focusNN "`n"

            targetCombo := ""
            method := ""

            ; --- 方法 A：探針模式 (Probe) ---
            for patternData in PatternList {
                if (patternData.map.Has(focusNN)) {
                    targetCombo := patternData.map[focusNN].img
                    method := "Probe (" patternData.name ")"
                    break
                }
            }

            ; --- 方法 B：Acc 模式 (Fallback) ---
            if (targetCombo == "" && EnableAccFallback) {
                msg .= "- 探針未命中，嘗試 Acc 模式...`n"
                pacsRoot := Acc.ElementFromHandle(targetHwnd)

                ; 取得座標中心
                ControlGetPos(,, &cW, &cH, targetFocusHwnd, "ahk_id " targetHwnd)
                pt := Buffer(8), NumPut("int", 0, pt, 0), NumPut("int", 0, pt, 4)
                DllCall("ClientToScreen", "ptr", targetFocusHwnd, "ptr", pt)
                tX := NumGet(pt, 0, "int") + (cW // 2)
                tY := NumGet(pt, 4, "int") + (cH // 2)

                focusedEl := Acc.ElementFromPoint(tX, tY)
                fullPath := GetRelativePath(focusedEl, pacsRoot)

                if (fullPath != "") {
                    pathParts := StrSplit(fullPath, ",")
                    if (pathParts.Length >= 2) {
                        targetIdx := pathParts.Length - 1
                        pathParts[targetIdx] := Integer(pathParts[targetIdx]) + 1

                        basePath := ""
                        Loop targetIdx {
                            basePath .= pathParts[A_Index] ","
                        }
                        comboPath := basePath . pathParts[pathParts.Length] . ",1,4,2,4"

                        try {
                            comboEl := pacsRoot[comboPath]
                            loc := comboEl.Location
                            targetCombo := ControlGetClassNN(WindowFromPoint(loc.x + 10, loc.y + 10))
                            method := "Acc Path Calculation"
                        }
                    }
                }
            }

            if (targetCombo == "") {
                throw Error("無法定位目標 ComboBox")
            }

            msg .= "- 定位成功: " targetCombo " (方法: " method ")`n"

            ; --- 執行設定數值 ---
            success := false

            ; 優先嘗試 ControlChooseIndex (v2 標準方法)
            try {
                ControlChooseIndex(targetNum, targetCombo, "ahk_id " targetHwnd)
                success := true
                msg .= "- ControlChooseIndex: 成功`n"
            } catch {
                msg .= "- ControlChooseIndex: 失敗，嘗試 SendMessage...`n"
                res := SendMessage(0x014E, targetNum - 1, 0, targetCombo, "ahk_id " targetHwnd)
                if (res != -1) {
                    success := true
                    msg .= "- CB_SETCURSEL: 成功`n"
                }
            }

            if (success) {
                ShowTip("✅ 跳至影像: " targetNum, 1000)
            } else {
                throw Error("所有寫入方法皆失敗")
            }

        } catch Error as e {
            msg .= "❌ 錯誤: " e.Message "`n"
            ShowTip("❌ 設定失敗", 2000)
        }

        if (ShowDebugQuickSet) {
            ShowDebugWindow(msg, "Ctrl+G 除錯資訊", 5)
        }
    }
}

WindowFromPoint(x, y) {
    return DllCall("WindowFromPoint", "int64", (x & 0xFFFFFFFF) | (y << 32), "ptr")
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

; ==============================================================================
; ★ 新增：自定義除錯視窗 (附帶一鍵複製功能)
; ==============================================================================
ShowDebugWindow(content, title := "Debug Info", timeout := 0) {
    debugGui := Gui("+AlwaysOnTop +Resize", title)
    debugGui.SetFont("s10", "Maple Mono CN")
    debugGui.BackColor := "F0F0F0"

    ; 使用唯讀的 Edit 控制項來裝載內容，並啟用垂直捲軸
    editCtrl := debugGui.Add("Edit", "w600 h400 ReadOnly Multi", content)

    ; 一鍵複製按鈕
    btnCopy := debugGui.Add("Button", "Default w120 h30", "📋 複製全部")
    btnCopy.OnEvent("Click", (*) => (
        A_Clipboard := content,
        ToolTip("已複製到剪貼簿"),
        SetTimer(() => ToolTip(), -2000)
    ))

    ; 關閉按鈕
    btnClose := debugGui.Add("Button", "x+10 w100 h30", "關閉")
    btnClose.OnEvent("Click", (*) => debugGui.Destroy())

    debugGui.Show()

    ; [修正] 顯示後取消 Edit 文字選取，避免藍底影響閱讀
    SendMessage(0x00B1, -1, 0, editCtrl.Hwnd) ; EM_SETSEL

    if (timeout > 0) {
        SetTimer(() => debugGui.Destroy(), -timeout * 1000)
    }
}
