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

; ★ 新增：Acc 慢速模式的總開關 (強烈建議設為 false，徹底杜絕卡死)
global EnableAccFallback := true

; ★ 新增：OCR 視覺化與文字除錯開關 (發布或實戰時請改為 false)
global DebugOCR := false

global GuiStatusMsg := "Ready"
global txtStatus := ""

; ★ 將分散的 Map 改為統一的 PatternList 陣列統一管理
global PatternList := []

GenerateMaps()
UpdateGUI()

; ==============================================================================
; ★ 映射表生成 (極簡陣列配置)
; ==============================================================================
GenerateMaps() {
    classPrefix := "Afx:00400000:b:00000000:00000013:00000000"
    global PatternList := []

    ; ★ 極簡配置：只需依序填入 [Focus起始, Combo起始, Afx起始]
    configs := [
        [3, 10, 31],
        [3, 10, 33],
        [3, 10, 37],
        [3, 10, 39],
        [3, 10, 41],
        [3, 10, 45],
        [3, 10, 47],
        [5, 57, 39],
        [5, 57, 45],
        [5, 57, 47],
    ]

    for idx, cfg in configs {
        pMap := Map()
        pName := "Pattern " . idx

        Loop 8 {
            i := A_Index - 1
            f := classPrefix . (cfg[1] + i)
            c := "ComboBox" . (cfg[2] + (i * 6))
            a := "AfxWnd140u" . (cfg[3] + (i * 3))
            pMap[f] := {img: c, srs: a, type: pName}
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

        ; ★ 改用動態生成的 PatternList 依序測試
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
                ; 這是極速過濾，避免對非 Series 區域執行 OCR
                try {
                    ctrlText := ControlGetText(candidate.srs, hwnd)
                    if (!InStr(ctrlText, "VMTool")) {
                        continue ; Text 不符，表示這不是我們要找的 Series 區域
                    }

                    ; ★★★ 進階幾何除錯 ★★★
                    if (DebugOCR) {
                        ; 1. 取得 Srs 區域 (Screen) -> 紅色
                        ControlGetPos(&cX, &cY, &cW, &cH, candidate.srs, hwnd)
                        ptC := Buffer(8), NumPut("int", cX, ptC, 0), NumPut("int", cY, ptC, 4)
                        DllCall("ClientToScreen", "ptr", hwnd, "ptr", ptC)
                        srsX := NumGet(ptC, 0, "int"), srsY := NumGet(ptC, 4, "int")

                        ; 2. 取得 Focus 區域 (Screen) -> 綠色
                        ControlGetPos(&fX, &fY, &fW, &fH, focusNN, hwnd)
                        ptF := Buffer(8), NumPut("int", fX, ptF, 0), NumPut("int", fY, ptF, 4)
                        DllCall("ClientToScreen", "ptr", hwnd, "ptr", ptF)
                        focX := NumGet(ptF, 0, "int"), focY := NumGet(ptF, 4, "int")

                        ; 畫出雙色框
                        ShowDebugRects([
                            {x: srsX, y: srsY, w: cW, h: cH, color: "Red"},  ; Target Srs
                            {x: focX, y: focY, w: fW, h: fH, color: "Green"} ; Current Focus
                        ])
                    }

                    ; 新增：空間防錯驗證，確保 srs 控制項真的落在這個 focus 視窗的範圍內
                    if (!VerifySpatialMatch(candidate.srs, focusNN, hwnd)) {
                        continue ; 發生 Cross-Talk (抓到別格的資訊)，跳過此 Pattern
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
                        ; OCR 不需要重複畫框，上面已經畫過了

                        ocrResult := OCR.FromRect(screenX, screenY, cW, cH, {scale: 2})
                        srsVal := ParseSrs(ocrResult.Text)

                        ; ★ 顯示原始 OCR 字串與解析結果，比對是否抓錯
                        if (DebugOCR) {
                            MsgBox("【OCR 除錯資訊】`n`nScreen X: " screenX "`nScreen Y: " screenY "`nWidth: " cW "`nHeight: " cH "`n`n[原始 OCR 抓取文字]:`n" ocrResult.Text "`n`n[ParseSrs 解析結果]: " srsVal "`n`n[幾何驗證]: 通過", "Debug OCR", "T5")
                        }
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
; ★ 工具：F12 探針 (動態顯示資訊 - 精準唯一命中判定)
; ==============================================================================
ProbeControl() {
    MouseGetPos(,, &hwnd, &ctrlClassNN)
    msg := "【探針資訊】`n當前指向 ClassNN: " ctrlClassNN "`n`n"

    matchFound := false

    ; ★ 動態驗證所有 Pattern，找出唯一真正符合條件的那一個
    for patternData in PatternList {
        pMap := patternData.map

        if (pMap.Has(ctrlClassNN)) {
            candidate := pMap[ctrlClassNN]

            ; 模擬前兩重驗證：檢查對應的 Img 與 Srs 控制項是否真實存在且符合條件
            imgVal := ""
            try {
                imgVal := ControlGetText(candidate.img, hwnd)
            }

            srsText := ""
            try {
                srsText := ControlGetText(candidate.srs, hwnd)
            }

            ; 判定條件：Img 必須是數字，且 Srs 必須包含 "VMTool"
            if (IsNumber(Trim(imgVal)) && InStr(srsText, "VMTool")) {

                ; ★★★ F12 進階幾何除錯 ★★★
                if (DebugOCR) {
                    ; 1. 取得 Srs 區域 (Screen) -> 紅色
                    ControlGetPos(&cX, &cY, &cW, &cH, candidate.srs, hwnd)
                    ptC := Buffer(8), NumPut("int", cX, ptC, 0), NumPut("int", cY, ptC, 4)
                    ; ★ 修正：補上 "ptr"
                    DllCall("ClientToScreen", "ptr", hwnd, "ptr", ptC)
                    srsX := NumGet(ptC, 0, "int"), srsY := NumGet(ptC, 4, "int")

                    ; 2. 取得 Focus 區域 (Screen) -> 綠色
                    ControlGetPos(&fX, &fY, &fW, &fH, ctrlClassNN, hwnd)
                    ptF := Buffer(8), NumPut("int", fX, ptF, 0), NumPut("int", fY, ptF, 4)
                    ; ★ 修正：補上 "ptr"
                    DllCall("ClientToScreen", "ptr", hwnd, "ptr", ptF)
                    focX := NumGet(ptF, 0, "int"), focY := NumGet(ptF, 4, "int")

                    ; 畫出雙色框
                    ShowDebugRects([
                        {x: srsX, y: srsY, w: cW, h: cH, color: "Red"},
                        {x: focX, y: focY, w: fW, h: fH, color: "Green"}
                    ])
                }

                ; ★ 增加空間幾何驗證
                spatialPass := VerifySpatialMatch(candidate.srs, ctrlClassNN, hwnd)

                if (DebugOCR) {
                     msg .= "【幾何驗證結果】: " (spatialPass ? "✅ 通過" : "❌ 失敗") "`n------------------`n"
                }

                if (!spatialPass) {
                    continue ; 雖然文字符合，但是是別格的控制項，直接跳過
                }

                msg .= "🎯 實際命中: " patternData.name "`n"
                msg .= "  - Img 控制項: " candidate.img " (數值: " Trim(imgVal) ")`n"
                msg .= "  - Srs 控制項: " candidate.srs " (標籤吻合，且通過空間驗證)`n"
                matchFound := true
                break ; 找到真正吻合的就跳出，確保最多只命中一個
            }
        }
    }

    if (!matchFound) {
        msg .= "❌ 狀態：未命中任何完整 Pattern 規則。`n"
        msg .= "💡 建議：若這是新的影像版型，請擷取上述 ClassNN 以便新增至 Pattern 陣列。"
    }

    MsgBox(Trim(msg, "`n"))
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
; ★ 方案 B 核心：限制 Acc 掃描範圍於當前視窗 (防止 VS Code/系統卡死)
; ==============================================================================
GetNoduleInfoFromFocus() {
    try {
        ; 1. 取得當前活動視窗 (PACS) 的 Handle 與 Acc 根節點
        pacsHwnd := WinActive("A")
        if (!pacsHwnd) {
            return {srs: "", img: "", error: "Acc: 找不到活動視窗"}
        }
        pacsRoot := Acc.ElementFromHandle(pacsHwnd) ; ★ 範圍鎖定，絕對不碰桌面

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

        ; 5. ★ 取得「相對於 PACS 視窗」的路徑
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
            imgEl := pacsRoot[imgPath]  ; ★ 直接從 pacsRoot 往下找
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
        try {
            srsEl := pacsRoot[srsPath]  ; ★ 直接從 pacsRoot 往下找
            loc := srsEl.Location
            if (loc.w > 0 && loc.h > 0) {
                ocrResult := OCR.FromRect(loc.x, loc.y, loc.w, loc.h, {scale: 2})
                srsVal := ParseSrs(ocrResult.Text)
            }
        }

        return {srs: srsVal, img: imgVal}
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
    btnX := 50
    btnCopy := MyGui.Add("Button", "x" btnX " y+10 w100 h30", "Copy Report")
    btnCopy.OnEvent("Click", CopyReport)
    btnClear := MyGui.Add("Button", "x+20 w100 h30", "Clear All")
    btnClear.OnEvent("Click", ClearAll)

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