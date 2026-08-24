; ==============================================================================
; ★ Nodule Tracker (PACS Workflow Optimizer)
; ==============================================================================
; Description : 自動抓取 PACS 影像中的 Series 與 Image 編號並分類肺葉。
;               支援智慧探針 (ClassNN) 與 UI 自動化 (Acc) 雙模切換。
;               具備 OCR 驗證、命中率動態排序及 MPR 影像自動補償。
; Author      : Tsai, I-Ta (放射科醫師)
; GitHub      : tsaiid
; License     : MIT License
; Version     : 2026.04.03 (Updated compensation logic)
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
;    - Lobe:Img    : 依肺葉分類複製影像編號 (e.g., RUL:15;22;LUL:10)。
;    - Img No      : 提取所有 Image Number，排序並以分號分隔複製。
;    - Import Clipboard : 從剪貼簿匯入 Lobe:Img 格式，未知 Series 預設為 4。
; 6. 特殊邏輯：
;    - 自動補償：偵測到 MPR/MIP/COR/SAG 序列時，影像編號自動 +1 (適配 PACS)。
;    - 動態排序：依據歷史命中率自動調整探針順序，提升抓取速度。
; ------------------------------------------------------------------------------

#Requires AutoHotkey v2.0
#Include <Acc.v2>
#Include <OCR.v2>
#Include <RisDialog.v2>
#Include <G3PacsNotify.v2>
#Include <G3PacsProbe.v2>

; ==============================================================================
; ★ 類別定義 (階段一：架構與狀態遷移)
; ==============================================================================
class NoduleTracker {
    ; --- UI 常數 (Static) ---
    static COL_LEFT_X := 10
    static COL_RIGHT_X := 170
    static COL_WIDTH := 140

    ; --- 資料屬性 (Data Properties) ---
    NoduleData := Map("RUL", [], "RML", [], "RLL", [], "LUL", [], "LLL", [])
    LobeOrder := ["RUL", "RML", "RLL", "LUL", "LLL"]
    PatternList := []

    ; --- 配置屬性 (Config Properties) ---
    EnableAccFallback := true
    EnableMprImgOffset := true
    DebugOCR := false
    ShowDebugQuickSet := false
    StatsFile := A_ScriptDir "\PatternStats.ini"

    ; --- UI 狀態屬性 (UI State Properties) ---
    GuiX := A_ScreenWidth - 350
    GuiY := 50
    MyGui := ""
    txtStatus := ""
    GuiStatusMsg := "Ready"
    debugGuis := []

    __New() {
        ; [新增] 設定工作列 (Tray) 上的 Icon
        iconPath := A_ScriptDir "\assets\NoduleTracker_icon.png"
        if FileExist(iconPath) {
            TraySetIcon(iconPath)
        }
        this.InitTrayMenu()

        DllCall("SetThreadDpiAwarenessContext", "ptr", -4, "ptr")
        CoordMode("Mouse", "Screen")
        CoordMode("ToolTip", "Screen")

        this.PatternList := G3PacsProbe.GetPatternList()
        this.OptimizePatternOrder() ; 初始化時依照歷史命中率重新排序 PatternList

        this.UpdateGUI()
    }

    InitGUI() {
        this.MyGui := Gui("+AlwaysOnTop +ToolWindow", "Nodule Tracker")
        this.MyGui.SetFont("s10", "Segoe UI")
        this.MyGui.BackColor := "F8FAFC"
        this.MyGui.OnEvent("Close", this.WindowClosed.Bind(this))

        ; --- 標題區 (靜態) ---
        this.MyGui.SetFont("s11 Bold", "Segoe UI")
        this.MyGui.Add("Text", "x10 w300 Center", "Nodule Tracker")

        ; --- 按鈕區 (靜態) ---
        this.MyGui.SetFont("s9 Norm", "Segoe UI")
        btnX := 15
        btnCopy := this.MyGui.Add("Button", "x" btnX " y+10 w85 h30", "Full")
        btnCopy.OnEvent("Click", this.CopyReport.Bind(this))
        btnCopyLocImg := this.MyGui.Add("Button", "x+10 w90 h30", "Lobe:Img")
        btnCopyLocImg.OnEvent("Click", this.CopyLocImg.Bind(this))
        btnCopyImg := this.MyGui.Add("Button", "x+10 w85 h30", "Img No")
        btnCopyImg.OnEvent("Click", this.CopyImgNo.Bind(this))
        btnImport := this.MyGui.Add("Button", "x" btnX " y+5 w280 h28", "Import Clipboard")
        btnImport.OnEvent("Click", this.ImportLobeImgFromClipboard.Bind(this))

        ; --- 上分隔線 (靜態) ---
        this.MyGui.Add("Text", "x10 y+15 w300 h1 0x10")
    }

    InitTrayMenu() {
        A_TrayMenu.Add()
        A_TrayMenu.Add("啟用 MPR/MIP/COR/SAG Img +1 補償", this.ToggleMprImgOffset.Bind(this))
        this.UpdateTrayMenu()
    }

    ToggleMprImgOffset(*) {
        this.EnableMprImgOffset := !this.EnableMprImgOffset
        this.UpdateTrayMenu()
        this.UpdateGUI()
    }

    UpdateTrayMenu() {
        itemName := "啟用 MPR/MIP/COR/SAG Img +1 補償"
        if (this.EnableMprImgOffset) {
            A_TrayMenu.Check(itemName)
        } else {
            A_TrayMenu.Uncheck(itemName)
        }
    }

    OptimizePatternOrder() {
        try {
            this.Sort(this.PatternList, (a, b) => b.hits - a.hits)
            /*
            for patternData in this.PatternList {
                patternData.hits := Integer(IniRead(this.StatsFile, "Hits", patternData.name, 0))
            }
            Loop this.PatternList.Length {
                i := A_Index
                Loop this.PatternList.Length - i {
                    j := A_Index
                    if (this.PatternList[j].hits < this.PatternList[j+1].hits) {
                        temp := this.PatternList[j]
                        this.PatternList[j] := this.PatternList[j+1]
                        this.PatternList[j+1] := temp
                    }
                }
            }
            */
        }
    }

    Sort(arr, compareFunc := "") {
        if (arr.Length < 2) {
            return arr
        }
        if (compareFunc == "") {
            compareFunc := (a, b) => (a > b ? 1 : a < b ? -1 : 0)
        }
        Loop arr.Length {
            i := A_Index
            Loop arr.Length - i {
                j := A_Index
                if (compareFunc(arr[j], arr[j+1]) > 0) {
                    temp := arr[j], arr[j] := arr[j+1], arr[j+1] := temp
                }
            }
        }
        return arr
    }

    RecordPatternHit(patternName) {
        try {
            hits := Integer(IniRead(this.StatsFile, "Hits", patternName, 0))
            IniWrite(hits + 1, this.StatsFile, "Hits", patternName)
            lastSort := IniRead(this.StatsFile, "Settings", "LastSortDate", "")
            currentDate := FormatTime(A_Now, "yyyyMMdd")
            if (lastSort != currentDate) {
                IniWrite(currentDate, this.StatsFile, "Settings", "LastSortDate")
                this.OptimizePatternOrder()
            }
        }
    }

    GetSmartInfo() {
        fastInfo := this.GetInfo_ByProbe()
        if (fastInfo.valid) {
            return fastInfo
        }
        lastError := fastInfo.HasOwnProp("error") ? fastInfo.error : "Unknown Probe Error"
        if (!this.EnableAccFallback) {
            return {srs: "", img: "", valid: false, error: "Probe 失敗 (" . lastError . ")，已停用 Acc"}
        }
        accInfo := this.GetNoduleInfoFromFocus()
        if (accInfo.srs != "" && accInfo.img != "") {
            accInfo.valid := true
            accInfo.method := "Acc"
            return accInfo
        }
        accError := accInfo.HasOwnProp("error") ? accInfo.error : "未知 Acc 錯誤"
        return {srs: "", img: "", valid: false, error: "Probe 失敗 (" . lastError . ")，Acc 也失敗 (" . accError . ")"}
    }

    GetInfo_ByProbe() {
        try {
            ; 優先從滑鼠位置獲取控制項與視窗 (與 ProbeControl 同步，解決焦點不一致問題)
            MouseGetPos(,, &hwnd, &focusNN)
            match := G3PacsProbe.GetSeriesMatchForFocusClassNN(focusNN, hwnd, this.PatternList)

            ; 如果滑鼠位置未命中，再嘗試當前視窗焦點
            if (!match) {
                hwnd := WinActive("A")
                if (focusHwnd := ControlGetFocus(hwnd)) {
                    focusNN := ControlGetClassNN(focusHwnd)
                    match := G3PacsProbe.GetSeriesMatchForFocusClassNN(focusNN, hwnd, this.PatternList)
                }
            }

            if (match) {
                candidate := match.candidate
                imgVal := ""
                try imgVal := ControlGetText(candidate.img, hwnd)

                if (!IsNumber(Trim(imgVal))) {
                    return {srs: "", img: "", valid: false, error: "ComboBox 無有效數字"}
                }
                imgInfo := this.ApplyImageGroupOffset(imgVal, candidate.img, candidate.imgGroup, hwnd, focusNN)
                imgVal := imgInfo.value

                descVal := match.desc
                minSeries := this.ExtractMinSeriesFromDesc(descVal)

                srsVal := ""
                try {
                    ControlGetPos(&cX, &cY, &cW, &cH, candidate.srs, hwnd)
                    pt := Buffer(8), NumPut("int", cX, pt, 0), NumPut("int", cY, pt, 4)
                    DllCall("ClientToScreen", "ptr", hwnd, "ptr", pt)
                    screenX := NumGet(pt, 0, "int"), screenY := NumGet(pt, 4, "int")

                    if (cW > 0 && cH > 0) {
                        scanW := (cW > 150) ? 150 : cW
                        ocrResult := this.CaptureSrsOcr(screenX, screenY, scanW, cH, minSeries)
                        srsVal := ocrResult.srs
                    }

                    if (this.DebugOCR) {
                        try {
                            ControlGetPos(&fX, &fY, &fW, &fH, focusNN, hwnd)
                            ptF := Buffer(8), NumPut("int", fX, ptF, 0), NumPut("int", fY, ptF, 4)
                            DllCall("ClientToScreen", "ptr", hwnd, "ptr", ptF)
                            focX := NumGet(ptF, 0, "int"), focY := NumGet(ptF, 4, "int")
                            this.ShowDebugRects([
                                {x: screenX, y: screenY, w: cW, h: cH, color: "Red"},
                                {x: focX, y: focY, w: fW, h: fH, color: "Green"}
                            ])
                        }
                    }
                } catch {
                    ; OCR 失敗仍可嘗試備援方案
                }

                if (srsVal != "") {
                    if (this.EnableMprImgOffset && descVal != "" && RegExMatch(descVal, "i)MPR|MIP|COR|SAG") && !RegExMatch(descVal, "i)t1|t2|dwi|adc|dual|stir|fl2d|pd")) {
                        if (IsNumber(imgVal)) {
                            imgVal := String(Integer(imgVal) + 1)
                        }
                    }
                    this.RecordPatternHit(match.name)
                    return {srs: srsVal, img: imgVal, desc: descVal, valid: true, method: candidate.type}
                }
            }
            return {srs: "", img: "", valid: false, error: "所有模式驗證失敗 (標籤不符或 OCR 無效)"}
        } catch Error as e {
            return {srs: "", img: "", valid: false, error: "Probe Runtime Error: " . e.Message}
        }
    }

    GetNoduleInfoFromFocus() {
        try {
            pacsHwnd := WinActive("A")
            if (!pacsHwnd) {
                return {srs: "", img: "", error: "Acc: 找不到活動視窗"}
            }
            pacsRoot := Acc.ElementFromHandle(pacsHwnd)
            focusHwnd := ControlGetFocus("ahk_id " pacsHwnd)
            if (!focusHwnd) {
                return {srs: "", img: "", error: "Acc: 無法取得焦點控制項"}
            }
            pt := Buffer(8), NumPut("int", 0, pt, 0), NumPut("int", 0, pt, 4)
            DllCall("ClientToScreen", "ptr", focusHwnd, "ptr", pt)
            screenX := NumGet(pt, 0, "int")
            screenY := NumGet(pt, 4, "int")
            ControlGetPos(,, &cW, &cH, focusHwnd, "ahk_id " pacsHwnd)
            targetX := screenX + (cW // 2)
            targetY := screenY + (cH // 2)
            focusedEl := Acc.ElementFromPoint(targetX, targetY)
            if (!focusedEl) {
                return {srs: "", img: "", error: "Acc: 無法從座標取得節點"}
            }
            fullPath := this.GetRelativePath(focusedEl, pacsRoot)
            if (fullPath == "") {
                return {srs: "", img: "", error: "Acc: 路徑不在當前視窗內或解析超時"}
            }
            pathParts := StrSplit(fullPath, ",")
            if (pathParts.Length < 2) {
                return {srs: "", img: "", error: "Acc: 路徑層級過淺"}
            }
            targetIdx := pathParts.Length - 1
            pathParts[targetIdx] := Integer(pathParts[targetIdx]) + 1
            basePath := ""
            Loop targetIdx {
                basePath .= pathParts[A_Index] ","
            }
            imgPath := basePath . pathParts[pathParts.Length] . ",1,4,2,4"
            try {
                imgEl := pacsRoot[imgPath]
                imgVal := Trim(imgEl.Value)
                try {
                    loc := imgEl.Location
                    imgNN := ControlGetClassNN(this.WindowFromPoint(loc.x + 5, loc.y + 5))
                    imgInfo := this.ApplyImageGroupOffsetFromAccPaths(imgVal, imgNN, pacsRoot, pathParts, targetIdx, pacsHwnd)
                    if (imgInfo.groupValue != "") {
                        imgVal := imgInfo.value
                    }
                }
            } catch {
                imgVal := ""
            }
            pathParts[targetIdx] := Integer(pathParts[targetIdx]) + 1
            basePath := ""
            Loop targetIdx {
                basePath .= pathParts[A_Index] ","
            }
            srsPath := basePath . pathParts[pathParts.Length]
            srsVal := ""
            descVal := ""
            try {
                descPath := srsPath . ",2,4"
                descEl := pacsRoot[descPath]
                descVal := Trim(descEl.Value)
                if (descVal == "")
                    descVal := Trim(descEl.Name)
            }
            minSeries := this.ExtractMinSeriesFromDesc(descVal)
            try {
                srsEl := pacsRoot[srsPath]
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
                    parsedVal := this.ParseSrs(rawText)
                    if (parsedVal != "" && (minSeries <= 0 || Integer(parsedVal) >= minSeries)) {
                        srsVal := parsedVal
                    }
                }
                if (srsVal == "") {
                    loc := srsEl.Location
                    if (loc.w > 0 && loc.h > 0) {
                        scanW := (loc.w > 150) ? 150 : loc.w
                        ocrResult := this.CaptureSrsOcr(loc.x, loc.y, scanW, loc.h, minSeries)
                        srsVal := ocrResult.srs
                    }
                }
            }
            if (this.EnableMprImgOffset && descVal != "" && RegExMatch(descVal, "i)MPR|MIP|COR|SAG") && !RegExMatch(descVal, "i)t1|t2|dwi|adc|dual|stir|fl2d|pd")) {
                if (IsNumber(imgVal)) {
                    imgVal := String(Integer(imgVal) + 1)
                }
            }
            return {srs: srsVal, img: imgVal, desc: descVal}
        } catch Error as e {
            return {srs: "", img: "", error: "Acc Error: " . e.Message}
        }
    }

    CaptureNodule(location) {
        try {
            info := this.GetSmartInfo()
            if (!info.valid) {
                errLog := info.HasOwnProp("error") ? info.error : "未知錯誤"
                this.GuiStatusMsg := "❌ " . errLog
                this.UpdateGUI()
                return
            }
            For existingItem in this.NoduleData[location] {
                if (existingItem.srs == info.srs && existingItem.img == info.img) {
                    G3PacsNotify.Show("⚠️ 已存在 (忽略)", 1000)
                    return
                }
            }
            this.NoduleData[location].Push(info)
            if (info.HasOwnProp("method") && info.method == "Acc") {
                this.GuiStatusMsg := "⚠️ Probe 失敗，Acc 抓取成功 (建議按 F12 擴充)"
                this.UpdateGUI()
            } else {
                this.GuiStatusMsg := ""
                this.UpdateGUI()
                ; 移除命中的 probe 資訊，只顯示 location 與影像編號
                G3PacsNotify.Show("✅ " location ": " info.srs "/" info.img, 1000)
            }
        } catch Error as e {
            this.GuiStatusMsg := "❌ Critical: " e.Message
            this.UpdateGUI()
        }
    }

    DirectCopy(location) {
        info := this.GetSmartInfo()
        if (!info.valid) {
            G3PacsNotify.Show("⚠️ 複製失敗: " . (info.HasOwnProp("error") ? info.error : "無法抓取"), 2500)
            return
        }
        reportStr := location . " of lung (Srs/Img: " . info.srs . "/" . info.img . ")"
        A_Clipboard := reportStr
        G3PacsNotify.Show("📋 Copied:`n" reportStr, 2000)
    }

    SimpleDirectCopy() {
        try {
            info := this.GetSmartInfo()
            if (!info.valid) {
                G3PacsNotify.Show("⚠️ 抓取失敗: " . (info.HasOwnProp("error") ? info.error : "無法抓取"), 2000)
                return
            }
            clipStr := A_Clipboard
            entries := []
            if (RegExMatch(clipStr, "i)^\s*\(Srs/Img:\s*(.+)\)\s*$", &match)) {
                existingEntries := match[1]
                Loop Parse, existingEntries, ";", " " {
                    if (RegExMatch(A_LoopField, "(\d+)/([\d,]+)", &m)) {
                        srs := m[1]
                        imgs := StrSplit(m[2], ",")
                        for img in imgs {
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
            isNewDup := false
            for e in entries {
                if (e.srs == info.srs && e.img == info.img) {
                    isNewDup := true
                    break
                }
            }
            if (!isNewDup)
                entries.Push({srs: info.srs, img: info.img})
            if (entries.Length > 1) {
                this.Sort(entries, (a, b) => (
                    s1 := Integer(a.srs), s2 := Integer(b.srs),
                    i1 := Integer(a.img), i2 := Integer(b.img),
                    s1 != s2 ? s1 - s2 : i1 - i2
                ))
            }
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
            G3PacsNotify.Show("📋 Copied:`n" reportStr, 2000)
        } catch Error as e {
            G3PacsNotify.Show("❌ " e.Message, 2000)
        }
    }

    ProbeControl() {
        MouseGetPos(,, &hwnd, &ctrlClassNN)
        focusHwnd := ControlGetFocus(hwnd)
        if (!focusHwnd) {
            focusHwnd := hwnd
        }
        msg := "【Pattern 探針資訊】`n當前指向 ClassNN: " ctrlClassNN "`n`n"
        focusNum := this.ExtractClassNNNumber(ctrlClassNN)
        accProbeNums := {focus: focusNum, img: "", srs: "", desc: ""}
        suggestedProbeReport := ""
        matchFound := false
        match := G3PacsProbe.GetSeriesMatchForFocusClassNN(ctrlClassNN, hwnd, this.PatternList)
        if (match) {
            candidate := match.candidate
            msg .= "🎯 探針命中: " this.FormatPatternName(match.name) (match.isOffset ? " (💡 PACS 控制項 +1 容錯命中)" : "") "`n"
            imgVal := ""
            try imgVal := ControlGetText(candidate.img, hwnd)
            imgMax := this.GetComboBoxItemCount(candidate.img, hwnd)
            imgInfo := this.ApplyImageGroupOffset(imgVal, candidate.img, candidate.imgGroup, hwnd, ctrlClassNN)
            srsText := ""
            try srsText := ControlGetText(candidate.srs, hwnd)
            descText := ""
            try descText := ControlGetText(candidate.desc, hwnd)
            minSeries := this.ExtractMinSeriesFromDesc(descText)
            msg .= "  - Img 控制項: " candidate.img " (數值: " Trim(imgVal) ") (最大: " (imgMax != "" ? imgMax : "無法讀取") ")"
            if (imgInfo.groupValue != "") {
                msg .= " (影像組: " imgInfo.groupClassNN "=" imgInfo.groupValue
                if (imgInfo.itemCount != "") {
                    msg .= ", 調整後: " imgInfo.value
                }
                msg .= ")"
            }
            msg .= "`n"
            msg .= "  - Srs 控制項: " candidate.srs " (文字: " Trim(srsText) ")`n"
            msg .= "  - Desc 控制項: " candidate.desc " (數值: " Trim(descText) ")`n"
            srsVal := ""
            try {
                ControlGetPos(&cX, &cY, &cW, &cH, candidate.srs, hwnd)
                pt := Buffer(8), NumPut("int", cX, pt, 0), NumPut("int", cY, pt, 4)
                DllCall("ClientToScreen", "ptr", hwnd, "ptr", pt)
                srsX := NumGet(pt, 0, "int"), srsY := NumGet(pt, 4, "int")
                if (cW > 0 && cH > 0) {
                    scanW := (cW > 150) ? 150 : cW
                    ocrResult := this.CaptureSrsOcr(srsX, srsY, scanW, cH, minSeries)
                    srsVal := ocrResult.srs
                    msg .= "  - OCR 解析 Series: [" (srsVal == "" ? "解析失敗" : srsVal) "]`n"
                    msg .= "    (方法: " ocrResult.method ")`n"
                    msg .= "    (原始文字: " StrReplace(ocrResult.Text, "`n", " ") ")`n"
                    msg .= this.BuildOcrDebugReport(srsX, srsY, scanW, cH, "  - Series OCR 多選項結果", minSeries)
                }
            }
            matchFound := true
        }
        if (!matchFound) {
            msg .= "❌ 狀態：未命中任何完整 Pattern 規則。`n"
            msg .= "  (請確認該控制項是否已加入 GenerateMaps 的配置中)`n"
            msg .= "{SuggestedProbeReport}`n"
        }
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
                fullPath := this.GetRelativePath(focusedEl, pacsRoot)
                if (fullPath == "") {
                    msg .= "❌ 無法解析相對路徑 (越界或超時)`n"
                } else {
                    msg .= "- 原始相依路徑: " fullPath "`n"
                    pathParts := StrSplit(fullPath, ",")
                    if (pathParts.Length >= 2) {
                        targetIdx := pathParts.Length - 1
                        pathParts[targetIdx] := Integer(pathParts[targetIdx]) + 1
                        imgPathParts := pathParts.Clone()
                        basePathImg := ""
                        Loop targetIdx {
                            basePathImg .= pathParts[A_Index] ","
                        }
                        imgPath := basePathImg . pathParts[pathParts.Length] . ",1,4,2,4"
                        pathParts[targetIdx] := Integer(pathParts[targetIdx]) + 1
                        basePathSrs := ""
                        Loop targetIdx {
                            basePathSrs .= pathParts[A_Index] ","
                        }
                        srsPath := basePathSrs . pathParts[pathParts.Length]
                        try {
                            imgEl := pacsRoot[imgPath]
                            msg .= "   ✅ 抓取 Img 值: [" Trim(imgEl.Value) "]`n"
                            loc := imgEl.Location
                            imgHwnd := this.WindowFromPoint(loc.x + 5, loc.y + 5)
                            imgNN := ControlGetClassNN(imgHwnd)
                            imgNum := this.ExtractClassNNNumber(imgNN)
                            accProbeNums.img := imgNum
                            msg .= "     (📍 Img ClassNN: " imgNN " -> 數字: " (imgNum != "" ? imgNum : "?") ")`n"
                            for imgGroupPath in this.GetImageGroupAccPaths(imgPathParts, targetIdx) {
                                try {
                                    imgGroupEl := pacsRoot[imgGroupPath]
                                    accText := this.GetAccElementText(imgGroupEl)
                                    msg .= "     (🧩 影像組候選 Acc path: " imgGroupPath " Value=[" accText.value "] Name=[" accText.name "])`n"
                                    imgInfo := this.ApplyImageGroupOffsetFromAcc(imgEl.Value, imgNN, imgGroupEl, hwnd)
                                    if (imgInfo.groupValue != "") {
                                        msg .= "     (✅ 影像組 Acc 命中: " imgGroupPath "=" imgInfo.groupValue
                                        if (imgInfo.itemCount != "") {
                                            msg .= ", 調整後 Img: " imgInfo.value
                                        }
                                        msg .= ")`n"
                                        break
                                    }
                                } catch Error as e {
                                    msg .= "     (🧩 影像組候選 Acc path: " imgGroupPath " 讀取失敗: " e.Message ")`n"
                                }
                            }
                        } catch {
                            msg .= "   ❌ 預測的 Img 路徑無效`n"
                        }
                        try {
                            srsEl := pacsRoot[srsPath]
                            loc := srsEl.Location
                            msg .= "   ✅ Srs 座標框: W" loc.w " H" loc.h "`n"
                            srsHwnd := this.WindowFromPoint(loc.x + 5, loc.y + 5)
                            srsNN := ControlGetClassNN(srsHwnd)
                            srsNum := this.ExtractClassNNNumber(srsNN)
                            accProbeNums.srs := srsNum
                            msg .= "     (📍 Srs ClassNN: " srsNN " -> 數字: " (srsNum != "" ? srsNum : "?") ")`n"
                            descVal := ""
                            try {
                                descPath := srsPath . ",2,4"
                                descEl := pacsRoot[descPath]
                                descVal := descEl.Value ? descEl.Value : descEl.Name
                                if (descVal != "") {
                                    msg .= "   📝 Series Desc: [" descVal "]`n"
                                    dLoc := descEl.Location
                                    dHwnd := this.WindowFromPoint(dLoc.x + 5, dLoc.y + 5)
                                    dNN := ControlGetClassNN(dHwnd)
                                    descNum := this.ExtractClassNNNumber(dNN)
                                    accProbeNums.desc := descNum
                                    msg .= "     (📍 Desc ClassNN: " dNN " -> 數字: " (descNum != "" ? descNum : "?") ")`n"
                                }
                            }
                            minSeries := this.ExtractMinSeriesFromDesc(descVal)
                            rawText := ""
                            try rawText := srsEl.Name
                            if (rawText == "") {
                                try rawText := srsEl.Value
                            }
                            srsVal := ""
                            if (rawText != "") {
                                msg .= "   🔍 Acc 屬性文字: [" rawText "]`n"
                                parsedVal := this.ParseSrs(rawText)
                                if (parsedVal != "" && (minSeries <= 0 || Integer(parsedVal) >= minSeries)) {
                                    srsVal := parsedVal
                                }
                            }
                            if (srsVal == "") {
                                msg .= "   ⚠️ 文字解析無結果，啟動 OCR...`n"
                                if (loc.w > 0 && loc.h > 0) {
                                    scanW := (loc.w > 150) ? 150 : loc.w
                                    ocrResult := this.CaptureSrsOcr(loc.x, loc.y, scanW, loc.h, minSeries)
                                    safeText := StrReplace(ocrResult.Text, "`n", " ")
                                    msg .= "   🔍 OCR 方法: [" ocrResult.method "]`n"
                                    msg .= "   🔍 OCR 原始文字: [" safeText "]`n"
                                    msg .= this.BuildOcrDebugReport(loc.x, loc.y, scanW, loc.h, "   - Srs OCR 多選項結果", minSeries)
                                    srsVal := ocrResult.srs
                                }
                            }
                            if (srsVal != "") {
                                msg .= "   🎯 解析結果: [" srsVal "]`n"
                            } else {
                                msg .= "   🎯 解析結果: [無法識別序列號]`n"
                            }
                            if (!matchFound) {
                                suggestedProbeReport := this.BuildSuggestedProbeReport(accProbeNums)
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
        msg := StrReplace(msg, "{SuggestedProbeReport}`n", suggestedProbeReport)
        this.ShowDebugWindow(Trim(msg, "`n"), "Pattern 探針資訊")
    }

    RunSmartBenchmark() {
        resultText := "★ 智慧抓取效能測試 (Pattern A-D) ★`n`n"
        hwnd := WinActive("A")
        focusHwnd := ControlGetFocus("A")
        if (!focusHwnd) {
            this.ShowDebugWindow("❌ 測試失敗：無法取得視窗焦點。", "效能測試錯誤")
            return
        }
        focusNN := ControlGetClassNN(focusHwnd)
        resultText .= "當前焦點: " . focusNN . "`n" . "----------------------------------`n"
        startProbe := A_TickCount
        info := this.GetInfo_ByProbe()
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
        startAcc := A_TickCount
        accInfo := this.GetNoduleInfoFromFocus()
        accTime := A_TickCount - startAcc
        resultText .= "🐢 基準方法 (Acc)`n"
        resultText .= "耗時: " . accTime . " ms`n"
        resultText .= "數值: Srs " . (accInfo.srs ? accInfo.srs : "N/A") . " / Img " . (accInfo.img ? accInfo.img : "N/A") . "`n"
        resultText .= "----------------------------------`n"
        if (info.valid) {
            speedup := Round(accTime / (probeTime > 0 ? probeTime : 1), 1)
            resultText .= "🏆 速度提升: " . speedup . " 倍"
        } else {
            resultText .= "💡 建議：請檢查 F12 探針資訊以修正 Pattern 映射表。"
        }
        this.ShowDebugWindow(resultText, "智慧抓取效能測試")
    }

    QuickSetImage() {
        targetHwnd := WinActive("A")
        MouseGetPos(&mX, &mY, &mHwnd, &mCtrlNN)
        targetFocusHwnd := ControlGetFocus("ahk_id " targetHwnd)
        if (!targetFocusHwnd && mHwnd == targetHwnd) {
            try targetFocusHwnd := ControlGetHwnd(mCtrlNN, "ahk_id " targetHwnd)
        }
        if (!targetFocusHwnd) {
            G3PacsNotify.Show("❌ 無法鎖定目標控制項", 1500)
            return
        }
        ; 偵測螢幕 DPI 與 Windows 文字縮放比率
        dpi := 0
        try dpi := G3PacsNotify._GetDpiAtPoint(mX, mY)
        if (!dpi && targetHwnd) {
            try dpi := DllCall("User32\GetDpiForWindow", "Ptr", targetHwnd, "UInt")
        }
        if (!dpi) {
            dpi := A_ScreenDPI
        }
        dpiScale := (dpi > 0) ? (dpi / 96) : 1.0

        textFactor := 1.0
        try {
            rawFactor := RegRead("HKEY_CURRENT_USER\Software\Microsoft\Accessibility", "TextScaleFactor", 100)
            if (IsNumber(rawFactor) && rawFactor > 100) {
                textFactor := rawFactor / 100
            }
        }

        effectiveScale := dpiScale * textFactor
        fontSize := Max(10, Round(12 * effectiveScale))
        pad := Round(2 * effectiveScale)
        editW := Round(68 * effectiveScale)
        editH := Round(32 * effectiveScale)
        totalW := editW + pad * 2
        totalH := editH + pad * 2

        offsetGap := Round(6 * effectiveScale)
        guiX := Round(mX - totalW - offsetGap)
        guiY := Round(mY - totalH - offsetGap)
        try {
            workArea := G3PacsNotify._GetMonitorWorkAreaAtPoint(mX, mY)
            ; 若超出左邊界，則改放置於游標右側避免遮擋
            if (guiX < workArea.left) {
                guiX := Round(mX + (16 * effectiveScale))
            }
            ; 若超出上邊界，則改放置於游標下方避免遮擋
            if (guiY < workArea.top) {
                guiY := Round(mY + (20 * effectiveScale))
            }
            ; 確保仍在螢幕工作區範圍內
            guiX := Max(workArea.left, Min(guiX, workArea.right - totalW))
            guiY := Max(workArea.top, Min(guiY, workArea.bottom - totalH))
        }

        inputGui := Gui("+AlwaysOnTop -Caption +Border -DPIScale", "Jump to Image")
        inputGui.MarginX := pad
        inputGui.MarginY := pad
        inputGui.SetFont("s" fontSize " Bold", "Segoe UI")
        inputGui.BackColor := "White"
        editNum := inputGui.Add("Edit", Format("w{1} h{2} Center Number Limit4", editW, editH))
        btnOk := inputGui.Add("Button", "Default w0 h0 Hidden", "OK")
        btnOk.OnEvent("Click", (*) => ProcessJump(editNum.Value))
        inputGui.OnEvent("Escape", (*) => inputGui.Destroy())
        inputGui.Show(Format("x{1} y{2}", guiX, guiY))
        editNum.Focus()
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
            if (targetNum < 1) {
                inputGui.Destroy()
                G3PacsNotify.Show("❌ 影像編號需大於 0", 1500)
                return
            }
            inputGui.Destroy()
            msg := "【Ctrl+G 執行除錯】`n目標影像編號: " targetNum "`n"
            try {
                focusNN := ControlGetClassNN(targetFocusHwnd)
                msg .= "- 目標焦點: " focusNN "`n"
                targetCombo := ""
                method := ""
                match := G3PacsProbe.GetSeriesMatchForFocusClassNN(focusNN, targetHwnd, this.PatternList)
                if (match) {
                    targetCombo := match.candidate.img
                    method := "Probe (" match.name ")" (match.isOffset ? " [PACS +1 容錯]" : "")
                }
                if (targetCombo == "" && this.EnableAccFallback) {
                    msg .= "- 探針未命中，嘗試 Acc 模式...`n"
                    pacsRoot := Acc.ElementFromHandle(targetHwnd)
                    ControlGetPos(,, &cW, &cH, targetFocusHwnd, "ahk_id " targetHwnd)
                    pt := Buffer(8), NumPut("int", 0, pt, 0), NumPut("int", 0, pt, 4)
                    DllCall("ClientToScreen", "ptr", targetFocusHwnd, "ptr", pt)
                    tX := NumGet(pt, 0, "int") + (cW // 2)
                    tY := NumGet(pt, 4, "int") + (cH // 2)
                    focusedEl := Acc.ElementFromPoint(tX, tY)
                    fullPath := this.GetRelativePath(focusedEl, pacsRoot)
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
                                targetCombo := ControlGetClassNN(this.WindowFromPoint(loc.x + 10, loc.y + 10))
                                method := "Acc Path Calculation"
                            }
                        }
                    }
                }
                if (targetCombo == "") {
                    throw Error("無法定位目標 ComboBox")
                }
                msg .= "- 定位成功: " targetCombo " (方法: " method ")`n"
                itemCount := SendMessage(0x0146, 0, 0, targetCombo, "ahk_id " targetHwnd) ; CB_GETCOUNT
                if (itemCount != -1) {
                    msg .= "- ComboBox 項目數: " itemCount "`n"
                    if (targetNum > itemCount) {
                        throw Error("影像編號超過範圍 (最大: " itemCount ")")
                    }
                } else {
                    msg .= "- ComboBox 項目數: 無法讀取，改用寫入後驗證`n"
                }
                success := false
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
                    selectedIndex := SendMessage(0x0147, 0, 0, targetCombo, "ahk_id " targetHwnd) ; CB_GETCURSEL
                    msg .= "- 寫入後選取 index: " selectedIndex "`n"
                    if (selectedIndex != targetNum - 1) {
                        throw Error("寫入後驗證失敗，目前選取: " (selectedIndex + 1))
                    }
                    G3PacsNotify.Show("✅ 跳至影像: " targetNum, 1000)
                } else {
                    throw Error("所有寫入方法皆失敗")
                }
            } catch Error as e {
                msg .= "❌ 錯誤: " e.Message "`n"
                G3PacsNotify.Show("❌ 設定失敗", 2000)
            }
            if (this.ShowDebugQuickSet) {
                this.ShowDebugWindow(msg, "Ctrl+G 除錯資訊")
            }
        }
    }

    ShowDebugRects(rectList) {
        if (this.debugGuis && this.debugGuis.Length > 0) {
            For dGui in this.debugGuis {
                try {
                    dGui.Destroy()
                }
            }
        }
        this.debugGuis := []
        For r in rectList {
            dGui := Gui("+AlwaysOnTop -Caption +ToolWindow -DPIScale +E0x20")
            dGui.BackColor := r.color
            dGui.Opt("+LastFound")
            WinSetTransparent(100)
            dGui.Show("x" r.x " y" r.y " w" r.w " h" r.h " NoActivate")
            this.debugGuis.Push(dGui)
        }
    }

    ; --- GUI 邏輯與功能方法 (階段三：封裝) ---

    SetGuiStatus(msg, color := "cRed") {
        this.GuiStatusMsg := msg
        if (this.txtStatus && Type(this.txtStatus) == "Gui.Text") {
            try {
                this.txtStatus.Opt(color)
                this.txtStatus.Value := msg
            }
        }
    }

    UpdateGUI() {
        ; --- 1. 紀錄位置並銷毀舊 GUI 避免畫面破碎 ---
        if (this.MyGui) {
            try {
                this.MyGui.GetPos(&currX, &currY)
                this.GuiX := currX
                this.GuiY := currY
            }
            this.MyGui.Destroy()
        }

        this.InitGUI()

        ; --- 2. 數據排序 ---
        For key, arr in this.NoduleData {
            this.SortNoduleData(arr)
        }

        ; --- 3. 繪製列表區 (動態) ---
        startY := 110
        this.MyGui.SetFont("s11 Bold", "Segoe UI")

        ; 右肺標題
        this.MyGui.Add("Text", "Section x" NoduleTracker.COL_LEFT_X " y" startY " w" NoduleTracker.COL_WIDTH " Center cBlue", "Right Lung")

        this.RenderSection("RUL", NoduleTracker.COL_LEFT_X)
        this.RenderSection("RML", NoduleTracker.COL_LEFT_X)
        this.RenderSection("RLL", NoduleTracker.COL_LEFT_X)

        ; 取得左欄最底部的 Y 座標
        dummyLeft := this.MyGui.Add("Text", "x" NoduleTracker.COL_LEFT_X " y+0 w0 h0", "")
        dummyLeft.GetPos(, &leftY,, &leftH)
        maxLeftY := leftY + leftH

        ; 左肺標題
        this.MyGui.SetFont("s11 Bold", "Segoe UI")
        this.MyGui.Add("Text", "ys x" NoduleTracker.COL_RIGHT_X " y" startY " w" NoduleTracker.COL_WIDTH " Center cBlue", "Left Lung")

        this.RenderSection("LUL", NoduleTracker.COL_RIGHT_X)

        ; 補償 RML 造成的偏移
        if (this.NoduleData["RML"].Length > 0) {
            this.MyGui.Add("Text", "x" NoduleTracker.COL_RIGHT_X " y+28", "")
        } else {
            this.MyGui.Add("Text", "x" NoduleTracker.COL_RIGHT_X " y+5", "")
        }
        this.RenderSection("LLL", NoduleTracker.COL_RIGHT_X)

        ; 取得右欄最底部的 Y 座標
        dummyRight := this.MyGui.Add("Text", "x" NoduleTracker.COL_RIGHT_X " y+0 w0 h0", "")
        dummyRight.GetPos(, &rightY,, &rightH)
        maxRightY := rightY + rightH

        ; 計算兩欄中的最大 Y 座標
        bottomY := (maxLeftY > maxRightY) ? maxLeftY : maxRightY

        ; --- 4. 狀態列與選項區 (動態) ---
        this.MyGui.Add("Text", "x10 y" (bottomY + 15) " w300 h1 0x10")

        this.MyGui.SetFont("s9 Bold", "Segoe UI")
        statusColor := InStr(this.GuiStatusMsg, "✅") ? "cBlue" : "cRed"
        displayMsg := (this.GuiStatusMsg == "") ? " " : this.GuiStatusMsg

        this.txtStatus := this.MyGui.Add("Text", "x10 y+5 w300 h30 Center " . statusColor, displayMsg)

        this.MyGui.SetFont("s8 Norm", "Segoe UI")
        chkDebug := this.MyGui.Add("Checkbox", "x10 y+2 Checked" (this.ShowDebugQuickSet ? "1" : "0"), "顯示 Ctrl+G 除錯資訊")
        chkDebug.OnEvent("Click", (ctrl, *) => this.ShowDebugQuickSet := ctrl.Value)

        chkOCR := this.MyGui.Add("Checkbox", "x10 y+2 Checked" (this.DebugOCR ? "1" : "0"), "顯示 OCR 除錯資訊")
        chkOCR.OnEvent("Click", (ctrl, *) => this.DebugOCR := ctrl.Value)

        chkMprOffset := this.MyGui.Add("Checkbox", "x10 y+2 Checked" (this.EnableMprImgOffset ? "1" : "0"), "啟用 MPR/MIP/COR/SAG Img +1 補償")
        chkMprOffset.OnEvent("Click", (ctrl, *) => (this.EnableMprImgOffset := ctrl.Value, this.UpdateTrayMenu()))

        this.MyGui.Show("x" this.GuiX " y" this.GuiY " NoActivate AutoSize")
        RisDialog.ApplyWindowStyle(this.MyGui.Hwnd)
    }

    RenderSection(label, xPos) {
        this.MyGui.SetFont("s10 Bold", "Segoe UI")
        this.MyGui.Add("Text", "x" xPos " y+5 w" NoduleTracker.COL_WIDTH " Center c003366", label)

        items := this.NoduleData[label]
        if (items.Length == 0) {
            this.MyGui.SetFont("s9 Norm cGray", "Segoe UI")
            this.MyGui.Add("Text", "xp y+2 w" NoduleTracker.COL_WIDTH " Center", "-")
        } else {
            this.MyGui.SetFont("s10 Norm cDefault", "Segoe UI")
            For index, item in items {
                displayText := item.srs . "/" . item.img
                textX := xPos + (NoduleTracker.COL_WIDTH / 2) - 35
                this.MyGui.Add("Text", "x" textX " y+5 w50 Right", displayText)

                btnDel := this.MyGui.Add("Button", "x+5 yp-3 w20 h20", "x")
                btnDel.OnEvent("Click", this.DeleteItem.Bind(this, label, index))
            }
        }
    }

    DeleteItem(location, index, *) {
        this.NoduleData[location].RemoveAt(index)
        this.UpdateGUI()
    }

    CopyImgNo(*) {
        imgMap := Map()
        for label, items in this.NoduleData {
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
            G3PacsNotify.Show("! 無資料可複製", 2000)
            return
        }

        ; 數值排序
        this.Sort(imgList)

        finalStr := ""
        for val in imgList {
            finalStr .= val . ";"
        }
        finalStr := Trim(finalStr, ";")

        A_Clipboard := finalStr
        G3PacsNotify.Show("📋 Copied Img No:`n" finalStr, 3000)
    }

    CopyLocImg(*) {
        lobeParts := []
        For label in this.LobeOrder {
            items := this.NoduleData[label]
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
                this.Sort(imgList)

                imgStr := ""
                for val in imgList {
                    imgStr .= val . ";"
                }
                imgStr := Trim(imgStr, ";")
                lobeParts.Push(label . ":" . imgStr)
            }
        }
        if (lobeParts.Length == 0) {
            G3PacsNotify.Show("! 無資料可複製", 2000)
            return
        }
        finalStr := ""
        for part in lobeParts {
            finalStr .= part . ";"
        }
        finalStr := Trim(finalStr, ";")
        A_Clipboard := finalStr
        G3PacsNotify.Show("📋 Copied Lobe:Img:`n" finalStr, 3000)
    }

    ImportLobeImgFromClipboard(*) {
        importedData := this.ParseLobeImgText(A_Clipboard, "4")
        importedCount := 0
        skippedCount := 0

        For label in this.LobeOrder {
            For item in importedData[label] {
                if (this.HasNodule(label, item.srs, item.img)) {
                    skippedCount += 1
                    continue
                }
                this.NoduleData[label].Push(item)
                importedCount += 1
            }
        }

        if (importedCount == 0) {
            this.GuiStatusMsg := skippedCount > 0 ? "⚠️ 剪貼簿資料已存在" : "⚠️ 剪貼簿無可匯入資料"
            this.UpdateGUI()
            G3PacsNotify.Show(this.GuiStatusMsg, 2000)
            return
        }

        this.GuiStatusMsg := "✅ 匯入 " importedCount " 筆剪貼簿資料"
        this.UpdateGUI()
        G3PacsNotify.Show(this.GuiStatusMsg, 2000)
    }

    ParseLobeImgText(text, defaultSrs := "4") {
        importedData := Map()
        For label in this.LobeOrder {
            importedData[label] := []
        }

        currentLobe := ""
        For token in StrSplit(StrReplace(text, "`r", ""), ";") {
            token := Trim(token, " `t`n")
            if (token == "") {
                continue
            }

            imageText := token
            if (RegExMatch(token, "i)^(RUL|RML|RLL|LUL|LLL)\s*:\s*(.*)$", &match)) {
                currentLobe := StrUpper(match[1])
                imageText := Trim(match[2], " `t`n")
            }

            if (currentLobe == "" || imageText == "" || !RegExMatch(imageText, "^\d+$")) {
                continue
            }

            importedData[currentLobe].Push({srs: defaultSrs, img: imageText})
        }

        return importedData
    }

    HasNodule(location, srs, img) {
        For existingItem in this.NoduleData[location] {
            if (existingItem.srs == srs && existingItem.img == img) {
                return true
            }
        }
        return false
    }

    ClearAll(*) {
        For key, arr in this.NoduleData {
            this.NoduleData[key] := []
        }
        this.GuiStatusMsg := "Ready"
        this.UpdateGUI()

        if (this.debugGuis && this.debugGuis.Length > 0) {
            For dGui in this.debugGuis {
                try {
                    dGui.Destroy()
                }
            }
            this.debugGuis := []
        }
    }

    CopyReport(*) {
        reportParts := []
        For label in this.LobeOrder {
            items := this.NoduleData[label]
            if (items.Length > 0) {
                this.SortNoduleData(items)
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
                    this.Sort(srsKeys, (a, b) => Integer(a) - Integer(b))
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
            G3PacsNotify.Show("! 無資料可複製", 2000)
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
        G3PacsNotify.Show("Copied:`n" finalStr, 3000)
    }

    SortNoduleData(arr) {
        this.Sort(arr, (a, b) => (
            s1 := Integer(a.srs), s2 := Integer(b.srs),
            i1 := Integer(a.img), i2 := Integer(b.img),
            s1 != s2 ? s1 - s2 : i1 - i2
        ))
    }

    WindowClosed(*) {
        try {
            if (this.MyGui && WinExist("ahk_id " this.MyGui.Hwnd)) {
                WinGetPos(&currentX, &currentY,,, "ahk_id " this.MyGui.Hwnd)
                this.GuiX := currentX
                this.GuiY := currentY
            }
        }
        For key, arr in this.NoduleData {
            this.NoduleData[key] := []
        }
    }

    ; --- 輔助方法 (Helpers) ---

    ExtractMinSeriesFromDesc(descText) {
        if (descText != "" && RegExMatch(Trim(descText), "^\(\s*(\d+)\s*\)", &match)) {
            return Integer(match[1])
        }
        return 0
    }

    ParseSrs(text) {
        splitText := StrSplit(text, ",")
        if (splitText.Length > 0) {
            if (RegExMatch(Trim(splitText[1]), "^(\d+)", &match)) {
                return match[1]
            }
        }
        return ""
    }

    CaptureOcr(x, y, w, h, scale := 3) {
        try {
            return OCR.FromRect(x, y, w, h, {scale: scale})
        } catch {
            return {Text: ""}
        }
    }

    CaptureOcrWithOptions(x, y, w, h, options) {
        try {
            return OCR.FromRect(x, y, w, h, options)
        } catch {
            return {Text: ""}
        }
    }

    CaptureSrsOcr(x, y, w, h, minSeries := 0) {
        variants := [
            {method: "scale=3", opts: {scale: 3}},
            {method: "scale=3 invert", opts: {scale: 3, invertcolors: true}},
            {method: "scale=2", opts: {scale: 2}},
            {method: "scale=4", opts: {scale: 4}},
            {method: "scale=3 grayscale", opts: {scale: 3, grayscale: true}},
            {method: "scale=4 grayscale", opts: {scale: 4, grayscale: true}},
            {method: "scale=4 invert", opts: {scale: 4, invertcolors: true}},
            {method: "scale=4 gray+invert", opts: {scale: 4, grayscale: true, invertcolors: true}}
        ]
        fallback := {Text: "", srs: "", method: variants[1].method}
        for variant in variants {
            result := this.CaptureOcrWithOptions(x, y, w, h, variant.opts)
            srsVal := this.ParseSrs(result.Text)
            if (variant.method == variants[1].method) {
                fallback := {Text: result.Text, srs: (minSeries <= 0 || (srsVal != "" && Integer(srsVal) >= minSeries)) ? srsVal : "", method: variant.method}
            }
            if (srsVal != "" && (minSeries <= 0 || Integer(srsVal) >= minSeries)) {
                return {Text: result.Text, srs: srsVal, method: variant.method}
            }
        }
        return fallback
    }

    BuildOcrDebugReport(x, y, w, h, label := "", minSeries := 0) {
        variants := [
            {name: "scale=2", opts: {scale: 2}},
            {name: "scale=3", opts: {scale: 3}},
            {name: "scale=4", opts: {scale: 4}},
            {name: "scale=3 grayscale", opts: {scale: 3, grayscale: true}},
            {name: "scale=4 grayscale", opts: {scale: 4, grayscale: true}},
            {name: "scale=3 invert", opts: {scale: 3, invertcolors: true}},
            {name: "scale=4 invert", opts: {scale: 4, invertcolors: true}},
            {name: "scale=4 gray+invert", opts: {scale: 4, grayscale: true, invertcolors: true}}
        ]
        msg := ""
        if (label != "") {
            msg .= label (minSeries > 0 ? " (驗證 Series >= " minSeries ")" : "") "`n"
        }
        msg .= "    區域: X" x " Y" y " W" w " H" h "`n"
        for variant in variants {
            result := this.CaptureOcrWithOptions(x, y, w, h, variant.opts)
            text := Trim(StrReplace(result.Text, "`n", " "))
            if (text == "") {
                text := "<empty>"
            }
            status := ""
            parsed := this.ParseSrs(result.Text)
            if (parsed != "") {
                if (minSeries > 0 && Integer(parsed) < minSeries) {
                    status := " -> 解析: " parsed " (<" minSeries " 排除)"
                } else {
                    status := " -> 解析: " parsed " ✅"
                }
            }
            msg .= "    - " variant.name ": [" text "]" status "`n"
        }
        return msg
    }

    ProbeComboBox(controlName, hwnd) {
        try {
            txt := ControlGetText(controlName, hwnd)
            return IsNumber(Trim(txt))
        } catch {
            return false
        }
    }

    GetRelativePath(targetEl, rootEl) {
        path := ""
        curr := targetEl
        startTime := A_TickCount
        maxDepth := 50
        timeoutMs := 2500
        try {
            Loop maxDepth {
                if (curr.IsEqual(rootEl)) {
                    break
                }
                if (!curr.Parent) {
                    return ""
                }
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

    ExtractClassNNNumber(classNN) {
        if (RegExMatch(classNN, "(\d+)$", &match)) {
            return Integer(match[1])
        }
        return ""
    }

    GetComboBoxItemCount(classNN, hwnd) {
        try {
            itemCount := SendMessage(0x0146, 0, 0, classNN, "ahk_id " hwnd) ; CB_GETCOUNT
            return itemCount != -1 ? itemCount : ""
        } catch {
            return ""
        }
    }

    GetImageGroupAccPaths(pathParts, targetIdx) {
        paths := []
        if (targetIdx < 1 || pathParts.Length < 2 || targetIdx > pathParts.Length) {
            return paths
        }

        ; Acc tree 的影像組節點是 Img row 的下一個 sibling；
        ; ClassNN 才是 ComboBox10 -> ComboBox13 的 +3 規則。
        groupIdx := Integer(pathParts[targetIdx]) + 1
        suffixes := ["4,4,4"]
        for suffix in suffixes {
            basePath := ""
            Loop targetIdx {
                if (A_Index == targetIdx) {
                    basePath .= groupIdx ","
                } else {
                    basePath .= pathParts[A_Index] ","
                }
            }
            paths.Push(basePath . suffix)
        }
        return paths
    }

    GetAccElementText(accEl) {
        text := {value: "", name: ""}
        try {
            text.value := Trim(accEl.Value)
        }
        try {
            text.name := Trim(accEl.Name)
        }
        return text
    }

    ApplyImageGroupOffset(imgVal, imgClassNN, imgGroupClassNN, hwnd, focusNN := "") {
        info := {value: Trim(imgVal), groupClassNN: imgGroupClassNN, groupValue: "", itemCount: "", adjusted: false}
        if (!IsNumber(info.value) || imgGroupClassNN == "") {
            return info
        }
        if (focusNN != "" && !G3PacsProbe.IsSpatialControlMatch(imgGroupClassNN, focusNN, hwnd)) {
            return info
        }

        try {
            info.groupValue := Trim(ControlGetText(imgGroupClassNN, hwnd))
        } catch {
            return info
        }
        if (info.groupValue == "" || StrUpper(info.groupValue) == "A" || !IsNumber(info.groupValue)) {
            return info
        }

        groupNum := Integer(info.groupValue)
        if (groupNum < 1) {
            return info
        }

        info.itemCount := this.GetComboBoxItemCount(imgClassNN, hwnd)
        if (info.itemCount == "") {
            return info
        }

        info.value := String(Integer(info.value) + ((groupNum - 1) * Integer(info.itemCount)))
        info.adjusted := groupNum > 1
        return info
    }

    ApplyImageGroupOffsetFromAcc(imgVal, imgClassNN, imgGroupEl, hwnd) {
        info := {value: Trim(imgVal), groupClassNN: "Acc", groupValue: "", itemCount: "", adjusted: false}
        if (!IsNumber(info.value) || !imgGroupEl) {
            return info
        }

        accText := this.GetAccElementText(imgGroupEl)
        info.groupValue := accText.value != "" ? accText.value : accText.name
        if (info.groupValue == "" || StrUpper(info.groupValue) == "A" || !IsNumber(info.groupValue)) {
            return info
        }

        groupNum := Integer(info.groupValue)
        if (groupNum < 1) {
            return info
        }

        info.itemCount := this.GetComboBoxItemCount(imgClassNN, hwnd)
        if (info.itemCount == "") {
            return info
        }

        info.value := String(Integer(info.value) + ((groupNum - 1) * Integer(info.itemCount)))
        info.adjusted := groupNum > 1
        return info
    }

    ApplyImageGroupOffsetFromAccPaths(imgVal, imgClassNN, pacsRoot, pathParts, targetIdx, hwnd) {
        info := {value: Trim(imgVal), groupClassNN: "Acc", groupValue: "", itemCount: "", adjusted: false, path: ""}
        for imgGroupPath in this.GetImageGroupAccPaths(pathParts, targetIdx) {
            try {
                imgGroupEl := pacsRoot[imgGroupPath]
                candidate := this.ApplyImageGroupOffsetFromAcc(imgVal, imgClassNN, imgGroupEl, hwnd)
                if (candidate.groupValue != "") {
                    candidate.path := imgGroupPath
                    return candidate
                }
            }
        }
        return info
    }

    BuildSuggestedProbeReport(nums) {
        if (nums.focus == "" || nums.img == "" || nums.srs == "" || nums.desc == "") {
            return ""
        }

        candidates := []
        for family in [{focus: 3, img: 10}, {focus: 5, img: 57}] {
            offset := nums.focus - family.focus
            if (offset >= 0 && offset < 8) {
                expectedImg := family.img + (offset * 6)
                if (nums.img == expectedImg) {
                    candidates.Push({text: "[" family.focus ", " family.img ", " (nums.srs - (offset * 3)) ", " (nums.desc - (offset * 5)) "]", offset: offset})
                }
            }
        }
        if (candidates.Length == 0) {
            directProbe := "[" nums.focus ", " nums.img ", " nums.srs ", " nums.desc "]"
            return "   ⚠️ 無法回推建議 probe: " directProbe " (警告：不符合 [3, 10, ?, ?] 或 [5, 57, ?, ?]，請人工確認)`n"
        }

        msg := "   💡 建議新增 probe: "
        for idx, candidate in candidates {
            msg .= (idx > 1 ? " 或 " : "") candidate.text
        }
        if (candidates[1].offset > 0) {
            msg .= " (註：從 focus " nums.focus " 回推)"
        }
        msg .= "`n"
        return msg
    }

    FormatPatternName(patternName) {
        if (RegExMatch(patternName, "^Pattern_(\d+)_(\d+)_(\d+)_(\d+)$", &match)) {
            return "[" match[1] ", " match[2] ", " match[3] ", " match[4] "]"
        }
        return patternName
    }


    WindowFromPoint(x, y) {
        return DllCall("WindowFromPoint", "int64", (x & 0xFFFFFFFF) | (y << 32), "ptr")
    }

    ShowDebugWindow(content, title := "Debug Info", timeout := 0) {
        debugGui := RisDialog.Create(title, "+AlwaysOnTop +ToolWindow +Resize", {MarginX: 14, MarginY: 12, FontSize: "s10", FontName: "Maple Mono CN"})
        editCtrl := debugGui.Add("Edit", "xm w600 h400 ReadOnly Multi -WantReturn", content)
        btnCopy := debugGui.Add("Button", "Default w120 xm y+12", "📋 複製全部")
        btnCopy.OnEvent("Click", (*) => (
            A_Clipboard := content,
            ToolTip("已複製到剪貼簿"),
            SetTimer(() => ToolTip(), -2000)
        ))
        btnClose := debugGui.Add("Button", "x+10 yp w100", "關閉")
        btnClose.OnEvent("Click", (*) => debugGui.Destroy())
        debugGui.OnEvent("Close", (*) => debugGui.Destroy())
        debugGui.OnEvent("Escape", (*) => debugGui.Destroy())
        RisDialog.ShowCenter(debugGui)
        SendMessage(0x00B1, 0, 0, editCtrl.Hwnd)
        btnCopy.Focus()
        if (timeout > 0) {
            SetTimer(() => debugGui.Destroy(), -timeout * 1000)
        }
    }
}

; 建立全域執行個體
global tracker := NoduleTracker()

; ==============================================================================
; ★ 快捷鍵區域
; ==============================================================================
#HotIf WinActive("ahk_exe G3PACS.exe") && WinActive("INFINITT PACS")

; F11: 效能測試
F11::tracker.RunSmartBenchmark()

; F12: 探針工具 (Debug用)
F12::tracker.ProbeControl()

!q::tracker.CaptureNodule("RUL")
!a::tracker.CaptureNodule("RML")
!z::tracker.CaptureNodule("RLL")
!w::tracker.CaptureNodule("LUL")
!s::tracker.CaptureNodule("LLL")

; ★ 新增：Ctrl+G 快速跳至指定影像編號 (Image Number)
^g::tracker.QuickSetImage()

; 使用大括號語法 (AHK v2)，並加入 vkE8 阻斷語言切換偵測

+!q:: {
    Send("{Blind}{vkE8}")
    tracker.DirectCopy("RUL")
    Send("{Blind}{vkE8}")
}

+!a:: {
    Send("{Blind}{vkE8}")
    tracker.DirectCopy("RML")
    Send("{Blind}{vkE8}")
}

+!z:: {
    Send("{Blind}{vkE8}")
    tracker.DirectCopy("RLL")
    Send("{Blind}{vkE8}")
}

+!w:: {
    Send("{Blind}{vkE8}")
    tracker.DirectCopy("LUL")
    Send("{Blind}{vkE8}")
}

+!s:: {
    Send("{Blind}{vkE8}")
    tracker.DirectCopy("LLL")
    Send("{Blind}{vkE8}")
}

m:: {
    ;Send("{Blind}{vkE8}")
    tracker.SimpleDirectCopy()
}

#HotIf
