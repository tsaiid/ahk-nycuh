#Requires AutoHotkey v2.0
#SingleInstance Force

; =================================================================
; 獨立工作清單背景更新排程腳本 (AutoWorklistUpdate.v2.ahk)
; 目的：供 Windows 工作排程器定時執行，支援 RDP 斷線/螢幕鎖定下的後台更新與讀取
; =================================================================

; 設定相對於此檔案的工作目錄，確保相對路徑與主程式一致 (指向專案根目錄)
SetWorkingDir(A_ScriptDir . "\..")

#Include ..\Lib\UIA.v2.ahk
#Include ..\Lib\RisWorklist.v2.ahk

; 執行日誌記錄
LogMessage("==============================================")
LogMessage("開始執行工作清單背景更新排程...")

; 立即執行更新流程
RunUpdate()

LogMessage("工作清單背景更新排程執行結束。")
LogMessage("==============================================")

/**
 * 記錄日誌訊息到日誌檔案與輸出調試視窗
 * @param {String} msg 訊息文字
 */
LogMessage(msg) {
    logDir := A_ScriptDir . "\..\logs"
    if !DirExist(logDir) {
        try DirCreate(logDir)
    }
    logFile := logDir . "\worklist-update.log"
    
    timestamp := FormatTime(, "yyyy-MM-dd HH:mm:ss")
    logLine := timestamp . " " . msg . "`n"
    
    OutputDebug("[RisAuto] " . logLine)
    FileAppend(logLine, logFile, "UTF-8")
}

RunUpdate() {
    winTitle := "工作清單(frmRIS)"
    
    if !WinExist(winTitle) {
        LogMessage("❌ 錯誤：找不到工作清單視窗 (frmRIS)")
        return
    }
    
    winHwnd := WinExist(winTitle)
    LogMessage("找到工作清單視窗，HWND: " . Format("0x{:X}", winHwnd))
    
    try {
        ; 1. 定位控制項 (包含快取與 Fallback 機制)
        ctrlMap := ResolveWorklistControls(winHwnd)
        
        btnHwnd := ctrlMap["RefreshButton"]
        erHwnd := ctrlMap["ER"]
        admHwnd := ctrlMap["ADM"]
        opdHwnd := ctrlMap["OPD"]
        
        LogMessage(Format("控制項定位結果 - RefreshButton: 0x{:X}, ER: 0x{:X}, ADM: 0x{:X}, OPD: 0x{:X}", btnHwnd, erHwnd, admHwnd, opdHwnd))
        
        ; 2. 觸發重新整理 (後台發送 BM_CLICK)
        LogMessage("正在發送更新訊息給重新整理按鈕 (BM_CLICK)...")
        PostMessage(0x00F5, 0, 0, , "ahk_id " . btnHwnd) ; BM_CLICK = 0x00F5
        
        ; 等待重新整理完成 (給予充足時間載入資料)
        LogMessage("等待工作清單重新整理資料 (2000ms)...")
        Sleep(2000)
        
        ; 3. 複製並解析資料 (透過後台選取與複製)
        categories := ["ER", "ADM", "OPD"]
        categoryData := Map()
        
        for cat in categories {
            hwnd := ctrlMap[cat]
            LogMessage(Format("正在擷取 [{1}] 表格資料 (HWND: 0x{2:X})...", cat, hwnd))
            categoryData[cat] := ExtractGridData(hwnd, cat)
            LogMessage(Format("[{1}] 擷取完成，成功解析出 {2} 筆資料", cat, categoryData[cat].Count))
        }
        
        ; 4. 組裝 JSON
        payload := RisWorklist.BuildJson(categories, categoryData)
        LogMessage("JSON Payload 組裝完成: " . payload.Json)
        
        if (payload.ValidDataCount == 0) {
            LogMessage("⚠️ 警告：統計資料總筆數為 0，略過上傳")
            return
        }
        
        LogMessage("開始上傳資料至 Webhook...")
        
        ; 5. 上傳 Webhook (isSilent 設為 false，並將 LogMessage 作為 notify 回呼)
        RisWorklist.PostDataToWebhook(payload.Json, false, LogMessage)
        
    } catch as err {
        LogMessage("❌ 執行更新發生嚴重異常: " . err.Message)
    }
}

/**
 * 尋找並解析工作清單的控制項 HWND
 * @param {Integer} winHwnd 工作清單視窗控制代碼
 * @returns {Map} 控制項 HWND 對照表
 */
ResolveWorklistControls(winHwnd) {
    cacheFile := "config\worklist-controls.ini"
    
    ; 1. 嘗試驗證現有快取
    cacheValid := false
    cachedHwndMap := Map()
    try {
        cachedWinHwnd := IniRead(cacheFile, "Window", "Hwnd", "")
        if (cachedWinHwnd != "" && Format("0x{:X}", winHwnd) == Format("0x{:X}", cachedWinHwnd)) {
            ctrlKeys := ["RefreshButton", "ER", "ADM", "OPD"]
            allExist := true
            for key in ctrlKeys {
                hCtrl := IniRead(cacheFile, "Controls", key, "")
                if (hCtrl == "" || !DllCall("IsWindow", "Ptr", hCtrl, "int")) {
                    allExist := false
                    break
                }
                ; 驗證是否同屬此視窗
                hRoot := DllCall("GetAncestor", "Ptr", hCtrl, "UInt", 3, "Ptr") ; GA_ROOTOWNER = 3
                if (hRoot != winHwnd) {
                    allExist := false
                    break
                }
                cachedHwndMap[key] := Number(hCtrl)
            }
            if (allExist) {
                cacheValid := true
            }
        }
    }
    
    if (cacheValid) {
        LogMessage("-> 讀取到有效的控制項 HWND 快取檔")
        return cachedHwndMap
    }
    
    LogMessage("⚠️ 快取無效或不存在，啟動動態控制項定位流程...")
    
    ; 2. 檢測 Session 是否有畫面
    isSessionActive := DllCall("User32\OpenInputDesktop", "uint", 0, "int", 0, "uint", 0, "ptr")
    if (isSessionActive) {
        DllCall("CloseDesktop", "ptr", isSessionActive)
        
        ; 使用 UIA 定位並更新快取
        try {
            LogMessage("Session 處於 Active 狀態，嘗試使用 UI Automation (UIA) 進行定位...")
            elWindow := UIA.ElementFromHandle(winHwnd)
            
            btnHwnd := elWindow.FindElement({AutomationId: "btnRefresh"}).NativeWindowHandle
            erHwnd := elWindow.FindElement({AutomationId: "dgvClassifyOPDE"}).NativeWindowHandle
            admHwnd := elWindow.FindElement({AutomationId: "dgvClassifyADM"}).NativeWindowHandle
            opdHwnd := elWindow.FindElement({AutomationId: "dgvClassifyOPDR"}).NativeWindowHandle
            
            resMap := Map(
                "RefreshButton", btnHwnd,
                "ER", erHwnd,
                "ADM", admHwnd,
                "OPD", opdHwnd
            )
            
            ; 寫入快取
            IniWrite(Format("0x{:X}", winHwnd), cacheFile, "Window", "Hwnd")
            for key, hwnd in resMap {
                IniWrite(Format("0x{:X}", hwnd), cacheFile, "Controls", key)
                ; 同步寫入相對座標供 Fallback 比對
                ControlGetPos(&x, &y, &w, &h, hwnd)
                IniWrite(x, cacheFile, "Positions", key . "_X")
                IniWrite(y, cacheFile, "Positions", key . "_Y")
                IniWrite(w, cacheFile, "Positions", key . "_W")
                IniWrite(h, cacheFile, "Positions", key . "_H")
            }
            LogMessage("✅ UIA 定位成功，快取已更新")
            return resMap
        } catch as err {
            LogMessage("⚠️ UIA 定位失敗: " . err.Message . "，轉用座標 Fallback 定位演算法...")
        }
    } else {
        LogMessage("💤 Session 鎖定中 (RDP 斷線)，使用座標 Fallback 定位演算法...")
    }
    
    ; 3. Fallback 座標排序與尺寸匹配定位 (在 Locked Session / 無 UIA 時觸發)
    resMap := Map()
    allCtrls := WinGetControls(winHwnd)
    dgvCandidates := []
    btnCandidates := []
    
    for ctrlName in allCtrls {
        try {
            hwnd := ControlGetHwnd(ctrlName, winHwnd)
            ControlGetPos(&x, &y, &w, &h, hwnd)
            
            ; 篩選 DataGridView 候選
            if (InStr(ctrlName, "WindowsForms10.Window") == 1) {
                if (w > 150 && h > 100) {
                    dgvCandidates.Push({hwnd: hwnd, x: x, y: y, w: w, h: h, name: ctrlName})
                }
            }
            ; 篩選 Button 候選
            else if (InStr(ctrlName, "WindowsForms10.BUTTON") == 1) {
                btnCandidates.Push({hwnd: hwnd, x: x, y: y, w: w, h: h, name: ctrlName})
            }
        }
    }
    
    LogMessage(Format("Fallback 掃描完成。找到 {1} 個 DataGridView 候選, {2} 個 Button 候選", dgvCandidates.Length, btnCandidates.Length))
    
    ; 定位三個 DataGridView (按 X 座標排序: ER -> ADM -> OPD)
    if (dgvCandidates.Length >= 3) {
        Loop dgvCandidates.Length - 1 {
            i := A_Index
            Loop dgvCandidates.Length - i {
                j := A_Index
                if (dgvCandidates[j].x > dgvCandidates[j+1].x) {
                    temp := dgvCandidates[j]
                    dgvCandidates[j] := dgvCandidates[j+1]
                    dgvCandidates[j+1] := temp
                }
            }
        }
        resMap["ER"] := dgvCandidates[1].hwnd
        resMap["ADM"] := dgvCandidates[2].hwnd
        resMap["OPD"] := dgvCandidates[3].hwnd
        
        LogMessage(Format("DGV 排序結果 - ER: 0x{1:X} (x:{2}), ADM: 0x{3:X} (x:{4}), OPD: 0x{5:X} (x:{6})", resMap["ER"], dgvCandidates[1].x, resMap["ADM"], dgvCandidates[2].x, resMap["OPD"], dgvCandidates[3].x))
    } else {
        throw Error("找不到足夠的 DataGridView 控制項 (僅找到 " . dgvCandidates.Length . " 個)")
    }
    
    ; 定位 RefreshButton
    if (btnCandidates.Length > 0) {
        foundBtn := 0
        
        ; A. 優先嘗試讀取按鈕文字
        for btn in btnCandidates {
            try {
                txt := ControlGetText(btn.hwnd)
                if (txt == "重新整理" || txt == "更新" || txt == "查詢" || InStr(txt, "Refresh")) {
                    foundBtn := btn.hwnd
                    LogMessage(Format("按鈕匹配：藉由文字 [{1}] 尋獲 HWND: 0x{2:X}", txt, foundBtn))
                    break
                }
            }
        }
        
        ; B. 讀取不到文字時，比對快取中記錄的位置大小
        if (!foundBtn) {
            cachedX := IniRead(cacheFile, "Positions", "RefreshButton_X", "")
            cachedY := IniRead(cacheFile, "Positions", "RefreshButton_Y", "")
            
            if (cachedX != "" && cachedY != "") {
                bestDiff := 999999
                for btn in btnCandidates {
                    diff := Abs(btn.x - Number(cachedX)) + Abs(btn.y - Number(cachedY))
                    if (diff < bestDiff) {
                        bestDiff := diff
                        foundBtn := btn.hwnd
                    }
                }
                LogMessage(Format("按鈕匹配：藉由快取位置 (cachedX:{1}, cachedY:{2}) 尋獲最接近 HWND: 0x{3:X}", cachedX, cachedY, foundBtn))
            }
        }
        
        ; C. 最底層 Fallback (取第一個按鈕)
        if (!foundBtn) {
            foundBtn := btnCandidates[1].hwnd
            LogMessage(Format("按鈕匹配：Fallback 取第一個按鈕 HWND: 0x{1:X}", foundBtn))
        }
        
        resMap["RefreshButton"] := foundBtn
    } else {
        throw Error("找不到任何 Button 控制項")
    }
    
    return resMap
}

/**
 * 擷取 DataGridView 資料 (優先嘗試 UIA，失敗時 Fallback 至剪貼簿)
 * @param {Integer} hwnd 控制項 HWND
 * @param {String} cat 分類名稱 (ER/ADM/OPD)
 * @returns {Map} 解析後的鍵值對
 */
ExtractGridData(hwnd, cat := "") {
    data := Map()
    
    ; 1. 優先使用 UIA 直接從 HWND 讀取 (免剪貼簿，且支援背景與鎖定 session)
    try {
        LogMessage(Format("  [{1}] 嘗試使用 UIA 從 HWND (0x{2:X}) 讀取資料...", cat, hwnd))
        elGrid := UIA.ElementFromHandle(hwnd)
        if (elGrid) {
            rowElements := elGrid.FindAll({Type: "Custom"})
            LogMessage(Format("  [{1}] UIA 尋獲 {2} 個 Row 元素", cat, rowElements.Length))
            
            if (rowElements.Length > 0) {
                walker := UIA.TreeWalkerTrue
                for row in rowElements {
                    keyEl := walker.TryGetFirstChildElement(row)
                    if (!keyEl)
                        continue
                    valEl := walker.TryGetNextSiblingElement(keyEl)
                    if (!valEl)
                        continue
                    
                    k := "", v := "0"
                    try k := keyEl.Value
                    try v := valEl.Value
                    
                    if (k != "") {
                        k := Trim(k)
                        v := Trim(v)
                        if (v == "")
                            v := 0
                        
                        ; 標準化 Key
                        k := StrReplace(k, "-", "_")
                        data[k] := v
                    }
                }
                
                if (data.Count > 0) {
                    LogMessage(Format("  [{1}] ✅ UIA 讀取成功，共 {2} 筆資料", cat, data.Count))
                    return data
                }
            }
        }
    } catch as err {
        LogMessage(Format("  [{1}] ⚠️ UIA 讀取失敗 (可能處於鎖定狀態且未渲染 UIA 節點): {2}", cat, err.Message))
    }
    
    ; 2. Fallback: 使用剪貼簿複製
    LogMessage(Format("  [{1}] 轉用剪貼簿複製 Fallback 流程...", cat))
    return ExtractGridDataFromClipboard(hwnd, cat)
}

/**
 * 透過剪貼簿複製讀取 DataGridView 資料
 * @param {Integer} hwnd 控制項 HWND
 * @param {String} cat 分類名稱 (ER/ADM/OPD)
 * @returns {Map} 解析後的鍵值對
 */
ExtractGridDataFromClipboard(hwnd, cat := "") {
    data := Map()
    
    ; 備份剪貼簿
    savedClip := ClipboardAll()
    A_Clipboard := ""
    
    try {
        ; 後台聚焦與發送複製快速鍵
        ControlFocus(hwnd)
        Sleep 50
        ControlSend("^a^c", hwnd)
        
        ; 等待複製完成，最多 1.5 秒
        if !ClipWait(1.5) {
            LogMessage(Format("  [{1}] ⚠️ 複製逾時，無法將資料載入剪貼簿 (HWND: 0x{2:X})", cat, hwnd))
            A_Clipboard := savedClip
            return data
        }
        
        clipText := A_Clipboard
        LogMessage(Format("  [{1}] 剪貼簿讀取字元數: {2}", cat, StrLen(clipText)))
        
        ; 還原剪貼簿
        A_Clipboard := savedClip
        
        ; 解析資料
        Loop Parse, clipText, "`n", "`r" {
            line := Trim(A_LoopField)
            if (line == "")
                continue
            
            fields := StrSplit(line, "`t")
            if (fields.Length >= 2) {
                k := Trim(fields[1])
                v := Trim(fields[2])
                
                ; 排除表格的 Header 列 (如果第二欄的值不是純數字，代表它不是資料列)
                if (!IsNumber(v)) {
                    continue
                }
                
                if (k != "") {
                    k := StrReplace(k, "-", "_") ; 標準化 Key
                    data[k] := v
                }
            }
        }
    } catch as err {
        LogMessage(Format("  [{1}] ❌ 後台複製出錯: {2}", cat, err.Message))
        A_Clipboard := savedClip
    }
    
    return data
}
