#Requires AutoHotkey v2.0
#Include <Acc.v2>
#Include <OCR.v2>  ; 引用您指定的 lib 名稱

; ==============================================================================
; ★ 關鍵修正：設定 DPI 感知，解決多螢幕 OCR 座標錯位問題
; ==============================================================================
DllCall("SetThreadDpiAwarenessContext", "ptr", -4, "ptr")

; ==============================================================================
; 全域變數與設定
; ==============================================================================
global NoduleData := Map(
    "RUL", [],
    "RML", [],
    "RLL", [],
    "LUL", [],
    "LLL", []
)

; 視窗座標記憶
global GuiX := 100
global GuiY := 150
global MyGui := ""

; 佈局設定
global COL_LEFT_X  := 10
global COL_RIGHT_X := 170
global COL_WIDTH   := 140

UpdateGUI()

; ==============================================================================
; 快捷鍵區域 (僅在 G3PACS 有效)
; ==============================================================================
#HotIf WinActive("ahk_exe G3PACS.exe")

; --- 儲存並更新 GUI (Alt + Key) ---
!q::CaptureNodule("RUL")
!a::CaptureNodule("RML")
!z::CaptureNodule("RLL")
!w::CaptureNodule("LUL")
!s::CaptureNodule("LLL")

; --- 直接複製報告 (Shift + Alt + Key) ---
+!q::DirectCopy("RUL")
+!a::DirectCopy("RML")
+!z::DirectCopy("RLL")
+!w::DirectCopy("LUL")
+!s::DirectCopy("LLL")

#HotIf

; ==============================================================================
; 視窗專用快捷鍵
; ==============================================================================
#HotIf WinActive("Nodule Tracker ahk_class AutoHotkeyGUI")
^c::CopyReport()
Esc::ClearAll()
#HotIf

; ==============================================================================
; 核心功能函數
; ==============================================================================

CaptureNodule(location) {
    try {
        info := GetNoduleInfoFromFocus() ; 取得 {srs, img}

        if (info.img == "" || info.srs == "") {
            ShowTip("⚠️ 抓取失敗 (數值為空)", 2000)
            return
        }

        ; 檢查重複
        For existingItem in NoduleData[location] {
            if (existingItem.srs == info.srs && existingItem.img == info.img) {
                ShowTip("⚠️ 已存在: " info.srs "/" info.img " (忽略)", 1000)
                return
            }
        }

        NoduleData[location].Push(info)
        UpdateGUI()
        ShowTip("✅ " location ": " info.srs "/" info.img, 1000)

    } catch Error as e {
        ShowTip("❌ 錯誤: " e.Message, 2000)
    }
}

DirectCopy(location) {
    try {
        info := GetNoduleInfoFromFocus()

        if (info.img == "" || info.srs == "") {
            ShowTip("⚠️ 抓取失敗 (數值為空)", 2000)
            return
        }

        reportStr := location . " of lung (Srs/Img: " . info.srs . "/" . info.img . ")"
        A_Clipboard := reportStr
        ShowTip("📋 Copied:`n" reportStr, 2000)

    } catch Error as e {
        ShowTip("❌ 錯誤: " e.Message, 2000)
    }
}

ShowTip(msg, duration) {
    ToolTip(msg)
    SetTimer(() => ToolTip(), -duration)
}

/**
 * 整合 ACC (Value) 與 OCR (Text) 來取得資訊
 * Image Number: 透過 ACC Value 取得 (路徑 ...8 -> ...9)
 * Series Number: 透過 OCR 辨識取得 (路徑 ...8 -> ...10)
 */
GetNoduleInfoFromFocus() {
    try {
        focusedEl := Acc.ElementFromPoint()
        if !focusedEl {
            throw Error("無法偵測到元素")
        }

        fullPath := GetFullPath(focusedEl)
        pathParts := StrSplit(fullPath, ",")

        if (pathParts.Length < 2)
            throw Error("路徑層級不足")

        ; 取得目標層級索引 (例如原本是 ...8,4 中的 8 的位置)
        targetIdx := pathParts.Length - 1

        ; ==================================================
        ; 1. 取得 Image Number (維持原有的 ACC Value 方法)
        ; ==================================================
        ; 邏輯：原本是 8，+1 變成 9
        pathParts[targetIdx] := Integer(pathParts[targetIdx]) + 1

        basePath := ""
        Loop targetIdx
            basePath .= pathParts[A_Index] ","

        ; Image Path: ...9,4,1,4,2,4
        imgPath := basePath . pathParts[pathParts.Length] . ",1,4,2,4"
        try {
            imgEl := Acc.GetRootElement()[imgPath]
            imgVal := Trim(imgEl.Value)
        } catch {
            imgVal := "" ; 容錯
        }

        ; ==================================================
        ; 2. 取得 Series Number (改用 OCR)
        ; ==================================================
        ; 邏輯：原本是 8，+2 變成 10 (因為上面已經 +1 變 9 了，所以這裡再 +1 就好)
        pathParts[targetIdx] := Integer(pathParts[targetIdx]) + 1

        basePath := ""
        Loop targetIdx
            basePath .= pathParts[A_Index] ","

        ; Series Path: ...10,4 (直接抓取顯示文字的容器)
        srsPath := basePath . pathParts[pathParts.Length]

        srsVal := ""
        try {
            srsEl := Acc.GetRootElement()[srsPath]
            loc := srsEl.Location

            if (loc.w > 0 && loc.h > 0) {
                ; 執行 OCR
                ; scale: 2 可以提高對小字體的辨識率
                ocrResult := OCR.FromRect(loc.x, loc.y, loc.w, loc.h, {scale: 2})
                rawText := ocrResult.Text ; 例如 "10 , tl_vibe..."

                ; 解析字串：取第一個逗號前的部分
                splitText := StrSplit(rawText, ",")
                if (splitText.Length > 0) {
                    firstPart := splitText[1]
                    ; 使用 RegEx 提取純數字 (避免 OCR 雜訊如空白)
                    if RegExMatch(firstPart, "(\d+)", &match) {
                        srsVal := match[1]
                    }
                }
            }
        } catch {
            srsVal := ""
        }

        return {srs: srsVal, img: imgVal}
    }
    catch {
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
                if child.IsEqual(curr) {
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
; GUI 介面繪製
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

    ; Title
    MyGui.SetFont("s11 Bold", "Segoe UI")
    MyGui.Add("Text", "w320 Center", "Nodule Tracker")

    ; Buttons
    MyGui.SetFont("s9 Norm", "Segoe UI")
    btnX := (320 - 220) / 2
    btnCopy := MyGui.Add("Button", "x" btnX " w100 h30", "Copy Report")
    btnCopy.OnEvent("Click", CopyReport)
    btnClear := MyGui.Add("Button", "x+20 w100 h30", "Clear All")
    btnClear.OnEvent("Click", ClearAll)
    MyGui.Add("Text", "x10 y+10 w300 h1 0x10")

    ; Sort Data
    For key, arr in NoduleData {
        SortNoduleData(arr)
    }

    ; Left Column (Right Lung)
    MyGui.SetFont("s11 Bold", "Segoe UI")
    MyGui.Add("Text", "Section x" COL_LEFT_X " y+10 w" COL_WIDTH " Center cBlue", "Right Lung")
    RenderSection("RUL", COL_LEFT_X)
    RenderSection("RML", COL_LEFT_X)
    RenderSection("RLL", COL_LEFT_X)

    ; Right Column (Left Lung)
    MyGui.SetFont("s11 Bold", "Segoe UI")
    MyGui.Add("Text", "ys x" COL_RIGHT_X " w" COL_WIDTH " Center cBlue", "Left Lung")
    RenderSection("LUL", COL_RIGHT_X)

    if (NoduleData["RML"].Length > 0)
        MyGui.Add("Text", "x" COL_RIGHT_X " y+28", "")
    else
        MyGui.Add("Text", "x" COL_RIGHT_X " y+5", "")

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
        MyGui.SetFont("s10 Norm cDefault", "Segoe UI") ; 確保顏色重置為黑色
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
                s := item.srs
                i := item.img
                if !seriesMap.Has(s)
                    seriesMap[s] := []
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
                For val in imgArr
                    imgStr .= val . ","
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
        Loop reportParts.Length - 1
            finalStr .= reportParts[A_Index] . ", "
        finalStr .= "and " . reportParts[reportParts.Length]
    }

    A_Clipboard := finalStr
    ShowTip("Copied:`n" finalStr, 3000)
}

SortNoduleData(arr) {
    if (arr.Length < 2)
        return
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