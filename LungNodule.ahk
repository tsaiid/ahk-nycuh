#Requires AutoHotkey v2.0
#Include <Acc.v2>

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

; 佈局設定：左右欄位的 X 座標與欄位寬度
global COL_LEFT_X  := 10
global COL_RIGHT_X := 170
global COL_WIDTH   := 140

UpdateGUI()

; ==============================================================================
; 快捷鍵區域 (僅在 G3PACS 有效)
; ==============================================================================
#HotIf WinActive("ahk_exe G3PACS.exe")

!q::CaptureNodule("RUL")
!a::CaptureNodule("RML")
!z::CaptureNodule("RLL")
!w::CaptureNodule("LUL")
!s::CaptureNodule("LLL")

#HotIf

; ==============================================================================
; 視窗專用快捷鍵 (僅在 Nodule Tracker 視窗有效)
; ==============================================================================
#HotIf WinActive("Nodule Tracker ahk_class AutoHotkeyGUI")

^c::CopyReport()  ; Ctrl + C 複製
Esc::ClearAll()   ; Esc 清除

#HotIf

; ==============================================================================
; 核心功能函數
; ==============================================================================
CaptureNodule(location) {
    try {
        ; 取得物件資訊 {srs: "xxx", img: "xxx"}
        info := GetNoduleInfoFromFocus()

        if (info.img == "" || info.srs == "") {
            ShowTip("⚠️ 抓取失敗 (數值或 Series 為空)", 2000)
            return
        }

        ; 檢查重複 (同時比對 srs 與 img)
        For existingItem in NoduleData[location] {
            if (existingItem.srs == info.srs && existingItem.img == info.img) {
                ShowTip("⚠️ 已存在: " info.srs "/" info.img " (忽略)", 1000)
                return
            }
        }

        ; 存入資料結構
        NoduleData[location].Push(info)
        UpdateGUI()

        ShowTip("✅ " location ": " info.srs "/" info.img, 1000)

    } catch Error as e {
        ShowTip("❌ 錯誤: " e.Message, 2000)
    }
}

ShowTip(msg, duration) {
    ToolTip(msg)
    SetTimer(() => ToolTip(), -duration)
}

/**
 * 從 Focus 元素推算 Image Number 與 Series Number
 * @returns {Object} {srs: string, img: string}
 */
GetNoduleInfoFromFocus() {
    try {
        focusedEl := Acc.ElementFromPoint()
        if !focusedEl {
            throw Error("無法偵測到元素 (請確認滑鼠位置)")
        }

        fullPath := GetFullPath(focusedEl)
        pathParts := StrSplit(fullPath, ",")

        if (pathParts.Length < 2)
            throw Error("路徑層級不足")

        ; 基準索引 (原本的 8 位置)
        targetIdx := pathParts.Length - 1

        ; 1. 取得 Image Number (Value)
        ; 路徑邏輯: ...8,4 -> ...9,4,1,4,2,4
        pathParts[targetIdx] := Integer(pathParts[targetIdx]) + 1 ; 8 -> 9

        basePath := ""
        Loop targetIdx {
            basePath .= pathParts[A_Index] ","
        }

        ; 組合 Image Path
        imgPath := basePath . pathParts[pathParts.Length] . ",1,4,2,4"
        imgEl := Acc.GetRootElement()[imgPath]
        imgVal := Trim(imgEl.Value)

        ; 2. 取得 Series Number (Name)
        ; 路徑邏輯: ...8,4 -> ...10,4,2,4 (注意: Series 跳過了中間的 1,4 層級)
        pathParts[targetIdx] := Integer(pathParts[targetIdx]) + 1 ; 9 -> 10

        basePath := ""
        Loop targetIdx {
            basePath .= pathParts[A_Index] ","
        }

        ; 組合 Series Path
        srsPath := basePath . pathParts[pathParts.Length] . ",2,4"
        srsEl := Acc.GetRootElement()[srsPath]
        srsNameRaw := srsEl.Name ; 例如 "(2) Thorax 5.00 Bl60 S2"

        ; Regex 提取括號內的數字
        srsVal := ""
        if RegExMatch(srsNameRaw, "^\((\d+)\)", &match) {
            srsVal := match[1]
        } else {
            ; 如果格式不符，嘗試直接抓取或報錯，這裡暫時回傳原始值或空
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

    ; --- 功能按鈕區 ---
    MyGui.SetFont("s11 Bold", "Segoe UI")
    MyGui.Add("Text", "w320 Center", "Nodule Tracker")

    MyGui.SetFont("s9 Norm", "Segoe UI")
    btnX := (320 - 220) / 2
    btnCopy := MyGui.Add("Button", "x" btnX " w100 h30", "Copy Report")
    btnCopy.OnEvent("Click", CopyReport)

    btnClear := MyGui.Add("Button", "x+20 w100 h30", "Clear All")
    btnClear.OnEvent("Click", ClearAll)

    MyGui.Add("Text", "x10 y+10 w300 h1 0x10")

    ; 排序資料 (Sort)
    For key, arr in NoduleData {
        SortNoduleData(arr)
    }

    ; ========================================================
    ; 雙欄排版區
    ; ========================================================

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

    ; 部位標題
    MyGui.SetFont("s10 Bold", "Segoe UI")
    MyGui.Add("Text", "x" xPos " y+5 w" COL_WIDTH " Center c003366", label)

    items := NoduleData[label]

    if (items.Length == 0) {
        MyGui.SetFont("s9 Norm cGray", "Segoe UI")
        MyGui.Add("Text", "xp y+2 w" COL_WIDTH " Center", "-")
    } else {
        MyGui.SetFont("s10 Norm", "Segoe UI")
        For index, item in items {
            ; 顯示格式: Srs/Img
            displayText := item.srs . "/" . item.img

            textX := xPos + (COL_WIDTH / 2) - 35 ; 稍微往左移一點留給長數字
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

            ; 分組：依據 Series Number
            ; seriesMap 結構: Map("2", [3, 5, 8], "4", [33, 35])
            seriesMap := Map()

            For item in items {
                s := item.srs
                i := item.img
                if !seriesMap.Has(s)
                    seriesMap[s] := []
                seriesMap[s].Push(i)
            }

            ; 處理 Series 順序 (取出 Key 並排序)
            srsKeys := []
            For k, v in seriesMap {
                srsKeys.Push(k)
            }
            ; 簡單氣泡排序 Series Key
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

            ; 組合該肺葉的字串： 2/3,5,8; 4/33,35
            lobeStr := ""
            For sKey in srsKeys {
                imgArr := seriesMap[sKey]
                ; 排序 Image
                imgStr := ""
                For val in imgArr ; imgArr 已經是依據 SortNoduleData 排好的，但為了保險可再排一次，這裡省略
                    imgStr .= val . ","
                imgStr := Trim(imgStr, ",")

                lobeStr .= sKey . "/" . imgStr . "; "
            }
            lobeStr := Trim(lobeStr, "; ")

            fullLobeReport := label . " (Srs/Img: " . lobeStr . ")"
            reportParts.Push(fullLobeReport)
        }
    }

    if (reportParts.Length == 0) {
        ShowTip("! 無資料可複製", 2000)
        return
    }

    ; 組合最終報告 (Oxford Comma)
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

/**
 * 對 NoduleData 的陣列進行排序
 * 規則：先比 Series (數值)，再比 Image (數值)
 */
SortNoduleData(arr) {
    if (arr.Length < 2)
        return

    ; 簡單氣泡排序 (Bubble Sort) 適用於少量資料
    Loop arr.Length {
        i := A_Index
        Loop arr.Length - i {
            j := A_Index
            item1 := arr[j]
            item2 := arr[j+1]

            s1 := Integer(item1.srs)
            s2 := Integer(item2.srs)
            i1 := Integer(item1.img)
            i2 := Integer(item2.img)

            swap := false
            if (s1 > s2) {
                swap := true
            } else if (s1 == s2) {
                if (i1 > i2)
                    swap := true
            }

            if (swap) {
                temp := arr[j]
                arr[j] := arr[j+1]
                arr[j+1] := temp
            }
        }
    }
}

/**
 * 當視窗關閉時觸發：記憶位置並清除資料
 */
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