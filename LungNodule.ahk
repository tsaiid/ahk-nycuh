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
        imgNum := GetImageNumberFromFocus()

        if (imgNum == "") {
            ShowTip("⚠️ 抓取失敗 (數值為空)", 2000)
            return
        }

        For existingNum in NoduleData[location] {
            if (existingNum == imgNum) {
                ShowTip("⚠️ 已存在: " imgNum " (忽略)", 1000)
                return
            }
        }

        NoduleData[location].Push(imgNum)
        UpdateGUI()

        ShowTip("✅ " location ": " imgNum, 1000)

    } catch Error as e {
        ShowTip("❌ 錯誤: " e.Message, 2000)
    }
}

ShowTip(msg, duration) {
    ToolTip(msg)
    SetTimer(() => ToolTip(), -duration)
}

GetImageNumberFromFocus() {
    try {
        focusedEl := Acc.ElementFromPoint()
        if !focusedEl {
            throw Error("無法偵測到元素 (請確認滑鼠位置)")
        }

        fullPath := GetFullPath(focusedEl)
        pathParts := StrSplit(fullPath, ",")

        if (pathParts.Length < 2)
            throw Error("路徑層級不足")

        targetIdx := pathParts.Length - 1
        pathParts[targetIdx] := Integer(pathParts[targetIdx]) + 1

        newPathBase := ""
        Loop targetIdx {
            newPathBase .= pathParts[A_Index] ","
        }

        finalPath := newPathBase . pathParts[pathParts.Length] . ",1,4,2,4"

        targetEl := Acc.GetRootElement()[finalPath]
        return Trim(targetEl.Value)
    }
    catch {
        return ""
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

    ; ★ 新增這一行：當使用者關閉視窗 (按 X) 時，觸發 WindowClosed 函數
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
        SortArrayNumeric(arr)
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

    nums := NoduleData[label]

    if (nums.Length == 0) {
        MyGui.SetFont("s9 Norm cGray", "Segoe UI")
        MyGui.Add("Text", "xp y+2 w" COL_WIDTH " Center", "-")
    } else {
        MyGui.SetFont("s10 Norm", "Segoe UI")
        For index, num in nums {
            textX := xPos + (COL_WIDTH / 2) - 25
            MyGui.Add("Text", "x" textX " y+5 w30 Right", num)
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
        nums := NoduleData[label]
        if (nums.Length > 0) {
            ; ★ 修改 3: 產生報告前確保數值由小到大排序
            SortArrayNumeric(nums)

            numStr := ""
            For n in nums {
                numStr .= n . ","
            }
            numStr := Trim(numStr, ",")
            part := label . " (Srs/Img: 4/" . numStr . ")"
            reportParts.Push(part)
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
        finalStr := Trim(finalStr, ", ") . ", and " . reportParts[reportParts.Length]
    }

    A_Clipboard := finalStr
    ShowTip("Copied:`n" finalStr, 3000)
}

/**
 * 輔助函數：將陣列內容進行數值排序 (由小到大)
 */
SortArrayNumeric(arr) {
    if (arr.Length < 2)
        return

    ; 將陣列轉為換行字串
    str := ""
    For item in arr
        str .= item . "`n"

    ; 使用 AHK 的 Sort 函數，選項 N 代表數值排序，D`n 代表以換行分隔
    sortedStr := Sort(Trim(str, "`n"), "N D`n")

    ; 清空原陣列並填回排序後的數值
    arr.Length := 0
    Loop Parse, sortedStr, "`n"
        arr.Push(A_LoopField)
}

/**
 * 當視窗關閉時觸發：記憶位置並清除資料
 */
WindowClosed(*) {
    ; 1. 趁視窗還沒完全消失，趕快記憶最後的位置
    try {
        if (MyGui && WinExist("ahk_id " MyGui.Hwnd)) {
            WinGetPos(&currentX, &currentY,,, "ahk_id " MyGui.Hwnd)
            global GuiX := currentX
            global GuiY := currentY
        }
    }

    ; 2. 清空資料 (不呼叫 UpdateGUI，因為視窗正要關閉)
    For key, arr in NoduleData {
        NoduleData[key] := []
    }
}