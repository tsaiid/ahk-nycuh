#Requires AutoHotkey v2.0

#Include RisDialog.v2.ahk
#Include Paste.v2.ahk

/**
 * RIS Hotstring 即時檢索浮動命令列 (Command Palette)
 * 支援熱字關鍵字即時搜尋、多行模板命中行摘要與即時預覽、Enter 直接貼入報告
 */
class RisHotstringPalette {
    static _cache := []
    static _isLoaded := false
    static _gui := 0
    static _editSearch := 0
    static _lvResults := 0
    static _previewBox := 0
    static _lblStatus := 0
    static _filteredItems := []
    static _parentWnd := 0

    static Hwnd => (RisHotstringPalette._gui ? RisHotstringPalette._gui.Hwnd : 0)

    /**
     * 顯示 Hotstring 檢索命令列視窗
     */
    static Show() {
        this._parentWnd := WinActive("A")

        if (!this._isLoaded) {
            this._LoadCache()
        }

        if (this._gui) {
            try {
                WinActivate("ahk_id " . this._gui.Hwnd)
                this._editSearch.Focus()
                return
            } catch {
                this._gui := 0
            }
        }

        targetMonitor := this._GetTargetMonitor()
        mLeft := 0, mTop := 0, mRight := 0, mBottom := 0
        MonitorGetWorkArea targetMonitor, &mLeft, &mTop, &mRight, &mBottom
        monWidth := mRight - mLeft
        monHeight := mBottom - mTop

        guiWidth := Round(monWidth * 2 / 3)
        if (guiWidth < 800) {
            guiWidth := Min(800, monWidth)
        }
        contentWidth := guiWidth - 32

        guiHeight := 605
        if (guiHeight > monHeight - 40) {
            guiHeight := monHeight - 40
        }

        guiX := mLeft + Floor((monWidth - guiWidth) / 2)
        guiY := mTop + Floor((monHeight - guiHeight) / 2)

        this._CreateGui(guiWidth, contentWidth)
        this._editSearch.Value := ""
        this.Filter("")

        this._gui.Show(Format("x{1} y{2} w{3} h{4}", guiX, guiY, guiWidth, guiHeight))
        RisDialog.ApplyWindowStyle(this._gui.Hwnd)
        this._editSearch.Focus()
    }

    /**
     * 關閉檢索視窗並還原焦點
     */
    static Close(*) {
        if (!this._gui) {
            return
        }

        try {
            this._gui.Destroy()
        }
        this._gui := 0

        if (this._parentWnd && WinExist("ahk_id " . this._parentWnd)) {
            try {
                WinActivate("ahk_id " . this._parentWnd)
            }
        }
    }

    /**
     * 強制重新讀取 Hotstrings 目錄下之檔案並重建快取
     */
    static ReloadCache() {
        this._isLoaded := false
        this._cache := []
        this._LoadCache()
    }

    /**
     * 搜尋與過濾核心邏輯
     * @param {String} keyword 關鍵字字串
     */
    static Filter(keyword) {
        kw := Trim(keyword)
        searchTerms := []
        for _, term in StrSplit(kw, " ") {
            t := Trim(term)
            if (t != "") {
                searchTerms.Push(t)
            }
        }

        matched := []
        totalCache := this._cache.Length

        for _, item in this._cache {
            if (searchTerms.Length == 0) {
                snippet := this._GetDefaultSnippet(item)
                matched.Push({ item: item, snippet: snippet, score: 0 })
                if (matched.Length >= 150) {
                    break
                }
                continue
            }

            allMatched := true
            for _, term in searchTerms {
                if (!InStr(item.trigger, term) && !InStr(item.replacement, term)) {
                    allMatched := false
                    break
                }
            }
            if (!allMatched) {
                continue
            }

            snippet := this._FindMatchingSnippet(item, searchTerms)
            score := this._CalculateScore(item, kw, searchTerms)
            matched.Push({ item: item, snippet: snippet, score: score })
        }

        if (searchTerms.Length > 0) {
            this._SortResults(matched)
            if (matched.Length > 150) {
                matched.Length := 150
            }
        }

        this._filteredItems := matched
        this._UpdateListView()
    }

    /**
     * 送出當前選取的項目並貼入原視窗
     */
    static SubmitSelection() {
        if (!this._gui) {
            return
        }

        row := this._lvResults.GetNext(0, "Focused")
        if (row == 0 && this._filteredItems.Length > 0) {
            row := 1
        }
        if (row <= 0 || row > this._filteredItems.Length) {
            return
        }

        selected := this._filteredItems[row].item
        parentWnd := this._parentWnd

        this.Close()

        if (parentWnd && WinExist("ahk_id " . parentWnd)) {
            try {
                WinActivate("ahk_id " . parentWnd)
                WinWaitActive("ahk_id " . parentWnd, , 1)
                Sleep 30
            }
        }

        if (selected.isFunction) {
            SendInput("{Raw}" . selected.trigger . "`t")
        } else {
            Paste(selected.replacement)
        }
    }

    /**
     * 選擇下一個項目
     */
    static SelectNext() {
        if (!this._gui || this._filteredItems.Length == 0) {
            return
        }

        currentRow := this._lvResults.GetNext(0, "Focused")
        nextRow := (currentRow <= 0) ? 1 : (currentRow >= this._filteredItems.Length) ? this._filteredItems.Length : (currentRow + 1)

        this._lvResults.Modify(0, "-Select -Focus")
        this._lvResults.Modify(nextRow, "+Select +Focus +Vis")
        this._UpdatePreview(nextRow)
    }

    /**
     * 選擇上一個項目
     */
    static SelectPrev() {
        if (!this._gui || this._filteredItems.Length == 0) {
            return
        }

        currentRow := this._lvResults.GetNext(0, "Focused")
        prevRow := (currentRow <= 1) ? 1 : (currentRow - 1)

        this._lvResults.Modify(0, "-Select -Focus")
        this._lvResults.Modify(prevRow, "+Select +Focus +Vis")
        this._UpdatePreview(prevRow)
    }

    ; =========================================================================
    ; 內部輔助方法 (Internal Helpers)
    ; =========================================================================

    static _GetReportMonospaceFont() {
        try {
            risCtrl := (%("RisController")%)
            if (HasProp(risCtrl, "EnforcedFontName") && risCtrl.EnforcedFontName != "") {
                return risCtrl.EnforcedFontName
            }
        }
        return "Maple Mono CN"
    }

    static _GetTargetMonitor() {
        if (this._parentWnd && WinExist("ahk_id " . this._parentWnd)) {
            try {
                wx := 0, wy := 0, ww := 0, wh := 0
                WinGetPos(&wx, &wy, &ww, &wh, "ahk_id " . this._parentWnd)
                cx := wx + ww / 2
                cy := wy + wh / 2

                monCount := MonitorGetCount()
                Loop monCount {
                    mLeft := 0, mTop := 0, mRight := 0, mBottom := 0
                    MonitorGetWorkArea A_Index, &mLeft, &mTop, &mRight, &mBottom
                    if (mLeft <= cx && cx <= mRight && mTop <= cy && cy <= mBottom) {
                        return A_Index
                    }
                }
            }
        }
        return GetCurrentMonitorIndex()
    }

    static _CreateGui(guiWidth := 1000, contentWidth := 0) {
        if (contentWidth <= 0) {
            contentWidth := guiWidth - 32
        }

        g := RisDialog.Create("RIS Hotstring Command Palette", "+AlwaysOnTop +ToolWindow -DPIScale")
        g.BackColor := "F4F5F7"
        g.MarginX := 16
        g.MarginY := 14
        g.OnEvent("Escape", ObjBindMethod(this, "Close"))
        g.OnEvent("Close", ObjBindMethod(this, "Close"))

        monoFont := this._GetReportMonospaceFont()

        ; 搜尋輸入框
        editSearch := g.Add("Edit", Format("x16 y14 w{1} h36 -E0x200 Border", contentWidth))
        editSearch.SetFont("s13", "Microsoft JhengHei UI")
        editSearch.OnEvent("Change", ObjBindMethod(this, "_OnSearchChange"))
        this._editSearch := editSearch

        ; 結果清單 (ListView)
        col1Width := 180
        col2Width := Max(200, contentWidth - col1Width - 25)
        lvResults := g.Add("ListView", Format("x16 y58 w{1} h250 -Hdr -Multi Grid", contentWidth), ["縮寫", "內容摘要 / 命中行"])
        lvResults.SetFont("s11", monoFont)
        lvResults.ModifyCol(1, col1Width)
        lvResults.ModifyCol(2, col2Width)
        lvResults.OnEvent("ItemSelect", ObjBindMethod(this, "_OnItemSelect"))
        lvResults.OnEvent("DoubleClick", ObjBindMethod(this, "_OnDoubleClick"))
        this._lvResults := lvResults

        ; 預覽區標籤與狀態
        titleW := Max(200, contentWidth - 200)
        statusX := 16 + contentWidth - 190
        lblTitle := g.Add("Text", Format("x16 y316 w{1} h20 c475569 BackgroundTrans", titleW), "📄 完整模板即時預覽 (Full Template Preview)")
        lblTitle.SetFont("s10 bold", "Microsoft JhengHei UI")

        lblStatus := g.Add("Text", Format("x{1} y316 w190 h20 Right c64748B BackgroundTrans", statusX), "")
        lblStatus.SetFont("s9", "Microsoft JhengHei UI")
        this._lblStatus := lblStatus

        ; 預覽多行內容區塊 (ReadOnly Edit, 自動換行, 無左右滾動)
        previewBox := g.Add("Edit", Format("x16 y340 w{1} h220 ReadOnly +Wrap +VScroll -E0x200 Border", contentWidth))
        previewBox.SetFont("s11", monoFont)
        this._previewBox := previewBox

        ; 底部操作提示
        hint := g.Add("Text", Format("x16 y570 w{1} h20 Center c64748B BackgroundTrans", contentWidth), "[Enter] 貼入報告    |    [↑ / ↓] 切換選取    |    [Esc] 關閉")
        hint.SetFont("s9", "Microsoft JhengHei UI")

        this._gui := g
    }

    static _OnSearchChange(*) {
        if (!this._gui) {
            return
        }
        this.Filter(this._editSearch.Value)
    }

    static _OnItemSelect(ctrl, itemIndex, selected) {
        if (selected && itemIndex > 0) {
            this._UpdatePreview(itemIndex)
        }
    }

    static _OnDoubleClick(*) {
        this.SubmitSelection()
    }

    static _UpdateListView() {
        if (!this._gui) {
            return
        }

        this._lvResults.Opt("-Redraw")
        this._lvResults.Delete()

        for _, entry in this._filteredItems {
            this._lvResults.Add(, entry.item.trigger, entry.snippet)
        }

        this._lvResults.Opt("+Redraw")

        count := this._filteredItems.Length
        this._lblStatus.Value := Format("共 {1} 項結果", count)

        if (count > 0) {
            this._lvResults.Modify(1, "+Select +Focus +Vis")
            this._UpdatePreview(1)
        } else {
            this._previewBox.Value := "(查無符合關鍵字的 Hotstring)"
        }
    }

    static _UpdatePreview(index) {
        if (!this._gui || index <= 0 || index > this._filteredItems.Length) {
            return
        }

        item := this._filteredItems[index].item
        this._previewBox.Value := item.replacement
    }

    static _GetDefaultSnippet(item) {
        if (!item.isMultiLine) {
            return item.replacement
        }
        for _, lineText in item.lines {
            trimmed := Trim(lineText)
            if (trimmed != "") {
                return trimmed
            }
        }
        return item.replacement
    }

    static _FindMatchingSnippet(item, searchTerms) {
        if (!item.isMultiLine) {
            return item.replacement
        }

        for _, lineText in item.lines {
            trimmed := Trim(lineText)
            if (trimmed == "") {
                continue
            }
            for _, term in searchTerms {
                if (InStr(trimmed, term)) {
                    return trimmed
                }
            }
        }

        return this._GetDefaultSnippet(item)
    }

    static _CalculateScore(item, fullKw, searchTerms) {
        score := 0
        trig := item.trigger

        if (trig = fullKw) {
            score += 200
        } else if (InStr(trig, fullKw) == 1) {
            score += 150
        } else if (InStr(trig, fullKw)) {
            score += 100
        }

        for _, term in searchTerms {
            if (InStr(trig, term) == 1) {
                score += 50
            } else if (InStr(trig, term)) {
                score += 30
            } else if (InStr(item.replacement, term)) {
                score += 10
            }
        }

        return score
    }

    static _SortResults(matched) {
        ; 簡易插入排序 (In-place insertion sort by score descending)
        len := matched.Length
        i := 2
        while (i <= len) {
            key := matched[i]
            j := i - 1
            while (j >= 1 && matched[j].score < key.score) {
                matched[j + 1] := matched[j]
                j -= 1
            }
            matched[j + 1] := key
            i += 1
        }
    }

    static _LoadCache() {
        this._cache := []

        baseDir := A_ScriptDir . "\Hotstrings"
        if (!DirExist(baseDir)) {
            baseDir := A_ScriptDir . "\..\Hotstrings"
        }
        if (!DirExist(baseDir)) {
            this._isLoaded := true
            return
        }

        Loop Files, baseDir . "\*.v2.ahk", "R" {
            fileName := A_LoopFileName
            if (fileName = "regex-hotstrings.v2.ahk") {
                continue
            }

            try {
                content := FileRead(A_LoopFileFullPath, "UTF-8")
            } catch {
                continue
            }

            this._ParseFileContent(content, fileName)
        }

        this._isLoaded := true
    }

    static _ParseFileContent(content, fileName) {
        lines := StrSplit(content, "`n", "`r")
        totalLines := lines.Length
        i := 1

        while (i <= totalLines) {
            line := lines[i]

            if (RegExMatch(line, "^\s*(:[^:]*):([^:]+)::(.*)$", &m)) {
                options := m[1]
                trigger := m[2]
                rawRest := RTrim(m[3], "`r`n")
                trimmedCheck := Trim(rawRest)

                ; 單行 Hotstring 判斷
                if (trimmedCheck != "" && SubStr(trimmedCheck, 1, 1) != ";") {
                    commentPos := InStr(rawRest, " `;")
                    if (commentPos > 0) {
                        replacement := RTrim(SubStr(rawRest, 1, commentPos - 1))
                    } else {
                        replacement := rawRest
                    }

                    ; 處理結尾轉義空白符號
                    if (SubStr(replacement, -2) = "`` ") {
                        replacement := SubStr(replacement, 1, -2) . " "
                    } else if (SubStr(replacement, -1) = "``") {
                        replacement := SubStr(replacement, 1, -1)
                    }

                    this._cache.Push({
                        trigger: trigger,
                        options: options,
                        replacement: replacement,
                        lines: [replacement],
                        isMultiLine: false,
                        isFunction: false,
                        file: fileName
                    })
                    i += 1
                    continue
                }

                ; 多行區塊 / 表單 Hotstring 判斷
                j := i + 1
                foundBrace := false
                while (j <= totalLines) {
                    tLine := Trim(lines[j])
                    if (tLine == "" || SubStr(tLine, 1, 1) == ";") {
                        j += 1
                        continue
                    }
                    if (InStr(tLine, "{")) {
                        foundBrace := true
                        break
                    }
                    break
                }

                if (!foundBrace) {
                    i += 1
                    continue
                }

                braceDepth := 0
                blockLines := []
                k := j
                while (k <= totalLines) {
                    curr := lines[k]
                    loop parse, curr {
                        if (A_LoopField == "{") {
                            braceDepth += 1
                        } else if (A_LoopField == "}") {
                            braceDepth -= 1
                        }
                    }
                    blockLines.Push(curr)
                    if (braceDepth <= 0) {
                        break
                    }
                    k += 1
                }

                blockText := ""
                for _, blkLine in blockLines {
                    blockText .= blkLine . "`n"
                }

                replacement := ""
                isFunc := false

                if (RegExMatch(blockText, "s)\r?\n\s*\(\r?\n([\s\S]*?)\r?\n\s*\)", &cm)) {
                    replacement := cm[1]
                } else if (RegExMatch(blockText, 's)MyForm\s*:=\s*"([^"]*)"', &fm)) {
                    replacement := fm[1]
                } else if (RegExMatch(blockText, 's)Paste\("([^"]*)"\)', &pm)) {
                    replacement := pm[1]
                } else if (RegExMatch(blockText, 'm)^\s*([a-zA-Z0-9_]+Form\([^)]*\))', &fnm)) {
                    isFunc := true
                    replacement := "[" . fnm[1] . "]"
                } else {
                    replacement := Trim(blockText)
                }

                repLines := StrSplit(replacement, "`n", "`r")
                cleanedLines := []
                for _, rLine in repLines {
                    cleanedLines.Push(rLine)
                }

                this._cache.Push({
                    trigger: trigger,
                    options: options,
                    replacement: replacement,
                    lines: cleanedLines,
                    isMultiLine: (cleanedLines.Length > 1),
                    isFunction: isFunc,
                    file: fileName
                })

                i := k + 1
                continue
            }

            i += 1
        }
    }
}

#HotIf (RisHotstringPalette.Hwnd && WinActive("ahk_id " . RisHotstringPalette.Hwnd))
Down::RisHotstringPalette.SelectNext()
Up::RisHotstringPalette.SelectPrev()
Enter::RisHotstringPalette.SubmitSelection()
Escape::RisHotstringPalette.Close()
#HotIf
