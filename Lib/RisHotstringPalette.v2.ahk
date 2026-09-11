#Requires AutoHotkey v2.0

#Include RisDialog.v2.ahk
#Include Paste.v2.ahk

/**
 * RIS Hotstring 即時檢索浮動命令列 (Command Palette)
 * 支援熱字關鍵字即時搜尋、HTML 即時 Highlighting、多行模板即時預覽、Enter 直接貼入報告
 */
class RisHotstringPalette {
    static _cache := []
    static _isLoaded := false
    static _gui := 0
    static _editSearch := 0
    static _browser := 0
    static _doc := 0
    static _filteredItems := []
    static _searchTerms := []
    static _selectedIndex := 1
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
                this._browser := 0
                this._doc := 0
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

        guiHeight := 592
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
        this._browser := 0
        this._doc := 0

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

        for _, item in this._cache {
            if (searchTerms.Length == 0) {
                snippet := this._GetDefaultSnippet(item)
                matched.Push({ item: item, snippet: snippet, score: 0 })
                if (matched.Length >= 100) {
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
            if (matched.Length > 100) {
                matched.Length := 100
            }
        }

        this._filteredItems := matched
        this._searchTerms := searchTerms
        this._selectedIndex := 1
        this._UpdateHtml()
    }

    /**
     * 送出當前選取的項目並貼入原視窗
     * @param {Integer} targetRow 指定選取的列 (預設 0 代表使用當前選中項)
     */
    static SubmitSelection(targetRow := 0) {
        if (!this._gui) {
            return
        }

        row := (targetRow > 0) ? targetRow : this._selectedIndex
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

        nextRow := (this._selectedIndex >= this._filteredItems.Length) ? this._filteredItems.Length : (this._selectedIndex + 1)
        this._SetSelectedIndex(nextRow)
    }

    /**
     * 選擇上一個項目
     */
    static SelectPrev() {
        if (!this._gui || this._filteredItems.Length == 0) {
            return
        }

        prevRow := (this._selectedIndex <= 1) ? 1 : (this._selectedIndex - 1)
        this._SetSelectedIndex(prevRow)
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

        ; 搜尋輸入框 (原生 Win32 Edit，輸入法 0 延遲)
        editSearch := g.Add("Edit", Format("x16 y14 w{1} h36 -E0x200 Border", contentWidth))
        editSearch.SetFont("s13", "Microsoft JhengHei UI")
        editSearch.OnEvent("Change", ObjBindMethod(this, "_OnSearchChange"))
        this._editSearch := editSearch

        ; ActiveX HTML 容器 (承載結果清單與完整預覽)
        browserCtrl := g.Add("ActiveX", Format("x16 y56 w{1} h526", contentWidth), "Shell.Explorer")
        this._browser := browserCtrl.Value
        try {
            this._browser.Silent := true
        }
        this._InitBrowserHtml()

        this._gui := g
    }

    static _InitBrowserHtml() {
        browser := this._browser
        browser.Navigate("about:blank")
        while (browser.Busy || browser.ReadyState < 4) {
            Sleep 10
        }

        monoFont := this._GetReportMonospaceFont()

        html := "<!doctype html><html><head><meta http-equiv='X-UA-Compatible' content='IE=edge'>"
            . "<meta charset='utf-8'><style>"
            . "* { box-sizing: border-box; }"
            . "html, body { margin:0; padding:0; background:#F4F5F7; color:#1E293B; font-family:'" monoFont "','Cascadia Code','Consolas',monospace; font-size:14px; user-select:none; overflow:hidden; }"
            . "#container { height:522px; padding:0; }"
            . "#resultsList { height:220px; overflow-y:auto; background:#FFFFFF; border:1px solid #CBD5E1; border-radius:4px; }"
            . ".item { display:block; padding:3px 8px; border-left:4px solid transparent; border-bottom:1px solid #F1F5F9; cursor:pointer; font-size:14px; line-height:1.35; white-space:nowrap; overflow:hidden; text-overflow:ellipsis; }"
            . ".item:hover { background:#F8FAFC; }"
            . ".item.selected { background:#E2E8F0; border-left:4px solid #0F766E; font-weight:bold; }"
            . ".trigger { display:inline-block; width:170px; color:#0F766E; font-weight:bold; vertical-align:top; overflow:hidden; text-overflow:ellipsis; white-space:nowrap; }"
            . ".snippet { display:inline; color:#334155; padding-left:8px; }"
            . ".preview-header { margin:8px 0 4px; font-weight:bold; font-size:13px; color:#475569; overflow:hidden; }"
            . ".preview-title { float:left; }"
            . ".preview-count { float:right; color:#64748B; font-weight:normal; }"
            . "#previewBox { height:250px; background:#FFFFFF; border:1px solid #CBD5E1; border-radius:4px; padding:8px 12px; overflow-y:auto; white-space:pre-wrap; word-wrap:break-word; word-break:break-word; font-family:'" monoFont "','Cascadia Code','Consolas',monospace; font-size:14px; line-height:1.5; color:#1E293B; clear:both; }"
            . ".hint { margin:8px 0 0; font-size:12px; color:#64748B; text-align:center; font-family:'Microsoft JhengHei UI',sans-serif; }"
            . ".hl { background-color:#FEF08A; color:#854D0E; font-weight:bold; padding:0 2px; border-radius:2px; }"
            . ".no-results { padding:28px 16px; text-align:center; color:#94A3B8; font-style:italic; }"
            . "</style>"
            . "<script>"
            . "function selectIndex(idx) {"
            . "  var items = document.getElementsByTagName('div');"
            . "  for (var i = 0; i < items.length; i++) {"
            . "    if (items[i].id && items[i].id.indexOf('item_') === 0) {"
            . "      items[i].className = 'item';"
            . "    }"
            . "  }"
            . "  var curr = document.getElementById('item_' + idx);"
            . "  if (curr) {"
            . "    curr.className = 'item selected';"
            . "    if (curr.scrollIntoView) {"
            . "      curr.scrollIntoView(false);"
            . "    }"
            . "  }"
            . "}"
            . "</script>"
            . "</head><body>"
            . "<div id='container'>"
            . "  <div id='resultsList'></div>"
            . "  <div class='preview-header'>"
            . "    <span class='preview-title'>📄 完整模板即時預覽 (Full Template Preview)</span>"
            . "    <span id='statusCount' class='preview-count'></span>"
            . "  </div>"
            . "  <div id='previewBox'></div>"
            . "  <div class='hint'>[Enter] 貼入報告    |    [↑ / ↓] 切換選取    |    [Esc] 關閉</div>"
            . "</div></body></html>"

        doc := browser.Document
        doc.Open()
        doc.Write(html)
        doc.Close()

        doc.parentWindow.ahkSelect := ObjBindMethod(this, "_OnHtmlSelect")
        doc.parentWindow.ahkSubmit := ObjBindMethod(this, "_OnHtmlSubmit")
        this._doc := doc
    }

    static _OnSearchChange(*) {
        if (!this._gui) {
            return
        }
        this.Filter(this._editSearch.Value)
    }

    static _OnHtmlSelect(idx) {
        if (!this._gui || idx <= 0 || idx > this._filteredItems.Length) {
            return
        }
        this._SetSelectedIndex(idx)
    }

    static _OnHtmlSubmit(idx) {
        this.SubmitSelection(idx)
    }

    static _SetSelectedIndex(idx) {
        if (!this._doc || this._filteredItems.Length == 0) {
            return
        }

        this._selectedIndex := idx
        try {
            this._doc.parentWindow.selectIndex(idx)
            box := this._doc.getElementById("previewBox")
            box.innerHTML := this._GetPreviewHtml(idx)
            box.scrollTop := 0
        }
    }

    static _UpdateHtml() {
        if (!this._doc) {
            return
        }

        count := this._filteredItems.Length
        terms := this._searchTerms
        listHtml := ""

        if (count == 0) {
            listHtml := "<div class='no-results'>(查無符合關鍵字的 Hotstring)</div>"
            try {
                this._doc.getElementById("resultsList").innerHTML := listHtml
                this._doc.getElementById("statusCount").innerText := "共 0 項結果"
                box := this._doc.getElementById("previewBox")
                box.innerHTML := "<span style='color:#94A3B8;font-style:italic;'>(查無符合關鍵字的 Hotstring)</span>"
                box.scrollTop := 0
            }
            return
        }

        for idx, entry in this._filteredItems {
            selectedClass := (idx == 1) ? " selected" : ""
            hlTrigger := this._Highlight(entry.item.trigger, terms)
            hlSnippet := this._Highlight(entry.snippet, terms)

            listHtml .= Format(
                "<div class='item{1}' id='item_{2}' onclick='window.ahkSelect({2})' ondblclick='window.ahkSubmit({2})'>"
                . "<span class='trigger'>{3}</span>"
                . "<span class='snippet'>{4}</span>"
                . "</div>",
                selectedClass, idx, hlTrigger, hlSnippet
            )
        }

        try {
            this._doc.getElementById("resultsList").innerHTML := listHtml
            this._doc.getElementById("statusCount").innerText := Format("共 {1} 項結果", count)
            box := this._doc.getElementById("previewBox")
            box.innerHTML := this._GetPreviewHtml(1)
            box.scrollTop := 0
        }
    }

    static _GetPreviewHtml(idx) {
        if (idx <= 0 || idx > this._filteredItems.Length) {
            return ""
        }
        item := this._filteredItems[idx].item
        html := this._Highlight(item.replacement, this._searchTerms)
        html := StrReplace(html, "`r`n", "<br>")
        html := StrReplace(html, "`n", "<br>")
        html := StrReplace(html, "`r", "<br>")
        html := StrReplace(html, "`t", "&nbsp;&nbsp;&nbsp;&nbsp;")
        html := StrReplace(html, "  ", "&nbsp;&nbsp;")
        html := StrReplace(html, "<br> ", "<br>&nbsp;")
        if (SubStr(html, 1, 1) == " ") {
            html := "&nbsp;" . SubStr(html, 2)
        }
        return html
    }

    static _EscapeHtml(text) {
        text := StrReplace(text, "&", "&amp;")
        text := StrReplace(text, "<", "&lt;")
        text := StrReplace(text, ">", "&gt;")
        text := StrReplace(text, '"', "&quot;")
        return text
    }

    static _Highlight(rawText, searchTerms) {
        if (searchTerms.Length == 0) {
            return this._EscapeHtml(rawText)
        }

        sortedTerms := this._SortTermsByLength(searchTerms)
        patternParts := []
        for _, term in sortedTerms {
            if (term != "") {
                patternParts.Push(this._EscapeRegEx(term))
            }
        }
        if (patternParts.Length == 0) {
            return this._EscapeHtml(rawText)
        }

        fullPattern := "i)(" . this._Join(patternParts, "|") . ")"
        result := ""
        lastPos := 1

        while RegExMatch(rawText, fullPattern, &m, lastPos) {
            matchPos := m.Pos(0)
            matchLen := m.Len(0)

            if (matchPos > lastPos) {
                result .= this._EscapeHtml(SubStr(rawText, lastPos, matchPos - lastPos))
            }

            result .= "<span class='hl' style='background-color:#FEF08A;color:#854D0E;font-weight:bold;padding:0 2px;'>" . this._EscapeHtml(m[0]) . "</span>"
            lastPos := matchPos + matchLen
        }

        if (lastPos <= StrLen(rawText)) {
            result .= this._EscapeHtml(SubStr(rawText, lastPos))
        }

        return result
    }

    static _SortTermsByLength(terms) {
        sorted := []
        for _, t in terms {
            sorted.Push(t)
        }
        len := sorted.Length
        i := 2
        while (i <= len) {
            key := sorted[i]
            j := i - 1
            while (j >= 1 && StrLen(sorted[j]) < StrLen(key)) {
                sorted[j + 1] := sorted[j]
                j -= 1
            }
            sorted[j + 1] := key
            i += 1
        }
        return sorted
    }

    static _Join(arr, delimiter := "") {
        res := ""
        for i, item in arr {
            res .= (i > 1 ? delimiter : "") . item
        }
        return res
    }

    static _EscapeRegEx(str) {
        return RegExReplace(str, "([\\.\$\*\+\?\(\)\[\]\{\}\|\^])", "\$1")
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
