#Requires AutoHotkey v2.0

class RisHotkeyHelp {
    static _gui := 0
    static _scale := 1.0

    static Show() {
        if this._gui {
            try {
                WinActivate(this._gui.Hwnd)
                return
            }
            this._gui := 0
        }

        ref := RisNotify._ResolveReferencePoint()
        this._scale := RisNotify._GetDpiAtPoint(ref.x, ref.y) / 96
        workArea := RisNotify._GetMonitorWorkAreaAtPoint(ref.x, ref.y)
        width := Min(Round(1220 * this._scale), Round((workArea.right - workArea.left) * 0.96))
        height := Min(Round(780 * this._scale), Round((workArea.bottom - workArea.top) * 0.88))
        x := workArea.left + Floor(((workArea.right - workArea.left) - width) / 2)
        y := workArea.top + Floor(((workArea.bottom - workArea.top) - height) / 2)
        theme := this._ChooseTheme(x, y, width, height)
        colors := this._GetColors(theme)

        g := Gui("+AlwaysOnTop -Caption +ToolWindow -DPIScale")
        g.BackColor := colors.Back
        g.MarginX := 0
        g.MarginY := 0
        g.OnEvent("Escape", ObjBindMethod(this, "Close"))
        g.OnEvent("Close", ObjBindMethod(this, "Close"))
        g.SetFont("s" Round(10 * this._scale) " c" colors.Text, "Microsoft JhengHei UI")

        pad := Round(22 * this._scale)
        titleH := Round(42 * this._scale)
        closeSize := Round(30 * this._scale)
        closeX := width - pad - closeSize
        closeY := Round(14 * this._scale)

        title := g.Add("Text", Format("x{1} y{2} w{3} h{4} BackgroundTrans", pad, Round(16 * this._scale), width - pad * 2 - closeSize, titleH), "NYCU RIS 快速鍵")
        title.SetFont("s" Round(17 * this._scale) " bold c" colors.Text, "Microsoft JhengHei UI")

        closeButton := g.Add("Text", Format("x{1} y{2} w{3} h{3} Center 0x200 BackgroundTrans", closeX, closeY, closeSize), "×")
        closeButton.SetFont("s" Round(18 * this._scale) " c" colors.Muted, "Segoe UI")
        closeButton.OnEvent("Click", ObjBindMethod(this, "Close"))

        contentY := Round(64 * this._scale)
        contentH := height - contentY - pad
        browserCtrl := g.Add("ActiveX", Format("x{1} y{2} w{3} h{4}", pad, contentY, width - pad * 2, contentH), "Shell.Explorer")
        this._WriteHelpHtml(browserCtrl.Value, colors)

        this._gui := g
        g.Show(Format("x{1} y{2} w{3} h{4}", x, y, width, height))
        this._ApplyVisualStyle(g.Hwnd)
    }

    static Close(*) {
        if !this._gui
            return

        try this._gui.Destroy()
        this._gui := 0
    }

    static _ChooseTheme(x, y, width, height) {
        brightness := RisNotify._GetAverageScreenBrightness(x, y, width, height)
        if (brightness = "")
            return "dark"

        return (brightness >= 128) ? "dark" : "light"
    }

    static _GetColors(theme) {
        if (theme = "light") {
            return {
                Back: "F8FAFC",
                TableBack: "FFFFFF",
                Border: "CBD5E1",
                Accent: "0F766E",
                ScrollThumb: "CBD5E1",
                ScrollTrack: "F8FAFC",
                Text: "111827",
                Muted: "475569"
            }
        }

        return {
            Back: "202020",
            TableBack: "111827",
            Border: "334155",
            Accent: "5EEAD4",
            ScrollThumb: "334155",
            ScrollTrack: "202020",
            Text: "FFFFFF",
            Muted: "CBD5E1"
        }
    }

    static _ApplyVisualStyle(hwnd) {
        try {
            WinGetPos(, , &w, &h, hwnd)
            r := Round(18 * this._scale)
            WinSetRegion("0-0 w" w " h" h " r" r "-" r, hwnd)
            WinSetTransparent(246, hwnd)
        }
    }

    static _WriteHelpHtml(browser, colors) {
        browser.Navigate("about:blank")
        while browser.Busy || browser.ReadyState < 4
            Sleep(10)

        doc := browser.Document
        doc.Open()
        doc.Write(this._BuildHtml(colors))
        doc.Close()
    }

    static _BuildHtml(colors) {
        columns := this._BuildColumnsHtml()

        return "<!doctype html><html><head><meta http-equiv='X-UA-Compatible' content='IE=edge'>"
            . "<meta charset='utf-8'><style>"
            . "html,body{margin:0;padding:0;background:#" colors.Back ";color:#" colors.Text ";font-family:'Microsoft JhengHei UI','Segoe UI',sans-serif;font-size:14px;}"
            . "body{overflow-y:auto;scrollbar-face-color:#" colors.ScrollThumb ";scrollbar-track-color:#" colors.ScrollTrack ";scrollbar-arrow-color:#" colors.Muted ";scrollbar-highlight-color:#" colors.ScrollThumb ";scrollbar-shadow-color:#" colors.ScrollThumb ";scrollbar-3dlight-color:#" colors.ScrollThumb ";scrollbar-darkshadow-color:#" colors.ScrollTrack ";}"
            . ".wrap{width:100%;border-collapse:separate;border-spacing:8px 0;table-layout:fixed;padding-bottom:12px;}"
            . ".col{width:25%;vertical-align:top;}"
            . ".section{margin:14px 4px 8px;color:#" colors.Muted ";font-size:12px;font-weight:700;}"
            . ".row{width:100%;margin:6px 2px;padding:0;border:1px solid #" colors.Border ";border-radius:10px;background:#" colors.TableBack ";border-spacing:0;}"
            . ".keys{width:94px;padding:9px 9px;font-family:'Cascadia Mono','Consolas',monospace;font-weight:700;color:#" colors.Accent ";white-space:normal;line-height:1.35;vertical-align:top;}"
            . ".action{padding:9px 10px 9px 0;line-height:1.45;vertical-align:top;}"
            . "</style></head><body><table class='wrap'><tr>" columns "</tr></table></body></html>"
    }

    static _BuildColumnsHtml() {
        hotkeys := this._GetHotkeys()
        columnCount := 4
        perColumn := Ceil(hotkeys.Length / columnCount)
        columns := ""

        loop columnCount {
            startIndex := (A_Index - 1) * perColumn + 1
            endIndex := Min(A_Index * perColumn, hotkeys.Length)
            columns .= "<td class='col'>" this._BuildRowsHtml(hotkeys, startIndex, endIndex) "</td>"
        }

        return columns
    }

    static _BuildRowsHtml(hotkeys, startIndex, endIndex) {
        rows := ""
        lastContext := ""

        loop endIndex - startIndex + 1 {
            item := hotkeys[startIndex + A_Index - 1]
            if (item.Context != lastContext) {
                rows .= "<div class='section'>" this._EscapeHtml(item.Context) "</div>"
                lastContext := item.Context
            }
            rows .= "<table class='row'><tr>"
                . "<td class='keys'>" this._EscapeHtml(item.Keys) "</td>"
                . "<td class='action'>" this._EscapeHtml(item.Action) "</td>"
                . "</tr></table>"
        }

        return rows
    }

    static _EscapeHtml(value) {
        value := StrReplace(value, "&", "&amp;")
        value := StrReplace(value, "<", "&lt;")
        value := StrReplace(value, ">", "&gt;")
        value := StrReplace(value, '"', "&quot;")
        return value
    }

    static _GetHotkeys() {
        return [
            {Context: "通用", Keys: "Alt+H", Action: "顯示快速鍵說明"},
            {Context: "通用", Keys: "Ctrl+Alt+R", Action: "重新載入腳本"},
            {Context: "通用", Keys: "Win+Ctrl+R", Action: "重新載入相似檢查分組設定"},
            {Context: "通用", Keys: "SC07B", Action: "映射為滑鼠左鍵"},
            {Context: "任何 RIS 報告", Keys: "Ctrl+W", Action: "刪除前一個字"},
            {Context: "標準 RIS 報告", Keys: "Alt+Up / Alt+Down", Action: "移動目前行"},
            {Context: "標準 RIS 報告", Keys: "滑鼠左鍵", Action: "處理三連點選取"},
            {Context: "標準 RIS 報告", Keys: "Ctrl+1", Action: "切到歷史資料，顯示全部"},
            {Context: "標準 RIS 報告", Keys: "Ctrl+2", Action: "切到歷史資料，依 modality 篩選"},
            {Context: "標準 RIS 報告", Keys: "Ctrl+3", Action: "切到 clinical data 第 2 頁"},
            {Context: "標準 RIS 報告", Keys: "Ctrl+4", Action: "切到病理線上資料"},
            {Context: "標準 RIS 報告", Keys: "Ctrl+Alt+2", Action: "選取 clinical data 第 2 分頁"},
            {Context: "標準 RIS 報告", Keys: "Ctrl+Alt+4", Action: "選取 clinical data 第 4 分頁"},
            {Context: "標準 RIS 報告", Keys: "Ctrl+Esc", Action: "附加上一份報告"},
            {Context: "標準 RIS 報告", Keys: "Alt+C", Action: "取消 AutoNext 並存檔"},
            {Context: "標準 RIS 報告", Keys: "Ctrl+S", Action: "勾選 AutoNext 並存檔"},
            {Context: "標準 RIS 報告", Keys: "Alt+Q", Action: "送出 Ctrl+E"},
            {Context: "標準 RIS 報告", Keys: "Alt+E", Action: "插入檢查名稱"},
            {Context: "標準 RIS 報告", Keys: "Alt+Shift+E", Action: "插入檢查名稱並產生 AI indication"},
            {Context: "標準 RIS 報告", Keys: "Alt+Shift+D", Action: "插入目前選取的歷史報告日期"},
            {Context: "標準 RIS 報告", Keys: "Alt+D", Action: "插入已複製的歷史報告日期"},
            {Context: "標準 RIS 報告", Keys: "Ctrl+Alt+E", Action: "插入目前選取的歷史報告名稱"},
            {Context: "標準 RIS 報告", Keys: "Ctrl+Shift+C", Action: "複製病理報告或病歷號"},
            {Context: "標準 RIS 報告", Keys: "Alt+Esc", Action: "搜尋並選取相似歷史報告"},
            {Context: "標準 RIS 報告", Keys: "SC079 / Ctrl+Alt+,", Action: "格式化 Findings"},
            {Context: "標準 RIS 報告", Keys: "SC070 / Ctrl+Alt+.", Action: "格式化 Impression"},
            {Context: "標準 RIS 報告", Keys: "Win+V", Action: "複製 Finding 到 Impression"},
            {Context: "RIS 輸入框", Keys: "Alt+I", Action: "產生並插入 AI indication"},
            {Context: "RIS 輸入框", Keys: "Alt+S", Action: "產生並插入 AI impression"},
            {Context: "RIS 輸入框", Keys: "Alt+Shift+S", Action: "重新產生並插入 AI impression"},
            {Context: "RIS 輸入框", Keys: "Alt+R", Action: "比較 OpenAI 與 Google AI 潤色結果"},
            {Context: "RIS 輸入框", Keys: "Alt+Shift+R", Action: "AI 潤色選取文字"},
            {Context: "RIS 輸入框", Keys: "Ctrl+A", Action: "移到行首"},
            {Context: "RIS 輸入框", Keys: "Ctrl+E", Action: "移到行尾"},
            {Context: "RIS 輸入框", Keys: "Alt+F / Alt+B", Action: "向右或向左移動一個字"},
            {Context: "RIS 輸入框", Keys: "Win+D", Action: "清空目前輸入框"},
            {Context: "RIS 輸入框", Keys: "Ctrl+C", Action: "複製選取文字或整行"},
            {Context: "RIS 輸入框", Keys: "Ctrl+X", Action: "剪下選取文字或整行"},
            {Context: "RIS 輸入框", Keys: "Ctrl+Y", Action: "刪除目前行"},
            {Context: "RIS 輸入框", Keys: "Ctrl+K", Action: "刪除游標後方文字"},
            {Context: "RIS 輸入框", Keys: "Shift+Up / Shift+Down", Action: "智慧延伸選取"},
            {Context: "RIS 輸入框", Keys: "Ctrl+Alt+O", Action: "重排選取文字"},
            {Context: "RIS 輸入框", Keys: "Ctrl+Alt+Shift+O", Action: "重排選取文字並保留 Series/Image 標記"},
            {Context: "RIS 輸入框", Keys: "Ctrl+Shift+*", Action: "移除編號並改用 *"},
            {Context: "RIS 輸入框", Keys: "Ctrl+Shift+-", Action: "移除編號並改用 -"},
            {Context: "RIS 輸入框", Keys: "Ctrl+Alt+Shift+-", Action: "移除編號並自動偵測項目符號"},
            {Context: "RIS 輸入框", Keys: "Ctrl+Shift+>", Action: "移除編號並改用 >"},
            {Context: "RIS 輸入框", Keys: "Ctrl+D", Action: "送出 Delete"},
            {Context: "RIS 輸入框", Keys: "Win+A", Action: "全選"},
            {Context: "RIS 輸入框", Keys: "Ctrl+Up / Ctrl+Down", Action: "智慧翻頁移動"},
            {Context: "RIS 輸入框", Keys: "Ctrl+Shift+Up / Ctrl+Shift+Down", Action: "智慧翻頁移動並延伸選取"},
            {Context: "RIS 輸入框", Keys: "Ctrl+Enter", Action: "在下方插入新行"},
            {Context: "RIS 輸入框", Keys: "Shift+Enter", Action: "在上方插入新行"},
            {Context: "RIS 輸入框", Keys: "XButton1", Action: "自動偵測項目符號並重排選取文字"},
            {Context: "RIS 輸入框", Keys: "XButton2", Action: "重排選取文字"},
            {Context: "RIS 智慧清單", Keys: "Enter / NumpadEnter", Action: "智慧清單換行"},
            {Context: "RIS 智慧清單", Keys: "Backspace", Action: "智慧清單退格"},
            {Context: "鍵盤配置", Keys: "F2", Action: "顯示目前鍵盤語言 ID"},
            {Context: "US 鍵盤", Keys: "Right Alt", Action: "啟用或切換 RIS 焦點"},
            {Context: "日文鍵盤", Keys: "SC029", Action: "啟用或切換 RIS 焦點"},
            {Context: "危急值視窗", Keys: "Alt+1..Alt+4", Action: "點擊對應位置按鈕"},
            {Context: "危急值視窗", Keys: "Alt+S", Action: "點擊 Save"},
            {Context: "危急值視窗", Keys: "Esc", Action: "點擊 Cancel"},
            {Context: "會診視窗", Keys: "Ctrl+T", Action: "補上 20 分鐘會診時間"}
        ]
    }
}
