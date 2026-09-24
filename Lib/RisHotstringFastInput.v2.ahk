#Requires AutoHotkey v2.0

#Include Paste.v2.ahk
#Include RisHotstringPalette.v2.ahk

/**
 * RIS Hotstring 快速輸入/貼上加速器 (Fast Paste Accelerator)
 * 啟動時掃描 Hotstrings 目錄，將單行靜態文字熱字動態升級為使用 Paste() 輸出，
 * 大幅消除長字串逐字鍵盤模擬的延遲與卡頓。
 */
class RisHotstringFastInput {
    static isEnabled := false
    static minLength := 20
    static conditionExpr := "IsAnyRisReportWindow()"
    static _originalHotstrings := Map()
    static _count := 0

    /**
     * 建立獨立的閉包回呼函式，避免迴圈變數捕捉衝突
     * @param {String} text 要貼上的文字
     * @returns {Func} 熱字觸發時呼叫的函式
     */
    static _MakePasteCallback(text) {
        return (*) => Paste(text)
    }

    /**
     * 規格化熱字觸發前綴 (確保選項與觸發字之間包含成對的冒號，例如 "::livok1" 或 ":c:RUL")
     * @param {String} options 剖析所得之選項字串 (如 ":" 或 ":c")
     * @param {String} trigger 觸發熱字 (如 "livok1")
     * @returns {String} 合法之 Hotstring 註冊字串
     */
    static _FormatFullTrigger(options, trigger) {
        opt := options
        if (StrSplit(opt, ":").Length - 1 == 1) {
            opt .= ":"
        }
        return opt . trigger
    }

    /**
     * 檢查熱字替換文字是否包含 AHK 特殊按鍵指令 (例如 {Left 11}, {Enter}, {Tab} 等)
     * @param {Object} item Hotstring 快取物件
     * @returns {Boolean} 若包含特殊按鍵指令回傳 true
     */
    static _HasKeyCommands(item) {
        ; 若已指定 Raw (R) 或 Text (T) 模式，則大括號視為純文字
        if (RegExMatch(item.options, "i)[RT]")) {
            return false
        }
        return RegExMatch(item.replacement, "\{[^}]+\}") > 0
    }

    /**
     * 啟用快速貼上加速
     * @param {Integer} minLength 觸發 Paste 的最小字串長度，預設 20 字元 (短於此長度維持原生鍵盤模擬)
     * @param {String} condition HotIf 條件表達式，留空時使用既有設定或預設 "IsAnyRisReportWindow()"
     * @returns {Integer} 成功加速的熱字數量
     */
    static Enable(minLength := 20, condition := "") {
        this.minLength := minLength
        if (condition != "") {
            this.conditionExpr := condition
        } else if (this.conditionExpr == "") {
            this.conditionExpr := "IsAnyRisReportWindow()"
        }
        condition := this.conditionExpr

        if (!RisHotstringPalette._isLoaded) {
            RisHotstringPalette._LoadCache()
        }

        if (condition != "") {
            HotIf condition
        } else {
            HotIf()
        }

        this._count := 0
        for item in RisHotstringPalette._cache {
            ; 略過函式呼叫與多行區塊 (多行已有自訂 Paste 邏輯)
            if (item.isFunction || item.isMultiLine) {
                continue
            }

            ; 略過包含特殊按鍵指令的熱字 (如 {Left 11}, {Enter}, {Tab} 等)
            if (this._HasKeyCommands(item)) {
                continue
            }

            ; 略過短字串 (除非包含換行符號)
            if (StrLen(item.replacement) < minLength && !InStr(item.replacement, "`n")) {
                continue
            }

            fullTrigger := this._FormatFullTrigger(item.options, item.trigger)
            rep := item.replacement

            ; 記錄原始文字以供 Disable 還原
            if (!this._originalHotstrings.Has(fullTrigger)) {
                this._originalHotstrings[fullTrigger] := rep
            }

            try {
                Hotstring(fullTrigger, this._MakePasteCallback(rep))
                this._count++
            }
        }

        HotIf()
        this.isEnabled := true
        return this._count
    }

    /**
     * 停用快速貼上加速，還原為原始鍵盤文字輸出
     * @param {String} condition HotIf 條件表達式
     */
    static Disable(condition := "") {
        if (!this.isEnabled) {
            return
        }

        if (condition == "") {
            condition := this.conditionExpr
        }

        if (condition != "") {
            HotIf condition
        } else {
            HotIf()
        }

        for fullTrigger, originalRep in this._originalHotstrings {
            try {
                Hotstring(fullTrigger, originalRep)
            }
        }

        HotIf()
        this.isEnabled := false
        this._count := 0
    }

    /**
     * 切換啟用/停用狀態
     * @param {String} condition HotIf 條件表達式
     * @returns {Boolean} 切換後的啟用狀態
     */
    static Toggle(condition := "") {
        if (this.isEnabled) {
            this.Disable(condition)
            try {
                TrayTip "Hotstring 快速貼上已停用", "已切換回原生鍵盤模擬模式", 1
            }
        } else {
            cnt := this.Enable(this.minLength, condition)
            try {
                TrayTip "Hotstring 快速貼上已啟用", "已加速 " . cnt . " 個熱字模板 (門檻 >= " . this.minLength . " 字元)", 1
            }
        }
        return this.isEnabled
    }
}
