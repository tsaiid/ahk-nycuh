#Requires AutoHotkey v2.0

class RisReportText {
    static FindContentRange(text, mode) {
        startPos := 0
        endPos := -1

        if (mode == "Advanced") {
            ; CT/MR 的起始關鍵字
            if RegExMatch(text, "m)FINDINGS:[ \t]*(?:\r?\n)?|The study shows:\r?\n\r?\n|show the following findings:\r?\n\r?\n|which revealed:\r?\n\r?\n", &match) {
                startPos := match.Pos + match.Len - 1

                ; CT/MR 特有的結尾偵測 (REMARKS/RECOMMENDATION)
                if RegExMatch(text, "im)^[ \t]*(REMARKS?:|RECOMMENDATION:)", &endMatch, startPos + 1) {
                    endPos := endMatch.Pos - 1
                    return {Start: startPos, End: endPos, TrailingNewlines: "`r`n`r`n"}
                }
                return {Start: startPos, End: endPos}
            }

            ; HCC staging form: format only the free-text Other findings section.
            if RegExMatch(text, "im)^\s*\d+\.\s*Other findings\s*:?[ \t]*(?:\r?\n)?", &match) {
                startPos := match.Pos + match.Len - 1

                if RegExMatch(text, "m)^[ \t]*=+[ \t]*$", &endMatch, startPos + 1) {
                    endPos := endMatch.Pos - 1
                    return {Start: startPos, End: endPos, TrailingNewlines: "`r`n`r`n"}
                }
                return {Start: startPos, End: endPos}
            }
        } else {
            ; Basic (CR/US) 的起始關鍵字
            if RegExMatch(text, "m)FINDINGS:\r?\n|:\s*\r?\n\s*\r?\n", &match) {
                startPos := match.Pos + match.Len - 1
                return {Start: startPos, End: -1} ; Basic 預設選到最後
            }
        }

        return false
    }

    static DeidentifyText(text) {
        if (text == "") {
            return ""
        }

        ; 1. 處理日期：民國/西元 (115/03/13, 2026-03-13, 115.3.13)
        text := RegExReplace(text, "i)(\d{3,4})[/\.\-]\d{1,2}[/\.\-]\d{1,2}", "[DATE]")

        ; 2. 處理連寫日期：(1150313), 20260313
        text := RegExReplace(text, "i)\(?\b(\d{7,8})\b\)?", "([DATE])")

        ; 3. 處理年齡：將 "29 歲 9 月" 轉換為 "20s-yo"
        ; 保留大致年齡段對放射科診斷有臨床價值，但移除精確月分
        if (RegExMatch(text, "(\d{1,2})\s*(歲|y|years?)", &m)) {
            matchedAge := Number(m[1])
            decade := Floor(matchedAge / 10) * 10
            text := RegExReplace(text, "\d{1,2}\s*(歲|y|years?)\s*(\d{1,2}\s*(月|m|months?))?", decade . "s-yo")
        }

        ; 4. 處理身分證字號 (台灣格式：首位字母 + 9位數字，涵蓋本國人 1/2 與外籍人士 8/9)
        text := RegExReplace(text, "i)[A-Z][1289]\d{8}", "[ID_REDACTED]")

        ; 5. 處理電話號碼 (09xx-xxx-xxx 或 02-xxxx-xxxx)
        text := RegExReplace(text, "i)0\d{1,2}-?\d{3,4}-?\d{3,4}", "[PHONE_REDACTED]")

        ; 6. 移除姓名標籤後的內容 (處理到行尾)
        text := RegExReplace(text, "i)(Name|姓名|Patient)\s*[:：]\s*\V+", "$1: [PATIENT_NAME]")

        return text
    }

    static GetExamType(examName) {
        if (InStr(examName, "CT") || InStr(examName, "電腦斷層")) {
            return "CT"
        }
        if (InStr(examName, "MR") || InStr(examName, "磁振造影")) {
            return "MR"
        }
        if (InStr(examName, "US") || InStr(examName, "超音波")) {
            return "US"
        }
        return "CR"
    }

    static ReorderSelectedText(selectedText, deOrder := false, keepEmptyLine := false, itemChar := "", discardSeIm := true, forceStartFromOne := false) {
        if (selectedText == "") {
            return ""
        }

        selectedText := StrReplace(selectedText, "`r`n", "`n")
        txtAry := StrSplit(selectedText, "`n")
        finalText := ""
        isSpine := false
        startLineNo := 1
        if (!forceStartFromOne && RegExMatch(selectedText, "^(\d+)", &existLineNo)) {
            startLineNo := existLineNo[1]
        }

        for index, line in txtAry {
            if (!RegExMatch(line, "^\s*$")) {
                tmpText := line
                if (RegExMatch(line, "^\s*[-\+\*]*\s*([Vv]arying degree|[Mm]ild).+causing:")) {
                    isSpine := true
                }

                if (!deOrder) {
                    orderChar := (itemChar != "" ? itemChar : startLineNo++ . ".")
                    if (isSpine && RegExMatch(line, "^\s*([-\+\*]*|-->)\s*([CcTtLl]\d{1,2}-.+$)", &m)) {
                        finalText .= "--> "
                        tmpText := m[2]
                    } else {
                        finalText .= orderChar . " "
                    }
                }
                if (itemChar == "" && discardSeIm) {
                    tmpText := RegExReplace(tmpText, "\s*\((Srs|Ser)\/Img:.+?\)", "")
                    tmpText := RegExReplace(tmpText, "Mark L\d+:\s*", "")
                }

                tmpText := RTrim(tmpText, " `t")

                if (tmpText != "") {
                    if RegExMatch(tmpText, "[,;]$") {
                        tmpText := SubStr(tmpText, 1, -1) . "."
                    } else if !RegExMatch(tmpText, "[.:?!]$") {
                        tmpText .= "."
                    }
                }

                finalText .= RegExReplace(tmpText, "^(\s*)((\d+\.)|([-\+\*>=])|(\(?\d+\)))?(\s*)(\w?)(.*)", "$u{7}${8}")
                finalText .= "`r`n"
            } else {
                if (keepEmptyLine) {
                    finalText .= "`r`n"
                }
            }
        }

        return RTrim(finalText, "`r`n")
    }

    static DetectItemChar(text, defaultChar := "-") {
        for line in StrSplit(text, "`n", "`r") {
            if (Trim(line, " `t") != "") {
                if RegExMatch(line, "^\s*([>\-=\+\*])", &match) {
                    return match[1]
                }
                return defaultChar
            }
        }

        return defaultChar
    }

    static GetNormalizedTrailingNewlines(text) {
        if !RegExMatch(text, "(\R+)$", &match) {
            return ""
        }

        trailingNewlines := StrReplace(match[1], "`r`n", "`n")
        trailingNewlines := StrReplace(trailingNewlines, "`r", "`n")
        return StrReplace(trailingNewlines, "`n", "`r`n")
    }
}
