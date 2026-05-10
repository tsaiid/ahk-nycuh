#Requires AutoHotkey v2.0

class RisAIText {
    static EscapeJsonString(text) {
        escaped := StrReplace(text, "\", "\\")
        escaped := StrReplace(escaped, "`"", "\`"")
        escaped := StrReplace(escaped, "`n", "\n")
        escaped := StrReplace(escaped, "`r", "\r")
        escaped := StrReplace(escaped, "`t", "\t")
        return escaped
    }

    static DecodeJsonEscapedText(text) {
        val := text

        ; 1. 還原 JSON 內的跳脫字元
        val := StrReplace(val, "\n", "`n")
        val := StrReplace(val, "\r", "`r")
        val := StrReplace(val, "\t", "`t")
        val := StrReplace(val, '\"', '"')
        val := StrReplace(val, "\\", "\")

        ; 2. 解碼 \uXXXX (Unicode)
        while RegExMatch(val, "i)\\u([0-9a-f]{4})", &m) {
            val := StrReplace(val, m[0], Chr(Integer("0x" . m[1])))
        }

        return val
    }

    static StripMarkdownCodeFence(text) {
        text := Trim(text, " `t`r`n")

        ; 清理 Markdown 標記
        if (RegExMatch(text, "s)^``````(?:\w+)?\R?(.*?)\R?``````$", &m)) {
            text := m[1]
        } else if (SubStr(text, 1, 1) == "``" && SubStr(text, -1) == "``") {
            text := SubStr(text, 2, StrLen(text) - 2)
        }

        return Trim(text, " `t`r`n")
    }

    static ExtractGoogleResponseText(responseText) {
        combinedText := ""
        searchPos := 1

        ; 遍歷所有 "text": "..." 區段
        while (searchPos := RegExMatch(responseText, 's)"text":\s*"(.*?)(?<!\\)"', &match, searchPos)) {
            val := match[1]

            ; 如果這段文字不是 thought 區塊，才加入最終結果
            context := SubStr(responseText, searchPos + match.Len, 100)
            if !RegExMatch(context, '^\s*,\s*"thought":\s*true') {
                combinedText .= val
            }

            searchPos += match.Len
        }

        if (combinedText == "") {
            throw Error("無法從 API 回應中提取有效文字。")
        }

        return combinedText
    }

    static ExtractOpenAIResponseText(responseText) {
        combinedText := ""
        searchPos := 1

        while (searchPos := RegExMatch(responseText, 's)"type"\s*:\s*"output_text".*?"text"\s*:\s*"(.*?)(?<!\\)"', &match, searchPos)) {
            combinedText .= match[1]
            searchPos += match.Len
        }

        if (combinedText == "") {
            searchPos := 1
            while (searchPos := RegExMatch(responseText, 's)"output_text"\s*:\s*"(.*?)(?<!\\)"', &match, searchPos)) {
                combinedText .= match[1]
                searchPos += match.Len
            }
        }

        if (combinedText == "") {
            throw Error("無法從 OpenAI 回應中提取有效文字。")
        }

        return combinedText
    }

    static ParseGoogleResponse(responseText) {
        text := this.ExtractGoogleResponseText(responseText)
        text := this.DecodeJsonEscapedText(text)
        return this.StripMarkdownCodeFence(text)
    }

    static ParseOpenAIResponse(responseText) {
        text := this.ExtractOpenAIResponseText(responseText)
        text := this.DecodeJsonEscapedText(text)
        return this.StripMarkdownCodeFence(text)
    }
}
