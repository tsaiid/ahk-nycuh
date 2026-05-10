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
}
