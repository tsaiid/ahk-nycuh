#Requires AutoHotkey v2.0

class RisDate {
    static ConvertRISDate(inputString) {
        cleanString := StrReplace(StrReplace(StrReplace(inputString, "/"), ":"), " ")
        if RegExMatch(cleanString, "^((?:19|20)\d{2})(\d{2})(\d{2})", &m) {
            return Format("{:04}-{:02}-{:02}", m[1], m[2], m[3])
        }
        if RegExMatch(cleanString, "^(\d{3})(\d{2})(\d{2})", &m) {
            return Format("{:04}-{:02}-{:02}", Integer(m[1]) + 1911, m[2], m[3])
        }
        return inputString
    }
}
