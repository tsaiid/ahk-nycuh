#Requires AutoHotkey v2.0

#Include TestLib.v2.ahk

RegisterTest("ExtractMinSeriesFromDesc extracts index from parentheses", Test_ExtractMinSeriesFromDesc)
RegisterTest("ParseSrs parses leading series number", Test_ParseSrs)
RegisterTest("minSeries filtering logic accepts only srs >= minSeries", Test_MinSeriesValidation)

ExtractMinSeriesFromDesc(descText) {
    if (descText != "" && RegExMatch(Trim(descText), "^\(\s*(\d+)\s*\)", &match)) {
        return Integer(match[1])
    }
    return 0
}

ParseSrs(text) {
    splitText := StrSplit(text, ",")
    if (splitText.Length > 0) {
        if (RegExMatch(Trim(splitText[1]), "^(\d+)", &match)) {
            return match[1]
        }
    }
    return ""
}

Test_ExtractMinSeriesFromDesc() {
    AssertEqual(7, ExtractMinSeriesFromDesc("(7) V phase  5.0  MPR  cor"))
    AssertEqual(1, ExtractMinSeriesFromDesc("(1) Plain CT"))
    AssertEqual(12, ExtractMinSeriesFromDesc("(12) CTA Chest 1.0"))
    AssertEqual(3, ExtractMinSeriesFromDesc("  ( 3 ) Axial "))
    AssertEqual(0, ExtractMinSeriesFromDesc("(0) Scout"))
    AssertEqual(0, ExtractMinSeriesFromDesc(""))
    AssertEqual(0, ExtractMinSeriesFromDesc("Plain CT without parens"))
    AssertEqual(0, ExtractMinSeriesFromDesc("(abc) Invalid format"))
}

Test_ParseSrs() {
    AssertEqual("11", ParseSrs("11 , V phase"))
    AssertEqual("11", ParseSrs("11, V phase"))
    AssertEqual("1", ParseSrs("1 V phase"))
    AssertEqual("105", ParseSrs("105 Coronal"))
    AssertEqual("", ParseSrs("VMTool"))
    AssertEqual("", ParseSrs(""))
}

Test_MinSeriesValidation() {
    minSeries := ExtractMinSeriesFromDesc("(7) V phase  5.0  MPR  cor")
    AssertEqual(7, minSeries)

    ; scale=3 gives "1" -> should be rejected because 1 < 7
    srsVal1 := ParseSrs("1 V phase")
    isValid1 := (srsVal1 != "" && (minSeries <= 0 || Integer(srsVal1) >= minSeries))
    AssertFalse(isValid1, "Series 1 should be rejected when minSeries is 7")

    ; scale=2 gives "11" -> should be accepted because 11 >= 7
    srsVal2 := ParseSrs("11 , V phase")
    isValid2 := (srsVal2 != "" && (minSeries <= 0 || Integer(srsVal2) >= minSeries))
    AssertTrue(isValid2, "Series 11 should be accepted when minSeries is 7")

    ; If minSeries is 0 (no desc bracket), "1" should be accepted
    minSeries0 := ExtractMinSeriesFromDesc("Plain")
    isValid3 := (srsVal1 != "" && (minSeries0 <= 0 || Integer(srsVal1) >= minSeries0))
    AssertTrue(isValid3, "Series 1 should be accepted when minSeries is 0")
}

RunRegisteredTests()
