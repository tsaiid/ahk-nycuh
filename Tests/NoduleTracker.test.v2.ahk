#Requires AutoHotkey v2.0

#Include TestLib.v2.ahk

RegisterTest("ExtractMinSeriesFromDesc extracts index from parentheses", Test_ExtractMinSeriesFromDesc)
RegisterTest("ParseSrs parses leading series number", Test_ParseSrs)
RegisterTest("minSeries filtering logic accepts only srs >= minSeries", Test_MinSeriesValidation)
RegisterTest("IsPositionInScreen detects in-bound and out-of-bound positions", Test_IsPositionInScreen)
RegisterTest("GetPrimaryTopRightPos calculates valid position on primary monitor", Test_GetPrimaryTopRightPos)

IsPositionInScreen(x, y, w := 320, h := 300) {
    if (x == "" || y == "") {
        return false
    }
    monitorCount := MonitorGetCount()
    loop monitorCount {
        if (!MonitorGetWorkArea(A_Index, &workLeft, &workTop, &workRight, &workBottom)) {
            continue
        }
        workW := workRight - workLeft
        workH := workBottom - workTop

        xValid := (w <= workW) ? (x >= workLeft && (x + w) <= workRight) : (x >= workLeft)
        yValid := (h <= workH) ? (y >= workTop && (y + h) <= workBottom) : (y >= workTop)

        if (xValid && yValid) {
            return true
        }
    }
    return false
}

GetPrimaryTopRightPos(w := 320, h := 300, &outX := 0, &outY := 0) {
    primaryIndex := 1
    try {
        primaryIndex := MonitorGetPrimary()
    }
    if (!MonitorGetWorkArea(primaryIndex, &workLeft, &workTop, &workRight, &workBottom)) {
        workLeft := 0, workTop := 0, workRight := A_ScreenWidth, workBottom := A_ScreenHeight
    }

    marginRight := 20
    marginTop := 50

    outX := workRight - w - marginRight
    if (outX < workLeft) {
        outX := workLeft
    }

    outY := workTop + marginTop
    if (outY + h > workBottom) {
        outY := Max(workTop, workBottom - h)
    }
    return {x: outX, y: outY}
}

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

Test_IsPositionInScreen() {
    primaryIndex := MonitorGetPrimary()
    MonitorGetWorkArea(primaryIndex, &workLeft, &workTop, &workRight, &workBottom)

    ; 正常在螢幕內的座標
    w := 320, h := 300
    validX := workLeft + 50
    validY := workTop + 50
    AssertTrue(IsPositionInScreen(validX, validY, w, h), "Valid position within work area should return true")

    ; 負數座標（超出左側/上方）
    AssertFalse(IsPositionInScreen(workLeft - 100, validY, w, h), "Negative/out-of-bounds X on left should return false")
    AssertFalse(IsPositionInScreen(validX, workTop - 100, w, h), "Out-of-bounds Y on top should return false")

    ; 超出右側/下方的極端座標
    AssertFalse(IsPositionInScreen(workRight + 500, validY, w, h), "X far beyond right bound should return false")
    AssertFalse(IsPositionInScreen(validX, workBottom + 500, w, h), "Y far beyond bottom bound should return false")

    ; 視窗超出右側邊界
    AssertFalse(IsPositionInScreen(workRight - 50, validY, w, h), "Window overflowing right edge should return false")

    ; 空座標
    AssertFalse(IsPositionInScreen("", 50, w, h), "Empty X should return false")
    AssertFalse(IsPositionInScreen(50, "", w, h), "Empty Y should return false")
}

Test_GetPrimaryTopRightPos() {
    primaryIndex := MonitorGetPrimary()
    MonitorGetWorkArea(primaryIndex, &workLeft, &workTop, &workRight, &workBottom)

    w := 320, h := 300
    pos := GetPrimaryTopRightPos(w, h, &outX, &outY)

    AssertEqual(pos.x, outX, "Returned object x should match outX")
    AssertEqual(pos.y, outY, "Returned object y should match outY")

    ; 檢查座標是否位於主螢幕 WorkArea 內部
    AssertTrue(outX >= workLeft, "Top-right X should be >= workLeft")
    AssertTrue(outX + w <= workRight, "Top-right X + w should be <= workRight")
    AssertTrue(outY >= workTop, "Top-right Y should be >= workTop")
    AssertTrue(outY + h <= workBottom, "Top-right Y + h should be <= workBottom")
    AssertTrue(IsPositionInScreen(outX, outY, w, h), "Calculated top-right position must pass IsPositionInScreen")
}

RunRegisteredTests()
