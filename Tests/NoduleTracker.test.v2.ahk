#Requires AutoHotkey v2.0

#Include TestLib.v2.ahk

RegisterTest("ParseSrs parses leading series number", Test_ParseSrs)
RegisterTest("Majority voting algorithm selects consensus series", Test_MajorityVotingAlgorithm)
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

ParseSrs(text) {
    splitText := StrSplit(text, ",")
    if (splitText.Length > 0) {
        if (RegExMatch(Trim(splitText[1]), "^(\d+)", &match)) {
            return match[1]
        }
    }
    return ""
}

Test_ParseSrs() {
    AssertEqual("11", ParseSrs("11 , V phase"))
    AssertEqual("11", ParseSrs("11, V phase"))
    AssertEqual("1", ParseSrs("1 V phase"))
    AssertEqual("105", ParseSrs("105 Coronal"))
    AssertEqual("", ParseSrs("VMTool"))
    AssertEqual("", ParseSrs(""))
}

Test_IsPositionInScreen() {
    primaryIndex := MonitorGetPrimary()
    MonitorGetWorkArea(primaryIndex, &workLeft, &workTop, &workRight, &workBottom)

    ; 計算跨所有螢幕的最大邊界，確保測試極端座標時真正超出所有螢幕
    maxRight := workRight
    minLeft := workLeft
    minTop := workTop
    maxBottom := workBottom
    loop MonitorGetCount() {
        if (MonitorGetWorkArea(A_Index, &mLeft, &mTop, &mRight, &mBottom)) {
            minLeft := Min(minLeft, mLeft)
            maxRight := Max(maxRight, mRight)
            minTop := Min(minTop, mTop)
            maxBottom := Max(maxBottom, mBottom)
        }
    }

    ; 正常在螢幕內的座標
    w := 320, h := 300
    validX := workLeft + 50
    validY := workTop + 50
    AssertTrue(IsPositionInScreen(validX, validY, w, h), "Valid position within work area should return true")

    ; 負數座標（超出最左側/最上方）
    AssertFalse(IsPositionInScreen(minLeft - 500, validY, w, h), "Negative/out-of-bounds X on left should return false")
    AssertFalse(IsPositionInScreen(validX, minTop - 500, w, h), "Out-of-bounds Y on top should return false")

    ; 超出右側/下方的極端座標
    AssertFalse(IsPositionInScreen(maxRight + 500, validY, w, h), "X far beyond right bound should return false")
    AssertFalse(IsPositionInScreen(validX, maxBottom + 500, w, h), "Y far beyond bottom bound should return false")

    ; 視窗超出右側邊界 (針對最右邊螢幕的邊界)
    AssertFalse(IsPositionInScreen(maxRight - 50, validY, w, h), "Window overflowing right edge should return false")

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

SelectBestSrsFromVotes(candidates, &outMethod := "") {
    votes := Map()
    for item in candidates {
        if (item.srs != "") {
            votes[item.srs] := votes.Has(item.srs) ? votes[item.srs] + 1 : 1
        }
    }
    if (votes.Count == 0) {
        return ""
    }
    bestSrs := ""
    maxVotes := 0
    for srsVal, count in votes {
        if (count > maxVotes) {
            maxVotes := count
            bestSrs := srsVal
        }
    }
    for item in candidates {
        if (item.srs == bestSrs) {
            outMethod := item.method " (投票: " maxVotes "/" candidates.Length ")"
            return item.srs
        }
    }
    return ""
}

Test_MajorityVotingAlgorithm() {
    ; 情境 1：全數一致（8 票全為 8）
    c1 := [
        {srs: "8", method: "scale=3"},
        {srs: "8", method: "scale=3 invert"},
        {srs: "8", method: "scale=2"},
        {srs: "8", method: "scale=4"},
        {srs: "8", method: "scale=3 grayscale"},
        {srs: "8", method: "scale=4 grayscale"},
        {srs: "8", method: "scale=4 invert"},
        {srs: "8", method: "scale=4 gray+invert"}
    ]
    res1 := SelectBestSrsFromVotes(c1, &m1)
    AssertEqual("8", res1, "Unanimous voting should pick 8")
    AssertEqual("scale=3 (投票: 8/8)", m1, "Method should report 8/8 votes")

    ; 情境 2：單一 scale 誤識（scale=2 為 18，其餘 7 個為 104）
    c2 := [
        {srs: "104", method: "scale=3"},
        {srs: "104", method: "scale=3 invert"},
        {srs: "18", method: "scale=2"},
        {srs: "104", method: "scale=4"},
        {srs: "104", method: "scale=3 grayscale"},
        {srs: "104", method: "scale=4 grayscale"},
        {srs: "104", method: "scale=4 invert"},
        {srs: "104", method: "scale=4 gray+invert"}
    ]
    res2 := SelectBestSrsFromVotes(c2, &m2)
    AssertEqual("104", res2, "Majority voting should pick 104 over 18")
    AssertEqual("scale=3 (投票: 7/8)", m2, "Method should report 7/8 votes")

    ; 情境 3：scale=3 漏字為 1，其他多數為 11
    c3 := [
        {srs: "1", method: "scale=3"},
        {srs: "11", method: "scale=2"},
        {srs: "11", method: "scale=4"},
        {srs: "11", method: "scale=3 grayscale"},
        {srs: "", method: "scale=4 grayscale"}
    ]
    res3 := SelectBestSrsFromVotes(c3, &m3)
    AssertEqual("11", res3, "Majority voting should pick 11 over 1")
}

RunRegisteredTests()
