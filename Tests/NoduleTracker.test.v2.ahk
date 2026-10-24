#Requires AutoHotkey v2.0

#Include TestLib.v2.ahk
#Include ..\Lib\G3PacsProbe.v2.ahk
#Include ..\NoduleTracker.ahk

RegisterTest("G3PacsProbe.GetPatternList initializes all patterns with hits property", Test_GetPatternList_HitsProperty)
RegisterTest("NoduleTracker.Sort sorts patterns in descending order of hits", Test_SortPatternsByHits)
RegisterTest("NoduleTracker.Sort correctly prioritizes real PatternList items by hits", Test_SortRealPatternListByHits)
RegisterTest("NoduleTracker.ParseSrs parses leading series number", Test_ParseSrs)
RegisterTest("NoduleTracker.SelectBestSrsFromVotes selects consensus series", Test_MajorityVotingAlgorithm)
RegisterTest("NoduleTracker.IsPositionInScreen detects in-bound and out-of-bound positions", Test_IsPositionInScreen)
RegisterTest("NoduleTracker.GetPrimaryTopRightPos calculates valid position on primary monitor", Test_GetPrimaryTopRightPos)
RegisterTest("NoduleTracker.IsMprSeries identifies MPR series and ignores MRI sequences", Test_IsMprSeries)
RegisterTest("NoduleTracker.CalculateQuickSetTarget applies reverse compensation for MPR", Test_CalculateQuickSetTarget)
RegisterTest("NoduleTracker.GetQuickSetRange calculates shifted range for MPR compensation", Test_GetQuickSetRange)
RegisterTest("NoduleTracker.FormatSentenceReport formats single nodule as full sentence", Test_FormatSentenceReport_Single)
RegisterTest("NoduleTracker.FormatSentenceReport formats multiple nodules across lobes with plural form", Test_FormatSentenceReport_MultipleLobes)
RegisterTest("NoduleTracker.FormatSentenceReport formats multiple nodules in single lobe with plural form", Test_FormatSentenceReport_SingleLobeMultiple)
RegisterTest("NoduleTracker.FormatSentenceReport formats 3 lobes with Oxford comma", Test_FormatSentenceReport_ThreeLobes)
RegisterTest("NoduleTracker.FormatSentenceReport returns empty string when no nodules", Test_FormatSentenceReport_Empty)
RegisterTest("NoduleTracker.FormatDirectCopy formats lobe or full sentence", Test_FormatDirectCopy)
RegisterTest("NoduleTracker.FormatHepaticDirectCopy formats liver segment or full sentence", Test_FormatHepaticDirectCopy)
RegisterTest("NoduleTracker initialization supports decoupled options without side effects", Test_InstanceInitialization_DecoupledOptions)
RegisterTest("NoduleTracker.BuildSuggestedProbeReport distinguishes existing vs new patterns", Test_BuildSuggestedProbeReport)

Test_GetPatternList_HitsProperty() {
    patterns := G3PacsProbe.GetPatternList()
    AssertTrue(patterns.Length > 0, "PatternList should not be empty")
    for pattern in patterns {
        AssertTrue(pattern.HasOwnProp("hits"), "Pattern " . pattern.name . " should have hits property")
        AssertEqual(0, pattern.hits, "Pattern " . pattern.name . " hits should default to 0")
    }
}

Test_SortPatternsByHits() {
    testPatterns := [
        {name: "p1", hits: 5},
        {name: "p2", hits: 20},
        {name: "p3", hits: 0},
        {name: "p4", hits: 15}
    ]
    NoduleTracker.Sort(testPatterns, (a, b) => (b.HasOwnProp("hits") ? b.hits : 0) - (a.HasOwnProp("hits") ? a.hits : 0))
    AssertEqual("p2", testPatterns[1].name, "First pattern should be highest hit (p2)")
    AssertEqual("p4", testPatterns[2].name, "Second pattern should be p4")
    AssertEqual("p1", testPatterns[3].name, "Third pattern should be p1")
    AssertEqual("p3", testPatterns[4].name, "Fourth pattern should be lowest hit (p3)")
}

Test_SortRealPatternListByHits() {
    patterns := G3PacsProbe.GetPatternList()
    AssertTrue(patterns.Length >= 3, "PatternList should have at least 3 patterns")
    patterns[1].hits := 10
    patterns[2].hits := 50
    patterns[3].hits := 5
    NoduleTracker.Sort(patterns, (a, b) => (b.HasOwnProp("hits") ? b.hits : 0) - (a.HasOwnProp("hits") ? a.hits : 0))
    AssertEqual(50, patterns[1].hits, "Top pattern should have 50 hits")
    AssertEqual(10, patterns[2].hits, "Second pattern should have 10 hits")
    AssertTrue(patterns[1].hits >= patterns[2].hits, "First hit should be >= second hit")
}

Test_ParseSrs() {
    AssertEqual("11", NoduleTracker.ParseSrs("11 , V phase"))
    AssertEqual("11", NoduleTracker.ParseSrs("11, V phase"))
    AssertEqual("1", NoduleTracker.ParseSrs("1 V phase"))
    AssertEqual("105", NoduleTracker.ParseSrs("105 Coronal"))
    AssertEqual("", NoduleTracker.ParseSrs("VMTool"))
    AssertEqual("", NoduleTracker.ParseSrs(""))
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
    m1 := ""
    res1 := NoduleTracker.SelectBestSrsFromVotes(c1, &m1)
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
    m2 := ""
    res2 := NoduleTracker.SelectBestSrsFromVotes(c2, &m2)
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
    m3 := ""
    res3 := NoduleTracker.SelectBestSrsFromVotes(c3, &m3)
    AssertEqual("11", res3, "Majority voting should pick 11 over 1")
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
    AssertTrue(NoduleTracker.IsPositionInScreen(validX, validY, w, h), "Valid position within work area should return true")

    ; 負數座標（超出最左側/最上方）
    AssertFalse(NoduleTracker.IsPositionInScreen(minLeft - 500, validY, w, h), "Negative/out-of-bounds X on left should return false")
    AssertFalse(NoduleTracker.IsPositionInScreen(validX, minTop - 500, w, h), "Out-of-bounds Y on top should return false")

    ; 超出右側/下方的極端座標
    AssertFalse(NoduleTracker.IsPositionInScreen(maxRight + 500, validY, w, h), "X far beyond right bound should return false")
    AssertFalse(NoduleTracker.IsPositionInScreen(validX, maxBottom + 500, w, h), "Y far beyond bottom bound should return false")

    ; 視窗超出右側邊界 (針對最右邊螢幕的邊界)
    AssertFalse(NoduleTracker.IsPositionInScreen(maxRight - 50, validY, w, h), "Window overflowing right edge should return false")

    ; 空座標
    AssertFalse(NoduleTracker.IsPositionInScreen("", 50, w, h), "Empty X should return false")
    AssertFalse(NoduleTracker.IsPositionInScreen(50, "", w, h), "Empty Y should return false")
}

Test_GetPrimaryTopRightPos() {
    primaryIndex := MonitorGetPrimary()
    MonitorGetWorkArea(primaryIndex, &workLeft, &workTop, &workRight, &workBottom)

    w := 320, h := 300
    pos := NoduleTracker.GetPrimaryTopRightPos(w, h, &outX, &outY)

    AssertEqual(pos.x, outX, "Returned object x should match outX")
    AssertEqual(pos.y, outY, "Returned object y should match outY")

    ; 檢查座標是否位於主螢幕 WorkArea 內部
    AssertTrue(outX >= workLeft, "Top-right X should be >= workLeft")
    AssertTrue(outX + w <= workRight, "Top-right X + w should be <= workRight")
    AssertTrue(outY >= workTop, "Top-right Y should be >= workTop")
    AssertTrue(outY + h <= workBottom, "Top-right Y + h should be <= workBottom")
    AssertTrue(NoduleTracker.IsPositionInScreen(outX, outY, w, h), "Calculated top-right position must pass IsPositionInScreen")
}

Test_IsMprSeries() {
    AssertTrue(NoduleTracker.IsMprSeries("(3) Axial MPR"), "Axial MPR should be MPR series")
    AssertTrue(NoduleTracker.IsMprSeries("COR MIP"), "COR MIP should be MPR series")
    AssertTrue(NoduleTracker.IsMprSeries("SAG MPR 2mm"), "SAG MPR should be MPR series")
    AssertFalse(NoduleTracker.IsMprSeries("(2) 1.25mm CT Axial"), "Standard Axial should not be MPR series")
    AssertFalse(NoduleTracker.IsMprSeries("T2 SAG"), "T2 SAG is MRI, should be ignored")
    AssertFalse(NoduleTracker.IsMprSeries("DWI COR"), "DWI COR is MRI, should be ignored")
    AssertFalse(NoduleTracker.IsMprSeries(""), "Empty string should not be MPR")
}

Test_CalculateQuickSetTarget() {
    ; 當選項啟用且為 MPR 序列：輸入 30 應跳到 29
    AssertEqual(29, NoduleTracker.CalculateQuickSetTarget(30, true, true), "MPR enabled: 30 should jump to 29")
    AssertEqual(1, NoduleTracker.CalculateQuickSetTarget(2, true, true), "MPR enabled: 2 should jump to 1")
    AssertEqual(1, NoduleTracker.CalculateQuickSetTarget(1, true, true), "MPR enabled: 1 should clamp to 1")

    ; 當選項未啟用：輸入 30 維持 30
    AssertEqual(30, NoduleTracker.CalculateQuickSetTarget(30, true, false), "MPR disabled: 30 should remain 30")

    ; 當非 MPR 序列：輸入 30 維持 30
    AssertEqual(30, NoduleTracker.CalculateQuickSetTarget(30, false, true), "Non-MPR: 30 should remain 30")

    ; 無效輸入
    AssertEqual(0, NoduleTracker.CalculateQuickSetTarget(0, true, true), "0 should be invalid (0)")
    AssertEqual(0, NoduleTracker.CalculateQuickSetTarget("abc", true, true), "Non-number should be invalid (0)")
}

Test_GetQuickSetRange() {
    ; 當選項啟用且為 MPR 序列：控制項 97 應顯示 2 ~ 98，最大合法輸入為 98
    r1 := NoduleTracker.GetQuickSetRange(97, true, true)
    AssertEqual(2, r1.min, "MPR enabled min should be 2")
    AssertEqual(98, r1.max, "MPR enabled max should be 98")
    AssertEqual("2 ~ 98", r1.text, "MPR enabled text should be 2 ~ 98")

    ; 當選項未啟用但為 MPR 序列：維持 1 ~ 97
    r2 := NoduleTracker.GetQuickSetRange(97, true, false)
    AssertEqual(1, r2.min, "MPR disabled min should be 1")
    AssertEqual(97, r2.max, "MPR disabled max should be 97")
    AssertEqual("1 ~ 97", r2.text, "MPR disabled text should be 1 ~ 97")

    ; 非 MPR 序列：維持 1 ~ 97
    r3 := NoduleTracker.GetQuickSetRange(97, false, true)
    AssertEqual(1, r3.min, "Non-MPR min should be 1")
    AssertEqual(97, r3.max, "Non-MPR max should be 97")
    AssertEqual("1 ~ 97", r3.text, "Non-MPR text should be 1 ~ 97")

    ; 無效或 0
    r4 := NoduleTracker.GetQuickSetRange(0, true, true)
    AssertEqual("", r4.text, "0 itemCount should have empty text")
    r5 := NoduleTracker.GetQuickSetRange(-1, true, true)
    AssertEqual("", r5.text, "Negative itemCount should have empty text")
}

Test_FormatSentenceReport_Single() {
    entries := [
        {label: "RUL", srsImg: "6/15", count: 1}
    ]
    expected := "A tiny nodule in the RUL of lung (Srs/Img: 6/15)."
    AssertEqual(expected, NoduleTracker.FormatSentenceReport(entries), "Single nodule format mismatch")
}

Test_FormatSentenceReport_MultipleLobes() {
    entries := [
        {label: "RUL", srsImg: "6/15", count: 1},
        {label: "RLL", srsImg: "6/30", count: 1}
    ]
    expected := "Tiny nodules in the RUL (Srs/Img: 6/15) and RLL (Srs/Img: 6/30) of lung."
    AssertEqual(expected, NoduleTracker.FormatSentenceReport(entries), "Multiple lobes format mismatch")
}

Test_FormatSentenceReport_SingleLobeMultiple() {
    entries := [
        {label: "RUL", srsImg: "6/15,20", count: 2}
    ]
    expected := "Tiny nodules in the RUL (Srs/Img: 6/15,20) of lung."
    AssertEqual(expected, NoduleTracker.FormatSentenceReport(entries), "Single lobe multiple nodules format mismatch")
}

Test_FormatSentenceReport_ThreeLobes() {
    entries := [
        {label: "RUL", srsImg: "6/15", count: 1},
        {label: "RML", srsImg: "6/20", count: 1},
        {label: "RLL", srsImg: "6/30", count: 1}
    ]
    expected := "Tiny nodules in the RUL (Srs/Img: 6/15), RML (Srs/Img: 6/20), and RLL (Srs/Img: 6/30) of lung."
    AssertEqual(expected, NoduleTracker.FormatSentenceReport(entries), "Three lobes format mismatch")
}

Test_FormatSentenceReport_Empty() {
    entries := []
    AssertEqual("", NoduleTracker.FormatSentenceReport(entries), "Empty entries should return empty string")
}

Test_FormatDirectCopy() {
    AssertEqual("RML of lung (Srs/Img: 5/27)", NoduleTracker.FormatDirectCopy("RML", "5", "27", false), "Direct copy format mismatch")
    AssertEqual("A tiny nodule in the RML of lung (Srs/Img: 5/27).", NoduleTracker.FormatDirectCopy("RML", "5", "27", true), "Direct sentence copy format mismatch")
}

Test_FormatHepaticDirectCopy() {
    AssertEqual("S1 of the liver (Srs/Img: 2/1)", NoduleTracker.FormatHepaticDirectCopy(1, "2", "1", false), "Direct copy format mismatch for S1")
    AssertEqual("A nodule is noted in S1 of the liver (Srs/Img: 2/1).", NoduleTracker.FormatHepaticDirectCopy(1, "2", "1", true), "Direct sentence copy format mismatch for S1")
    AssertEqual("S8 of the liver (Srs/Img: 3/45)", NoduleTracker.FormatHepaticDirectCopy(8, "3", "45", false), "Direct copy format mismatch for S8")
    AssertEqual("A nodule is noted in S8 of the liver (Srs/Img: 3/45).", NoduleTracker.FormatHepaticDirectCopy(8, "3", "45", true), "Direct sentence copy format mismatch for S8")
    AssertEqual("S2 of the liver (Srs/Img: 4/12)", NoduleTracker.FormatHepaticDirectCopy("S2", "4", "12", false), "Direct copy format mismatch for S2 string input")
}

Test_InstanceInitialization_DecoupledOptions() {
    inst := NoduleTracker({initTray: false, showGui: false, setDpi: 0})
    AssertFalse(inst.hasTrayMenu, "hasTrayMenu should be false when initTray is false")
    AssertEqual("", inst.MyGui, "MyGui should remain empty when showGui is false")
    AssertTrue(inst.PatternList.Length > 0, "PatternList should be initialized")

    inst.UpdateTrayMenu()

    customMenu := Menu()
    inst.InitTrayMenu(customMenu)
    AssertTrue(inst.hasTrayMenu, "hasTrayMenu should be true after InitTrayMenu")
    AssertTrue(IsObject(inst.trayMenu), "trayMenu should be an object")

    ; Clean up menu callbacks to prevent circular reference on exit
    inst.trayMenu := ""
    customMenu.Delete()
}

Test_BuildSuggestedProbeReport() {
    inst := NoduleTracker({initTray: false, showGui: false, setDpi: 0})

    ; 1. 已存在的規則 [3, 10, 39, 101] 應標記為已存在
    existingNums := {focus: 3, img: 10, srs: 39, desc: 101}
    reportExist := inst.BuildSuggestedProbeReport(existingNums)
    AssertTrue(InStr(reportExist, "已存在於 G3PacsProbe.Configs 中") > 0, "Existing probe should be recognized as already in Configs")
    AssertFalse(InStr(reportExist, "💡 建議新增 probe") > 0, "Existing probe should not suggest adding")

    ; 2. 真正未配置的新規則應建議新增
    newNums := {focus: 3, img: 10, srs: 999, desc: 999}
    reportNew := inst.BuildSuggestedProbeReport(newNums)
    AssertTrue(InStr(reportNew, "💡 建議新增 probe: [3, 10, 999, 999]") > 0, "New probe should be suggested")

    ; 3. 欄位不完整時應回傳空字串
    emptyNums := {focus: 3, img: 10, srs: "", desc: 101}
    AssertEqual("", inst.BuildSuggestedProbeReport(emptyNums), "Incomplete nums should return empty report")
}

RunRegisteredTests()
