#Requires AutoHotkey v2.0

#Include TestLib.v2.ahk
#Include ..\Lib\RisEditControl.v2.ahk

RegisterTest("RisEditControl._GetSmartListNextPrefix continues non-contiguous spine levels", Test_RisEditControl_SpineNonContiguous)
RegisterTest("RisEditControl._GetSmartListNextPrefix continues contiguous spine levels", Test_RisEditControl_SpineContiguous)
RegisterTest("RisEditControl._GetSmartListNextPrefix preserves list markers with spine levels", Test_RisEditControl_SpineMarkers)
RegisterTest("RisEditControl._GetSmartListNextPrefix rejects invalid, reversed, or terminal spine levels", Test_RisEditControl_SpineInvalid)
RegisterTest("RisEditControl._GetSmartListPreviousPrefix derives previous spine level from first start", Test_RisEditControl_SpinePrevious)

Test_RisEditControl_SpineNonContiguous() {
    ; 使用者提報案例：胸腰交界跳過 T12-L1
    lineInfo1 := {Text: "T10-T11, T11-T12, L1-L2: mild indentation on anterior dural sac."}
    AssertEqual("L2-L3: ", RisEditControl._GetSmartListNextPrefix(lineInfo1), "should continue to L2-L3 after L1-L2")

    ; 頸椎跳節案例：跳過 C4-C5
    lineInfo2 := {Text: "C3-C4, C5-C6: mild disc herniation."}
    AssertEqual("C6-C7: ", RisEditControl._GetSmartListNextPrefix(lineInfo2), "should continue to C6-C7 after C5-C6")

    ; 腰椎跳節案例：跳過 L2-L3
    lineInfo3 := {Text: "L1-L2, L3-L4: degenerative changes."}
    AssertEqual("L4-L5: ", RisEditControl._GetSmartListNextPrefix(lineInfo3), "should continue to L4-L5 after L3-L4")
}

Test_RisEditControl_SpineContiguous() {
    ; 單節段
    lineInfo1 := {Text: "L1-L2: mild bulging"}
    AssertEqual("L2-L3: ", RisEditControl._GetSmartListNextPrefix(lineInfo1), "single level should continue")

    ; 連續多節段
    lineInfo2 := {Text: "C2-C3, C3-C4, C4-C5: disc bulge"}
    AssertEqual("C5-C6: ", RisEditControl._GetSmartListNextPrefix(lineInfo2), "contiguous multiple levels should continue")

    ; 跨交界連續節段 T12-L1 至 L1-L2
    lineInfo3 := {Text: "T12-L1, L1-L2: disc bulge"}
    AssertEqual("L2-L3: ", RisEditControl._GetSmartListNextPrefix(lineInfo3), "junction contiguous levels should continue")
}

Test_RisEditControl_SpineMarkers() {
    ; 符號清單
    lineInfo1 := {Text: "- T10-T11, T11-T12, L1-L2: mild indentation"}
    AssertEqual("- L2-L3: ", RisEditControl._GetSmartListNextPrefix(lineInfo1), "dash bullet should be preserved")

    ; 數字清單
    lineInfo2 := {Text: "1. T10-T11, T11-T12, L1-L2: mild indentation"}
    AssertEqual("2. L2-L3: ", RisEditControl._GetSmartListNextPrefix(lineInfo2), "numbered list should increment")
}

Test_RisEditControl_SpineInvalid() {
    ; 終點 L5-S1 不延續
    lineInfo1 := {Text: "L5-S1: disc herniation"}
    AssertEqual("", RisEditControl._GetSmartListNextPrefix(lineInfo1), "L5-S1 should terminate spine continuation")

    ; 多節段終點 L5-S1 不延續
    lineInfo2 := {Text: "L3-L4, L5-S1: disc herniation"}
    AssertEqual("", RisEditControl._GetSmartListNextPrefix(lineInfo2), "ending at L5-S1 should terminate")

    ; 倒序不延續
    lineInfo3 := {Text: "L1-L2, T11-T12: test"}
    AssertEqual("", RisEditControl._GetSmartListNextPrefix(lineInfo3), "reversed order should be rejected")

    ; 重複不延續
    lineInfo4 := {Text: "L1-L2, L1-L2: test"}
    AssertEqual("", RisEditControl._GetSmartListNextPrefix(lineInfo4), "duplicate segment should be rejected")

    ; 非相鄰單一節段（如跨節 T10-L2）
    lineInfo5 := {Text: "T10-L2: test"}
    AssertEqual("", RisEditControl._GetSmartListNextPrefix(lineInfo5), "non-adjacent single segment should be rejected")
}

Test_RisEditControl_SpinePrevious() {
    ; 向上插入取第一段的前一節
    lineInfo1 := {Text: "T10-T11, T11-T12, L1-L2: mild indentation"}
    AssertEqual("T9-T10: ", RisEditControl._GetSmartListPreviousPrefix(lineInfo1), "should derive T9-T10 from first start T10")

    ; C1-C2 沒有上一節
    lineInfo2 := {Text: "C1-C2: test"}
    AssertEqual("", RisEditControl._GetSmartListPreviousPrefix(lineInfo2), "C1 has no previous vertebra")
}

RunRegisteredTests()

