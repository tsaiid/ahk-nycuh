#Requires AutoHotkey v2.0

#Include TestLib.v2.ahk
#Include ..\Lib\RisEditControl.v2.ahk

RegisterTest("RisEditControl._GetSmartListNextPrefix continues non-contiguous spine levels", Test_RisEditControl_SpineNonContiguous)
RegisterTest("RisEditControl._GetSmartListNextPrefix continues contiguous spine levels", Test_RisEditControl_SpineContiguous)
RegisterTest("RisEditControl._GetSmartListNextPrefix preserves list markers with spine levels", Test_RisEditControl_SpineMarkers)
RegisterTest("RisEditControl._GetSmartListNextPrefix rejects invalid, reversed, or terminal spine levels", Test_RisEditControl_SpineInvalid)
RegisterTest("RisEditControl._GetSmartListPreviousPrefix derives previous spine level from first start", Test_RisEditControl_SpinePrevious)
RegisterTest("RisEditControl.KillLine bounds to ContentEnd without deleting subsequent lines", Test_RisEditControl_KillLine)
RegisterTest("RisEditControl.SelectLineForRemoval preserves previous line character on last line", Test_RisEditControl_SelectLineForRemoval)

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

Test_RisEditControl_KillLine() {
    testGui := Gui()
    editCtrl := testGui.Add("Edit", "w300 h200 Multi")
    hEdit := editCtrl.Hwnd

    ; 情境 1: 測試 LF (\n) 換行
    ; 原有 Bug: InStr 只找 `r，找不到時 targetPos 變成全文長度，整段被刪光
    ; 修復後: 只刪除到行尾 ContentEnd，保留後續行
    ControlSetText("First line`nSecond line", hEdit)
    RisEditControl.SetSel(hEdit, 5, 5) ; 游標在 "First" 之後
    RisEditControl.KillLine(hEdit)
    actualText := ControlGetText(hEdit)
    AssertTrue(InStr(actualText, "Second line") > 0, "KillLine should not delete subsequent lines on LF")

    ; 情境 2: 測試 CRLF (`r`n) 換行
    ControlSetText("First line`r`nSecond line", hEdit)
    RisEditControl.SetSel(hEdit, 5, 5)
    RisEditControl.KillLine(hEdit)
    AssertEqual("First`r`nSecond line", ControlGetText(hEdit), "KillLine should delete to CRLF without deleting subsequent lines")

    ; 情境 3: 單行無換行
    ControlSetText("Only line", hEdit)
    RisEditControl.SetSel(hEdit, 4, 4)
    RisEditControl.KillLine(hEdit)
    AssertEqual("Only", ControlGetText(hEdit), "KillLine on single line should delete to end of line")

    testGui.Destroy()
}

Test_RisEditControl_SelectLineForRemoval() {
    testGui := Gui()
    editCtrl := testGui.Add("Edit", "w300 h200 Multi")
    hEdit := editCtrl.Hwnd

    ; 情境 1: 針對最後一行，前行為 CRLF
    ControlSetText("Line 1`r`nLine 2", hEdit)
    boundsCRLF := RisEditControl.GetLogicalLineBoundaries(hEdit, 9) ; 游標在 Line 2
    RisEditControl.SelectLineForRemoval(hEdit, boundsCRLF)
    selCRLF := RisEditControl.GetSel(hEdit)
    AssertEqual(6, selCRLF.Start, "CRLF: removal start should be bounds.Start - 2")
    AssertEqual(14, selCRLF.End, "CRLF: removal end should be bounds.FullEnd")

    ; 情境 2: 針對最後一行，前行為純 LF (\n)
    ControlSetText("Line 1`nLine 2", hEdit)
    boundsLF := RisEditControl.GetLogicalLineBoundaries(hEdit, 8)
    RisEditControl.SelectLineForRemoval(hEdit, boundsLF)
    selLF := RisEditControl.GetSel(hEdit)
    AssertTrue(selLF.Start >= 6, "LF: removal start should never eat previous line's character (start >= 6)")

    ; 情境 3: 刪除最後一行後，前一行結尾字元不可被截斷
    ControlSetText("Line 1`r`nLine 2", hEdit)
    RisEditControl.SetSel(hEdit, 9, 9)
    RisEditControl.DeleteCurrentLine(hEdit)
    AssertEqual("Line 1", ControlGetText(hEdit), "DeleteCurrentLine on last line with CRLF should preserve 'Line 1'")

    testGui.Destroy()
}

RunRegisteredTests()

