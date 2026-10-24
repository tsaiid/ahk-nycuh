#Requires AutoHotkey v2.0

#Include TestLib.v2.ahk
#Include ..\Lib\G3PacsProbe.v2.ahk

RegisterTest("G3PacsProbe button arrays definition check", Test_ButtonArraysDefinition)
RegisterTest("G3PacsProbe.IsTextMatch with string target", Test_IsTextMatch_String)
RegisterTest("G3PacsProbe.IsTextMatch with array target and trims", Test_IsTextMatch_Array)
RegisterTest("G3PacsProbe.IsPacsImageViewActive handles non-existent window safely", Test_IsPacsImageViewActive_Safe)
RegisterTest("G3PacsProbe.ClickButtonIfTextMatches handles inactive window safely", Test_ClickButtonIfTextMatches_Safe)
RegisterTest("G3PacsProbe.GetDpiScale returns valid positive scale", Test_GetDpiScale_Valid)
RegisterTest("G3PacsProbe.HasConfig detects existing and tolerant patterns", Test_HasConfig_Patterns)
RegisterTest("G3PacsProbe.GetMatchingDiagnostics handles non-configured and configured patterns safely", Test_GetMatchingDiagnostics_Safe)

Test_ButtonArraysDefinition() {
    AssertEqual(4, G3PacsProbe.SyncButtons.Length, "SyncButtons should contain 4 buttons")
    AssertEqual("Button1", G3PacsProbe.SyncButtons[1])
    AssertEqual("Button85", G3PacsProbe.SyncButtons[2])
    AssertEqual("Button90", G3PacsProbe.SyncButtons[3])
    AssertEqual("Button102", G3PacsProbe.SyncButtons[4])

    AssertEqual(4, G3PacsProbe.DiffExamSyncButtons.Length, "DiffExamSyncButtons should contain 4 buttons")
    AssertEqual("Button2", G3PacsProbe.DiffExamSyncButtons[1])
    AssertEqual("Button86", G3PacsProbe.DiffExamSyncButtons[2])
    AssertEqual("Button91", G3PacsProbe.DiffExamSyncButtons[3])
    AssertEqual("Button103", G3PacsProbe.DiffExamSyncButtons[4])

    AssertEqual(2, G3PacsProbe.ScoutLineButtons.Length, "ScoutLineButtons should contain 2 buttons")
    AssertEqual("Button4", G3PacsProbe.ScoutLineButtons[1])
    AssertEqual("Button78", G3PacsProbe.ScoutLineButtons[2])
}

Test_IsTextMatch_String() {
    AssertTrue(G3PacsProbe.IsTextMatch(" Scout lines", " Scout lines"), "Exact match should succeed")
    AssertTrue(G3PacsProbe.IsTextMatch("Scout lines", " Scout lines"), "Trimmed match should succeed")
    AssertTrue(G3PacsProbe.IsTextMatch(" Scout lines", "Scout lines"), "Trimmed match reverse should succeed")
    AssertFalse(G3PacsProbe.IsTextMatch("Other text", " Scout lines"), "Different text should return false")
}

Test_IsTextMatch_Array() {
    targets := [" Auto sync", "自動同步"]
    AssertTrue(G3PacsProbe.IsTextMatch(" Auto sync", targets), "Exact match in array should succeed")
    AssertTrue(G3PacsProbe.IsTextMatch("Auto sync", targets), "Trimmed match in array should succeed")
    AssertTrue(G3PacsProbe.IsTextMatch("自動同步", targets), "Chinese exact match in array should succeed")
    AssertFalse(G3PacsProbe.IsTextMatch("Manual sync", targets), "Unmatched text in array should return false")
}

Test_IsPacsImageViewActive_Safe() {
    ; In headless / non-PACS environment, querying non-existent window should return false without exception
    isActive := G3PacsProbe.IsPacsImageViewActive("__NON_EXISTENT_WINDOW_TITLE_12345__")
    AssertFalse(isActive, "Should safely return false when target window is inactive")
}

Test_ClickButtonIfTextMatches_Safe() {
    ; Safe fallback without throwing exception
    btn := G3PacsProbe.ClickButtonIfTextMatches(["Button1"], [" Auto sync"], "__NON_EXISTENT_WINDOW_TITLE_12345__")
    AssertEqual("", btn, "Should return empty string when window inactive")
}

Test_GetDpiScale_Valid() {
    scale := G3PacsProbe.GetDpiScale(0)
    AssertTrue(scale >= 1.0, "DpiScale should be at least 1.0")
}

Test_HasConfig_Patterns() {
    AssertTrue(G3PacsProbe.HasConfig([3, 10, 39, 101]), "[3, 10, 39, 101] should be recognized as existing")
    AssertTrue(G3PacsProbe.HasConfig([3, 10, 39, 102]), "[3, 10, 39, 102] should be recognized under +1 fault tolerance")
    AssertTrue(G3PacsProbe.HasConfig([5, 57, 29, 185]), "[5, 57, 29, 185] should be recognized as existing")
    AssertTrue(G3PacsProbe.HasConfig([5, 57, 29, 186]), "[5, 57, 29, 186] should be recognized under +1 fault tolerance")
    AssertFalse(G3PacsProbe.HasConfig([999, 999, 999, 999]), "Non-existent config should return false")
    AssertFalse(G3PacsProbe.HasConfig("invalid"), "Non-array should return false")
}

Test_GetMatchingDiagnostics_Safe() {
    diagUnk := G3PacsProbe.GetMatchingDiagnostics("AfxNonExistentClassNN", 0)
    AssertTrue(InStr(diagUnk, "未包含在任何已配置 Pattern") > 0, "Unconfigured ClassNN should be reported as not included")

    diagCandidate := G3PacsProbe.GetMatchingDiagnostics(G3PacsProbe.ClassPrefix . "3", 0)
    AssertTrue(InStr(diagCandidate, "候選規則") > 0, "Known focus ClassNN should list candidates")
    AssertTrue(InStr(diagCandidate, "Pattern_3_10_") > 0, "Candidate report should include Pattern name")
}

RunRegisteredTests()
