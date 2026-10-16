#Requires AutoHotkey v2.0

#Include TestLib.v2.ahk
#Include ..\Lib\RisController.v2.ahk

RegisterTest("RisController public methods exist", Test_RisController_MethodsExist)
RegisterTest("RisController AI cache ReqNo binding and reset", Test_RisController_AiCacheReqNoBinding)

Test_RisController_MethodsExist() {
    AssertTrue(HasMethod(RisController, "GenerateAndInsertIndication"), "GenerateAndInsertIndication method must exist")
    AssertTrue(HasMethod(RisController, "GenerateAndInsertImpression"), "GenerateAndInsertImpression method must exist")
    AssertTrue(HasMethod(RisController, "InsertExamNameAtCaret"), "InsertExamNameAtCaret method must exist")
    AssertTrue(HasMethod(RisController, "GetIndicationFollowupSuffix"), "GetIndicationFollowupSuffix method must exist")
    AssertTrue(HasMethod(RisController, "FormatFindingText"), "FormatFindingText method must exist")
    AssertTrue(HasMethod(RisController, "FormatImpressionText"), "FormatImpressionText method must exist")
    AssertTrue(HasMethod(RisController, "CopyFindingToImpression"), "CopyFindingToImpression method must exist")
    AssertTrue(HasMethod(RisController, "CopyFindingAndImpression"), "CopyFindingAndImpression method must exist")
    AssertTrue(HasMethod(RisController, "JumpToFindings"), "JumpToFindings method must exist")
    AssertTrue(HasMethod(RisController, "CompareSelectionWithAI"), "CompareSelectionWithAI method must exist")
    AssertTrue(HasMethod(RisController, "PolishSelectionWithAI"), "PolishSelectionWithAI method must exist")
    AssertTrue(HasMethod(RisController, "InsertSelectedHistoryDate"), "InsertSelectedHistoryDate method must exist")
    AssertTrue(HasMethod(RisController, "InsertCopiedReportDate"), "InsertCopiedReportDate method must exist")
    AssertTrue(HasMethod(RisController, "InsertSelectedHistoryName"), "InsertSelectedHistoryName method must exist")
    AssertTrue(HasMethod(RisController, "CopyPathologyReportOrMRN"), "CopyPathologyReportOrMRN method must exist")
    AssertTrue(HasMethod(RisController, "ClearCurrentEdit"), "ClearCurrentEdit method must exist")
    AssertTrue(HasMethod(RisEditControl, "ClearAll"), "ClearAll method must exist")
}

Test_RisController_AiCacheReqNoBinding() {
    ; 1. 驗證快取寫入包含 hwnd 與 reqNo 屬性
    dummyHwnd := 99999
    dummyReqNo := "ACC20261016001"
    RisController._CacheIndicationResult("INDICATION: Test finding", 200, 100, "Google", "gemini-1.5", dummyHwnd, dummyReqNo)

    AssertTrue(RisController._aiCache.Has("_AI_Indication"), "Cache should contain _AI_Indication")
    cached := RisController._aiCache["_AI_Indication"]
    AssertEqual(dummyHwnd, cached.hwnd, "Cache hwnd should match")
    AssertEqual(dummyReqNo, cached.reqNo, "Cache reqNo should match")

    ; 2. 驗證非 RIS 視窗環境下取用快取時，因 HWND 不合而主動使快取失效並清除
    validCached := RisController._GetValidCachedIndication()
    AssertEqual(0, validCached, "Valid cache should return 0 when window/identity mismatches")
    AssertFalse(RisController._aiCache.Has("_AI_Indication"), "Invalid cache should be deleted automatically")

    ; 3. 驗證視窗生命週期重置
    RisController._CacheIndicationResult("INDICATION: Test 2", 200, 100, "Google", "gemini-1.5", 88888, "ACC888")
    RisController._isIndicationPending := true
    RisController._pendingIndicationInsert := true
    RisController._pendingIndicationContext := { Hwnd: 88888, ReqNo: "ACC888" }

    RisController._ResetWindowScopedCaches(77777)
    AssertFalse(RisController._aiCache.Has("_AI_Indication"), "aiCache should be empty after reset")
    AssertFalse(RisController._isIndicationPending, "isIndicationPending should be reset")
    AssertFalse(RisController._pendingIndicationInsert, "pendingIndicationInsert should be reset")
    AssertEqual(0, RisController._pendingIndicationContext.Hwnd, "pendingIndicationContext.Hwnd should be 0")
}

RunRegisteredTests()
