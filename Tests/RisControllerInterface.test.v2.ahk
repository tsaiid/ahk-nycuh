#Requires AutoHotkey v2.0

#Include TestLib.v2.ahk
#Include ..\Lib\RisController.v2.ahk

RegisterTest("RisController public methods exist", Test_RisController_MethodsExist)

Test_RisController_MethodsExist() {
    AssertTrue(HasMethod(RisController, "GenerateAndInsertIndication"), "GenerateAndInsertIndication method must exist")
    AssertTrue(HasMethod(RisController, "GenerateAndInsertImpression"), "GenerateAndInsertImpression method must exist")
    AssertTrue(HasMethod(RisController, "InsertExamNameAtCaret"), "InsertExamNameAtCaret method must exist")
    AssertTrue(HasMethod(RisController, "GetIndicationFollowupSuffix"), "GetIndicationFollowupSuffix method must exist")
    AssertTrue(HasMethod(RisController, "FormatFindingText"), "FormatFindingText method must exist")
    AssertTrue(HasMethod(RisController, "FormatImpressionText"), "FormatImpressionText method must exist")
    AssertTrue(HasMethod(RisController, "CopyFindingToImpression"), "CopyFindingToImpression method must exist")
    AssertTrue(HasMethod(RisController, "CompareSelectionWithAI"), "CompareSelectionWithAI method must exist")
    AssertTrue(HasMethod(RisController, "PolishSelectionWithAI"), "PolishSelectionWithAI method must exist")
    AssertTrue(HasMethod(RisController, "InsertSelectedHistoryDate"), "InsertSelectedHistoryDate method must exist")
    AssertTrue(HasMethod(RisController, "InsertCopiedReportDate"), "InsertCopiedReportDate method must exist")
    AssertTrue(HasMethod(RisController, "InsertSelectedHistoryName"), "InsertSelectedHistoryName method must exist")
    AssertTrue(HasMethod(RisController, "CopyPathologyReportOrMRN"), "CopyPathologyReportOrMRN method must exist")
}

RunRegisteredTests()
