#Requires AutoHotkey v2.0

#Include TestLib.v2.ahk
#Include ..\Lib\RisReportText.v2.ahk

RegisterTest("RisReportText.GetExamType detects CT", Test_RisReportText_GetExamTypeCt)
RegisterTest("RisReportText.DeidentifyText removes common PHI", Test_RisReportText_DeidentifyText)
RegisterTest("RisReportText.ReorderSelectedText normalizes list items", Test_RisReportText_ReorderSelectedText)
RegisterTest("RisReportText.ExtractPastFinding strips header/indication", Test_RisReportText_ExtractPastFinding)

Test_RisReportText_GetExamTypeCt() {
    AssertEqual("CT", RisReportText.GetExamType("Abdomen CT with contrast"))
    AssertEqual("CT", RisReportText.GetExamType("腹部電腦斷層"))
}

Test_RisReportText_DeidentifyText() {
    inputText := "Name: Wang Test`r`nID: A123456789`r`nDate: 115/03/13`r`nPhone: 0912345678"
    actualText := RisReportText.DeidentifyText(inputText)

    AssertTrue(InStr(actualText, "Name: [PATIENT_NAME]"), "name should be redacted")
    AssertTrue(InStr(actualText, "[ID_REDACTED]"), "ID should be redacted")
    AssertTrue(InStr(actualText, "[DATE]"), "date should be redacted")
    AssertTrue(InStr(actualText, "[PHONE_REDACTED]"), "phone should be redacted")
}

Test_RisReportText_ReorderSelectedText() {
    inputText := "- mild fatty liver`r`n- renal cyst"
    expectedText := "1. Mild fatty liver.`r`n2. Renal cyst."

    AssertEqual(expectedText, RisReportText.ReorderSelectedText(inputText))
}

Test_RisReportText_ExtractPastFinding() {
    structuredInput := "low dose lung CT (without contrast):`r`n`r`nINDICATION: 健檢檢查`r`n`r`nFINDINGS:`r`n- Fibrotic scarring`r`n- Mild fatty liver."
    expectedStructured := "- Fibrotic scarring`r`n- Mild fatty liver."
    AssertEqual(expectedStructured, RisReportText.ExtractPastFinding(structuredInput))

    structuredWithExtraNewlines := "low dose lung CT (without contrast):`r`n`r`nINDICATION: 健檢檢查`r`n`r`nFINDINGS:`r`n`r`n`r`n- Fibrotic scarring`r`n- Mild fatty liver."
    AssertEqual(expectedStructured, RisReportText.ExtractPastFinding(structuredWithExtraNewlines))

    plainInput := "- Fibrotic scarring`r`n- Mild fatty liver."
    AssertEqual(plainInput, RisReportText.ExtractPastFinding(plainInput))
}

RunRegisteredTests()

