#Requires AutoHotkey v2.0

#Include TestLib.v2.ahk
#Include ..\Lib\RisReportText.v2.ahk

RegisterTest("RisReportText.GetExamType detects CT", Test_RisReportText_GetExamTypeCt)
RegisterTest("RisReportText.DeidentifyText removes common PHI", Test_RisReportText_DeidentifyText)
RegisterTest("RisReportText.ReorderSelectedText normalizes list items", Test_RisReportText_ReorderSelectedText)
RegisterTest("RisReportText.ExtractPastFinding strips header/indication", Test_RisReportText_ExtractPastFinding)
RegisterTest("RisReportText.ExtractTotalCalciumScore extracts Agatston scores", Test_RisReportText_ExtractTotalCalciumScore)
RegisterTest("RisReportText.IsCalciumScoreExam detects calcium score exams", Test_RisReportText_IsCalciumScoreExam)
RegisterTest("RisReportText.GetCalciumScoreSeverity classifies scores", Test_RisReportText_GetCalciumScoreSeverity)
RegisterTest("RisReportText.FormatCalciumScoreImpression formats templates correctly", Test_RisReportText_FormatCalciumScoreImpression)

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

Test_RisReportText_ExtractTotalCalciumScore() {
    sample1 := "- Total Calcium Score (Equivalent Agatston Score) is 0`r`n LM calcium score is 0"
    AssertEqual("0", RisReportText.ExtractTotalCalciumScore(sample1))

    sample2 := "Total Calcium Score (Equivalent Agatston Score) is 158.4"
    AssertEqual("158.4", RisReportText.ExtractTotalCalciumScore(sample2))

    sample3 := "Total Calcium Score: 320"
    AssertEqual("320", RisReportText.ExtractTotalCalciumScore(sample3))

    sample4 := "Total Score is 45"
    AssertEqual("45", RisReportText.ExtractTotalCalciumScore(sample4))

    sampleEmpty := "No mention of score here."
    AssertEqual("", RisReportText.ExtractTotalCalciumScore(sampleEmpty))
}

Test_RisReportText_IsCalciumScoreExam() {
    AssertTrue(RisReportText.IsCalciumScoreExam("calcium score of coronary artery"))
    AssertTrue(RisReportText.IsCalciumScoreExam("CALCIUM SCORE OF CORONARY ARTERY"))
    AssertTrue(RisReportText.IsCalciumScoreExam("檢查項目: calcium score of coronary artery"))
    AssertTrue(RisReportText.IsCalciumScoreExam("CT calcium score"))
    AssertTrue(RisReportText.IsCalciumScoreExam("冠狀動脈鈣化評分"))
    AssertFalse(RisReportText.IsCalciumScoreExam("Chest CT without contrast"))
    AssertFalse(RisReportText.IsCalciumScoreExam("CT BRAIN"))
}

Test_RisReportText_GetCalciumScoreSeverity() {
    AssertEqual("No identifiable", RisReportText.GetCalciumScoreSeverity(0))
    AssertEqual("Minimal identifiable", RisReportText.GetCalciumScoreSeverity(5))
    AssertEqual("Mild", RisReportText.GetCalciumScoreSeverity(50))
    AssertEqual("Moderate", RisReportText.GetCalciumScoreSeverity(123))
    AssertEqual("Significant", RisReportText.GetCalciumScoreSeverity(450))
    AssertEqual("1st", RisReportText.GetOrdinal(1))
    AssertEqual("2nd", RisReportText.GetOrdinal(2))
    AssertEqual("3rd", RisReportText.GetOrdinal(3))
    AssertEqual("11th", RisReportText.GetOrdinal(11))
    AssertEqual("21st", RisReportText.GetOrdinal(21))
    AssertEqual("98th", RisReportText.GetOrdinal(98))
}

Test_RisReportText_FormatCalciumScoreImpression() {
    mesaNormal := {
        IsSuccess: true,
        IsOutOfRange: false,
        Percentile: "98",
        NonZeroProbability: "16%"
    }
    expectedNormal := "Total Calcium Score (Equivalent Agatston Score) is 123 (Moderate calcification; 98th percentile compared to age-, sex-, and race-matched peers in MESA; probability of non-zero CAC: 16%)."
    AssertEqual(expectedNormal, RisReportText.FormatCalciumScoreImpression(123, mesaNormal))

    mesaZero := {
        IsSuccess: true,
        IsOutOfRange: false,
        Percentile: "",
        NonZeroProbability: "16%"
    }
    expectedZero := "Total Calcium Score (Equivalent Agatston Score) is 0 (No identifiable calcification; probability of non-zero CAC in demographic peers is 16%)."
    AssertEqual(expectedZero, RisReportText.FormatCalciumScoreImpression(0, mesaZero))

    mesaOut := {
        IsSuccess: true,
        IsOutOfRange: true,
        Percentile: "",
        NonZeroProbability: ""
    }
    expectedOut := "Total Calcium Score (Equivalent Agatston Score) is 50 (Mild calcification; patient is outside MESA reference age range of 45-84 years)."
    AssertEqual(expectedOut, RisReportText.FormatCalciumScoreImpression(50, mesaOut))
}

RunRegisteredTests()

