#Requires AutoHotkey v2.0

#Include TestLib.v2.ahk
#Include ..\Lib\RisMesaService.v2.ahk

RegisterTest("RisMesaService.ParseResultHtml parses female age 50 score 123", Test_RisMesaService_ParseResultHtml_Female50Score123)
RegisterTest("RisMesaService.ParseResultHtml handles zero score", Test_RisMesaService_ParseResultHtml_ScoreZero)
RegisterTest("RisMesaService.ParseResultHtml handles out of range html", Test_RisMesaService_ParseResultHtml_ServerOutOfRange)
RegisterTest("RisMesaService.ExtractHiddenField extracts ASP.NET tokens", Test_RisMesaService_ExtractHiddenField)
RegisterTest("RisMesaService.Query handles out of range inputs offline", Test_RisMesaService_OutOfRangeInputs)

Test_RisMesaService_ParseResultHtml_Female50Score123() {
    html := '
(
<html>
<body>
    <span id="percLabel">98</span>
    <span id="Label10">16%</span>
</body>
</html>
)'
    res := RisMesaService.ParseResultHtml(html, 50, "Female", 123, "Chinese")
    AssertTrue(res.IsSuccess, "ParseResultHtml should succeed")
    AssertFalse(res.IsOutOfRange, "Should not be out of range")
    AssertEqual("98", res.Percentile, "Percentile should be 98")
    AssertEqual("16%", res.NonZeroProbability, "NonZeroProbability should be 16%")
    AssertEqual("Female", res.Gender)
    AssertEqual(50, res.Age)
    AssertEqual(123, res.TotalScore)
    AssertEqual("MESA CAC Percentile: 98% (Estimated non-zero CAC probability: 16% for age 50, Female, Chinese)", res.Summary)
}

Test_RisMesaService_ParseResultHtml_ScoreZero() {
    html := '
(
<html>
<body>
    <span id="Label10">25.3%</span>
</body>
</html>
)'
    res := RisMesaService.ParseResultHtml(html, 60, "Male", 0, "Chinese")
    AssertTrue(res.IsSuccess, "ParseResultHtml should succeed")
    AssertFalse(res.IsOutOfRange, "Should not be out of range")
    AssertEqual("", res.Percentile, "Percentile should be empty for score 0")
    AssertEqual("25.3%", res.NonZeroProbability, "NonZeroProbability should be parsed")
    AssertEqual("CAC score is 0 (No identifiable calcification; estimated non-zero CAC probability for age 60, Male, Chinese is 25.3%).", res.Summary)
}

Test_RisMesaService_ParseResultHtml_ServerOutOfRange() {
    html := '
(
<html>
<body>
    <div>The requested parameters are Out of Range for standard computation.</div>
</body>
</html>
)'
    res := RisMesaService.ParseResultHtml(html, 85, "Female", 50, "Chinese")
    AssertTrue(res.IsSuccess, "ParseResultHtml should succeed")
    AssertTrue(res.IsOutOfRange, "Should indicate out of range")
    AssertEqual("", res.Percentile, "Percentile should be empty")
    AssertEqual("Patient is outside MESA reference range (Age 45-84, Score <= 10000).", res.Summary)
}

Test_RisMesaService_ExtractHiddenField() {
    html := '
(
<form>
    <input type="hidden" name="__VIEWSTATE" id="__VIEWSTATE" value="viewStateToken123" />
    <input type="hidden" value="generatorToken456" name="__VIEWSTATEGENERATOR" id="__VIEWSTATEGENERATOR" />
</form>
)'
    AssertEqual("viewStateToken123", RisMesaService.ExtractHiddenField(html, "__VIEWSTATE"))
    AssertEqual("generatorToken456", RisMesaService.ExtractHiddenField(html, "__VIEWSTATEGENERATOR"))
    AssertEqual("", RisMesaService.ExtractHiddenField(html, "__NONEXISTENT"))
}

Test_RisMesaService_OutOfRangeInputs() {
    ; 年齡小於 45
    res1 := RisMesaService.Query(30, "M", 50, "1")
    AssertTrue(res1.IsSuccess, "Query should handle out of range gracefully")
    AssertTrue(res1.IsOutOfRange, "Age 30 should be out of range")
    AssertEqual("Patient age (30) or score (50) is outside MESA reference range (Age 45-84, Score <= 10000).", res1.Summary)

    ; 年齡大於 84
    res2 := RisMesaService.Query(85, "F", 100, "1")
    AssertTrue(res2.IsOutOfRange, "Age 85 should be out of range")

    ; 分數大於 10000
    res3 := RisMesaService.Query(55, "M", 10001, "1")
    AssertTrue(res3.IsOutOfRange, "Score 10001 should be out of range")

    ; 分數小於 0
    res4 := RisMesaService.Query(55, "M", -5, "1")
    AssertTrue(res4.IsOutOfRange, "Score -5 should be out of range")
}

RunRegisteredTests()
