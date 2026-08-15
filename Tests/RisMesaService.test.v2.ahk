#Requires AutoHotkey v2.0

#Include TestLib.v2.ahk
#Include ..\Lib\RisMesaService.v2.ahk

RegisterTest("RisMesaService.Query calculates female age 50 score 123", Test_RisMesaService_Female50Score123)
RegisterTest("RisMesaService.Query handles out of range age", Test_RisMesaService_OutOfRange)

Test_RisMesaService_Female50Score123() {
    res := RisMesaService.Query(50, "F", 123, "1")
    AssertTrue(res.IsSuccess, "Query should succeed")
    AssertEqual("98", res.Percentile, "Percentile should be 98")
    AssertEqual("16%", res.NonZeroProbability, "NonZeroProbability should be 16%")
    AssertEqual("Female", res.Gender)
    AssertEqual(50, res.Age)
    AssertEqual(123, res.TotalScore)
}

Test_RisMesaService_OutOfRange() {
    res := RisMesaService.Query(30, "M", 50, "1")
    AssertTrue(res.IsSuccess, "Query should handle out of range gracefully")
    AssertTrue(res.IsOutOfRange, "Age 30 should be out of range")
}

RunRegisteredTests()
