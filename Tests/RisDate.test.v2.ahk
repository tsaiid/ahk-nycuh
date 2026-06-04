#Requires AutoHotkey v2.0

#Include TestLib.v2.ahk
#Include ..\Lib\RisDate.v2.ahk

RegisterTest("RisDate.ConvertRISDate converts western compact date", Test_RisDate_ConvertWesternCompactDate)
RegisterTest("RisDate.ConvertRISDate converts ROC compact date", Test_RisDate_ConvertRocCompactDate)
RegisterTest("RisDate.ConvertRISDate preserves unknown input", Test_RisDate_PreserveUnknownInput)

Test_RisDate_ConvertWesternCompactDate() {
    AssertEqual("2026-03-13", RisDate.ConvertRISDate("20260313"))
}

Test_RisDate_ConvertRocCompactDate() {
    AssertEqual("2026-03-13", RisDate.ConvertRISDate("1150313"))
}

Test_RisDate_PreserveUnknownInput() {
    AssertEqual("not a date", RisDate.ConvertRISDate("not a date"))
}

RunRegisteredTests()
