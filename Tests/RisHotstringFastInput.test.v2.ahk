#Requires AutoHotkey v2.0

#Include TestLib.v2.ahk
#Include ..\Lib\RisHotstringFastInput.v2.ahk

; 定義靜態條件以供測試 HotIf
IsTestConditionActive() {
    return true
}

#HotIf IsTestConditionActive()
::hftest_short::short
::hftest_long::This is a long sentence for testing hotstring fast input.
::hftest_multiline_nl::line1`nline2
#HotIf

RegisterTest("RisHotstringFastInput._FormatFullTrigger normalizes colons", Test_FormatFullTrigger)
RegisterTest("RisHotstringFastInput._HasKeyCommands detects special keys", Test_HasKeyCommands)
RegisterTest("RisHotstringFastInput.Enable filters by minLength, newline, and special keys", Test_EnableFiltering)
RegisterTest("RisHotstringFastInput.Disable restores state", Test_Disable)
RegisterTest("RisHotstringFastInput.Toggle switches state", Test_Toggle)
RegisterTest("RisHotstringFastInput.Toggle without arguments preserves condition", Test_TogglePreservesCondition)

Test_FormatFullTrigger() {
    AssertEqual("::livok1", RisHotstringFastInput._FormatFullTrigger(":", "livok1"), "Single colon option should become double colon")
    AssertEqual(":c:RUL", RisHotstringFastInput._FormatFullTrigger(":c", "RUL"), "Option without trailing colon should append colon")
    AssertEqual("::livok1", RisHotstringFastInput._FormatFullTrigger("::", "livok1"), "Double colon option should remain double colon")
    AssertEqual(":c:RUL", RisHotstringFastInput._FormatFullTrigger(":c:", "RUL"), "Complete option should remain unchanged")
}

Test_HasKeyCommands() {
    AssertTrue(RisHotstringFastInput._HasKeyCommands({ options: ":", replacement: "up to  cm in size{Left 11}" }), "Should detect {Left 11}")
    AssertTrue(RisHotstringFastInput._HasKeyCommands({ options: ":", replacement: "First line{Enter}Second line" }), "Should detect {Enter}")
    AssertTrue(RisHotstringFastInput._HasKeyCommands({ options: ":", replacement: "Findings.+{Tab}Impression" }), "Should detect {Tab}")
    AssertFalse(RisHotstringFastInput._HasKeyCommands({ options: ":", replacement: "Normal report text without keys." }), "Normal text should be false")
    AssertFalse(RisHotstringFastInput._HasKeyCommands({ options: ":R:", replacement: "Raw text {Left 11}" }), "Raw mode with R should be false")
    AssertFalse(RisHotstringFastInput._HasKeyCommands({ options: ":t:", replacement: "Text mode {Enter}" }), "Text mode with t should be false")
}

Test_EnableFiltering() {
    RisHotstringPalette._cache := [
        {
            trigger: "hftest_short",
            options: ":",
            replacement: "short",
            lines: ["short"],
            isMultiLine: false,
            isFunction: false,
            file: "test.ahk"
        },
        {
            trigger: "hftest_long",
            options: ":",
            replacement: "This is a long sentence for testing hotstring fast input.",
            lines: ["This is a long sentence for testing hotstring fast input."],
            isMultiLine: false,
            isFunction: false,
            file: "test.ahk"
        },
        {
            trigger: "hftest_multiline_nl",
            options: ":",
            replacement: "line1`nline2",
            lines: ["line1", "line2"],
            isMultiLine: false,
            isFunction: false,
            file: "test.ahk"
        },
        {
            trigger: "hftest_func",
            options: ":",
            replacement: "[SomeFunc()]",
            lines: ["[SomeFunc()]"],
            isMultiLine: false,
            isFunction: true,
            file: "test.ahk"
        },
        {
            trigger: "hftest_block",
            options: ":",
            replacement: "multi`nline`nform",
            lines: ["multi", "line", "form"],
            isMultiLine: true,
            isFunction: false,
            file: "test.ahk"
        },
        {
            trigger: "su",
            options: ":",
            replacement: "up to  cm in size{Left 11}",
            lines: ["up to  cm in size{Left 11}"],
            isMultiLine: false,
            isFunction: false,
            file: "test.ahk"
        }
    ]
    RisHotstringPalette._isLoaded := true

    RisHotstringFastInput._originalHotstrings := Map()
    count := RisHotstringFastInput.Enable(20, "IsTestConditionActive()")

    AssertEqual(2, count, "Should enable fast input for 2 items (skipping short, func, block, and su)")
    AssertTrue(RisHotstringFastInput.isEnabled, "isEnabled should be true")
    AssertEqual(2, RisHotstringFastInput._count, "_count should be 2")
    AssertTrue(RisHotstringFastInput._originalHotstrings.Has("::hftest_long"), "Should record hftest_long")
    AssertTrue(RisHotstringFastInput._originalHotstrings.Has("::hftest_multiline_nl"), "Should record hftest_multiline_nl")
    AssertFalse(RisHotstringFastInput._originalHotstrings.Has("::hftest_short"), "Should not record hftest_short")
    AssertFalse(RisHotstringFastInput._originalHotstrings.Has("::su"), "Should not record su because it has {Left 11}")
}

Test_Disable() {
    RisHotstringFastInput.Disable("IsTestConditionActive()")
    AssertFalse(RisHotstringFastInput.isEnabled, "isEnabled should be false after disable")
    AssertEqual(0, RisHotstringFastInput._count, "_count should be 0 after disable")
}

Test_Toggle() {
    newState := RisHotstringFastInput.Toggle("IsTestConditionActive()")
    AssertTrue(newState, "Toggle from disabled should return true")
    AssertTrue(RisHotstringFastInput.isEnabled, "isEnabled should be true")

    newState2 := RisHotstringFastInput.Toggle("IsTestConditionActive()")
    AssertFalse(newState2, "Toggle from enabled should return false")
    AssertFalse(RisHotstringFastInput.isEnabled, "isEnabled should be false")
}

Test_TogglePreservesCondition() {
    RisHotstringFastInput.Enable(20, "IsTestConditionActive()")
    AssertTrue(RisHotstringFastInput.isEnabled, "isEnabled should be true initially")
    AssertEqual("IsTestConditionActive()", RisHotstringFastInput.conditionExpr, "conditionExpr should be IsTestConditionActive()")

    ; 第一次 Toggle（停用，不傳 condition 參數，模擬 Tray 選單行為）
    RisHotstringFastInput.Toggle()
    AssertFalse(RisHotstringFastInput.isEnabled, "isEnabled should be false after toggle disable")
    AssertEqual("IsTestConditionActive()", RisHotstringFastInput.conditionExpr, "conditionExpr should still be IsTestConditionActive()")

    ; 第二次 Toggle（重新啟用，不傳 condition 參數，模擬 Tray 選單行為）
    RisHotstringFastInput.Toggle()
    AssertTrue(RisHotstringFastInput.isEnabled, "isEnabled should be true after toggle enable")
    AssertEqual("IsTestConditionActive()", RisHotstringFastInput.conditionExpr, "conditionExpr should not be wiped to empty string")
    AssertTrue(RisHotstringFastInput._count > 0, "count should be > 0 when re-enabled")

    ; 清理狀態
    RisHotstringFastInput.Disable()
}

RunRegisteredTests()
