#Requires AutoHotkey v2.0

#Include TestLib.v2.ahk
#Include ..\Lib\RisAIDebugGui.v2.ahk

RegisterTest("RisAIDebugGui.ApplyPolishProviderChoice invokes choice callbacks", Test_RisAIDebugGui_ApplyPolishProviderChoice)
RegisterTest("RisAIDebugGui.ApplyPolishProviderChoice handles uninitialized callbacks safely", Test_RisAIDebugGui_ApplyPolishProviderChoice_Safe)
RegisterTest("RisAIDebugGui._AlignParagraphs handles multi-paragraph and fallback properly", Test_RisAIDebugGui_AlignParagraphs)

Test_RisAIDebugGui_ApplyPolishProviderChoice() {
    choice0Called := false
    choice1Called := false
    choice2Called := false
    choice3Called := false

    RisAIDebugGui.applyOriginalChoiceFunc := () => (choice0Called := true)
    RisAIDebugGui.applyFirstChoiceFunc := () => (choice1Called := true)
    RisAIDebugGui.applySecondChoiceFunc := () => (choice2Called := true)
    RisAIDebugGui.applyCustomChoiceFunc := () => (choice3Called := true)

    try {
        RisAIDebugGui.ApplyPolishProviderChoice(0)
        AssertTrue(choice0Called, "Choice 0 (Original) callback should be executed")
        AssertFalse(choice1Called, "Choice 1 should not be executed yet")

        RisAIDebugGui.ApplyPolishProviderChoice(1)
        AssertTrue(choice1Called, "Choice 1 (OpenAI) callback should be executed")
        AssertFalse(choice2Called, "Choice 2 should not be executed yet")

        RisAIDebugGui.ApplyPolishProviderChoice(2)
        AssertTrue(choice2Called, "Choice 2 (Google) callback should be executed")
        AssertFalse(choice3Called, "Choice 3 should not be executed yet")

        RisAIDebugGui.ApplyPolishProviderChoice(3)
        AssertTrue(choice3Called, "Choice 3 (Custom) callback should be executed")
    } finally {
        RisAIDebugGui.applyOriginalChoiceFunc := 0
        RisAIDebugGui.applyFirstChoiceFunc := 0
        RisAIDebugGui.applySecondChoiceFunc := 0
        RisAIDebugGui.applyCustomChoiceFunc := 0
    }
}

Test_RisAIDebugGui_ApplyPolishProviderChoice_Safe() {
    RisAIDebugGui.applyOriginalChoiceFunc := 0
    RisAIDebugGui.applyFirstChoiceFunc := 0
    RisAIDebugGui.applySecondChoiceFunc := 0
    RisAIDebugGui.applyCustomChoiceFunc := 0

    ; Should not throw when callbacks are not registered or index is invalid
    RisAIDebugGui.ApplyPolishProviderChoice(0)
    RisAIDebugGui.ApplyPolishProviderChoice(1)
    RisAIDebugGui.ApplyPolishProviderChoice(2)
    RisAIDebugGui.ApplyPolishProviderChoice(3)
    RisAIDebugGui.ApplyPolishProviderChoice(4)
    AssertTrue(true, "Handled safely without error")
}

Test_RisAIDebugGui_AlignParagraphs() {
    ; 1. Double-newline multi-paragraph
    origDbl := "First para.`r`n`r`nSecond para."
    openAIDbl := "OpenAI first.`r`n`r`nOpenAI second."
    googleDbl := "Google first.`r`n`r`nGoogle second."
    res1 := RisAIDebugGui._AlignParagraphs(origDbl, openAIDbl, googleDbl)
    AssertTrue(res1.IsMulti, "Should detect multi-paragraph with double newlines")
    AssertEqual(2, res1.Orig.Length, "Should have 2 paragraphs")
    AssertEqual("First para.", res1.Orig[1], "First original paragraph match")
    AssertEqual("`r`n`r`n", res1.Separator, "Separator should be double newline")

    ; 2. Single-newline multi-line (e.g. bullet points)
    origLines := "- Point 1`r`n- Point 2`r`n- Point 3"
    openAILines := "- Point 1 AI`r`n- Point 2 AI`r`n- Point 3 AI"
    googleLines := "- Point 1 G`r`n- Point 2 G`r`n- Point 3 G"
    res2 := RisAIDebugGui._AlignParagraphs(origLines, openAILines, googleLines)
    AssertTrue(res2.IsMulti, "Should detect multi-paragraph with single newlines")
    AssertEqual(3, res2.Orig.Length, "Should have 3 items")
    AssertEqual("`r`n", res2.Separator, "Separator should be single newline")

    ; 3. Paragraph count mismatch -> Fallback to single block
    origMismatch := "Para 1`r`n`r`nPara 2"
    openAIMismatch := "Para 1 AI`r`n`r`nPara 2 AI"
    googleMismatch := "Single combined paragraph from Google."
    res3 := RisAIDebugGui._AlignParagraphs(origMismatch, openAIMismatch, googleMismatch)
    AssertFalse(res3.IsMulti, "Mismatch should fallback to single block")
    AssertEqual(1, res3.Orig.Length, "Fallback length should be 1")
    AssertEqual(origMismatch, res3.Orig[1], "Fallback original should match full text")

    ; 4. Single paragraph -> Fallback to single block
    origSingle := "Single sentence."
    openAISingle := "Single sentence polished."
    googleSingle := "Single sentence by Google."
    res4 := RisAIDebugGui._AlignParagraphs(origSingle, openAISingle, googleSingle)
    AssertFalse(res4.IsMulti, "Single paragraph should not be multi-mode")
    AssertEqual(1, res4.Orig.Length, "Length should be 1")
}

RunRegisteredTests()
