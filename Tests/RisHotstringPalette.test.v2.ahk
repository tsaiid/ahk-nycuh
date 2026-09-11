#Requires AutoHotkey v2.0

#Include TestLib.v2.ahk
#Include ..\Lib\RisHotstringPalette.v2.ahk

RegisterTest("RisHotstringPalette._ParseFileContent parses single-line hotstrings", Test_ParseSingleLine)
RegisterTest("RisHotstringPalette._ParseFileContent parses multi-line continuation section", Test_ParseMultiLineContinuation)
RegisterTest("RisHotstringPalette._ParseFileContent parses function call hotstring", Test_ParseFunctionCall)
RegisterTest("RisHotstringPalette._FindMatchingSnippet returns line matching keyword", Test_FindMatchingSnippet_Keyword)
RegisterTest("RisHotstringPalette._FindMatchingSnippet returns default line when trigger matches", Test_FindMatchingSnippet_Trigger)
RegisterTest("RisHotstringPalette.Filter matches words and ranks trigger exact match first", Test_FilterAndRanking)
RegisterTest("RisHotstringPalette._CreateGui initializes controls without option errors", Test_CreateGui)

Test_ParseSingleLine() {
    sample := "
    (
    ::ggo::ground-glass opacity `` 
    :c:RUL::right upper lobe
    ::tb::tuberculosis ; pulmonary infection
    )"
    RisHotstringPalette._cache := []
    RisHotstringPalette._ParseFileContent(sample, "test.ahk")

    AssertEqual(3, RisHotstringPalette._cache.Length, "Should parse 3 items")
    AssertEqual("ggo", RisHotstringPalette._cache[1].trigger, "Item 1 trigger")
    AssertEqual("ground-glass opacity ", RisHotstringPalette._cache[1].replacement, "Item 1 replacement with trailing space")
    AssertEqual(false, RisHotstringPalette._cache[1].isMultiLine, "Item 1 isMultiLine")

    AssertEqual("RUL", RisHotstringPalette._cache[2].trigger, "Item 2 trigger")
    AssertEqual("right upper lobe", RisHotstringPalette._cache[2].replacement, "Item 2 replacement")

    AssertEqual("tb", RisHotstringPalette._cache[3].trigger, "Item 3 trigger")
    AssertEqual("tuberculosis", RisHotstringPalette._cache[3].replacement, "Item 3 replacement stripped inline comment")
}

Test_ParseMultiLineContinuation() {
    sample := "::ccttrok0::`n"
        . "{`n"
        . '    MyForm := "`n'
        . "  (`n"
        . "No pneumothorax or hemothorax.`n"
        . "No lung contusion, pneumothorax, or hemothorax.`n"
        . '  )"`n'
        . "    Paste(MyForm)`n"
        . "}"
    RisHotstringPalette._cache := []
    RisHotstringPalette._ParseFileContent(sample, "test.ahk")

    AssertEqual(1, RisHotstringPalette._cache.Length, "Should parse 1 multi-line item")
    item := RisHotstringPalette._cache[1]
    AssertEqual("ccttrok0", item.trigger, "Multi-line trigger")
    AssertTrue(item.isMultiLine, "Should be multi-line")
    AssertEqual(2, item.lines.Length, "Should have 2 lines")
    AssertEqual("No pneumothorax or hemothorax.", Trim(item.lines[1]), "First line")
    AssertEqual("No lung contusion, pneumothorax, or hemothorax.", Trim(item.lines[2]), "Second line")
}

Test_ParseFunctionCall() {
    sample := "
    (
    ::fsg::
    {
        parentWnd := WinExist("A")
        Fleischner2017Form()
    }
    )"
    RisHotstringPalette._cache := []
    RisHotstringPalette._ParseFileContent(sample, "test.ahk")

    AssertEqual(1, RisHotstringPalette._cache.Length, "Should parse 1 function item")
    item := RisHotstringPalette._cache[1]
    AssertEqual("fsg", item.trigger, "Function trigger")
    AssertTrue(item.isFunction, "Should be marked as function")
    AssertEqual("[Fleischner2017Form()]", item.replacement, "Function replacement indicator")
}

Test_FindMatchingSnippet_Keyword() {
    item := {
        trigger: "ccttrok",
        replacement: "No pneumothorax or hemothorax.`nNo lung contusion.`nThe heart appears normal.",
        lines: [
            "No pneumothorax or hemothorax.",
            "No lung contusion.",
            "The heart appears normal."
        ],
        isMultiLine: true
    }

    snippet := RisHotstringPalette._FindMatchingSnippet(item, ["contusion"])
    AssertEqual("No lung contusion.", snippet, "Should return the line containing 'contusion'")

    snippet2 := RisHotstringPalette._FindMatchingSnippet(item, ["heart"])
    AssertEqual("The heart appears normal.", snippet2, "Should return the line containing 'heart'")
}

Test_FindMatchingSnippet_Trigger() {
    item := {
        trigger: "ccttrok",
        replacement: "No pneumothorax or hemothorax.`nNo lung contusion.`nThe heart appears normal.",
        lines: [
            "No pneumothorax or hemothorax.",
            "No lung contusion.",
            "The heart appears normal."
        ],
        isMultiLine: true
    }

    ; When searching for trigger keyword, not in lines
    snippet := RisHotstringPalette._FindMatchingSnippet(item, ["ccttr"])
    AssertEqual("No pneumothorax or hemothorax.", snippet, "Should default to first non-empty line")
}

Test_FilterAndRanking() {
    RisHotstringPalette._cache := [
        {
            trigger: "lung",
            options: "",
            replacement: "general lung finding",
            lines: ["general lung finding"],
            isMultiLine: false,
            isFunction: false,
            file: "test.ahk"
        },
        {
            trigger: "cctali",
            options: "",
            replacement: "acute lung injury and pulmonary edema",
            lines: ["acute lung injury and pulmonary edema"],
            isMultiLine: false,
            isFunction: false,
            file: "test.ahk"
        },
        {
            trigger: "lung_nodule",
            options: "",
            replacement: "suspicious lung nodule",
            lines: ["suspicious lung nodule"],
            isMultiLine: false,
            isFunction: false,
            file: "test.ahk"
        }
    ]

    RisHotstringPalette.Filter("lung")
    AssertEqual(3, RisHotstringPalette._filteredItems.Length, "All 3 items contain 'lung'")
    AssertEqual("lung", RisHotstringPalette._filteredItems[1].item.trigger, "Exact trigger match 'lung' should be ranked first")
}

Test_CreateGui() {
    RisHotstringPalette._CreateGui()
    try {
        AssertTrue(RisHotstringPalette._gui != 0, "Gui should be created")
        AssertTrue(RisHotstringPalette.Hwnd > 0, "Hwnd should exist")
    } finally {
        RisHotstringPalette.Close()
    }
}

RunRegisteredTests()
