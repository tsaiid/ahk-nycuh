#Requires AutoHotkey v2.0

#Include TestLib.v2.ahk
#Include ..\Lib\RisHotstringPalette.v2.ahk

RegisterTest("RisHotstringPalette._ParseFileContent parses single-line hotstrings", Test_ParseSingleLine)
RegisterTest("RisHotstringPalette._ParseFileContent parses multi-line continuation section", Test_ParseMultiLineContinuation)
RegisterTest("RisHotstringPalette._ParseFileContent parses function call hotstring", Test_ParseFunctionCall)
RegisterTest("RisHotstringPalette._FindMatchingSnippet returns line matching keyword", Test_FindMatchingSnippet_Keyword)
RegisterTest("RisHotstringPalette._FindMatchingSnippet returns default line when trigger matches", Test_FindMatchingSnippet_Trigger)
RegisterTest("RisHotstringPalette.Filter matches words and ranks trigger exact match first", Test_FilterAndRanking)
RegisterTest("RisHotstringPalette._EscapeHtml escapes special HTML characters", Test_EscapeHtml)
RegisterTest("RisHotstringPalette._Highlight wraps matched keywords in span tags", Test_Highlight)
RegisterTest("RisHotstringPalette._GetPreviewHtml converts newlines to br and preserves indentation", Test_GetPreviewHtml)
RegisterTest("RisHotstringPalette._CreateGui initializes controls without option errors", Test_CreateGui)
RegisterTest("RisHotstringPalette._CreateGui initializes controls with custom height", Test_CreateGui_CustomHeight)
RegisterTest("RisHotstringPalette.CopySelection handles Edit and HTML selection gracefully", Test_CopySelection)
RegisterTest("RisHotstringPalette._ParseFileContent parses same-line brace hotstrings", Test_ParseSameLineBrace)
RegisterTest("RisHotstringPalette._SortResults sorts matches in descending score order", Test_SortResultsDescending)
RegisterTest("RisHotstringPalette._FlushPendingSearch executes pending search immediately", Test_FlushPendingSearch)

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

Test_EscapeHtml() {
    raw := '<div class="test">& "hello"</div>'
    escaped := RisHotstringPalette._EscapeHtml(raw)
    AssertEqual("&lt;div class=&quot;test&quot;&gt;&amp; &quot;hello&quot;&lt;/div&gt;", escaped, "HTML characters should be escaped")
}

Test_Highlight() {
    text := "No pulmonary nodule or consolidation."
    highlighted := RisHotstringPalette._Highlight(text, ["nodule", "pulmonary"])
    expected := "No <span class='hl' style='background-color:#FEF08A;color:#854D0E;font-weight:bold;padding:0 2px;'>pulmonary</span> <span class='hl' style='background-color:#FEF08A;color:#854D0E;font-weight:bold;padding:0 2px;'>nodule</span> or consolidation."
    AssertEqual(expected, highlighted, "Matched words should be wrapped in highlight span tags")
}

Test_GetPreviewHtml() {
    RisHotstringPalette._filteredItems := [
        {
            item: {
                trigger: "testtrig",
                replacement: "Heading 1`r`nHeading 2`n  * Indented detail",
                lines: ["Heading 1", "Heading 2", "  * Indented detail"],
                isMultiLine: true,
                isFunction: false,
                file: "test.ahk"
            },
            snippet: "Heading 1",
            score: 100
        }
    ]
    RisHotstringPalette._searchTerms := ["Heading"]

    preview := RisHotstringPalette._GetPreviewHtml(1)
    hl := "<span class='hl' style='background-color:#FEF08A;color:#854D0E;font-weight:bold;padding:0 2px;'>Heading</span>"
    expected := hl . " 1<br>" . hl . " 2<br>&nbsp;&nbsp;* Indented detail"
    AssertEqual(expected, preview, "Preview HTML should convert newlines to <br> and indentations to &nbsp;")
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

Test_CreateGui_CustomHeight() {
    RisHotstringPalette._CreateGui(1000, 968, 800)
    try {
        AssertTrue(RisHotstringPalette._gui != 0, "Gui should be created with custom height")
        AssertTrue(RisHotstringPalette.Hwnd > 0, "Hwnd should exist")
    } finally {
        RisHotstringPalette.Close()
    }
}

Test_CopySelection() {
    RisHotstringPalette._CreateGui()
    try {
        ; 測試 1: 無選取文字時執行不應拋出例外
        RisHotstringPalette.CopySelection()

        ; 測試 2: Edit 控制項選取並複製
        editCtrl := RisHotstringPalette._editSearch
        editCtrl.Value := "CopyTestSearch"
        editCtrl.Focus()
        SendMessage(0x00B1, 0, 8, editCtrl.Hwnd) ; EM_SETSEL: 選取 "CopyTest"
        A_Clipboard := ""
        RisHotstringPalette.CopySelection()
        AssertEqual("CopyTest", A_Clipboard, "Should copy selected text from edit control")
    } finally {
        RisHotstringPalette.Close()
    }
}

Test_ParseSameLineBrace() {
    sample := "::sk:: {`n"
        . '    MyForm := "`n'
        . "(`n"
        . "The bowel gas pattern is unremarkable.`n"
        . "No subphrenic free air.`n"
        . ')"`n'
        . "    Paste(MyForm)`n"
        . "}"
    RisHotstringPalette._cache := []
    RisHotstringPalette._ParseFileContent(sample, "test.ahk")

    AssertEqual(1, RisHotstringPalette._cache.Length, "Should parse 1 item with same-line brace")
    item := RisHotstringPalette._cache[1]
    AssertEqual("sk", item.trigger, "Trigger should be 'sk'")
    AssertTrue(item.isMultiLine, "Should be multi-line")
    AssertEqual(2, item.lines.Length, "Should have 2 lines")
    AssertEqual("The bowel gas pattern is unremarkable.", Trim(item.lines[1]), "First line")
    AssertEqual("No subphrenic free air.", Trim(item.lines[2]), "Second line")
}

Test_SortResultsDescending() {
    items := [
        { score: 10, name: "low" },
        { score: 250, name: "high" },
        { score: 50, name: "mid" },
        { score: 150, name: "midhigh" }
    ]
    RisHotstringPalette._SortResults(items)
    AssertEqual(250, items[1].score, "First item should have highest score")
    AssertEqual(150, items[2].score, "Second item score")
    AssertEqual(50, items[3].score, "Third item score")
    AssertEqual(10, items[4].score, "Fourth item score")
}

Test_FlushPendingSearch() {
    RisHotstringPalette._cache := [
        { trigger: "sk", options: "", replacement: "kidney shadow unremarkable", lines: ["kidney shadow unremarkable"], isMultiLine: false, isFunction: false, file: "test.ahk" },
        { trigger: "sono", options: "", replacement: "ultrasound study", lines: ["ultrasound study"], isMultiLine: false, isFunction: false, file: "test.ahk" }
    ]
    RisHotstringPalette._CreateGui()
    try {
        editCtrl := RisHotstringPalette._editSearch
        editCtrl.Value := "sk"
        RisHotstringPalette._OnSearchChange()
        AssertTrue(RisHotstringPalette._debounceTimer != 0, "Debounce timer should be active")

        RisHotstringPalette._FlushPendingSearch()
        AssertEqual(0, RisHotstringPalette._debounceTimer, "Debounce timer should be cleared after flush")
        AssertEqual(1, RisHotstringPalette._filteredItems.Length, "Should match 1 item")
        AssertEqual("sk", RisHotstringPalette._filteredItems[1].item.trigger, "Top item should be 'sk'")
    } finally {
        RisHotstringPalette.Close()
    }
}

RunRegisteredTests()
