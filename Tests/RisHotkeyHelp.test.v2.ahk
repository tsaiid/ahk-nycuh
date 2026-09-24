#Requires AutoHotkey v2.0

#Include TestLib.v2.ahk
#Include ..\Lib\RisNotify.v2.ahk
#Include ..\Lib\RisHotkeyHelp.v2.ahk

RegisterTest("RisHotkeyHelp._GetHotkeys contains valid entries and includes Ctrl+Shift+F", Test_RisHotkeyHelp_GetHotkeys)
RegisterTest("RisHotkeyHelp._BuildHtml generates valid HTML without throwing", Test_RisHotkeyHelp_BuildHtml)

Test_RisHotkeyHelp_GetHotkeys() {
    hotkeys := RisHotkeyHelp._GetHotkeys()
    AssertTrue(hotkeys.Length > 0, "Should have hotkey entries")

    foundPalette := false
    for _, item in hotkeys {
        AssertTrue(item.HasOwnProp("Context") && item.Context != "", "Entry should have non-empty Context")
        AssertTrue(item.HasOwnProp("Keys") && item.Keys != "", "Entry should have non-empty Keys")
        AssertTrue(item.HasOwnProp("Action") && item.Action != "", "Entry should have non-empty Action")

        if (item.Keys = "Ctrl+Shift+F") {
            foundPalette := true
            AssertEqual("任何 RIS 報告", item.Context, "Ctrl+Shift+F should belong to 任何 RIS 報告 context")
        }

        if (item.Keys = "Ctrl+Shift+D") {
            foundSplit := true
            AssertEqual("RIS 輸入框", item.Context, "Ctrl+Shift+D should belong to RIS 輸入框 context")
        }
    }

    AssertTrue(foundPalette, "Ctrl+Shift+F should be registered in RisHotkeyHelp")
    AssertTrue(foundSplit, "Ctrl+Shift+D should be registered in RisHotkeyHelp")
}

Test_RisHotkeyHelp_BuildHtml() {
    colors := RisHotkeyHelp._GetColors("light")
    html := RisHotkeyHelp._BuildHtml(colors)
    AssertTrue(InStr(html, "Ctrl+Shift+F") > 0, "Generated HTML should contain Ctrl+Shift+F")
    AssertTrue(InStr(html, "Hotstring 搜尋命令列") > 0, "Generated HTML should contain Hotstring 搜尋命令列")
    AssertTrue(InStr(html, "Ctrl+Shift+D") > 0, "Generated HTML should contain Ctrl+Shift+D")
    AssertTrue(InStr(html, "將選取文字拆分為一句一行") > 0, "Generated HTML should contain 將選取文字拆分為一句一行")
}

RunRegisteredTests()
