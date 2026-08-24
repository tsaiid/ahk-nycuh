#Requires AutoHotkey v2.0

#Include TestLib.v2.ahk
#Include ..\Lib\RisDialog.v2.ahk

RegisterTest("RisDialog.Create sets default theme attributes", Test_RisDialog_Create_DefaultTheme)
RegisterTest("RisDialog.ApplyTheme supports custom options", Test_RisDialog_ApplyTheme_CustomOptions)
RegisterTest("RisDialog.GetWindowsBuildNumber returns positive integer", Test_RisDialog_GetWindowsBuildNumber)
RegisterTest("RisDialog.ApplyWindowStyle executes safely without throwing", Test_RisDialog_ApplyWindowStyle)
RegisterTest("RisDialog.CloseAndRestoreFocus destroys GUI", Test_RisDialog_CloseAndRestoreFocus)

Test_RisDialog_Create_DefaultTheme() {
    testGui := RisDialog.Create("Test Dialog")
    try {
        AssertEqual("F4F5F7", testGui.BackColor, "Default BackColor")
        AssertEqual(16, testGui.MarginX, "Default MarginX")
        AssertEqual(14, testGui.MarginY, "Default MarginY")
    } finally {
        testGui.Destroy()
    }
}

Test_RisDialog_ApplyTheme_CustomOptions() {
    testGui := Gui()
    try {
        RisDialog.ApplyTheme(testGui, {
            BackColor: "FFFFFF",
            MarginX: 20,
            MarginY: 10,
            FontSize: "s12",
            FontName: "Arial"
        })
        AssertEqual("FFFFFF", testGui.BackColor, "Custom BackColor")
        AssertEqual(20, testGui.MarginX, "Custom MarginX")
        AssertEqual(10, testGui.MarginY, "Custom MarginY")
    } finally {
        testGui.Destroy()
    }
}

Test_RisDialog_GetWindowsBuildNumber() {
    buildNum := RisDialog.GetWindowsBuildNumber()
    AssertTrue(buildNum > 0, "Build number should be a positive integer")
}

Test_RisDialog_ApplyWindowStyle() {
    testGui := RisDialog.Create("Style Test")
    try {
        testGui.Show("Hide")
        RisDialog.ApplyWindowStyle(testGui.Hwnd)
        AssertTrue(true, "ApplyWindowStyle should execute safely")
    } finally {
        testGui.Destroy()
    }
}

Test_RisDialog_CloseAndRestoreFocus() {
    DetectHiddenWindows(true)
    testGui := RisDialog.Create("Close Test")
    testGui.Show("Hide")
    hwnd := testGui.Hwnd
    AssertTrue(WinExist("ahk_id " . hwnd) > 0, "Window should exist before close")

    RisDialog.CloseAndRestoreFocus(testGui)
    AssertFalse(WinExist("ahk_id " . hwnd), "Window should be destroyed after close")
}

RunRegisteredTests()
