#Requires AutoHotkey v2.0

#Include TestLib.v2.ahk
#Include ..\Lib\Paste.v2.ahk
#Include ..\Lib\RisController.v2.ahk

RegisterTest("Paste: 空字串直接返回且不拋錯", Test_Paste_EmptyText)
RegisterTest("Paste: 使用 HWND 直通 EditPaste", Test_Paste_DirectEditPaste_Hwnd)
RegisterTest("Paste: 使用 GuiControl 物件直通 EditPaste", Test_Paste_DirectEditPaste_GuiControl)
RegisterTest("Paste: 使用 UIA Element Mock 直通 EditPaste", Test_Paste_DirectEditPaste_UiaMock)
RegisterTest("Paste: 換行符號規格化為 CRLF", Test_Paste_ConvertCRLF)
RegisterTest("Paste: 非 Win32 UIA 控制項略過直通並呼叫 SetFocus", Test_Paste_NonWin32UiaMock_FallbackFocus)
RegisterTest("RisController: PasteTo 委託至 Paste() 直通輸出", Test_RisController_PasteTo_Delegation)

Test_Paste_EmptyText() {
    Paste("")
    Paste("", false, 0)
    AssertTrue(true, "空字串應直接返回")
}

Test_Paste_DirectEditPaste_Hwnd() {
    testGui := Gui()
    try {
        editCtrl := testGui.Add("Edit", "w200 h100", "")
        testGui.Show("Hide")

        Paste("Direct edit test via HWND", true, editCtrl.Hwnd)
        AssertEqual("Direct edit test via HWND", editCtrl.Value, "Edit 控制項應接收到文字")
    } finally {
        testGui.Destroy()
    }
}

Test_Paste_DirectEditPaste_GuiControl() {
    testGui := Gui()
    try {
        editCtrl := testGui.Add("Edit", "w200 h100", "")
        testGui.Show("Hide")

        Paste("Gui control object test", true, editCtrl)
        AssertEqual("Gui control object test", editCtrl.Value, "傳入 GuiControl 物件應直通 EditPaste")
    } finally {
        testGui.Destroy()
    }
}

Test_Paste_DirectEditPaste_UiaMock() {
    testGui := Gui()
    try {
        editCtrl := testGui.Add("Edit", "w200 h100", "")
        testGui.Show("Hide")

        mockUia := {NativeWindowHandle: editCtrl.Hwnd, FrameworkId: "Win32"}
        Paste("UIA mock test", true, mockUia)
        AssertEqual("UIA mock test", editCtrl.Value, "傳入 UIA mock 物件應能正確提取 HWND 直通 EditPaste")
    } finally {
        testGui.Destroy()
    }
}

Test_Paste_ConvertCRLF() {
    testGui := Gui()
    try {
        editCtrl := testGui.Add("Edit", "w200 h100", "")
        testGui.Show("Hide")

        sampleLF := "Line1`nLine2`nLine3"
        Paste(sampleLF, true, editCtrl.Hwnd)
        rawText := ControlGetText(editCtrl.Hwnd)
        AssertEqual("Line1`r`nLine2`r`nLine3", rawText, "LF 換行應被規格化為 CRLF")
    } finally {
        testGui.Destroy()
    }
}

Test_Paste_NonWin32UiaMock_FallbackFocus() {
    focused := false
    _ThrowFocusError(*) {
        throw Error("Simulated focus failure")
    }

    failedFocusMock := {
        FrameworkId: "WPF",
        NativeWindowHandle: 12345,
        SetFocus: _ThrowFocusError
    }

    ; 當非 Win32 控制項且 SetFocus 失敗時，應捕捉錯誤呼叫通知並安全退出，不中斷主執行緒
    Paste("Testing non-win32 control fallback", true, failedFocusMock)
    AssertTrue(true, "SetFocus 失敗時應安全攔截並退出")

    successFocusMock := {
        FrameworkId: "WPF",
        NativeWindowHandle: 0,
        SetFocus: (*) => (focused := true)
    }
    ; 使用 minLength > 字串長度（觸發 SendText 支線，避免在測試環境模擬剪貼簿）
    Paste("A", true, successFocusMock, 100)
    AssertTrue(focused, "非 Win32 控制項應被觸發 SetFocus()")
}

Test_RisController_PasteTo_Delegation() {
    testGui := Gui()
    try {
        editCtrl := testGui.Add("Edit", "w200 h100", "")
        testGui.Show("Hide")

        mockEl := {NativeWindowHandle: editCtrl.Hwnd, FrameworkId: "Win32"}
        RisController.PasteTo(mockEl, "RisController PasteTo delegated")
        AssertEqual("RisController PasteTo delegated", editCtrl.Value, "RisController.PasteTo 應正確委託 Paste() 輸出文字")
    } finally {
        testGui.Destroy()
    }
}

RunRegisteredTests()
