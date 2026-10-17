#Requires AutoHotkey v2.0

#Include TestLib.v2.ahk
#Include ..\Lib\ImeGuard.v2.ahk

RegisterTest("WithImeGuard executes action and returns result", Test_WithImeGuard_ReturnsResult)
RegisterTest("WithImeGuard propagates exceptions while executing finally block", Test_WithImeGuard_PropagatesException)

Test_WithImeGuard_ReturnsResult() {
    executed := false
    result := WithImeGuard(() => (executed := true, 42))
    AssertTrue(executed, "Action should be executed")
    AssertEqual(42, result, "Return value should be passed through")
}

ThrowTestError() {
    throw ValueError("Test error")
}

Test_WithImeGuard_PropagatesException() {
    RunProtected() {
        WithImeGuard(ThrowTestError)
    }
    err := AssertThrows(ValueError, RunProtected)
    AssertEqual("Test error", err.Message)
}

RunRegisteredTests()
