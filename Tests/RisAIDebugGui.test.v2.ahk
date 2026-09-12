#Requires AutoHotkey v2.0

#Include TestLib.v2.ahk
#Include ..\Lib\RisAIDebugGui.v2.ahk

RegisterTest("RisAIDebugGui.ApplyPolishProviderChoice invokes choice callbacks", Test_RisAIDebugGui_ApplyPolishProviderChoice)
RegisterTest("RisAIDebugGui.ApplyPolishProviderChoice handles uninitialized callbacks safely", Test_RisAIDebugGui_ApplyPolishProviderChoice_Safe)

Test_RisAIDebugGui_ApplyPolishProviderChoice() {
    choice1Called := false
    choice2Called := false

    RisAIDebugGui.applyFirstChoiceFunc := () => (choice1Called := true)
    RisAIDebugGui.applySecondChoiceFunc := () => (choice2Called := true)

    try {
        RisAIDebugGui.ApplyPolishProviderChoice(1)
        AssertTrue(choice1Called, "Choice 1 callback should be executed")
        AssertFalse(choice2Called, "Choice 2 callback should not be executed yet")

        RisAIDebugGui.ApplyPolishProviderChoice(2)
        AssertTrue(choice2Called, "Choice 2 callback should be executed")
    } finally {
        RisAIDebugGui.applyFirstChoiceFunc := 0
        RisAIDebugGui.applySecondChoiceFunc := 0
    }
}

Test_RisAIDebugGui_ApplyPolishProviderChoice_Safe() {
    RisAIDebugGui.applyFirstChoiceFunc := 0
    RisAIDebugGui.applySecondChoiceFunc := 0

    ; Should not throw when callbacks are not registered or index is invalid
    RisAIDebugGui.ApplyPolishProviderChoice(1)
    RisAIDebugGui.ApplyPolishProviderChoice(2)
    RisAIDebugGui.ApplyPolishProviderChoice(3)
    AssertTrue(true, "Handled safely without error")
}

RunRegisteredTests()
