#Requires AutoHotkey v2.0

#Include TestLib.v2.ahk
#Include ..\Lib\RisAIProviderPolicy.v2.ahk
#Include ..\Lib\RisAIText.v2.ahk
#Include ..\Lib\RisAIOrchestration.v2.ahk

RegisterTest("RisAIOrchestration.NormalizeImpressionResult removes trailing whitespace", Test_RisAIOrchestration_NormalizeImpressionResult)

Test_RisAIOrchestration_NormalizeImpressionResult() {
    inputText := "  1. First finding.  `n2. Second finding.`t`n  "
    expectedText := "1. First finding.`r`n2. Second finding."

    AssertEqual(expectedText, RisAIOrchestration.NormalizeImpressionResult(inputText))
}

RunRegisteredTests()
