#Requires AutoHotkey v2.0

#Include TestLib.v2.ahk
#Include ..\Lib\RisAIText.v2.ahk
#Include ..\Lib\RisAIPayload.v2.ahk

RegisterTest("RisAIPayload.BuildOpenAIPayload includes temperature when reasoning is none", Test_RisAIPayload_BuildOpenAIPayload_WithNoneReasoning)
RegisterTest("RisAIPayload.BuildOpenAIPayload omits temperature when reasoning is low", Test_RisAIPayload_BuildOpenAIPayload_WithLowReasoning)
RegisterTest("RisAIPayload.BuildGoogleUrl omits API key from url query string", Test_RisAIPayload_BuildGoogleUrl_OmitsApiKey)

Test_RisAIPayload_BuildOpenAIPayload_WithNoneReasoning() {
    options := {
        Model: "gpt-6-luna",
        ReasoningEffort: "none",
        Temperature: 0.2
    }
    payload := RisAIPayload.BuildOpenAIPayload("Hello", options)
    AssertTrue(InStr(payload, '"temperature":0.2') > 0, "Payload should include temperature when reasoning is none")
    AssertTrue(InStr(payload, '"reasoning":{"effort":"none"}') > 0, "Payload should include reasoning effort none")
}

Test_RisAIPayload_BuildOpenAIPayload_WithLowReasoning() {
    options := {
        Model: "gpt-6-luna",
        ReasoningEffort: "low",
        Temperature: 0.2
    }
    payload := RisAIPayload.BuildOpenAIPayload("Hello", options)
    AssertTrue(InStr(payload, '"temperature"') == 0, "Payload must omit temperature when reasoning effort is low")
    AssertTrue(InStr(payload, '"reasoning":{"effort":"low"}') > 0, "Payload should include reasoning effort low")
}

Test_RisAIPayload_BuildGoogleUrl_OmitsApiKey() {
    options := {
        Model: "gemini-2.5-flash",
        APIKey: "secret-key-12345"
    }
    url := RisAIPayload.BuildGoogleUrl(options)
    AssertEqual("https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent", url)
    AssertTrue(InStr(url, "secret-key-12345") == 0, "Url must NOT contain API key")
    AssertTrue(InStr(url, "?key=") == 0, "Url must NOT contain ?key= query string")
}

RunRegisteredTests()
