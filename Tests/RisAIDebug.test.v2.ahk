#Requires AutoHotkey v2.0

#Include TestLib.v2.ahk
#Include ..\Lib\RisAIText.v2.ahk
#Include ..\Lib\RisAIPayload.v2.ahk
#Include ..\Lib\RisAIRequestBuilder.v2.ahk
#Include ..\Lib\RisAIDebug.v2.ahk

RegisterTest("RisAIDebug.MaskApiKey masks short and long keys properly", Test_MaskApiKey)
RegisterTest("RisAIDebug.SanitizeUrl strips sensitive query params", Test_SanitizeUrl)
RegisterTest("RisAIDebug.BuildGoogleCurlCommand uses header and avoids leaking key", Test_BuildGoogleCurlCommand)
RegisterTest("RisAIRequestBuilder.BuildGoogleRequest preserves APIKey", Test_BuildGoogleRequest_PreservesApiKey)

Test_MaskApiKey() {
    AssertEqual("", RisAIDebug.MaskApiKey(""))
    AssertEqual("********", RisAIDebug.MaskApiKey("12345678"))
    AssertEqual("********", RisAIDebug.MaskApiKey("short"))
    AssertEqual("AIza...abcd", RisAIDebug.MaskApiKey("AIzaSyD-1234567890abcd"))
}

Test_SanitizeUrl() {
    cleanUrl := "https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent"
    AssertEqual(cleanUrl, RisAIDebug.SanitizeUrl(cleanUrl))

    dirtyUrl1 := cleanUrl . "?key=AIzaSySecretKey"
    AssertEqual(cleanUrl . "?key=[REDACTED]", RisAIDebug.SanitizeUrl(dirtyUrl1))

    dirtyUrl2 := cleanUrl . "?foo=bar&key=AIzaSySecretKey&baz=qux"
    AssertEqual(cleanUrl . "?foo=bar&key=[REDACTED]&baz=qux", RisAIDebug.SanitizeUrl(dirtyUrl2))
}

Test_BuildGoogleCurlCommand() {
    rawUrl := "https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=RawSecret1234"
    payload := '{"contents":[]}'
    curlCmd := RisAIDebug.BuildGoogleCurlCommand(rawUrl, payload, "RawSecret1234")

    ; 驗證金鑰沒有明文出現
    AssertTrue(InStr(curlCmd, "RawSecret1234") == 0, "curl command must NOT contain plaintext key")
    ; 驗證 URL 已被消毒
    AssertTrue(InStr(curlCmd, "[REDACTED]") > 0, "curl uri must have redacted query key")
    ; 驗證具備 x-goog-api-key header 與變數
    AssertTrue(InStr(curlCmd, 'x-goog-api-key: $apiKey') > 0, "curl command must include x-goog-api-key header")
    AssertTrue(InStr(curlCmd, '$apiKey = if ($env:GOOGLE_API_KEY)') > 0, "curl command should define $apiKey from environment or fallback")
}

Test_BuildGoogleRequest_PreservesApiKey() {
    options := {
        Model: "gemini-2.5-flash",
        APIKey: "my-test-api-key-999",
        APIKeyName: "CustomKey",
        Temperature: 0.2,
        ThinkingLevel: "",
        TopP: 0.95,
        EnableGoogleSearch: false
    }
    request := RisAIRequestBuilder.BuildGoogleRequest("Test prompt", options, 5)

    AssertEqual("my-test-api-key-999", request.APIKey)
    AssertEqual("CustomKey", request.APIKeyName)
    AssertEqual("gemini-2.5-flash", request.Model)
    AssertTrue(InStr(request.Url, "my-test-api-key-999") == 0, "Url should not have key")
}

RunRegisteredTests()
