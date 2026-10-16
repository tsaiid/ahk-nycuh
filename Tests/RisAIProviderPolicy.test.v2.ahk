#Requires AutoHotkey v2.0

#Include TestLib.v2.ahk
#Include ..\Lib\RisAIProviderPolicy.v2.ahk

RegisterTest("RisAIProviderPolicy.ResolveProvider normalizes and fallbacks", Test_ResolveProvider)
RegisterTest("RisAIProviderPolicy.ResolveModelList resolves provider-specific, models, model, and fallback", Test_ResolveModelList)
RegisterTest("RisAIProviderPolicy.ShouldRetryModelStatus recognizes 429, 500, 502, 503, 504", Test_ShouldRetryModelStatus)
RegisterTest("RisAIProviderPolicy.ParseConfigBool parses booleans, numbers, and strings", Test_ParseConfigBool)
RegisterTest("RisAIProviderPolicy.GoogleModelSupportsThinkingLevel checks model compatibility", Test_GoogleModelSupportsThinkingLevel)

Test_ResolveProvider() {
    AssertEqual("openai", RisAIProviderPolicy.ResolveProvider({ Provider: "OpenAI" }), "Should normalize provider to lowercase")
    AssertEqual("google", RisAIProviderPolicy.ResolveProvider({ Provider: "  Google  " }), "Should trim whitespace")
    AssertEqual("google", RisAIProviderPolicy.ResolveProvider({}), "Should default to google")
    AssertEqual("google", RisAIProviderPolicy.ResolveProvider(0), "Should default to google when aiConfig is 0")
    AssertEqual("custom", RisAIProviderPolicy.ResolveProvider({}, "custom"), "Should use custom defaultProvider when specified")
}

Test_ResolveModelList() {
    ; 1. 依 Provider 指定專屬清單
    cfg1 := {
        OpenAIModels: ["gpt-6-luna", "", "  ", "gpt-5"],
        GoogleModels: ["gemma-4-31b-it", "gemini-3.1-flash-lite"]
    }
    openAIList := RisAIProviderPolicy.ResolveModelList(cfg1, "openai")
    AssertEqual(2, openAIList.Length, "Should resolve 2 OpenAI models, filtering empty ones")
    AssertEqual("gpt-6-luna", openAIList[1])
    AssertEqual("gpt-5", openAIList[2])

    googleList := RisAIProviderPolicy.ResolveModelList(cfg1, "google")
    AssertEqual(2, googleList.Length, "Should resolve 2 Google models")
    AssertEqual("gemma-4-31b-it", googleList[1])
    AssertEqual("gemini-3.1-flash-lite", googleList[2])

    ; 2. Fallback 至 Models 清單
    cfg2 := { Models: ["general-model-1", "general-model-2"] }
    listFromModels := RisAIProviderPolicy.ResolveModelList(cfg2, "google")
    AssertEqual(2, listFromModels.Length)
    AssertEqual("general-model-1", listFromModels[1])

    ; 3. Fallback 至單一 Model 字串
    cfg3 := { Model: "single-model-only" }
    listFromSingle := RisAIProviderPolicy.ResolveModelList(cfg3, "openai")
    AssertEqual(1, listFromSingle.Length)
    AssertEqual("single-model-only", listFromSingle[1])

    ; 4. Fallback 至 fallbackModel 參數
    listFromFallback := RisAIProviderPolicy.ResolveModelList({}, "openai", "fallback-model")
    AssertEqual(1, listFromFallback.Length)
    AssertEqual("fallback-model", listFromFallback[1])

    ; 5. 全部皆無時回傳空陣列
    emptyList := RisAIProviderPolicy.ResolveModelList({}, "openai")
    AssertEqual(0, emptyList.Length)
}

Test_ShouldRetryModelStatus() {
    ; 重試/降級狀態碼 (Rate limit / Server errors / Gateway timeouts)
    retryCodes := [429, 500, 502, 503, 504]
    for code in retryCodes {
        AssertTrue(RisAIProviderPolicy.ShouldRetryModelStatus(code), "Status " . code . " should be retryable")
    }

    ; 非重試狀態碼 (正常或客戶端錯誤)
    nonRetryCodes := [200, 201, 400, 401, 403, 404, 408, 413, 501]
    for code in nonRetryCodes {
        AssertFalse(RisAIProviderPolicy.ShouldRetryModelStatus(code), "Status " . code . " should NOT be retryable")
    }
}

Test_ParseConfigBool() {
    ; 空字串回傳預設值
    AssertEqual(false, RisAIProviderPolicy.ParseConfigBool("", false))
    AssertEqual(true, RisAIProviderPolicy.ParseConfigBool("", true))

    ; 數值型態
    AssertEqual(true, RisAIProviderPolicy.ParseConfigBool(1))
    AssertEqual(true, RisAIProviderPolicy.ParseConfigBool("1"))
    AssertEqual(false, RisAIProviderPolicy.ParseConfigBool(0))
    AssertEqual(false, RisAIProviderPolicy.ParseConfigBool("0"))
    AssertEqual(true, RisAIProviderPolicy.ParseConfigBool(2))

    ; 字串真值 (true, yes, on，大小寫不拘)
    AssertEqual(true, RisAIProviderPolicy.ParseConfigBool("true"))
    AssertEqual(true, RisAIProviderPolicy.ParseConfigBool("True"))
    AssertEqual(true, RisAIProviderPolicy.ParseConfigBool("YES"))
    AssertEqual(true, RisAIProviderPolicy.ParseConfigBool("  on  "))

    ; 字串假值 (false, no, off，大小寫不拘)
    AssertEqual(false, RisAIProviderPolicy.ParseConfigBool("false"))
    AssertEqual(false, RisAIProviderPolicy.ParseConfigBool("False"))
    AssertEqual(false, RisAIProviderPolicy.ParseConfigBool("NO"))
    AssertEqual(false, RisAIProviderPolicy.ParseConfigBool("  off  "))

    ; 未知字串回傳預設值
    AssertEqual(false, RisAIProviderPolicy.ParseConfigBool("unknown", false))
    AssertEqual(true, RisAIProviderPolicy.ParseConfigBool("unknown", true))
}

Test_GoogleModelSupportsThinkingLevel() {
    AssertFalse(RisAIProviderPolicy.GoogleModelSupportsThinkingLevel("gemini-2.5-flash-lite"), "gemini-2.5-flash-lite should not support thinking level")
    AssertFalse(RisAIProviderPolicy.GoogleModelSupportsThinkingLevel("GEMINI-2.5-FLASH-LITE"), "Should be case-insensitive")
    AssertFalse(RisAIProviderPolicy.GoogleModelSupportsThinkingLevel("  gemini-2.5-flash-lite  "), "Should trim whitespace")

    AssertTrue(RisAIProviderPolicy.GoogleModelSupportsThinkingLevel("gemma-4-31b-it"), "gemma-4-31b-it should support thinking level")
    AssertTrue(RisAIProviderPolicy.GoogleModelSupportsThinkingLevel("gemini-3.1-flash-lite"), "gemini-3.1-flash-lite should support thinking level")
    AssertTrue(RisAIProviderPolicy.GoogleModelSupportsThinkingLevel("gpt-6-luna"), "Other models should return true by default")
}

RunRegisteredTests()
