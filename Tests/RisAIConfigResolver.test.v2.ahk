#Requires AutoHotkey v2.0

#Include TestLib.v2.ahk
#Include ..\Lib\RisAIProviderPolicy.v2.ahk
#Include ..\Lib\RisAIConfigResolver.v2.ahk

RegisterTest("RisAIConfigResolver.ResolveOpenAIOptions model fallback hierarchy", Test_ResolveOpenAIOptions_ModelFallback)
RegisterTest("RisAIConfigResolver.ResolveGoogleOptions model fallback hierarchy", Test_ResolveGoogleOptions_ModelFallback)

Test_ResolveOpenAIOptions_ModelFallback() {
    tempIni := A_Temp . "\test_openai_cfg_" . A_TickCount . ".ini"
    IniWrite("test-api-key", tempIni, "OpenAI", "APIKey")

    try {
        baseCfg := {
            ConfigFile: tempIni,
            APIKey: "test-api-key",
            BaseUrl: "https://api.openai.com/v1/responses",
            Model: "cfg-default-model",
            Temperature: "0.2",
            ReasoningEffort: "none"
        }

        ; 1. 無 override 且無 aiConfig：應使用 cfg.Model
        opts1 := RisAIConfigResolver.ResolveOpenAIOptions(baseCfg)
        AssertEqual("cfg-default-model", opts1.Model, "Should fallback to cfg.Model when no override or aiConfig")

        ; 2. aiConfig 提供 Model，無 override：應使用 aiConfig.Model
        aiCfg := { Model: "aiconfig-model", Temperature: "0.5" }
        opts2 := RisAIConfigResolver.ResolveOpenAIOptions(baseCfg, aiCfg)
        AssertEqual("aiconfig-model", opts2.Model, "Should use aiConfig.Model when present")

        ; 3. 有 modelOverride：優先權最高，覆寫 aiConfig.Model 與 cfg.Model
        opts3 := RisAIConfigResolver.ResolveOpenAIOptions(baseCfg, aiCfg, "override-model")
        AssertEqual("override-model", opts3.Model, "Should prioritize modelOverride over all else")
    } finally {
        if FileExist(tempIni) {
            FileDelete(tempIni)
        }
    }
}

Test_ResolveGoogleOptions_ModelFallback() {
    tempIni := A_Temp . "\test_google_cfg_" . A_TickCount . ".ini"
    IniWrite("test-google-key", tempIni, "GoogleAI", "APIKey")

    try {
        baseCfg := {
            ConfigFile: tempIni,
            APIKey: "test-google-key",
            Model: "google-cfg-model",
            Temperature: "0.2",
            TopP: "0.95",
            EnableGoogleSearch: false
        }

        ; 1. 無 override 且無 aiConfig：應使用 cfg.Model
        opts1 := RisAIConfigResolver.ResolveGoogleOptions(baseCfg)
        AssertEqual("google-cfg-model", opts1.Model, "Should fallback to cfg.Model")

        ; 2. aiConfig 提供 Model，無 override：應使用 aiConfig.Model
        aiCfg := { Model: "google-aiconfig-model" }
        opts2 := RisAIConfigResolver.ResolveGoogleOptions(baseCfg, aiCfg)
        AssertEqual("google-aiconfig-model", opts2.Model, "Should use aiConfig.Model")

        ; 3. 有 modelOverride：優先權最高
        opts3 := RisAIConfigResolver.ResolveGoogleOptions(baseCfg, aiCfg, "google-override-model")
        AssertEqual("google-override-model", opts3.Model, "Should prioritize modelOverride")
    } finally {
        if FileExist(tempIni) {
            FileDelete(tempIni)
        }
    }
}

RunRegisteredTests()
