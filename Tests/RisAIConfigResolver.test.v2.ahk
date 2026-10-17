#Requires AutoHotkey v2.0

#Include TestLib.v2.ahk
#Include ..\Lib\RisAIProviderPolicy.v2.ahk
#Include ..\Lib\RisAIConfigResolver.v2.ahk

RegisterTest("RisAIConfigResolver.GetGoogleConfig reads defaults and custom values", Test_GetGoogleConfig)
RegisterTest("RisAIConfigResolver.GetOpenAIConfig reads defaults and custom values", Test_GetOpenAIConfig)
RegisterTest("RisAIConfigResolver.ResolveOpenAIOptions model fallback hierarchy", Test_ResolveOpenAIOptions_ModelFallback)
RegisterTest("RisAIConfigResolver.ResolveGoogleOptions model fallback hierarchy", Test_ResolveGoogleOptions_ModelFallback)
RegisterTest("RisAIConfigResolver.ResolveGoogleOptions parameters and thinking level filtering", Test_ResolveGoogleOptions_Parameters)
RegisterTest("RisAIConfigResolver.ResolveOpenAIOptions parameters resolution", Test_ResolveOpenAIOptions_Parameters)
RegisterTest("RisAIConfigResolver.ResolveAPIKey custom name, fallback, and missing error", Test_ResolveAPIKey)

Test_GetGoogleConfig() {
    tempIni := A_Temp . "\test_get_google_cfg_" . A_TickCount . ".ini"

    try {
        ; 1. 空檔案或無區段時的預設值
        cfgDefault := RisAIConfigResolver.GetGoogleConfig(tempIni)
        AssertEqual("gemma-3-27b-it", cfgDefault.Model, "Default model should be gemma-3-27b-it")
        AssertEqual("0.2", cfgDefault.Temperature, "Default temperature should be 0.2")
        AssertEqual("0.95", cfgDefault.TopP, "Default TopP should be 0.95")
        AssertEqual(false, cfgDefault.EnableGoogleSearch, "Default EnableGoogleSearch should be false")

        ; 2. 寫入自訂值
        IniWrite("custom-key", tempIni, "GoogleAI", "APIKey")
        IniWrite("custom-google-model", tempIni, "GoogleAI", "Model")
        IniWrite("0.7", tempIni, "GoogleAI", "Temperature")
        IniWrite("0.8", tempIni, "GoogleAI", "TopP")
        IniWrite("true", tempIni, "GoogleAI", "EnableGoogleSearch")

        cfgCustom := RisAIConfigResolver.GetGoogleConfig(tempIni)
        AssertEqual("custom-key", cfgCustom.APIKey)
        AssertEqual("custom-google-model", cfgCustom.Model)
        AssertEqual("0.7", cfgCustom.Temperature)
        AssertEqual("0.8", cfgCustom.TopP)
        AssertEqual(true, cfgCustom.EnableGoogleSearch)
    } finally {
        if FileExist(tempIni) {
            FileDelete(tempIni)
        }
    }
}

Test_GetOpenAIConfig() {
    tempIni := A_Temp . "\test_get_openai_cfg_" . A_TickCount . ".ini"

    try {
        ; 1. 空檔案或無區段時的預設值
        cfgDefault := RisAIConfigResolver.GetOpenAIConfig(tempIni)
        AssertEqual("gpt-6-luna", cfgDefault.Model, "Default model should be gpt-6-luna")
        AssertEqual("https://api.openai.com/v1/responses", cfgDefault.BaseUrl)
        AssertEqual("0.2", cfgDefault.Temperature)
        AssertEqual("none", cfgDefault.ReasoningEffort)

        ; 2. 寫入自訂值
        IniWrite("custom-openai-key", tempIni, "OpenAI", "APIKey")
        IniWrite("custom-openai-model", tempIni, "OpenAI", "Model")
        IniWrite("https://custom.endpoint/v1", tempIni, "OpenAI", "BaseUrl")
        IniWrite("0.5", tempIni, "OpenAI", "Temperature")
        IniWrite("low", tempIni, "OpenAI", "ReasoningEffort")

        cfgCustom := RisAIConfigResolver.GetOpenAIConfig(tempIni)
        AssertEqual("custom-openai-key", cfgCustom.APIKey)
        AssertEqual("custom-openai-model", cfgCustom.Model)
        AssertEqual("https://custom.endpoint/v1", cfgCustom.BaseUrl)
        AssertEqual("0.5", cfgCustom.Temperature)
        AssertEqual("low", cfgCustom.ReasoningEffort)
    } finally {
        if FileExist(tempIni) {
            FileDelete(tempIni)
        }
    }
}

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

Test_ResolveGoogleOptions_Parameters() {
    tempIni := A_Temp . "\test_google_params_" . A_TickCount . ".ini"
    IniWrite("test-google-key", tempIni, "GoogleAI", "APIKey")

    try {
        baseCfg := {
            ConfigFile: tempIni,
            APIKey: "test-google-key",
            Model: "gemma-4-31b-it",
            Temperature: "0.2",
            TopP: "0.95",
            EnableGoogleSearch: false
        }

        ; 1. 不支援思考等級之模型 (gemini-2.5-flash-lite) 應清空 ThinkingLevel
        aiCfgLite := {
            Model: "gemini-2.5-flash-lite",
            ThinkingLevel: "MINIMAL",
            Temperature: "0.5",
            TopP: "0.8",
            EnableGoogleSearch: true
        }
        optsLite := RisAIConfigResolver.ResolveGoogleOptions(baseCfg, aiCfgLite)
        AssertEqual("", optsLite.ThinkingLevel, "gemini-2.5-flash-lite should have ThinkingLevel stripped")
        AssertEqual("0.5", optsLite.Temperature, "Temperature should be overridden by aiConfig")
        AssertEqual("0.8", optsLite.TopP, "TopP should be overridden by aiConfig")
        AssertEqual(true, optsLite.EnableGoogleSearch, "EnableGoogleSearch should be overridden by aiConfig")

        ; 2. 支援思考等級之模型應保留 ThinkingLevel
        aiCfgSupported := {
            Model: "gemini-3.1-flash-lite",
            ThinkingLevel: "MINIMAL"
        }
        optsSupported := RisAIConfigResolver.ResolveGoogleOptions(baseCfg, aiCfgSupported)
        AssertEqual("MINIMAL", optsSupported.ThinkingLevel, "Supported model should retain ThinkingLevel")
        AssertEqual("0.2", optsSupported.Temperature, "Unspecified temperature should fallback to cfg")
    } finally {
        if FileExist(tempIni) {
            FileDelete(tempIni)
        }
    }
}

Test_ResolveOpenAIOptions_Parameters() {
    tempIni := A_Temp . "\test_openai_params_" . A_TickCount . ".ini"
    IniWrite("test-openai-key", tempIni, "OpenAI", "APIKey")

    try {
        baseCfg := {
            ConfigFile: tempIni,
            APIKey: "test-openai-key",
            BaseUrl: "https://api.openai.com/v1/responses",
            Model: "gpt-6-luna",
            Temperature: "0.2",
            ReasoningEffort: "none"
        }

        aiCfg := {
            Temperature: "0.4",
            ReasoningEffort: "medium"
        }
        opts := RisAIConfigResolver.ResolveOpenAIOptions(baseCfg, aiCfg)
        AssertEqual("https://api.openai.com/v1/responses", opts.BaseUrl)
        AssertEqual("0.4", opts.Temperature, "Temperature should be overridden by aiConfig")
        AssertEqual("medium", opts.ReasoningEffort, "ReasoningEffort should be overridden by aiConfig")
    } finally {
        if FileExist(tempIni) {
            FileDelete(tempIni)
        }
    }
}

Test_ResolveAPIKey() {
    tempIni := A_Temp . "\test_apikey_res_" . A_TickCount . ".ini"

    try {
        ; 1. 寫入特定 APIKeyName
        IniWrite("special-indication-key", tempIni, "GoogleAI", "IndicationAPIKey")
        IniWrite("default-google-key", tempIni, "GoogleAI", "APIKey")

        baseCfg := RisAIConfigResolver.GetGoogleConfig(tempIni)

        ; 當 aiConfig 指定自訂 APIKeyName 且存在於 INI
        aiCfgSpecific := { APIKeyName: "IndicationAPIKey" }
        res1 := RisAIConfigResolver.ResolveGoogleOptions(baseCfg, aiCfgSpecific)
        AssertEqual("IndicationAPIKey", res1.APIKeyName)
        AssertEqual("special-indication-key", res1.APIKey)

        ; 當 aiConfig 指定之自訂 key 不存在，但 cfg.APIKey 存在，應 fallback 至預設 APIKey
        aiCfgNonExistent := { APIKeyName: "NonExistentKey" }
        res2 := RisAIConfigResolver.ResolveGoogleOptions(baseCfg, aiCfgNonExistent)
        AssertEqual("APIKey", res2.APIKeyName)
        AssertEqual("default-google-key", res2.APIKey)

        ; 2. 若兩者皆不存在，應拋出 Error
        baseCfgEmpty := {
            ConfigFile: tempIni,
            APIKey: "",
            Model: "gemma-3-27b-it",
            Temperature: "0.2",
            TopP: "0.95",
            EnableGoogleSearch: false
        }
        AssertThrows(Error, () => RisAIConfigResolver.ResolveGoogleOptions(baseCfgEmpty, { APIKeyName: "MissingKey" }))
    } finally {
        if FileExist(tempIni) {
            FileDelete(tempIni)
        }
    }
}

RunRegisteredTests()
