#Requires AutoHotkey v2.0

class RisAIConfigResolver {
    static GetGoogleConfig(configFile := "config\private.ini") {
        return {
            ConfigFile: configFile,
            APIKey: IniRead(configFile, "GoogleAI", "APIKey", ""),
            Model: IniRead(configFile, "GoogleAI", "Model", "gemma-3-27b-it"),
            Temperature: IniRead(configFile, "GoogleAI", "Temperature", "0.2"),
            TopP: IniRead(configFile, "GoogleAI", "TopP", "0.95"),
            EnableGoogleSearch: RisAIProviderPolicy.ParseConfigBool(IniRead(configFile, "GoogleAI", "EnableGoogleSearch", "false"), false)
        }
    }

    static GetOpenAIConfig(configFile := "config\private.ini") {
        return {
            ConfigFile: configFile,
            APIKey: IniRead(configFile, "OpenAI", "APIKey", ""),
            BaseUrl: IniRead(configFile, "OpenAI", "BaseUrl", "https://api.openai.com/v1/responses"),
            Model: IniRead(configFile, "OpenAI", "Model", "gpt-5.4-nano"),
            Temperature: IniRead(configFile, "OpenAI", "Temperature", "0.2"),
            ReasoningEffort: IniRead(configFile, "OpenAI", "ReasoningEffort", "none")
        }
    }

    static ResolveGoogleOptions(cfg, aiConfig := 0, modelOverride := "") {
        modelName := (IsObject(aiConfig) && aiConfig.HasOwnProp("Model")) ? aiConfig.Model : ""
        temperature := (IsObject(aiConfig) && aiConfig.HasOwnProp("Temperature")) ? aiConfig.Temperature : ""
        thinkingLevel := (IsObject(aiConfig) && aiConfig.HasOwnProp("ThinkingLevel")) ? aiConfig.ThinkingLevel : ""
        topP := (IsObject(aiConfig) && aiConfig.HasOwnProp("TopP")) ? aiConfig.TopP : ""
        enableGoogleSearch := (IsObject(aiConfig) && aiConfig.HasOwnProp("EnableGoogleSearch")) ? aiConfig.EnableGoogleSearch : cfg.EnableGoogleSearch
        apiKey := this.ResolveAPIKey(cfg, aiConfig, "GoogleAI")
        resolvedModel := (modelOverride != "") ? modelOverride : ((modelName != "") ? modelName : cfg.Model)
        if (thinkingLevel != "" && !RisAIProviderPolicy.GoogleModelSupportsThinkingLevel(resolvedModel)) {
            thinkingLevel := ""
        }

        return {
            APIKey: apiKey.Value,
            APIKeyName: apiKey.Name,
            Model: resolvedModel,
            Temperature: (temperature != "") ? temperature : cfg.Temperature,
            ThinkingLevel: thinkingLevel,
            TopP: (topP != "") ? topP : cfg.TopP,
            EnableGoogleSearch: RisAIProviderPolicy.ParseConfigBool(enableGoogleSearch, false)
        }
    }

    static ResolveOpenAIOptions(cfg, aiConfig := 0, modelOverride := "") {
        temperature := (IsObject(aiConfig) && aiConfig.HasOwnProp("Temperature")) ? aiConfig.Temperature : cfg.Temperature
        reasoningEffort := (IsObject(aiConfig) && aiConfig.HasOwnProp("ReasoningEffort")) ? aiConfig.ReasoningEffort : cfg.ReasoningEffort
        apiKey := this.ResolveAPIKey(cfg, aiConfig, "OpenAI")

        return {
            APIKey: apiKey.Value,
            APIKeyName: apiKey.Name,
            BaseUrl: cfg.BaseUrl,
            Model: modelOverride,
            Temperature: temperature,
            ReasoningEffort: reasoningEffort
        }
    }

    static ResolveAPIKey(cfg, aiConfig, sectionName) {
        keyName := "APIKey"
        if (IsObject(aiConfig) && aiConfig.HasOwnProp("APIKeyName") && aiConfig.APIKeyName != "") {
            keyName := aiConfig.APIKeyName
        }

        apiKey := IniRead(cfg.ConfigFile, sectionName, keyName, "")
        if (apiKey != "") {
            return {
                Name: keyName,
                Value: apiKey
            }
        }

        if (keyName != "APIKey" && cfg.APIKey != "") {
            return {
                Name: "APIKey",
                Value: cfg.APIKey
            }
        }

        throw Error("請在 " . cfg.ConfigFile . " 中設定 [" . sectionName . "] " . keyName)
    }
}
