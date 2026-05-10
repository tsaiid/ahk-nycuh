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
