#Requires AutoHotkey v2.0

class RisAIConfigResolver {
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
