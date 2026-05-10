#Requires AutoHotkey v2.0

class RisAIProviderPolicy {
    static ResolveModelList(aiConfig, providerName, fallbackModel := "") {
        models := []
        providerName := StrLower(Trim(providerName))
        providerProp := (providerName == "openai") ? "OpenAIModels" : "GoogleModels"

        if (IsObject(aiConfig) && aiConfig.HasOwnProp(providerProp)) {
            for _, modelName in aiConfig.%providerProp% {
                modelName := Trim(modelName)
                if (modelName != "") {
                    models.Push(modelName)
                }
            }
        }

        if (models.Length == 0 && IsObject(aiConfig) && aiConfig.HasOwnProp("Models")) {
            for _, modelName in aiConfig.Models {
                modelName := Trim(modelName)
                if (modelName != "") {
                    models.Push(modelName)
                }
            }
        }

        if (models.Length == 0 && IsObject(aiConfig) && aiConfig.HasOwnProp("Model")) {
            modelName := Trim(aiConfig.Model)
            if (modelName != "") {
                models.Push(modelName)
            }
        }

        if (models.Length == 0 && fallbackModel != "") {
            models.Push(fallbackModel)
        }

        return models
    }

    static ShouldRetryModelStatus(status) {
        return status == 500 || status == 503
    }

    static ParseConfigBool(value, defaultValue := false) {
        if (value == "") {
            return defaultValue
        }

        if IsNumber(value) {
            return Number(value) != 0
        }

        normalized := StrLower(Trim(value))
        if (normalized == "true" || normalized == "yes" || normalized == "on") {
            return true
        }
        if (normalized == "false" || normalized == "no" || normalized == "off") {
            return false
        }

        return defaultValue
    }

    static GoogleModelSupportsThinkingLevel(modelName) {
        unsupportedModels := Map(
            "gemini-2.5-flash-lite", false
        )

        return !unsupportedModels.Has(StrLower(Trim(modelName)))
    }
}
