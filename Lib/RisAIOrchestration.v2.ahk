#Requires AutoHotkey v2.0

class RisAIOrchestration {
    static NormalizeResult(result) {
        result := StrReplace(result, "`r`n", "`n")
        result := StrReplace(result, "`n", "`r`n")
        return result
    }

    static FormatCompleteNotify(title, apiKeyName, modelName, detail := "") {
        text := title . "`r`nAPI Key: " . apiKeyName . "`r`nModel: " . modelName
        if (detail != "") {
            text .= "`r`n" . detail
        }
        return text
    }

    static NormalizePolishResult(result, trailingNewlines := "") {
        result := this.NormalizeResult(result)
        if (trailingNewlines != "") {
            result .= trailingNewlines
        }
        return result
    }

    static CreateRequest(promptText, aiConfig, extraFields := 0) {
        request := {
            Prompt: promptText,
            Config: aiConfig
        }

        if IsObject(extraFields) {
            for key, value in extraFields.OwnProps() {
                request.%key% := value
            }
        }

        return request
    }

    static FormatPolishComparisonDebugInfo(response) {
        return {
            APIKeyName: response.APIKeyName,
            Model: response.Model,
            ApiTime: response.ApiTime . " ms"
        }
    }

    static BuildRefineProviderSuccessResult(displayName, response, trailingNewlines := "") {
        response.Result := this.NormalizePolishResult(response.Result, trailingNewlines)
        return {
            Success: true,
            DisplayName: displayName,
            Text: response.Result,
            DebugInfo: this.FormatPolishComparisonDebugInfo(response)
        }
    }

    static BuildRefineProviderFailureResult(displayName, message) {
        return {
            Success: false,
            DisplayName: displayName,
            Text: "[" . displayName . " 失敗]`r`n" . message,
            DebugInfo: {
                APIKeyName: "-",
                Model: "-",
                ApiTime: "failed"
            }
        }
    }
}
