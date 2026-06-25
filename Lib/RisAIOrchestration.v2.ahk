#Requires AutoHotkey v2.0

class RisAIOrchestration {
    static NormalizeResult(result) {
        result := StrReplace(result, "`r`n", "`n")
        result := StrReplace(result, "`n", "`r`n")
        return result
    }

    static NormalizeImpressionResult(result) {
        result := this.NormalizeResult(result)
        result := RegExReplace(result, "m)[ `t]+(?=`r?$)", "")
        return Trim(result, " `t`r`n")
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

    static BuildRequestResult(response, apiTime) {
        return {
            Result: response.Result,
            ApiTime: apiTime,
            APIKeyName: response.APIKeyName,
            Model: response.Model,
            Provider: response.Provider
        }
    }

    static CloneConfigWithProvider(aiConfig, providerName) {
        cloned := {}
        for key, value in aiConfig.OwnProps() {
            cloned.%key% := value
        }
        cloned.Provider := providerName
        return cloned
    }

    static BuildRefineRequest(baseConfig, selectedText, providerName) {
        conf := this.CloneConfigWithProvider(baseConfig, providerName)
        prompt := conf.SystemPrompt . "`n`nInput Text:`n" . selectedText

        return this.CreateRequest(prompt, conf)
    }

    static ShouldRetryRefineProviderTask(task) {
        return task.HasOwnProp("LastHttpStatus")
            && RisAIProviderPolicy.ShouldRetryModelStatus(task.LastHttpStatus)
            && task.ModelIndex < task.Models.Length
    }

    static BuildTransportResponse(req) {
        return {
            Status: req.Status,
            ResponseText: req.ResponseText
        }
    }

    static ParseProviderResponse(providerName, responseText) {
        return (providerName == "openai")
            ? RisAIText.ParseOpenAIResponse(responseText)
            : RisAIText.ParseGoogleResponse(responseText)
    }

    static BuildProviderResponseResult(parsed, request, apiTime, providerName) {
        return {
            Result: parsed,
            ApiTime: apiTime,
            APIKeyName: request.APIKeyName,
            Model: request.Model,
            Provider: providerName
        }
    }

    static BuildProviderCallResult(parsed, request, providerName) {
        return {
            Result: parsed,
            APIKeyName: request.APIKeyName,
            Model: request.Model,
            Provider: providerName
        }
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
