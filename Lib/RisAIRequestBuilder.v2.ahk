#Requires AutoHotkey v2.0

class RisAIRequestBuilder {
    static BuildGoogleRequest(promptText, options, configTime) {
        payloadStart := A_TickCount

        return {
            Url: RisAIPayload.BuildGoogleUrl(options),
            Payload: RisAIPayload.BuildGooglePayload(promptText, options),
            APIKeyName: options.APIKeyName,
            Model: options.Model,
            Metrics: {
                ConfigReadTime: configTime,
                PayloadBuildTime: A_TickCount - payloadStart
            }
        }
    }

    static BuildOpenAIRequest(promptText, options, configTime) {
        payloadStart := A_TickCount

        return {
            Url: options.BaseUrl,
            Payload: RisAIPayload.BuildOpenAIPayload(promptText, options),
            APIKey: options.APIKey,
            APIKeyName: options.APIKeyName,
            Model: options.Model,
            Metrics: {
                ConfigReadTime: configTime,
                PayloadBuildTime: A_TickCount - payloadStart
            }
        }
    }
}
