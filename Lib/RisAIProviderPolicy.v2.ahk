#Requires AutoHotkey v2.0

class RisAIProviderPolicy {
    static ShouldRetryModelStatus(status) {
        return status == 500 || status == 503
    }

    static GoogleModelSupportsThinkingLevel(modelName) {
        unsupportedModels := Map(
            "gemini-2.5-flash-lite", false
        )

        return !unsupportedModels.Has(StrLower(Trim(modelName)))
    }
}
