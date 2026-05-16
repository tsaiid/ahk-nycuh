#Requires AutoHotkey v2.0

class RisAIPayload {
    static BuildGoogleUrl(options) {
        return "https://generativelanguage.googleapis.com/v1beta/models/" . options.Model . ":generateContent?key=" . options.APIKey
    }

    static BuildGooglePayload(promptText, options, inlineImage := 0) {
        escapedPrompt := RisAIText.EscapeJsonString(promptText)
        generationConfig := '"temperature": ' . options.Temperature . ','
        if (options.ThinkingLevel != "") {
            generationConfig .= '"thinkingConfig": {"thinkingLevel": "' . RisAIText.EscapeJsonString(options.ThinkingLevel) . '"},'
        }
        generationConfig .= '"topP": ' . options.TopP

        parts := '[{"text": "' . escapedPrompt . '"}'
        if IsObject(inlineImage) {
            parts .= ',{"inline_data": {"mime_type": "' . RisAIText.EscapeJsonString(inlineImage.MimeType) . '","data": "' . inlineImage.Base64 . '"}}'
        }
        parts .= ']'

        payload := '{'
            . '"contents": [{'
                . '"role": "user",'
                . '"parts": ' . parts
            . '}],'
            . '"generationConfig": {' . generationConfig . '}'

        if (options.EnableGoogleSearch) {
            payload .= ',"tools": [{"googleSearch": {}}]'
        }

        return payload . '}'
    }

    static BuildOpenAIPayload(promptText, options) {
        escapedPrompt := RisAIText.EscapeJsonString(promptText)
        payload := '{'
            . '"model":"' . RisAIText.EscapeJsonString(options.Model) . '",'
            . '"input":[{"role":"user","content":[{"type":"input_text","text":"' . escapedPrompt . '"}]}],'
            . '"reasoning":{"effort":"' . RisAIText.EscapeJsonString(options.ReasoningEffort) . '"},'
            . '"temperature":' . options.Temperature
        return payload . '}'
    }
}
