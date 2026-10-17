#Requires AutoHotkey v2.0

class RisAIDebug {
    static EscapePowerShellSingleQuotedString(text) {
        return StrReplace(text, "'", "''")
    }

    static MaskApiKey(apiKey) {
        if (apiKey == "") {
            return ""
        }
        len := StrLen(apiKey)
        if (len <= 8) {
            return "********"
        }
        return SubStr(apiKey, 1, 4) . "..." . SubStr(apiKey, len - 3)
    }

    static SanitizeUrl(url) {
        return RegExReplace(url, "i)([?&]key=)[^&]+", "$1[REDACTED]")
    }

    static BuildGoogleCurlCommand(url, payload, apiKey := "") {
        sanitizedUrl := this.SanitizeUrl(url)
        escapedUrl := this.EscapePowerShellSingleQuotedString(sanitizedUrl)

        apiKeyVar := "$apiKey = if ($env:GOOGLE_API_KEY) { $env:GOOGLE_API_KEY } else { '<GOOGLE_API_KEY>' }`r`n"

        return "$uri = '" . escapedUrl . "'`r`n"
            . apiKeyVar
            . "$body = @'`r`n"
            . payload . "`r`n"
            . "'@`r`n"
            . "$sw = [Diagnostics.Stopwatch]::StartNew()`r`n"
            . "$response = curl.exe -sS -X POST -H 'Content-Type: application/json' -H `"x-goog-api-key: $apiKey`" --data-binary $body $uri`r`n"
            . "$sw.Stop()`r`n"
            . "$response`r`n"
            . '"ElapsedMs=$($sw.ElapsedMilliseconds)"'
    }

    static LogGoogleBlockingMetrics(metrics, status := "") {
        statusText := (status != "") ? ", status=" . status : ""
        OutputDebug(Format(
            "[RisController] GoogleAI blocking metrics: config={}ms, payload={}ms, wait={}ms, parse={}ms{}`n",
            metrics.ConfigReadTime,
            metrics.PayloadBuildTime,
            metrics.WaitForResponseTime,
            metrics.ResponseParseTime,
            statusText
        ))
    }
}
