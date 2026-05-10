#Requires AutoHotkey v2.0

class RisAIDebug {
    static EscapePowerShellSingleQuotedString(text) {
        return StrReplace(text, "'", "''")
    }

    static BuildGoogleCurlCommand(url, payload) {
        escapedUrl := this.EscapePowerShellSingleQuotedString(url)

        return "$uri = '" . escapedUrl . "'`r`n"
            . "$body = @'`r`n"
            . payload . "`r`n"
            . "'@`r`n"
            . "$sw = [Diagnostics.Stopwatch]::StartNew()`r`n"
            . "$response = curl.exe -sS -X POST -H 'Content-Type: application/json' --data-binary $body $uri`r`n"
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
