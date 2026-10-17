#Requires AutoHotkey v2.0

class RisAITransport {
    static WaitForResponse(req, timeoutMs := 60000) {
        startTime := A_TickCount
        while !req.WaitForResponse(0.01) {
            if (timeoutMs > 0 && (A_TickCount - startTime >= timeoutMs)) {
                try {
                    req.Abort()
                }
                throw Error("AI 請求逾時 (超過 " . timeoutMs . " 毫秒)")
            }
            Sleep(10)
        }
    }

    static SendGoogleAsync(requestOrUrl, payload := "", apiKey := "") {
        req := ComObject("WinHttp.WinHttpRequest.5.1")
        req.SetTimeouts(5000, 5000, 10000, 60000)

        if IsObject(requestOrUrl) {
            url := requestOrUrl.Url
            body := requestOrUrl.Payload
            key := requestOrUrl.HasOwnProp("APIKey") ? requestOrUrl.APIKey : ""
        } else {
            url := requestOrUrl
            body := payload
            key := apiKey
        }

        req.Open("POST", url, True)
        req.SetRequestHeader("Content-Type", "application/json")
        if (key != "") {
            req.SetRequestHeader("x-goog-api-key", key)
        }
        req.Send(body)
        return req
    }

    static SendGoogle(requestOrUrl, payloadOrTimeout := "", maybeTimeout := 60000) {
        if IsObject(requestOrUrl) {
            timeoutMs := (payloadOrTimeout != "") ? payloadOrTimeout : 60000
            req := this.SendGoogleAsync(requestOrUrl)
        } else {
            timeoutMs := maybeTimeout
            req := this.SendGoogleAsync(requestOrUrl, payloadOrTimeout)
        }
        this.WaitForResponse(req, timeoutMs)

        return {
            Status: req.Status,
            ResponseText: req.ResponseText
        }
    }

    static SendOpenAIAsync(request) {
        req := ComObject("WinHttp.WinHttpRequest.5.1")
        req.SetTimeouts(5000, 5000, 10000, 60000)
        req.Open("POST", request.Url, True)
        req.SetRequestHeader("Content-Type", "application/json")
        req.SetRequestHeader("Authorization", "Bearer " . request.APIKey)
        req.Send(request.Payload)
        return req
    }

    static SendOpenAI(request, timeoutMs := 60000) {
        req := this.SendOpenAIAsync(request)
        this.WaitForResponse(req, timeoutMs)

        return {
            Status: req.Status,
            ResponseText: req.ResponseText
        }
    }
}
