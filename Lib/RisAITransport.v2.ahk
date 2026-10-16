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

    static SendGoogleAsync(url, payload) {
        req := ComObject("WinHttp.WinHttpRequest.5.1")
        req.SetTimeouts(5000, 5000, 10000, 60000)
        req.Open("POST", url, True)
        req.SetRequestHeader("Content-Type", "application/json")
        req.Send(payload)
        return req
    }

    static SendGoogle(url, payload, timeoutMs := 60000) {
        req := this.SendGoogleAsync(url, payload)
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
