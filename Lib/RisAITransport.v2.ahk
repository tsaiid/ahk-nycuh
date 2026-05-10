#Requires AutoHotkey v2.0

class RisAITransport {
    static WaitForResponse(req) {
        while !req.WaitForResponse(0.01) {
            Sleep(10)
        }
    }

    static SendGoogleAsync(url, payload) {
        req := ComObject("WinHttp.WinHttpRequest.5.1")

        req.Open("POST", url, True)
        req.SetRequestHeader("Content-Type", "application/json")
        req.Send(payload)
        return req
    }

    static SendGoogle(url, payload) {
        req := this.SendGoogleAsync(url, payload)
        this.WaitForResponse(req)

        return {
            Status: req.Status,
            ResponseText: req.ResponseText
        }
    }

    static SendOpenAIAsync(request) {
        req := ComObject("WinHttp.WinHttpRequest.5.1")
        req.Open("POST", request.Url, True)
        req.SetRequestHeader("Content-Type", "application/json")
        req.SetRequestHeader("Authorization", "Bearer " . request.APIKey)
        req.Send(request.Payload)
        return req
    }

    static SendOpenAI(request) {
        req := this.SendOpenAIAsync(request)
        this.WaitForResponse(req)

        return {
            Status: req.Status,
            ResponseText: req.ResponseText
        }
    }
}
