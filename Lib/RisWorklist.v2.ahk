#Requires AutoHotkey v2.0

class RisWorklist {
    static BuildJson(categories, categoryData) {
        jsonStr := "{"
        validDataCount := 0

        for i, cat in categories {
            gridData := categoryData.Has(cat) ? categoryData[cat] : Map()
            validDataCount += gridData.Count

            if (i > 1)
                jsonStr .= ", "
            jsonStr .= '"' . cat . '": {'

            isFirstProp := true
            for k, v in gridData {
                if (!isFirstProp)
                    jsonStr .= ", "
                safeKey := StrReplace(k, "-", "_")
                valStr := IsNumber(v) ? v : '"' . v . '"'
                jsonStr .= Format('"{1}": {2}', safeKey, valStr)
                isFirstProp := false
            }
            jsonStr .= "}"
        }
        jsonStr .= "}"

        return {
            Json: jsonStr,
            ValidDataCount: validDataCount
        }
    }

    static ExtractGridData(elWindow, gridSelector) {
        data := Map()
        try {
            try {
                elGrid := elWindow.FindElement(gridSelector)
            } catch {
                return data
            }

            try {
                rowElements := elGrid.FindAll({Type: "Custom"})
            } catch {
                return data
            }

            if (rowElements.Length == 0)
                return data

            walker := UIA.TreeWalkerTrue

            for row in rowElements {
                keyEl := walker.TryGetFirstChildElement(row)
                if (!keyEl)
                    continue

                valEl := walker.TryGetNextSiblingElement(keyEl)
                if (!valEl)
                    continue

                k := "", v := "0"
                try k := keyEl.Value
                try v := valEl.Value

                if (k != "") {
                    k := Trim(k)
                    v := Trim(v)
                    if (v == "")
                        v := 0
                    data[k] := v
                }
            }
        } catch {
        }
        return data
    }

    static PostDataToWebhook(jsonStr, isSilent := false, notify := 0) {
        notifyFn := IsObject(notify) ? notify : (*) => 0
        logFn := (msg) => (isSilent ? OutputDebug("[RisPost] " . msg . "`n") : notifyFn(msg))

        configFile := "config\private.ini"
        url  := IniRead(configFile, "n8n", "WebhookURL", "")
        user := IniRead(configFile, "n8n", "Username", "")
        pass := IniRead(configFile, "n8n", "Password", "")

        if (url == "") {
            logFn("❌ 錯誤：找不到 WebhookURL 設定")
            return
        }

        try {
            req := ComObject("WinHttp.WinHttpRequest.5.1")
            req.Open("POST", url, False)
            req.SetRequestHeader("Content-Type", "application/json")

            if (user != "" && pass != "") {
                authStr := this.Base64Encode(user . ":" . pass)
                req.SetRequestHeader("Authorization", "Basic " . authStr)
            }

            req.Send(jsonStr)

            if (req.Status == 200) {
                logFn("✅ 資料已上傳至 n8n")
            } else {
                logFn("❌ 上傳失敗 (Status: " . req.Status . ")")
            }
        } catch as err {
            logFn("❌ 網路錯誤: " . err.Message)
        }
    }

    static Base64Encode(text) {
        buf := Buffer(StrPut(text, "UTF-8"))
        StrPut(text, buf, "UTF-8")

        flags := 0x40000001

        reqSize := 0
        DllCall("Crypt32\CryptBinaryToStringW", "Ptr", buf, "UInt", buf.Size - 1, "UInt", flags, "Ptr", 0, "UInt*", &reqSize)

        outBuf := Buffer(reqSize * 2)
        DllCall("Crypt32\CryptBinaryToStringW", "Ptr", buf, "UInt", buf.Size - 1, "UInt", flags, "Ptr", outBuf, "UInt*", &reqSize)

        return StrGet(outBuf)
    }
}
