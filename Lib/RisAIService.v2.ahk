#Requires AutoHotkey v2.0

/**
 * 負責 AI provider 的調度、重試、快取與傳輸協調
 */
class RisAIService {
    static _aiProviderTimeoutMs := 90000
    static _aiProviderPollIntervalMs := 25

    /**
     * 發送 AI 請求 (自動選擇 provider)
     * @param promptText Prompt 內容
     * @param aiConfig 覆蓋設定
     * @param options 選項 { Notify: func, ShowCurl: bool, ShowGoogleDebugCurl: func, GetGoogleConfig: func, GetOpenAIConfig: func }
     */
    static Call(promptText, aiConfig := 0, options := 0) {
        defaultProvider := IniRead("config\private.ini", "AI", "Provider", "google")
        providerName := RisAIProviderPolicy.ResolveProvider(aiConfig, defaultProvider)

        switch providerName {
            case "google":
                return this.CallGoogle(promptText, aiConfig, options)
            case "openai":
                return this.CallOpenAI(promptText, aiConfig, options)
            default:
                throw Error("不支援的 AI provider: " . providerName)
        }
    }

    static CallGoogle(promptText, aiConfig := 0, options := 0) {
        return this.CallGoogleWithInlineImage(promptText, 0, aiConfig, options)
    }

    static CallGoogleWithInlineImage(promptText, inlineImage := 0, aiConfig := 0, options := 0) {
        notify := (IsObject(options) && options.HasOwnProp("Notify")) ? options.Notify : (*) => 0
        showCurl := (IsObject(options) && options.HasOwnProp("ShowCurl")) ? options.ShowCurl : false

        models := this._ResolveGoogleAIModelList(aiConfig)
        healthPolicy := RisAIModelHealth.GetGooglePolicy()
        models := RisAIModelHealth.PrioritizeGoogleModels(models, healthPolicy)
        lastError := ""

        for index, modelName in models {
            request := this._BuildGoogleAIRequest(promptText, aiConfig, modelName, options, inlineImage)
            waitStart := A_TickCount
            response := RisAITransport.SendGoogle(request.Url, request.Payload)
            request.Metrics.WaitForResponseTime := A_TickCount - waitStart
            this._MaybeShowGoogleDebugCurl(showCurl, options, request.Url, request.Payload, response, request)

            if (response.Status != 200) {
                request.Metrics.ResponseParseTime := 0
                RisAIDebug.LogGoogleBlockingMetrics(request.Metrics, response.Status)
                RisAIModelHealth.RecordGoogleHttpError(request.Model, healthPolicy)
                lastError := "HTTP " . response.Status . " (" . request.Model . ") - " . response.ResponseText
                if (RisAIProviderPolicy.ShouldRetryModelStatus(response.Status) && index < models.Length) {
                    notify("AI model 發生 HTTP " . response.Status . "，改用 " . models[index + 1] . " 重試", 2500)
                    OutputDebug("[RisAIService] GoogleAI retry with fallback model after HTTP " . response.Status . ": " . request.Model . "`n")
                    continue
                }
                throw Error(lastError)
            }

            parseStart := A_TickCount
            parsed := RisAIOrchestration.ParseProviderResponse("google", response.ResponseText)
            request.Metrics.ResponseParseTime := A_TickCount - parseStart
            RisAIDebug.LogGoogleBlockingMetrics(request.Metrics, response.Status)
            return RisAIOrchestration.BuildProviderCallResult(parsed, request, "google")
        }

        throw Error(lastError)
    }

    static CallOpenAI(promptText, aiConfig := 0, options := 0) {
        notify := (IsObject(options) && options.HasOwnProp("Notify")) ? options.Notify : (*) => 0

        models := this._ResolveOpenAIModelList(aiConfig)
        lastError := ""

        for index, modelName in models {
            request := this._BuildOpenAIRequest(promptText, aiConfig, modelName)
            waitStart := A_TickCount
            response := RisAITransport.SendOpenAI(request)
            request.Metrics.WaitForResponseTime := A_TickCount - waitStart

            if (response.Status != 200) {
                lastError := "HTTP " . response.Status . " (" . request.Model . ") - " . response.ResponseText
                if (RisAIProviderPolicy.ShouldRetryModelStatus(response.Status) && index < models.Length) {
                    notify("AI model 發生 HTTP " . response.Status . "，改用 " . models[index + 1] . " 重試", 2500)
                    continue
                }
                throw Error(lastError)
            }

            parsed := RisAIOrchestration.ParseProviderResponse("openai", response.ResponseText)
            return RisAIOrchestration.BuildProviderCallResult(parsed, request, "openai")
        }

        throw Error(lastError)
    }

    /**
     * 並行執行多個 provider 的潤色請求 (用於比對)
     */
    static RunRefineProvidersParallel(selectedText, providerSpecs, trailingNewlines := "", options := 0) {
        notify := (IsObject(options) && options.HasOwnProp("Notify")) ? options.Notify : (*) => 0
        showCurl := (IsObject(options) && options.HasOwnProp("ShowCurl")) ? options.ShowCurl : false

        results := Map()
        tasks := []

        for _, spec in providerSpecs {
            try {
                tasks.Push(this._StartRefineProviderTask(selectedText, spec.Provider, spec.DisplayName, options))
            } catch as err {
                results[spec.Provider] := RisAIOrchestration.BuildRefineProviderFailureResult(spec.DisplayName, err.Message)
            }
        }

        loop {
            pendingCount := 0

            for _, task in tasks {
                if (task.Done) {
                    continue
                }

                if (A_TickCount - task.TaskStartedAt > this._aiProviderTimeoutMs) {
                    task.Done := true
                    try task.Req.Abort()
                    results[task.Provider] := RisAIOrchestration.BuildRefineProviderFailureResult(task.DisplayName, "AI request timeout")
                    continue
                }

                if (!task.Req.WaitForResponse(0)) {
                    pendingCount += 1
                    continue
                }

                task.CompletedAt := A_TickCount
                try {
                    response := this._FinalizeRefineProviderTask(task, task.Options)
                    task.Done := true
                    results[task.Provider] := RisAIOrchestration.BuildRefineProviderSuccessResult(task.DisplayName, response, trailingNewlines)
                } catch as err {
                    if (RisAIOrchestration.ShouldRetryRefineProviderTask(task)) {
                        try {
                            this._StartNextRefineProviderRequest(task)
                            pendingCount += 1
                            continue
                        } catch as retryErr {
                            task.Done := true
                            results[task.Provider] := RisAIOrchestration.BuildRefineProviderFailureResult(task.DisplayName, retryErr.Message)
                            continue
                        }
                    }

                    task.Done := true
                    results[task.Provider] := RisAIOrchestration.BuildRefineProviderFailureResult(task.DisplayName, err.Message)
                }
            }

            if (pendingCount == 0) {
                break
            }

            Sleep(this._aiProviderPollIntervalMs)
        }

        return results
    }

    static _StartRefineProviderTask(selectedText, providerName, displayName, options := 0) {
        request := RisAIOrchestration.BuildRefineRequest(RisConfig.AI.Refine, selectedText, providerName)
        providerName := StrLower(Trim(providerName))

        switch providerName {
            case "google":
                models := this._ResolveGoogleAIModelList(request.Config, options)
                healthPolicy := RisAIModelHealth.GetGooglePolicy()
                models := RisAIModelHealth.PrioritizeGoogleModels(models, healthPolicy)
                task := {
                    Provider: providerName,
                    DisplayName: displayName,
                    Prompt: request.Prompt,
                    Config: request.Config,
                    Options: options,
                    Models: models,
                    ModelIndex: 0,
                    TaskStartedAt: A_TickCount,
                    Done: false,
                    HealthPolicy: healthPolicy
                }
                this._StartNextRefineProviderRequest(task)
                return task
            case "openai":
                models := this._ResolveOpenAIModelList(request.Config, options)
                task := {
                    Provider: providerName,
                    DisplayName: displayName,
                    Prompt: request.Prompt,
                    Config: request.Config,
                    Options: options,
                    Models: models,
                    ModelIndex: 0,
                    TaskStartedAt: A_TickCount,
                    Done: false
                }
                this._StartNextRefineProviderRequest(task)
                return task
            default:
                throw Error("不支援的 AI provider: " . providerName)
        }
    }

    static _StartNextRefineProviderRequest(task) {
        task.ModelIndex += 1
        if (task.ModelIndex > task.Models.Length) {
            throw Error("沒有可用的 " . task.DisplayName . " fallback model")
        }

        modelName := task.Models[task.ModelIndex]
        options := task.HasOwnProp("Options") ? task.Options : 0
        if (task.Provider == "google") {
            providerRequest := this._BuildGoogleAIRequest(task.Prompt, task.Config, modelName, options)
            startedAt := A_TickCount
            req := RisAITransport.SendGoogleAsync(providerRequest.Url, providerRequest.Payload)
        } else if (task.Provider == "openai") {
            providerRequest := this._BuildOpenAIRequest(task.Prompt, task.Config, modelName, options)
            startedAt := A_TickCount
            req := RisAITransport.SendOpenAIAsync(providerRequest)
        } else {
            throw Error("不支援的 AI provider: " . task.Provider)
        }

        task.Request := providerRequest
        task.Req := req
        task.StartedAt := startedAt
        task.CompletedAt := 0
        task.LastHttpStatus := ""
        task.LastError := ""
    }

    static _FinalizeRefineProviderTask(task, options := 0) {
        notify := (IsObject(options) && options.HasOwnProp("Notify")) ? options.Notify : (*) => 0
        showCurl := (IsObject(options) && options.HasOwnProp("ShowCurl")) ? options.ShowCurl : false

        request := task.Request
        response := RisAIOrchestration.BuildTransportResponse(task.Req)

        completedAt := task.HasOwnProp("CompletedAt") ? task.CompletedAt : A_TickCount
        providerLatency := completedAt - task.StartedAt
        request.Metrics.WaitForResponseTime := providerLatency
        if (task.Provider == "google") {
            this._MaybeShowGoogleDebugCurl(showCurl, options, request.Url, request.Payload, response, request)
        }

        if (response.Status != 200) {
            task.LastHttpStatus := response.Status
            task.LastError := "HTTP " . response.Status . " (" . request.Model . ") - " . response.ResponseText
            if (task.Provider == "google") {
                request.Metrics.ResponseParseTime := 0
                RisAIDebug.LogGoogleBlockingMetrics(request.Metrics, response.Status)
                RisAIModelHealth.RecordGoogleHttpError(request.Model, task.HealthPolicy)
            }
            throw Error(task.LastError)
        }

        parseStart := A_TickCount
        parsed := RisAIOrchestration.ParseProviderResponse(task.Provider, response.ResponseText)
        request.Metrics.ResponseParseTime := A_TickCount - parseStart

        if (task.Provider == "google") {
            RisAIDebug.LogGoogleBlockingMetrics(request.Metrics, response.Status)
        }

        return RisAIOrchestration.BuildProviderResponseResult(parsed, request, providerLatency, task.Provider)
    }

    ; --- 內部 Helper (Config & Request Building) ---

    static _MaybeShowGoogleDebugCurl(showCurl, options, url, payload, response, request) {
        if (!showCurl || !IsObject(options) || !options.HasOwnProp("ShowGoogleDebugCurl")) {
            return
        }

        options.ShowGoogleDebugCurl.Call(url, payload, response, request)
    }

    static _GetGoogleAIConfig(options := 0) {
        if (IsObject(options) && options.HasOwnProp("GetGoogleConfig")) {
            return options.GetGoogleConfig.Call()
        }

        return RisAIConfigResolver.GetGoogleConfig()
    }

    static _ResolveGoogleAIModelList(aiConfig := 0, options := 0) {
        cfg := this._GetGoogleAIConfig(options)
        return RisAIProviderPolicy.ResolveModelList(aiConfig, "google", cfg.Model)
    }

    static _ResolveGoogleAIOptions(aiConfig := 0, modelOverride := "", options := 0) {
        cfg := this._GetGoogleAIConfig(options)
        return RisAIConfigResolver.ResolveGoogleOptions(cfg, aiConfig, modelOverride)
    }

    static _BuildGoogleAIRequest(promptText, aiConfig := 0, modelOverride := "", optionsSource := 0, inlineImage := 0) {
        configStart := A_TickCount
        options := this._ResolveGoogleAIOptions(aiConfig, modelOverride, optionsSource)
        configTime := A_TickCount - configStart

        return RisAIRequestBuilder.BuildGoogleRequest(promptText, options, configTime, inlineImage)
    }

    static _GetOpenAIConfig(options := 0) {
        if (IsObject(options) && options.HasOwnProp("GetOpenAIConfig")) {
            return options.GetOpenAIConfig.Call()
        }

        return RisAIConfigResolver.GetOpenAIConfig()
    }

    static _ResolveOpenAIModelList(aiConfig := 0, options := 0) {
        cfg := this._GetOpenAIConfig(options)
        return RisAIProviderPolicy.ResolveModelList(aiConfig, "openai", cfg.Model)
    }

    static _ResolveOpenAIOptions(aiConfig := 0, modelOverride := "", options := 0) {
        cfg := this._GetOpenAIConfig(options)
        return RisAIConfigResolver.ResolveOpenAIOptions(cfg, aiConfig, modelOverride)
    }

    static _BuildOpenAIRequest(promptText, aiConfig := 0, modelOverride := "", optionsSource := 0) {
        configStart := A_TickCount
        options := this._ResolveOpenAIOptions(aiConfig, modelOverride, optionsSource)
        configTime := A_TickCount - configStart

        return RisAIRequestBuilder.BuildOpenAIRequest(promptText, options, configTime)
    }
}
