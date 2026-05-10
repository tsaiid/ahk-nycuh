#Requires AutoHotkey v2.0

class RisAIModelHealth {
    static _googleModelHttpErrors := Map()

    static GetGooglePolicy() {
        if (RisConfig.HasOwnProp("GoogleAIModelHealth")) {
            policy := RisConfig.GoogleAIModelHealth
            return {
                ErrorWindowMs: policy.HasOwnProp("ErrorWindowMs") ? policy.ErrorWindowMs : 3600000,
                ErrorThreshold: policy.HasOwnProp("ErrorThreshold") ? policy.ErrorThreshold : 3
            }
        }

        return {
            ErrorWindowMs: 3600000,
            ErrorThreshold: 3
        }
    }

    static PruneGoogleHttpErrors(modelName, nowTick, policy) {
        if (!this._googleModelHttpErrors.Has(modelName)) {
            this._googleModelHttpErrors[modelName] := []
            return this._googleModelHttpErrors[modelName]
        }

        pruned := []
        for _, errorTick in this._googleModelHttpErrors[modelName] {
            if (nowTick - errorTick <= policy.ErrorWindowMs) {
                pruned.Push(errorTick)
            }
        }

        this._googleModelHttpErrors[modelName] := pruned
        return pruned
    }

    static RecordGoogleHttpError(modelName, policy) {
        nowTick := A_TickCount
        errors := this.PruneGoogleHttpErrors(modelName, nowTick, policy)
        errors.Push(nowTick)
        this._googleModelHttpErrors[modelName] := errors
        OutputDebug(Format(
            "[RisController] GoogleAI model HTTP error count: model={}, count={}, window={}ms`n",
            modelName,
            errors.Length,
            policy.ErrorWindowMs
        ))
    }

    static IsGoogleModelDegraded(modelName, policy) {
        errors := this.PruneGoogleHttpErrors(modelName, A_TickCount, policy)
        return errors.Length >= policy.ErrorThreshold
    }

    static PrioritizeGoogleModels(models, policy) {
        preferredModels := []
        degradedModels := []

        for _, modelName in models {
            if (this.IsGoogleModelDegraded(modelName, policy)) {
                degradedModels.Push(modelName)
            } else {
                preferredModels.Push(modelName)
            }
        }

        if (preferredModels.Length == 0) {
            OutputDebug("[RisController] GoogleAI all models are degraded; preserve configured model order`n")
            return models
        }

        for _, modelName in degradedModels {
            preferredModels.Push(modelName)
        }

        return preferredModels
    }
}
