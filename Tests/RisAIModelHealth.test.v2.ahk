#Requires AutoHotkey v2.0

#Include TestLib.v2.ahk
#Include ..\Lib\RisAIModelHealth.v2.ahk

RegisterTest("RisAIModelHealth.GetGooglePolicy returns policy with defaults or config", Test_GetGooglePolicy)
RegisterTest("RisAIModelHealth.RecordGoogleHttpError filters status and records retryable errors", Test_RecordGoogleHttpError)
RegisterTest("RisAIModelHealth.PruneGoogleHttpErrors removes ticks outside ErrorWindowMs", Test_PruneGoogleHttpErrors)
RegisterTest("RisAIModelHealth.IsGoogleModelDegraded checks error threshold", Test_IsGoogleModelDegraded)
RegisterTest("RisAIModelHealth.PrioritizeGoogleModels prioritizes healthy over degraded", Test_PrioritizeGoogleModels)
RegisterTest("RisAIModelHealth.PrioritizeGoogleModels preserves order if all degraded", Test_PrioritizeGoogleModels_AllDegraded)

Test_GetGooglePolicy() {
    policy := RisAIModelHealth.GetGooglePolicy()
    AssertTrue(IsObject(policy), "Policy should be an object")
    AssertTrue(policy.HasOwnProp("ErrorWindowMs"), "Policy should have ErrorWindowMs")
    AssertTrue(policy.HasOwnProp("ErrorThreshold"), "Policy should have ErrorThreshold")
    AssertTrue(policy.ErrorWindowMs > 0, "ErrorWindowMs should be positive")
    AssertTrue(policy.ErrorThreshold > 0, "ErrorThreshold should be positive")
}

Test_RecordGoogleHttpError() {
    testPolicy := {
        ErrorWindowMs: 60000,
        ErrorThreshold: 2
    }

    testModelClientErr := "test-model-400-" . A_TickCount
    ; 傳入 400 客戶端錯誤，不應計入錯誤
    RisAIModelHealth.RecordGoogleHttpError(testModelClientErr, testPolicy, 400)
    RisAIModelHealth.RecordGoogleHttpError(testModelClientErr, testPolicy, 404)
    AssertFalse(RisAIModelHealth.IsGoogleModelDegraded(testModelClientErr, testPolicy), "Client errors (400, 404) should not degrade model")

    testModelRetryable := "test-model-retryable-" . A_TickCount
    ; 傳入 429 速率限制
    RisAIModelHealth.RecordGoogleHttpError(testModelRetryable, testPolicy, 429)
    AssertFalse(RisAIModelHealth.IsGoogleModelDegraded(testModelRetryable, testPolicy), "Should not degrade after 1 error (threshold is 2)")
    ; 傳入 503 伺服器錯誤
    RisAIModelHealth.RecordGoogleHttpError(testModelRetryable, testPolicy, 503)
    AssertTrue(RisAIModelHealth.IsGoogleModelDegraded(testModelRetryable, testPolicy), "Should degrade after 2 retryable errors")
}

Test_PruneGoogleHttpErrors() {
    testModel := "test-model-prune-" . A_TickCount
    policy := {
        ErrorWindowMs: 2000,
        ErrorThreshold: 2
    }

    ; 初始化錯誤記錄，包含過期與未過期的 tick
    ; 假設當前時間為 10000，window 為 2000 (有效區間: 8000 ~ 10000)
    nowTick := 10000
    errors := [5000, 7000, 8500, 9500]
    RisAIModelHealth._googleModelHttpErrors[testModel] := errors

    pruned := RisAIModelHealth.PruneGoogleHttpErrors(testModel, nowTick, policy)
    AssertEqual(2, pruned.Length, "Should prune ticks older than nowTick - ErrorWindowMs")
    AssertEqual(8500, pruned[1], "First retained tick should be 8500")
    AssertEqual(9500, pruned[2], "Second retained tick should be 9500")
}

Test_IsGoogleModelDegraded() {
    testModel := "test-model-degraded-" . A_TickCount
    policy := {
        ErrorWindowMs: 60000,
        ErrorThreshold: 3
    }

    RisAIModelHealth.RecordGoogleHttpError(testModel, policy, 500)
    AssertFalse(RisAIModelHealth.IsGoogleModelDegraded(testModel, policy), "1 error should not degrade when threshold is 3")

    RisAIModelHealth.RecordGoogleHttpError(testModel, policy, 502)
    AssertFalse(RisAIModelHealth.IsGoogleModelDegraded(testModel, policy), "2 errors should not degrade when threshold is 3")

    RisAIModelHealth.RecordGoogleHttpError(testModel, policy, 504)
    AssertTrue(RisAIModelHealth.IsGoogleModelDegraded(testModel, policy), "3 errors should degrade when threshold is 3")
}

Test_PrioritizeGoogleModels() {
    testPolicy := {
        ErrorWindowMs: 60000,
        ErrorThreshold: 1
    }

    mHealthy1 := "healthy-model-1-" . A_TickCount
    mHealthy2 := "healthy-model-2-" . A_TickCount
    mDegraded := "degraded-model-" . A_TickCount

    RisAIModelHealth.RecordGoogleHttpError(mDegraded, testPolicy, 429)

    ; 初始清單：降級模型在最前
    models := [mDegraded, mHealthy1, mHealthy2]
    prioritized := RisAIModelHealth.PrioritizeGoogleModels(models, testPolicy)

    AssertEqual(3, prioritized.Length, "Model count should remain the same")
    AssertEqual(mHealthy1, prioritized[1], "First healthy model should come first")
    AssertEqual(mHealthy2, prioritized[2], "Second healthy model should come second")
    AssertEqual(mDegraded, prioritized[3], "Degraded model should be deprioritized to end")
}

Test_PrioritizeGoogleModels_AllDegraded() {
    testPolicy := {
        ErrorWindowMs: 60000,
        ErrorThreshold: 1
    }

    mDegraded1 := "all-degraded-1-" . A_TickCount
    mDegraded2 := "all-degraded-2-" . A_TickCount

    RisAIModelHealth.RecordGoogleHttpError(mDegraded1, testPolicy, 500)
    RisAIModelHealth.RecordGoogleHttpError(mDegraded2, testPolicy, 500)

    ; 當所有模型均降級時，應保留原始配置順序
    models := [mDegraded1, mDegraded2]
    prioritized := RisAIModelHealth.PrioritizeGoogleModels(models, testPolicy)

    AssertEqual(2, prioritized.Length, "Model count should remain the same")
    AssertEqual(mDegraded1, prioritized[1], "Should preserve first model order")
    AssertEqual(mDegraded2, prioritized[2], "Should preserve second model order")
}

RunRegisteredTests()
