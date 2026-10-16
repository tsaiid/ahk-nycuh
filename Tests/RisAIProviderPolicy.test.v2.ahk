#Requires AutoHotkey v2.0

#Include TestLib.v2.ahk
#Include ..\Lib\RisAIProviderPolicy.v2.ahk
#Include ..\Lib\RisAIModelHealth.v2.ahk

RegisterTest("RisAIProviderPolicy.ShouldRetryModelStatus recognizes 429, 500, 502, 503, 504", Test_ShouldRetryModelStatus)
RegisterTest("RisAIModelHealth records retryable status and ignores client errors", Test_ModelHealth_StatusFiltering)

Test_ShouldRetryModelStatus() {
    ; 重試/降級狀態碼 (Rate limit / Server errors / Gateway timeouts)
    retryCodes := [429, 500, 502, 503, 504]
    for code in retryCodes {
        AssertTrue(RisAIProviderPolicy.ShouldRetryModelStatus(code), "Status " . code . " should be retryable")
    }

    ; 非重試狀態碼 (正常或客戶端錯誤)
    nonRetryCodes := [200, 201, 400, 401, 403, 404, 408, 413, 501]
    for code in nonRetryCodes {
        AssertFalse(RisAIProviderPolicy.ShouldRetryModelStatus(code), "Status " . code . " should NOT be retryable")
    }
}

Test_ModelHealth_StatusFiltering() {
    testPolicy := {
        ErrorWindowMs: 60000,
        ErrorThreshold: 2
    }

    testModelClientErr := "test-model-400-" . A_TickCount
    ; 傳入 400 客戶端錯誤，不應觸發健康度降級計數
    RisAIModelHealth.RecordGoogleHttpError(testModelClientErr, testPolicy, 400)
    RisAIModelHealth.RecordGoogleHttpError(testModelClientErr, testPolicy, 400)
    AssertFalse(RisAIModelHealth.IsGoogleModelDegraded(testModelClientErr, testPolicy), "HTTP 400 should not degrade model")

    testModelServerErr := "test-model-429-" . A_TickCount
    ; 傳入 429 速率限制，達到閾值應觸發健康度降級
    RisAIModelHealth.RecordGoogleHttpError(testModelServerErr, testPolicy, 429)
    AssertFalse(RisAIModelHealth.IsGoogleModelDegraded(testModelServerErr, testPolicy), "Should not degrade after 1 error (threshold is 2)")
    RisAIModelHealth.RecordGoogleHttpError(testModelServerErr, testPolicy, 429)
    AssertTrue(RisAIModelHealth.IsGoogleModelDegraded(testModelServerErr, testPolicy), "Should degrade after 2 errors")

    ; 測試降級後的排序優先度重排
    models := [testModelServerErr, "healthy-model"]
    prioritized := RisAIModelHealth.PrioritizeGoogleModels(models, testPolicy)
    AssertEqual("healthy-model", prioritized[1], "Healthy model should be prioritized first")
    AssertEqual(testModelServerErr, prioritized[2], "Degraded model should be pushed to end")
}

RunRegisteredTests()
