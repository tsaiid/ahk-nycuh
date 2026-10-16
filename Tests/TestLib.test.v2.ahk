#Requires AutoHotkey v2.0

#Include TestLib.v2.ahk

RegisterTest("AssertEqual passes on exact match", Test_AssertEqual_PassesOnExactMatch)
RegisterTest("AssertEqual throws on case mismatch", Test_AssertEqual_ThrowsOnCaseMismatch)
RegisterTest("AssertEqual throws on value mismatch", Test_AssertEqual_ThrowsOnValueMismatch)
RegisterTest("AssertThrows catches expected error class", Test_AssertThrows_CatchesExpectedErrorClass)
RegisterTest("AssertThrows catches any error when type omitted", Test_AssertThrows_CatchesAnyErrorWhenTypeOmitted)
RegisterTest("AssertThrows catches error by type name string", Test_AssertThrows_CatchesErrorByTypeNameString)
RegisterTest("AssertThrows fails when no exception thrown", Test_AssertThrows_FailsWhenNoExceptionThrown)
RegisterTest("AssertThrows fails on mismatched error type", Test_AssertThrows_FailsOnMismatchedErrorType)
RegisterTest("AssertThrows throws TypeError for non-callable", Test_AssertThrows_ThrowsTypeErrorForNonCallable)
RegisterTest("AssertTrue and AssertFalse behavior", Test_AssertTrue_And_AssertFalse)

Test_AssertEqual_PassesOnExactMatch() {
    AssertEqual("Hello", "Hello", "Exact strings should match")
    AssertEqual(123, 123, "Numbers should match")
    AssertEqual(true, true, "Booleans should match")
    AssertEqual("", "", "Empty strings should match")
}

Test_AssertEqual_ThrowsOnCaseMismatch() {
    ThrowCaseMismatch() {
        AssertEqual("CaseSensitive", "casesensitive")
    }
    AssertThrows(Error, ThrowCaseMismatch, "Should throw on case mismatch")

    ThrowUpperLower() {
        AssertEqual("CT", "ct")
    }
    AssertThrows(Error, ThrowUpperLower, "CT vs ct should throw")
}

Test_AssertEqual_ThrowsOnValueMismatch() {
    ThrowNumMismatch() {
        AssertEqual(1, 2)
    }
    AssertThrows(Error, ThrowNumMismatch, "1 vs 2 should throw")

    ThrowStrMismatch() {
        AssertEqual("apple", "banana")
    }
    AssertThrows(Error, ThrowStrMismatch, "apple vs banana should throw")
}

Test_AssertThrows_CatchesExpectedErrorClass() {
    ThrowValueError() {
        throw ValueError("invalid parameter")
    }
    err := AssertThrows(ValueError, ThrowValueError, "Should catch ValueError")
    AssertEqual("invalid parameter", err.Message)
}

Test_AssertThrows_CatchesAnyErrorWhenTypeOmitted() {
    ThrowCustomError() {
        throw Error("generic failure")
    }
    ; Single-argument form
    err1 := AssertThrows(ThrowCustomError)
    AssertEqual("generic failure", err1.Message)

    ; Callback + message form
    err2 := AssertThrows(ThrowCustomError, "Should catch with message")
    AssertEqual("generic failure", err2.Message)
}

Test_AssertThrows_CatchesErrorByTypeNameString() {
    ThrowTypeError() {
        throw TypeError("type mismatch")
    }
    err := AssertThrows("TypeError", ThrowTypeError, "Should catch by string type name")
    AssertEqual("type mismatch", err.Message)
}

Test_AssertThrows_FailsWhenNoExceptionThrown() {
    NoopAction() {
        return 42
    }
    RunAssertThrowsOnNoop() {
        AssertThrows(NoopAction)
    }
    AssertThrows(Error, RunAssertThrowsOnNoop, "AssertThrows should fail if action does not throw")
}

Test_AssertThrows_FailsOnMismatchedErrorType() {
    ThrowTargetError() {
        throw TargetError("window not found")
    }
    RunAssertThrowsWithWrongType() {
        AssertThrows(ValueError, ThrowTargetError)
    }
    AssertThrows(Error, RunAssertThrowsWithWrongType, "AssertThrows should fail on type mismatch")
}

Test_AssertThrows_ThrowsTypeErrorForNonCallable() {
    RunAssertThrowsOnNonCallable() {
        AssertThrows("not_a_function")
    }
    AssertThrows(TypeError, RunAssertThrowsOnNonCallable, "Non-callable should trigger TypeError")
}

Test_AssertTrue_And_AssertFalse() {
    AssertTrue(1 == 1, "1 == 1 is true")
    AssertFalse(1 == 2, "1 == 2 is false")

    FailAssertTrue() {
        AssertTrue(false)
    }
    AssertThrows(Error, FailAssertTrue)

    FailAssertFalse() {
        AssertFalse(true)
    }
    AssertThrows(Error, FailAssertFalse)
}

RunRegisteredTests()
