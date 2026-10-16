#Requires AutoHotkey v2.0

global TestCases := []

RegisterTest(name, callback) {
    global TestCases
    TestCases.Push({Name: name, Callback: callback})
}

AssertEqual(expected, actual, message := "") {
    if !(expected == actual) {
        prefix := message != "" ? message . ": " : ""
        throw Error(prefix . "expected " . FormatValue(expected) . ", got " . FormatValue(actual))
    }
}

AssertTrue(condition, message := "") {
    if (!condition) {
        throw Error(message != "" ? message : "expected condition to be true")
    }
}

AssertFalse(condition, message := "") {
    if (condition) {
        throw Error(message != "" ? message : "expected condition to be false")
    }
}

AssertThrows(arg1, arg2 := unset, arg3 := "") {
    if !IsSet(arg2) {
        expectedType := ""
        action := arg1
        message := ""
    } else if (arg1 is Class) {
        expectedType := arg1
        action := arg2
        message := arg3
    } else if (HasMethod(arg2, "Call")) {
        expectedType := arg1
        action := arg2
        message := arg3
    } else {
        expectedType := ""
        action := arg1
        message := arg2
    }

    prefix := message != "" ? message . ": " : ""

    if !HasMethod(action, "Call") {
        throw TypeError(prefix . "expected action to be callable")
    }

    try {
        action.Call()
    } catch as err {
        if !_IsMatchingError(err, expectedType) {
            expectedName := _GetErrorTypeName(expectedType)
            actualName := Type(err)
            errMsg := HasProp(err, "Message") ? err.Message : String(err)
            throw Error(prefix . "expected exception of type " . expectedName . ", but caught " . actualName . " (" . errMsg . ")")
        }
        return err
    }

    throw Error(prefix . "expected exception to be thrown, but none was thrown")
}

_IsMatchingError(err, expectedType) {
    if (expectedType == "") {
        return true
    }
    if (expectedType is Class) {
        return err is expectedType
    }
    if (expectedType is String) {
        return Type(err) == expectedType
    }
    return false
}

_GetErrorTypeName(expectedType) {
    if (expectedType is Class) {
        return HasProp(expectedType, "Prototype") && HasProp(expectedType.Prototype, "__Class")
            ? expectedType.Prototype.__Class
            : "Class"
    }
    return String(expectedType)
}

RunRegisteredTests() {
    global TestCases
    failedCount := 0

    for testCase in TestCases {
        try {
            testCase.Callback.Call()
            WriteTestLine("PASS " . testCase.Name)
        } catch as err {
            failedCount += 1
            WriteTestLine("FAIL " . testCase.Name)
            WriteTestLine("  " . err.Message)
        }
    }

    totalCount := TestCases.Length
    passedCount := totalCount - failedCount
    WriteTestLine(Format("{1} passed, {2} failed, {3} total", passedCount, failedCount, totalCount))

    if (failedCount > 0) {
        ExitApp 1
    }

    ExitApp 0
}

FormatValue(value) {
    if IsObject(value) {
        return "<object>"
    }
    return '"' . value . '"'
}

WriteTestLine(text) {
    FileAppend(text . "`n", "*", "UTF-8")
}
