#Requires AutoHotkey v2.0

global TestCases := []

RegisterTest(name, callback) {
    global TestCases
    TestCases.Push({Name: name, Callback: callback})
}

AssertEqual(expected, actual, message := "") {
    if (expected != actual) {
        prefix := message != "" ? message . ": " : ""
        throw Error(prefix . "expected " . FormatValue(expected) . ", got " . FormatValue(actual))
    }
}

AssertTrue(condition, message := "") {
    if (!condition) {
        throw Error(message != "" ? message : "expected condition to be true")
    }
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
