#Requires AutoHotkey v2.0

#Include TestLib.v2.ahk
#Include ..\Lib\RisAIText.v2.ahk

RegisterTest("RisAIText.EscapeJsonString escapes special characters", Test_EscapeJsonString)
RegisterTest("RisAIText.DecodeJsonEscapedText unescapes special and unicode characters", Test_DecodeJsonEscapedText)
RegisterTest("RisAIText.StripMarkdownCodeFence handles fenced code and inline backticks", Test_StripMarkdownCodeFence)
RegisterTest("RisAIText.ExtractGoogleResponseText extracts text and skips thought blocks", Test_ExtractGoogleResponseText)
RegisterTest("RisAIText.ExtractOpenAIResponseText extracts output_text and throws on empty", Test_ExtractOpenAIResponseText)
RegisterTest("RisAIText.ParseGoogleResponse end-to-end integration", Test_ParseGoogleResponse)
RegisterTest("RisAIText.ParseOpenAIResponse end-to-end integration", Test_ParseOpenAIResponse)

Test_EscapeJsonString() {
    input := 'Line1' . "`r`n" . 'Line2' . "`t" . '"quoted\"'
    expected := 'Line1\r\nLine2\t\"quoted\\\"'
    AssertEqual(expected, RisAIText.EscapeJsonString(input), "Special characters should be escaped")
    AssertEqual("simple text", RisAIText.EscapeJsonString("simple text"), "Normal text unchanged")
}

Test_DecodeJsonEscapedText() {
    input := 'Line1\r\nLine2\t\"quoted\\\" \u4e2d\u6587'
    expected := 'Line1' . "`r`n" . 'Line2' . "`t" . '"quoted\" 中文'
    AssertEqual(expected, RisAIText.DecodeJsonEscapedText(input), "JSON escaped characters and Unicode should be decoded")
}

Test_StripMarkdownCodeFence() {
    fencedWithLang := "``````markdown`r`n1. First finding`r`n2. Second finding`r`n``````"
    AssertEqual("1. First finding`r`n2. Second finding", RisAIText.StripMarkdownCodeFence(fencedWithLang), "Should strip triple backticks with language tag")

    fencedNoLang := "```````nPlain text content`n``````"
    AssertEqual("Plain text content", RisAIText.StripMarkdownCodeFence(fencedNoLang), "Should strip triple backticks without language tag")

    inlineTick := "``inline text``"
    AssertEqual("inline text", RisAIText.StripMarkdownCodeFence(inlineTick), "Should strip single backticks")

    plain := "  hello world  `t`r`n"
    AssertEqual("hello world", RisAIText.StripMarkdownCodeFence(plain), "Should trim whitespace")
}

Test_ExtractGoogleResponseText() {
    rawGoogleJson := '
    (
    {
      "candidates": [
        {
          "content": {
            "parts": [
              {
                "text": "Internal chain of thought",
                "thought": true
              },
              {
                "text": "Actual impression part 1 "
              },
              {
                "text": "and part 2."
              }
            ]
          }
        }
      ]
    }
    )'
    extracted := RisAIText.ExtractGoogleResponseText(rawGoogleJson)
    AssertEqual("Actual impression part 1 and part 2.", extracted, "Should skip thought and combine parts")

    invalidJson := '{"candidates": []}'
    AssertThrows(Error, () => RisAIText.ExtractGoogleResponseText(invalidJson), "Should throw when no text found")
}

Test_ExtractOpenAIResponseText() {
    rawJson1 := '
    (
    {
      "output": [
        {
          "type": "output_text",
          "text": "OpenAI impression line 1"
        }
      ]
    }
    )'
    extracted1 := RisAIText.ExtractOpenAIResponseText(rawJson1)
    AssertEqual("OpenAI impression line 1", extracted1, "Should extract text with type output_text")

    rawJson2 := '{"output_text": "Direct output text"}'
    extracted2 := RisAIText.ExtractOpenAIResponseText(rawJson2)
    AssertEqual("Direct output text", extracted2, "Should fallback to output_text field")

    invalidJson := '{"error": "no text"}'
    AssertThrows(Error, () => RisAIText.ExtractOpenAIResponseText(invalidJson), "Should throw when no output text found")
}

Test_ParseGoogleResponse() {
    rawGoogleJson := '
    (
    {
      "candidates": [
        {
          "content": {
            "parts": [
              {
                "text": "``````markdown\n1. No acute ICH.\n2. Normal CT.\n``````"
              }
            ]
          }
        }
      ]
    }
    )'
    parsed := RisAIText.ParseGoogleResponse(rawGoogleJson)
    AssertEqual("1. No acute ICH.`n2. Normal CT.", parsed, "Should extract, unescape, and strip markdown code fence")
}

Test_ParseOpenAIResponse() {
    rawJson := '{"output_text": "``````\nIndication: \u80ba\u708e\n``````"}'
    parsed := RisAIText.ParseOpenAIResponse(rawJson)
    AssertEqual("Indication: 肺炎", parsed, "Should decode unicode and strip code fence")
}

RunRegisteredTests()
