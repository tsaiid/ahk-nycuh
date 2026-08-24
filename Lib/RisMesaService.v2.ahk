#Requires AutoHotkey v2.0

/**
 * 負責向 MESA (Multi-Ethnic Study of Atherosclerosis) 官方線上計算器發送請求並解析結果
 * 網址: https://tools.mesa-nhlbi.org/Calcium/input.aspx
 */
class RisMesaService {
    static CalculatorUrl := "https://tools.mesa-nhlbi.org/Calcium/input.aspx"

    /**
     * 連線至 MESA 計算器計算百分位數與 non-zero 機率
     * @param age 年齡 (45-84)
     * @param sex 性別 "M" 或 "F" (或 "1", "0", "male", "female")
     * @param totalScore Agatston 分數 (0-10000)
     * @param race 種族代碼預設 "1" (Chinese: 1, Black: 0, Hispanic: 2, White: 3)
     * @returns { IsSuccess: bool, IsOutOfRange: bool, Age: int, Gender: string, Race: string, TotalScore: number, Percentile: string, NonZeroProbability: string, Summary: string, Error: string }
     */
    static Query(age, sex, totalScore, race := "1") {
        numAge := IsInteger(age) ? Integer(age) : (RegExMatch(String(age), "(\d+)", &mAge) ? Integer(mAge[1]) : 0)
        numScore := IsNumber(totalScore) ? Number(totalScore) : (RegExMatch(String(totalScore), "(\d+(?:\.\d+)?)", &mScore) ? Number(mScore[1]) : 0)
        scoreStr := (IsInteger(numScore) || numScore == Floor(numScore)) ? String(Integer(numScore)) : RegExReplace(Format("{:.2f}", numScore), "\.?0+$", "")

        ; 性別代碼：0 為 female, 1 為 male
        genderCode := "0"
        if (sex == "M" || sex == "1" || InStr(sex, "Male") || InStr(sex, "男")) {
            genderCode := "1"
        }
        genderLabel := (genderCode == "1") ? "Male" : "Female"

        ; 檢查是否超出 MESA 適用年齡/分數範圍 (45-84 歲, 分數 0-10000)
        if (numAge < 45 || numAge > 84 || numScore < 0 || numScore > 10000) {
            return {
                IsSuccess: true,
                IsOutOfRange: true,
                Age: numAge,
                Gender: genderLabel,
                Race: "Chinese",
                TotalScore: numScore,
                Percentile: "",
                NonZeroProbability: "",
                Summary: "Patient age (" . numAge . ") or score (" . scoreStr . ") is outside MESA reference range (Age 45-84, Score <= 10000).",
                Error: ""
            }
        }

        try {
            ; 1. GET 取得 ASP.NET ViewState 與 EventValidation
            reqGet := ComObject("WinHttp.WinHttpRequest.5.1")
            reqGet.Open("GET", this.CalculatorUrl, false)
            reqGet.SetRequestHeader("User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64)")
            reqGet.Send()

            if (reqGet.Status != 200) {
                throw Error("MESA calculator GET 失敗 (HTTP " . reqGet.Status . ")")
            }

            getHtml := reqGet.ResponseText
            viewState := this._ExtractHiddenField(getHtml, "__VIEWSTATE")
            viewStateGen := this._ExtractHiddenField(getHtml, "__VIEWSTATEGENERATOR")
            eventValidation := this._ExtractHiddenField(getHtml, "__EVENTVALIDATION")

            if (viewState == "") {
                throw Error("無法從 MESA 網頁解析 __VIEWSTATE")
            }

            ; 2. POST 發送計算請求
            postData := "__VIEWSTATE=" . this._UrlEncode(viewState)
                . "&__VIEWSTATEGENERATOR=" . this._UrlEncode(viewStateGen)
                . "&__EVENTVALIDATION=" . this._UrlEncode(eventValidation)
                . "&Age=" . this._UrlEncode(String(numAge))
                . "&gender=" . this._UrlEncode(genderCode)
                . "&Race=" . this._UrlEncode(String(race))
                . "&Score=" . this._UrlEncode(scoreStr)
                . "&Calculate=Calculate"

            reqPost := ComObject("WinHttp.WinHttpRequest.5.1")
            reqPost.Open("POST", this.CalculatorUrl, false)
            reqPost.SetRequestHeader("User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64)")
            reqPost.SetRequestHeader("Content-Type", "application/x-www-form-urlencoded")
            reqPost.Send(postData)

            if (reqPost.Status != 200) {
                throw Error("MESA calculator POST 失敗 (HTTP " . reqPost.Status . ")")
            }

            resHtml := reqPost.ResponseText

            percentile := ""
            if RegExMatch(resHtml, 'i)id="percLabel"[^>]*>([^<]+)<', &mPerc) {
                percentile := Trim(mPerc[1])
            }

            nonZeroProb := ""
            if RegExMatch(resHtml, 'i)id="Label10"[^>]*>([^<]+)<', &mProb) {
                nonZeroProb := Trim(mProb[1], " `t`r`n.%") . "%"
            }

            isOutOfRange := InStr(resHtml, "Out of Range") && (percentile == "")

            summary := ""
            if (percentile != "") {
                summary := "MESA CAC Percentile: " . percentile . "% (Estimated non-zero CAC probability: " . nonZeroProb . " for age " . numAge . ", " . genderLabel . ", Chinese)"
            } else if (isOutOfRange) {
                summary := "Patient is outside MESA reference range (Age 45-84, Score <= 10000)."
            } else if (numScore == 0) {
                summary := "CAC score is 0 (No identifiable calcification; estimated non-zero CAC probability for age " . numAge . ", " . genderLabel . ", Chinese is " . nonZeroProb . ")."
            }

            return {
                IsSuccess: true,
                IsOutOfRange: isOutOfRange,
                Age: numAge,
                Gender: genderLabel,
                Race: "Chinese",
                TotalScore: numScore,
                Percentile: percentile,
                NonZeroProbability: nonZeroProb,
                Summary: summary,
                Error: ""
            }

        } catch as err {
            return {
                IsSuccess: false,
                IsOutOfRange: false,
                Age: numAge,
                Gender: genderLabel,
                Race: "Chinese",
                TotalScore: numScore,
                Percentile: "",
                NonZeroProbability: "",
                Summary: "MESA online calculation unavailable: " . err.Message,
                Error: err.Message
            }
        }
    }

    static _ExtractHiddenField(html, name) {
        if RegExMatch(html, 'i)name="' . name . '"[^>]*value="([^"]*)"', &m) {
            return m[1]
        }
        if RegExMatch(html, 'i)value="([^"]*)"[^>]*name="' . name . '"', &m) {
            return m[1]
        }
        return ""
    }

    static _UrlEncode(str) {
        static doc := 0
        if (!doc) {
            doc := ComObject("HTMLFile")
            doc.write('<meta http-equiv="X-UA-Compatible" content="IE=edge">')
        }
        return doc.parentWindow.encodeURIComponent(str)
    }
}
