; RIS specific functions

SleepThenTab(sleepTime = 400, shiftTab = False)
{
  Sleep %sleepTime%
  If (shiftTab) {
    Send +{Tab}
  } Else {
    Send {Tab}
  }
}

; need to reimplant
GetGenderFromRIS() {

}

GetAgeFromRIS() {

}

ConvertRISDate(inputString) {
  ; 1. 從輸入字串 (yyymmddhhmm) 中提取各個部分
  minguoYear := SubStr(inputString, 1, 3)  ; yyy (例如: 114)
  month := SubStr(inputString, 4, 2)       ; mm (例如: 10)
  day := SubStr(inputString, 6, 2)         ; dd (例如: 15)

  ; 2. 將民國年轉換為西元年 (民國年 + 1911 = 西元年)
  ; AHK v1 會自動將字串 "114" 視為數字 114 進行計算
  gregorianYear := minguoYear + 1911

  ; 3. 組合並返回 yyyy-mm-dd 格式的字串
  ; 使用 . 運算子來連接字串
  outputDate := gregorianYear . "-" . month . "-" . day

  Return outputDate
}