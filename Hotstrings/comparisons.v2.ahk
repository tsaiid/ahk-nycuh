; Comparisons
; need other string and date lib
#Include ..\Lib\RisController.v2.ahk
#Include ..\Lib\Paste.v2.ahk

StringWithPrevExamDate(baseText)
{
    ; 1. 問 RisController 要後綴
    suffix := RisController.GetComparisonSuffix()

    ; 2. 組合 (注意：如果 suffix 有值，前面會自動帶有 " dated ...")
    finalText := baseText . suffix . "."

    ; 3. 貼上
    Paste(finalText)
}

::nic:: {
    StringWithPrevExamDate("No obvious interval changes compared to the previous study")
}

::nip:: {
    StringWithPrevExamDate("No obvious improvement compared to the previous study")
}

::mip:: {
    StringWithPrevExamDate("Mild improvement compared to the previous study")
}

::pc:: {
    StringWithPrevExamDate("Progressive changes compared to the previous study")
}

::mpc:: {
    StringWithPrevExamDate("Mild progressive changes compared to the previous study")
}

::rc:: {
    StringWithPrevExamDate("Regressive changes compared to the previous study")
}

::mrc:: {
    StringWithPrevExamDate("Mild regressive changes compared to the previous study")
}