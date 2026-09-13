#Requires AutoHotkey v2.0

#Include TestLib.v2.ahk
#Include ..\Lib\RisController.v2.ahk

RegisterTest("RisController._NormalizePatientSex handles standard sex strings", Test_NormalizePatientSex)
RegisterTest("RisController male rules detect female organs, surgeries, and acronyms", Test_MalePatientDetectsFemaleTerms)
RegisterTest("RisController male rules do not falsely flag cervical spine or normal findings", Test_MalePatientNoFalsePositives)
RegisterTest("RisController female rules detect male organs, surgeries, and acronyms", Test_FemalePatientDetectsMaleTerms)
RegisterTest("RisController female rules do not falsely flag normal findings", Test_FemalePatientNoFalsePositives)

CheckRulesMatch(rules, text) {
    for rule in rules {
        if RegExMatch(text, "i)" . rule.Pattern, &m) {
            return {Matched: true, Term: m[0], Label: rule.Label}
        }
    }
    return {Matched: false, Term: "", Label: ""}
}

Test_NormalizePatientSex() {
    AssertEqual("M", RisController._NormalizePatientSex("男"))
    AssertEqual("M", RisController._NormalizePatientSex("Male"))
    AssertEqual("M", RisController._NormalizePatientSex("M"))
    AssertEqual("M", RisController._NormalizePatientSex("  男  "))

    AssertEqual("F", RisController._NormalizePatientSex("女"))
    AssertEqual("F", RisController._NormalizePatientSex("Female"))
    AssertEqual("F", RisController._NormalizePatientSex("F"))
    AssertEqual("F", RisController._NormalizePatientSex("  女  "))

    AssertEqual("", RisController._NormalizePatientSex(""))
    AssertEqual("", RisController._NormalizePatientSex("Unknown"))
}

Test_MalePatientDetectsFemaleTerms() {
    maleRules := RisController._GetSexSpecificTermRules("M")

    sampleTexts := [
        "Small uterine myoma noted.",
        "Right ovarian cyst.",
        "Bilateral fallopian tube dilation.",
        "Right adnexal cystic lesion.",
        "Thickened endometrium.",
        "Intramural myometrial mass.",
        "Suspected cervical carcinoma.",
        "Status post hysterectomy.",
        "s/p bilateral oophorectomy.",
        "Left salpingectomy noted.",
        "s/p salpingo-oophorectomy.",
        "History of myomectomy.",
        "Radical trachelectomy.",
        "Tubal ligation clips in place.",
        "s/p anterior colporrhaphy.",
        "Partial vaginectomy.",
        "Prior vulvectomy.",
        "Status post TAH-BSO.",
        "s/p TVH.",
        "s/p TLH.",
        "s/p LAVH.",
        "s/p BSO.",
        "Left USO.",
        "s/p D&C for abnormal bleeding.",
        "Vaginal vault recurrence.",
        "Vulvar mass."
    ]

    for text in sampleTexts {
        result := CheckRulesMatch(maleRules, text)
        AssertTrue(result.Matched, "Male rule should detect female term in: " . text)
    }
}

Test_MalePatientNoFalsePositives() {
    maleRules := RisController._GetSexSpecificTermRules("M")

    safeTexts := [
        "Degenerative spondylosis of cervical spine.",
        "Cervical lymphadenopathy without acute infection.",
        "Normal liver, spleen, pancreas, and bilateral kidneys.",
        "No evidence of distant metastasis or lymph node enlargement."
    ]

    for text in safeTexts {
        result := CheckRulesMatch(maleRules, text)
        AssertFalse(result.Matched, "Male rule should NOT flag safe text: " . text)
    }
}

Test_FemalePatientDetectsMaleTerms() {
    femaleRules := RisController._GetSexSpecificTermRules("F")

    sampleTexts := [
        "Enlarged prostate gland with calcification.",
        "Status post radical prostatectomy.",
        "s/p TURP for urinary obstruction.",
        "s/p TUR-P with bladder neck contracture.",
        "s/p HoLEP.",
        "s/p ThuLEP.",
        "s/p TUIP.",
        "Impression: 1. BPH.",
        "Normal bilateral seminal vesicles.",
        "Normal right testis.",
        "s/p right orchiectomy.",
        "Left orchidectomy.",
        "History of bilateral orchiopexy.",
        "Scrotal swelling with fluid collection.",
        "Penile shaft lesion.",
        "Circumcision performed years ago.",
        "Right epididymal head cyst.",
        "Calcification along vas deferens.",
        "s/p bilateral vasectomy.",
        "Left spermatic cord lipoma.",
        "s/p varicocelectomy.",
        "s/p right hydrocelectomy."
    ]

    for text in sampleTexts {
        result := CheckRulesMatch(femaleRules, text)
        AssertTrue(result.Matched, "Female rule should detect male term in: " . text)
    }
}

Test_FemalePatientNoFalsePositives() {
    femaleRules := RisController._GetSexSpecificTermRules("F")

    safeTexts := [
        "Normal liver and gallbladder.",
        "No hydronephrosis or renal stone.",
        "Bilateral lungs are clear without active consolidation.",
        "No acute intracranial hemorrhage or territorial infarction."
    ]

    for text in safeTexts {
        result := CheckRulesMatch(femaleRules, text)
        AssertFalse(result.Matched, "Female rule should NOT flag safe text: " . text)
    }
}

RunRegisteredTests()
