#Requires AutoHotkey v2.0

#Include ..\Lib\Paste.v2.ahk
#Include lib\ris-common.v2.ahk

::0sg-ptccd:: {
    MyForm := "
(
PTCCD is indicated and has been arranged.
)"
    Paste(MyForm)
}

::1sg-ptgbd:: {
    currDateStr := FormatTime(, "M/d tt")
    MyForm := Format("
(
PTCCD was performed in {1}. A 8 Fr pigtail drain was inserted. 10 ml of aspirated bile was collected for Lab exam.
)",
        currDateStr)
    Paste(MyForm)
}

:*:0sg-ld::Percutaneous drainage for liver abscess has been arranged.

::0ctg-lb::
{
    MyForm := "
  (
CT-guide lung biopsy is indicated and has been scheduled on / PM. If specimen for tissue culture is needed, please prepare other specimen collecting bottles and send to CT room with the patient. Otherwise, only specimen immersed in formalin will be harvested.
  )"
    Paste(MyForm)
}

::1ctg-lb::
{
    currDateStr := FormatTime(, "M/d tt")
    MyForm := Format("
  (
CT guide lung biopsy was performed in {1}. Please follow up CXR if pneumothorax develops or progresses.
  )",
        currDateStr)
    Paste(MyForm)
}

::1ctg-b::
{
    currDateStr := FormatTime(, "M/d tt")
    MyForm := Format("
  (
CT guide biopsy was performed in {1}. Please keep bed rest and check if internal bleeding occurs.
  )",
        currDateStr)
    Paste(MyForm)
}

::1ctg-d::
{
    currDateStr := FormatTime(, "M/d tt")
    MyForm := Format("
  (
CT guide drainage was performed in {1}. A 8 Fr pigtail drain was inserted. 10 ml of aspirated pus was collected for Lab exam.
  )",
        currDateStr)
    Paste(MyForm)
}

::1xahaic:: {
    currDateStr := FormatTime(, "M/d tt")
    MyForm := Format("
(
The temporary HAIC catheter was placed in {1}. Please keep infusion of the catheter to prevent clotting. If oozing from the puncture area occurs, please check KUB to make sure the catheter tip location is proper. If further TAE with Lipiodol after this HAIC session is needed, please arrange the exam.
)",
        currDateStr)
    Paste(MyForm)
}

::1xadj:: {
    currDateStr := FormatTime(, "M/d tt")
    MyForm := Format("
(
Antegrade double-J ureteral stent placement was performed in {1}. Please keep PCN drainage if hematuria persists. For PCN removal, if needed, please clamp the PCN first, if no discomfort nor fever for hours to a day, arrange antegrade pyelography to check the patency of ureteral stent. If patent, I will remove the PCND at that time.
)",
        currDateStr)
    Paste(MyForm)
}

::1xapcn:: {
    currDateStr := FormatTime(, "M/d tt")
    MyForm := Format("
(
Percutaneous nephrostomy (8F pigtail) was performed in {1}.
)", currDateStr)
    Paste(MyForm)
}

::0xale::Angiography of lower extremity has been arranged.

::1xale:: {
    currDateStr := FormatTime(, "M/d tt")
    MyForm := Format("
(
Aortography and angiography of lower extremity were performed in {1}.
Angiography of lower extremity was performed in {1}.
Please keep external compression on the puncture site and check if bleeding or hematoma occurs.
)",
        currDateStr)
    Paste(MyForm)
}

::1xaptcd:: {
    currDateStr := FormatTime(, "M/d tt")
    MyForm := Format("
(
PTCD was performed in {1}. A 8 Fr pigtail drain was inserted through left / right IHD.
)", currDateStr)
    Paste(MyForm)
}
