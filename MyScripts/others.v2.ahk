
;; Limitations
::olnd::However, the obstruction level cannot be demonstrated in this study.
::motion::* Obvious motion artifacts may limit the interpretation.
::mart::* Obvious metallic artifact may limit the evaluation.
::ubl::(* Limited evaluation due to collapsed UB.)
::gil::(* limited evaluation due to peristalsis, susceptibility artifact from gas, etc.)
::ncl::
 {
  MyForm := "
(
* The evaluation is limited due to absence of contrast enhancement, especially for solid organs and vascular structure.
* The detection of tiny or occult metastasis is limited due to absence of contrast enhancement.
* The detection of tiny or occult residual/recurrent tumor and the evaluation of vascular structure are limited due to absence of contrast enhancement.
)"
  Paste(MyForm)
 }

::sgo::suggestive of `
::obv::obvious `
::fn::FOOTNOTE:{Enter}[{^}1]: `

::ar::
 {
  currDateStr := FormatTime(, "yyyy/M/d")
  MyForm := Format("
(
----
Additional report on {1}:


)", currDateStr)
  Paste(MyForm)
 }

; common hotstrings
::rcs::renal cysts
::rcs1::renal cysts, size up to `
::rss::renal stones
::lrc::A -cm renal cyst at the left kidney.
::rrc::A -cm renal cyst at the right kidney.

; 資源共享
::share::
 {
  MyForm := "
(
The study has been uploaded to our PACS system.
Original report has been attached as a picture file.
For second opinion, please submit a formal consultation request to our department.

IMPRESSION:
The study has been uploaded to our PACS system.
)"
  Paste(MyForm)
 }
