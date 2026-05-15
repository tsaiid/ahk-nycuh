#Requires AutoHotkey v2.0

#Include ..\Lib\RisController.v2.ahk
#Include ..\Lib\Paste.v2.ahk
#Include lib\ris-common.v2.ahk

; My Neuro Forms

;; Templates
;#Include MyScripts\neuro-lai.ahk
#Include neuro\neuro-brain.v2.ahk
#Include neuro\neuro-neck.v2.ahk
#Include neuro\neuro-orbital.v2.ahk
#Include neuro\neuro-spine.v2.ahk

;; Hotstrings
::ath::Atherosclerotic changes with calcification of intracranial portion of vertebrobasilar arteries and bilateral internal carotid arteries.
::athc::Atherosclerotic changes with calcification of intracranial portion of bilateral internal carotid arteries.
::athv::Atherosclerotic changes with calcification of intracranial portion of vertebrobasilar arteries.
::math::Mild atherosclerotic changes with calcification of intracranial portion of vertebrobasilar arteries and bilateral internal carotid arteries.
::mathc::Mild atherosclerotic changes with calcification of intracranial portion of bilateral internal carotid arteries.
::mathv::Mild atherosclerotic changes with calcification of intracranial portion of vertebrobasilar arteries.
::sae::Bilateral periventricular low density presents. Subcortical arteriosclerotic encephalopathy (leukoaraiosis) is considered.
::msae::Mild bilateral periventricular low density presents. Mild subcortical arteriosclerotic encephalopathy (leukoaraiosis) is considered.
::sae2::Presence of mild/moderate/severe confluent symmetric periventricular hyperintensity on T2WI and FLAIR suggests subcortical arteriosclerotic encephalopathy (leukoaraiosis).
::msae2::Presence of mild confluent symmetric periventricular hyperintensity on T2WI and FLAIR suggests mild subcortical arteriosclerotic encephalopathy (leukoaraiosis).
::sae3::Presence of mild confluent symmetric periventricular hyperintensity on T2WI and FLAIR noted, the subcortical arteriosclerotic encephalopathy (leukoaraiosis) considered. Several tiny hyperintensities in the bilateral subcortical and deep white matter regions on T2WI and FLAIR, which may be gliosis, demyelination or tiny old ischemia or tiny previous brain insult.
::sae4::Presence of bilateral confluent periventricular, and several small deep and subcortical white matter and pons hyperintensities on T2WI and FLAIR suggests subcortical arteriosclerotic encephalopathy (leukoaraiosis).
::ubo::Several nonspecific tiny hyperintensities in the bilateral subcortical and deep white matter regions on T2WI and FLAIR (unidentified bright objects).
::ubo1::Presence of several tiny hyperintensities in the periventricular white matter regions on T2WI and FLAIR, which may be gliosis, demyelination or tiny old ischemia or tiny previous brain insult.
::ubo2::Leukoaraiosis (some tiny/small hyperintensities on T2WI and FLAIR image in the periventricular and subcortical white matter regions) are mostly due to aging process and/or small vessel ischemic disease.
::ifch::from the imaging findings and clinical history,
::mrmast::Increased signal intensity in the -------------, mastoid air sinus on T2WI, in favor of mastoiditis or fluid collection in the mastoid.
::ctmast::Presence of soft tissue density in the ---- mastoid air sinus, R/O mastoiditis or fluid collection in the mastoid.
::lmast::Presence of soft tissue density in the left mastoid air sinus, R/O mastoiditis or fluid collection in the mastoid.
::rmast::Presence of soft tissue density in the right mastoid air sinus, R/O mastoiditis or fluid collection in the mastoid.
::ctcmast::Underdeveloped pneumatization and sclerotic changes of right/left/bilateral mastoid air cells, suspicious chronic mastoiditis.
::atrv::Diffuse atrophy of bilateral cerebral hemispheres, with compensatory dilatation of bilateral lateral ventricles, deepening and widening of cortical sulci.
::matrv::Mild enlargement of the ventricular system, in favor of mild brain atrophy.
::atrvs::Enlargement of the ventricular system and sulcal widening of bilateral cerebral hemispheres noted, in favor of brain atrophy.
::matrvs::Mild enlargement of the ventricular system and sulcal widening of bilateral cerebral hemispheres noted, in favor of mild brain atrophy.
::atrssa::Sulcal widening of bilateral cerebral hemispheres and enlargement of the subarachnoid space noted, in favor of mild brain atrophy.
::atrvsa::Mild enlargement of the intraventricular system with widening of the subarachnoid space of bilateral cerebral hemispheres, in favor of aging process and mild brain atrophy.
::atrs::Atrophy of bilateral cerebral cortices, with deepening and widening of fissures and cortical sulci.
::matrs::Mild atrophy of bilateral cerebral cortices, with deepening and widening of fissures and cortical sulci.
::matrs2::Mild brain atrophy with enlarged subarachnoid spaces of bilateral cerebral convexities.
::ctabi::A small ill-defined low density in the left frontal subcortical white matter, could be old or recent ischemic infarction. Clinical correlation is suggested.
::ctpbi::Presence of tiny/small low density involved ___ basal ganglion, ___ thalamus, and ___ periventricular white matter, previous brain ischemic insult, such as old tiny ischemic infarction considered.
::ctpbi2::Presence of brain tissue loss changes involving right temporal, right parietal, and left frontal regions, with compensatory dilatation of left lateral ventricle, in favor of previous brain insult, such as old ischemic infarction.
::mrpbi::Presence of several tiny hyperintensities in the periventricular white matter regions on T2WI and FLAIR, which may be gliosis, demyelination or tiny old ischemia or tiny previous brain insult.
::ctepvs::Small low density near right/left side of anterior commissure, in favor of enlarged perivascular space.
::mrepvs::Some tiny T2 hyperintensity spots in both basal ganglia, in favor of enlarged perivascular spaces.
::necsprt::Diffuse thickening and enhancement over the epiglottis, arytenoepiglottic folds, and posterior pharyngeal wall over the oropharynx and hypopharynx, c/w post-radiation changes.
::mrnecsprt::Presence of high signal intensity on T1WI over the C-spine, compatible with post radiation changes.
::mrnecspc::Mild mucosal and submucosal soft tissue thickening over the nasopharyngeal and oropharyngeal regions, in favor of post-treatment changes. Follow-up is suggested.
::mrns::No evidence of high signal intensity on DWI and lower apparent diffusion coefficient suggesting acute ischemia infarction in the brain and brainstem noted.
::rtpaok::No focal hypodensity or early ischemic change is identified. ASPECTS: 10.
::noaa::No evidence of aneurysm or arterial-venous malformation (AVM) noted near the circle of Willis regions.
::sdsa::Suggest correlate with DSA if clinically indicated.
::bbgt1h::Symmetrical T1-hyperintensity involving the bilateral globus pallidus, cerebral peduncles, and the dorsal aspect of pons. These areas show no obvious abnormal signal intensity on T2WI. Chronic hepatic encephalopathy is suspected. Suggest clinical correlation if chronic liver disease presents. DDx (less likely): hyperalimentation, Wilson disease, hyper-/hypoparathyroidism.
::pfsdh::Presence of fusiform high-attenuation lesion extending over the right anterior falx cerebri, suggestive of an acute parafalcine subdural hematoma.
::lfpca::left fetal type posterior cerebral artery
::rfpca::right fetal type posterior cerebral artery
::bfpca::bilateral fetal type posterior cerebral arteries
::psok::
{
    MyForm := "
  (
The visible paranasal sinuses are clear.
The paranasal sinuses are clear.
The visible paranasal sinuses and mastoids are unremarkable.
  )"
    Paste(MyForm)
}

::rfvps::status post ventriculo-peritoneal shunt from right frontal area, with tip at left lateral ventricle.
::lgwd::loss of gray-white matter differentiation
::mdr::midline deviation to right side.
::mdl::midline deviation to left side.
::rnth::Right nasal turbinate hypertrophy.
::lnth::Left nasal turbinate hypertrophy.
::bnth::Bilateral nasal turbinate hypertrophy.
::ctmcm::Prominent retrocerebellar cerebrospinal fluid space with normal vermis, 4th ventricle, and cerebellar hemispheres, in favor of mega cisterna magna.
::epc::endplate changes
::ctnph::Enlarged lateral and third ventricles, with relatively normal 4th ventricle. Ventricular enlargement out of proportion to cortical sulcal enlargement, and presence of bilateral periventricular low density. Normal pressure hydrocephalus may be suspected. DDx: normal aging brain.
::ctnph1::However, as the ventricular enlargement is slightly out of proportion to cortical sulcal enlargement, superimposed communicating hydrocephalus cannot be excluded. Clinical correlation is suggested.
::rposts::Soft tissue swelling in the right periorbital region.
::mrposts::Mild soft tissue swelling in the right periorbital region.
::lposts::Soft tissue swelling in the left periorbital region.
::mlposts::Mild soft tissue swelling in the left periorbital region.
::f-p::fronto-parietal `
::f-t-p::fronto-temporo-parietal `
::f-t::fronto-temporal `
::p-o::parieto-occipital `
::t-p::temporo-parietal `
::t-o::temporo-occipital `
::f-::frontal `
::p-::parietal `
::t-::temporal `
::o-::occipital `
::rd::restricted diffusion
::dds::disc desiccation
::dbd::diffuse bulging disc
::dpd::diffuse protrusion disc
::dhd::diffuse herniation disc
::bgd::disc bulging
::prd::disc protrusion
::hrd::disc herniation
::postcen::posterocentral `
::postlat::posterolateral `
::lpl::left posterolateral `
::rpl::right posterolateral `
::lfh::ligament flavum hypertrophy
::fjh::facet joint hypertrophy
::uvjh::uncovertebral joint hypertrophy
::retr::retrolisthesis
::dwiok::No evidence of high signal intensity on DWI and lower apparent diffusion coefficient suggesting acute ischemia infarction in the brain and brainstem noted.
::dvsok::The major dural venous sinuses are patent.
::mrnocva::No evidence of high signal intensity on DWI and lower apparent diffusion coefficient suggesting acute ischemia infarction in the brain and brainstem noted.
::ctnocva::No definite low density or loss of gray-white matter differentiation in the brain parenchyma.
::nocva::No evidence of acute ischemia infarction in the brain and brainstem noted.
::noich::No definite acute intracranial parenchymal hemorrhage, subarachnoid hemorrhage, epidural or subdural hematoma in the brain noted.
::noich0::No definite acute ICH, SAH, EDH, or SDH.
::lka::leukoaraiosis
::mlka::mild leukoaraiosis
::oli::old lacunar infarcts
::oii::old ischemic infarcts
::ali::acute lacunar infarcts
::aii::acute ischemic infarctss
::gpcal::Calcifications are noted at the bilateral globus pallidi, probably aging related.
::mgpcal::Mild calcifications are noted at the bilateral globus pallidi, probably aging related.
::gpdncal::Calcifications are noted at the bilateral globus pallidi and dentate nuclei, probably aging related.
::mgpdncal::Mild calcifications are noted at the bilateral globus pallidi and dentate nuclei, probably aging related.
::mpt::mucoperiosteal thickening
::npok::Symmetrical soft tissue thickening of the nasopharynx, in favor of adenoidal hyperplasia.
::c1::Cavum septum pellucidum.
::c2::Cavum septum pellucidum and cavum vergae.
::c3::Cavum veli interpositi.
::mrbccvm::A 1.3-cm nodular lesion at right frontal subcortical region, with a popcorn appearance, a rim of signal loss, and blooming effect on T2*WI. The postcontrast images show mild central enhancement. A cavernoma (cavernous venous malformation) is suspected.
::lv::lateral ventricle
::3v::third ventricle
::4v::fourth ventricle
::casc::cerebral atherosclerosis
::mcasc::mild cerebral atherosclerosis
::necmraok::The bilateral extracranial carotid and vertebral arteries are unremarkable.
::cbnic::The intracranial condition shows no obvious changes compared to the previous study.
::mriess::The sella is slightly enlarged, the pituitary gland is flattened against the sellar floor, and there is downward traction of chiasm. Empty sella syndrome may be suspected.
::msr9::midline shift to right, about 9 mm at the midthalamic level.
::msl9::midline shift to left, about 9 mm at the midthalamic level.
::msr::midline shift to right
::msl::midline shift to left
::noms::No midline shift is noted.
::sgh::subgaleal hematoma
::ntci::No definite traumatic intracranial injury.
::nacf::No definite acute intracranial findings.
::csptrok::No obvious fracture or dislocation of the cervical spine.
::icr::intracranial
::mritumorcpok::No evidence of abnormal tumor mass lesion over the skull base, bilateral CP angle cistern, and bilateral pre-pontine cistern region noted.
::riol::Status post right intraocular lens (IOL) implantation.
::liol::Status post left intraocular lens (IOL) implantation.
::biol::Status post bilateral intraocular lens (IOL) implantations.
::sba::Presence of skull base artifact with superimpose of bil. frontal base, bil. temporal base, and posterior fossa.
::lioso::Hyperdense material in the left vitreous chamber, highly suggestive of status post vitrectomy with silicone oil injection.
::rioso::Hyperdense material in the right vitreous chamber, highly suggestive of status post vitrectomy with silicone oil injection.

::li1::
{
    MyForm := "
  (
A small low density in the left basal ganglion region, in favor of old lacunar infarct.
A small low density in the right basal ganglion region, in favor of old lacunar infarct.
A small old lacunar infarct in the left basal ganglion.
A small old lacunar infarct in the right basal ganglion.
  )"
    Paste(MyForm)
}

::lis::
{
    MyForm := "
  (
Several small low densities in the bilateral basal ganglion regions, in favor of old lacunar infarcts.
Several tiny low densities in the left basal ganglion region, in favor of old lacunar infarcts.
Several tiny low densities in the right basal ganglion region, in favor of old lacunar infarcts.
Presence of tiny old ischemic infarction involving bilateral basal ganglion region.
Old lacunar infarcts in the bilateral basal ganglia.
  )"
    Paste(MyForm)
}

GetUnremarkableNeckFindings(searchText) {
    findings := []

    ; 1. Pharynx/Larynx
    if (!HasPositiveFinding(searchText, ["nasopharynx", "oropharynx", "hypopharynx", "larynx", "epiglottis", "vocal cord", "pharyngeal"])) {
        findings.Push("The nasopharynx, oropharynx, hypopharynx, and larynx are unremarkable.")
    }

    ; 2. Lymph Nodes
    if (!HasPositiveFinding(searchText, ["lymph node", "LN", "LAP", "lymphadenopathy", "submandibular", "submental", "carotid space", "posterior cervical", "supraclavicular"])) {
        findings.Push("No bulky lymph nodes in the bilateral submandibular and submental, carotid, posterior cervical spaces, and supraclavicular fossae noted.")
    }

    ; 3. Glands
    if (!HasPositiveFinding(searchText, ["parotid", "submandibular gland", "thyroid"])) {
        findings.Push("No particular findings of parotid gland, submandibular gland, and thyroid gland.")
    }

    ; 4. Sinuses/Mastoid
    if (!HasPositiveFinding(searchText, ["sinus", "maxillary", "frontal", "ethmoid", "sphenoid", "mastoid"])) {
        findings.Push("The paranasal sinuses and mastoid air cells are clear.")
    }

    ; 5. Brain/Lungs
    if (!HasPositiveFinding(searchText, ["brain", "lung"])) {
        findings.Push("The visible brain and lungs show no remarkable findings.")
    }

    return findings
}

::necok::
{
    ; 1. 取得文本
    searchText := RisController.GetFindingSearchText()

    ; 2. 呼叫 Helper 取得未提及之檢查項
    findings := GetUnremarkableNeckFindings(searchText)

    ; 3. 輸出
    if (findings.Length > 0) {
        output := ""
        loop findings.Length {
            output .= findings[A_Index] . (A_Index == findings.Length ? "" : "`n")
        }
        Paste(output)
    }
}

::neclapok::
{
    MyForm := "
  (
No bulky lymph nodes in the bilateral submandibular and submental, carotid, posterior cervical spaces, and supraclavicular fossae noted.
  )"
    Paste(MyForm)
}
::neclapok0::No bulky neck lymphadenopathy.

::ctps::
{
    MyForm := "
  (
Mild mucoperiosteal thickening and soft tissue density in the bilateral frontal, bilateral ethmoid, bilateral maxillary, and bilateral sphenoid sinuses, in favor of mild sinusitis.
  )"
    Paste(MyForm)
}
::mctps::Mild mucoperiosteal thickening and soft tissue density in the paranasal sinuses, in favor of mild sinusitis.
::mctps1::Small soft tissue density in the paranasal sinuses, in favor of mild sinusitis.

::mrps::
{
    MyForm := "
  (
Presence of hyperintensity on T2WI in the right/left/bilateral sphenoid, right/left/bilateral ethmoid, right/left/bilateral maxillary, right/left/bilateral frontal sinuses, in favor of mild sinusitis.
  )"
    Paste(MyForm)
}
::mmrps::
{
    MyForm := "
  (
Mild mucoperiosteal thickening and hyperintensity on T2WI in the paranasal sinuses, in favor of mild sinusitis.
  )"
    Paste(MyForm)
}

::bctaok::
{
    MyForm := "
  (
Pre and post-contrast CTA scan of brain:

Brain CT:
- No definite area of obvious abnormal density in the brain noted.
- No definite acute intracranial parenchymal hemorrhage or acute epidural or acute subdural hematoma in the brain noted.
- The bilateral lateral ventricles show symmetrical without dilatation.
- No definite bony lesion in the skull.
- The visible paranasal sinuses and mastoids are unremarkable.

Brain CTA:
- Bilateral CCA, carotid bulbs and distal ICAs are patent.
- Bilateral intracranial VA, and BA are patent.
- Major branches of bilateral ACA, MCA and PCA are patent.
- No definite vascular anomaly, such as aneurysm, AVM or vessel spasm are noted.
- The major dural venous sinuses are patent.

Multiphase CTA:
- No abnormal collateral or perfusion defect is noted.
  )"
    Paste(MyForm)
}

::mraok0::MRA shows no significant vascular stenosis in the major intracranial arteries or aneurysm near the circle of Willis regions.
::mraok::
{
    MyForm := "
  (
MRA shows no significant vascular stenosis in the major intracranial arteries or aneurysm near the circle of Willis regions.

MRA shows no evidence of aneurysm or arterial-venous malformation (AVM) noted near the circle of Willis regions.

MRA shows no evidence of occlusion or high grade stenosis in the intracranial portion of the internal carotid artery and basilar artery, and main trunk of the bilateral anterior cerebral arteries and middle cerebral arteries.

No significant vascular stenosis in the major intracranial arteries.

No significant vascular stenosis in the major intracranial arteries or aneurysm near the circle of Willis regions.
  )"
    Paste(MyForm)
}

::ctp1::
{
    MyForm := "
  (
- Mild mucoperiosteal thickening and soft tissue density in the bilateral frontal, bilateral ethmoid, bilateral maxillary, and bilateral sphenoid sinuses, in favor of mild sinusitis.
- Mild nasal septum deviation to left side.
- Bilateral nasal turbinate hypertrophy.
- Bilateral palatine tonsil enlargement.
- Tiny calcifications in the bilateral palatine tonsils.
- The bony structure is unremarkable.
- The mastoid air cells are well-aerated.
- Symmetric nasopharynx without mass lesion.
  )"
    RisController.PasteToFinding(MyForm)

    MyForm := "
  (
C/W chronic paranasal sinusitis.
  )"
    RisController.PasteToImpression(MyForm)
}

::ctpok::
{
    MyForm := "
  (
- The bilateral frontal, bilateral ethmoid, bilateral maxillary, and bilateral sphenoid sinuses are clear.
- Mild nasal septum deviation to right side.
- Bilateral nasal turbinate hypertrophy.
- The bony structure is unremarkable.
- The mastoid air cells are well-aerated.
  )"
    RisController.PasteToFinding(MyForm)

    MyForm := "
  (
No evidence of paranasal sinusitis.
  )"
    RisController.PasteToImpression(MyForm)
}

::mriiac::
{
    MyForm := "
  (
MRI of internal auditory canal with:
- T2 SPACE axial, coronal, oblique sagittal
- T1WI axial, coronal
- T2WI axial, coronal, sagittal
MRI of whole brain with:
- T2WI axial
- 3D T1+C axial, coronal, sagittal

COMPARISON: no

FINDINGS:
- Normal appearance of bilateral cochlea, vestibules, semicircular canals, and IACs.
- Normal appearance of bilateral vestibulocochlear nerves.
- Normal appearance of bilateral vestibular aqueducts.

- Normal appearance of bilateral mastoid air cells.
- Increased signal intensity in the left mastoid air sinus on T2WI, in favor of mastoiditis or fluid collection in the mastoid.

- No definite abnormal signal intensity tumor mass lesion in the brain noted including supratentorial cerebral hemisphere and infratentorial cerebellum and brain stem region.
- No definite abnormal signal intensity changes in the brain parenchyma.
- The bilateral lateral ventricles showed symmetrical without dilatation.

- Post contrast study shows no abnormal enhancing lesion in the brain and no abnormal leptomeningeal enhancement noted.
- The major dural venous sinuses are patent.
  )"
    RisController.PasteToFinding(MyForm)

    MyForm := "
  (
1. No evidence of cochlear aplasia, cochlear nerve agenesis, or other malformations.
2. Left mastoiditis or fluid collection.
  )"
    RisController.PasteToImpression(MyForm)
}

::mricvaok::
{
    MyForm := "
  (
No evidence of high signal intensity on DWI and lower apparent diffusion coefficient suggesting acute ischemia infarction in the brain and brainstem noted.

No definite abnormal signal intensity changes or tumor mass lesion in the brain noted.

The bilateral lateral ventricles showed symmetrical without dilatation.

The visible paranasal sinuses, mastoids and skull base are unremarkable.

MRA shows no significant vascular stenosis in the major intracranial arteries or aneurysm near the circle of Willis regions.
  )"
    RisController.PasteToFinding(MyForm)

    MyForm := "
  (
No evidence of high signal intensity on DWI suggesting acute or recent ischemia infarction in the brain noted.
No remarkable intracranial findings.
  )"
    RisController.PasteToImpression(MyForm)
}

::mrbok0::Post contrast study shows no abnormal enhancing lesion in the brain and no abnormal leptomeningeal enhancement noted.
::mrbok::
{
    MyForm := "
  (
No evidence of high signal intensity on DWI and lower apparent diffusion coefficient suggesting acute ischemia infarction in the brain and brainstem noted.

No definite abnormal signal intensity or tumor mass lesion in the brain noted.

The bilateral lateral ventricles show symmetrical without dilatation.

Post contrast study shows no abnormal enhancing lesion in the brain and no abnormal leptomeningeal enhancement noted.

The major dural venous sinuses are patent.

The visible paranasal sinuses, mastoids and skull base are unremarkable.

MRA shows no significant vascular stenosis in the major intracranial arteries or aneurysm near the circle of Willis region.

  )"
    RisController.PasteToFinding(MyForm)

    MyForm := "
  (
No remarkable intracranial findings.
No evidence of intracranial metastasis.
  )"
    RisController.PasteToImpression(MyForm)
}

::mriich::
{
    MyForm := "
  (
The MR of the brain performed with Sagittal T1WI
Axial T1WI, T2WI, GET2*WI, FLAIR (Fluid Attenuated Inversion Recovery)
Diffusion-weighted Imaging (DWI) and Apparent Diffusion Coefficient (ADC) map
Intracranial MRA with 3D TOF and focus on the circle of Willis showed:

Previous MRI of brain: none

Findings:
- A 5.9 x 3 x 3.7 cm hematoma, in subacute stage, with peripheral hemosiderin deposition, in the left temporo-parieto-occipital area.
- No evidence of mass lesion nearby that may be contribute to the ICH. (*noncontrast-enhanced study has lower sensitivity for subtle lesion)
- No evidence of microhemorrhage nor cortical superficial siderosis on the GET2*WI, that suggestive of cerebral amyloid angiopathy.
- MRA shows no evidence of aneurysm or arterial-venous malformation (AVM) noted near the circle of Willis regions.

- No other high signal intensity on DWI and lower apparent diffusion coefficient suggesting acute ischemia infarction in the brain and brainstem noted.
- Presence of mild confluent symmetric periventricular hyperintensity on T2WI and FLAIR noted, the subcortical arteriosclerotic encephalopathy considered.
- Presence of small hyperintensities in the right occipital cortical region on T2WI and FLAIR, which may be gliosis, or old ischemia or previous brain insult.
- Mild dilatation of intraventricular system with widening of subarachnoid space of bilateral cerebral hemispheres, in favor of aging process and mild brain atrophy.
- Incidental finding of right side fetal type posterior cerebral artery (PCA) from right internal carotid artery (ICA) with hypoplasia changes at P1 segment of right PCA noted.
- Presence of asymmetry of bilateral vertebral arteries, and more prominence over the right side, and patent flow of bilateral vertebral arteries noted, normal variation in favor.
  )"
    RisController.PasteToFinding(MyForm)

    MyForm := "
  (
Subacute ICH in the left temporo-parieto-occipital area. No definite etiology identified.
  )"
    RisController.PasteToImpression(MyForm)
}

;; MR Brachial Plexus
::mribp::
{
    MyForm := "
  (
MRI of Brachial Plexus:
- Cor T1WI, T2WI, T2WI+FS, T2 SPACE
- Sag T2WI
MRI of C-spine:
- Sag T2WI
- Axi T1WI, T2WI

COMPARISON: No

FINDINGS:
The bilateral brachial plexus are symmetrical with no abnormal signal intensity changes.

Multiple muscle groups show edematous changes, including bilateral rotator cuffs, extensors and flexors of left arm, c/w rhabdomyolysis.

C2-C3: postero-central protrusion disc, causing mild indentation of the anterior dural sac.
C3-C4: diffuse bulging disc, spondylosis, with ligamentum flavum hypertrophy, causing moderate to severe spinal stenosis.
C4-C5: diffuse bulging disc, spondylosis, with ligamentum flavum hypertrophy, causing moderate spinal stenosis.
C5-C6: diffuse bulging disc, causing mild indentation of the anterior dural sac.
C6-C7: diffuse bulging disc, spondylosis, disc space narrowing, endplate changes, with ligamentum flavum hypertrophy, causing moderate spinal stenosis.
C7-T1: diffuse bulging disc, causing mild spinal stenosis.

No abnormal signal intensity changes within the cervical spinal cord.
  )"
    Paste(MyForm)
}

;; Neuro CTA - PCU
::nctapcu::
{
    MyForm := "
  (
CTA of the neck and brain was performed before and after IV contrast agent administration
Scanning range: aortic arch to cranial vault.
Axial, 3D MPR, MIP (and VRT) images:

COMPARISON: nil

FINDINGS:
- No abnormal density or space-occupying lesion in the brain parenchyma.
- No abnormal enhancing lesion in the brain.
- The major dural venous sinuses are patent.
- Cavum septum pellucidum and cavum vergae.
- No abnormal dilatation of the ventricular system and no midline shift.

- Major intracranial arteries, including anterior, middle and posterior cerebral arteries, bilateral internal carotid arteries, basilar arteries, and bilateral vertebral arteries, are patent. No evidence of luminal stenosis is noted.
- Bilateral common carotid, carotid bulbs, external and internal carotid arteries are patent. No evidence of luminal stenosis is noted.
- No definite aneurysm, arteriovenous malformation, or other vascular abnormality.

- Nasal septum deviation to the right.
- Hypertrophy of the nasal turbinates, especially the left ones.
- The nasopharynx, oropharynx, hypopharynx, and larynx are unremarkable.
- No particular findings of parotid glands, submandibular glands, and thyroid gland.
- No bulky lymph nodes or mass lesion over the neck.
- Normal pneumatization of the paranasal sinuses and mastoid air cells.
- The visible lungs are unremarkable.
  )"
    RisController.PasteToFinding(MyForm)

    MyForm := "
  (
The major neck and intracranial arteries are patent, without vascular anomaly nor luminal stenosis.
  )"
    RisController.PasteToImpression(MyForm)
}

::br::
{
    GenerateBrainImpression(false)
}

::mbr::
{
    GenerateBrainImpression(true)
}

; --- Core Logic ---

GenerateBrainImpression(isMild := false)
{
    ; 1. 取得文本內容 (沿用既有邏輯)
    searchText := RisController.GetFindingSearchText()

    ; 若完全無內容，直接 Return 或給出預設
    if (searchText == "") {
        Paste("Unremarkable study.")
        return
    }

    ; 2. 準備收集陽性發現的陣列
    positiveFindings := []

    ; --- 檢查邏輯 (使用 HasPositiveFinding 進行負面表列過濾) ---

    ; (1) Brain Atrophy
    ; 關鍵字：atrophy (萎縮), involution (退化), volume loss (體積減少)
    if (HasPositiveFinding(searchText, ["atrophy", "atrophic", "involution", "volume loss"])) {
        positiveFindings.Push("brain atrophy")
    }

    ; (2) Leukoaraiosis (白質病變/小血管病變)
    ; 關鍵字：leukoaraiosis, white matter (通常指 deep white matter ischemic change), SVD, demyelination
    if (HasPositiveFinding(searchText, ["leukoaraiosis", "small vessel disease", "small vascular ischemic disease", "microangiopathy"])) {
        positiveFindings.Push("leukoaraiosis")
    }

    ; (3) Old Lacunar Infarcts (陳舊性小洞性梗塞)
    ; 關鍵字：lacunar, lacuna (放在 Old Insults 之前先偵測，避免被混淆)
    if (HasPositiveFinding(searchText, ["lacunar", "lacuna"])) {
        positiveFindings.Push("old lacunar infarcts")
    }

    ; (4) Old Insults (陳舊性腦損傷/軟化)
    ; 關鍵字：encephalomalacia (腦軟化), gliosis (膠質增生), old insult, old infarct
    ; 註：這裡排除 lacunar，避免重複，但若同時有大片軟化和小洞，兩者都會被列出
    if (HasPositiveFinding(searchText, ["encephalomalacia", "gliosis", "old insult", "ischemic insult"])) {
        positiveFindings.Push("old insults")
    }

    ; (5) Cerebral Atherosclerosis (腦動脈硬化)
    ; 關鍵字：atherosclero (包含 atherosclerosis, atherosclerotic), dolichoectasia (延長擴張)
    if (HasPositiveFinding(searchText, ["atherosclerosis", "atherosclerotic", "dolichoectasia"])) {
        positiveFindings.Push("cerebral atherosclerosis")
    }

    ; 3. 組合字串
    if (positiveFindings.Length == 0) {
        ; 如果沒有抓到上述特定病徵
        Paste("No specific intracranial abnormality.")
        return
    }

    ; 4. 格式化輸出
    mainStr := FormatList(positiveFindings)

    ; 根據參數決定是否加入 "Mild " 前綴
    prefix := isMild ? "Mild " : ""
    combinedStr := prefix . mainStr

    ; 將首字母大寫 (Capitalize first letter)
    finalStr := StrUpper(SubStr(combinedStr, 1, 1)) . SubStr(combinedStr, 2) . "."

    Paste(finalStr)
}
