#Requires AutoHotkey v2.0

/**
 * RIS AI 助手設定檔
 *
 * 在此管理所有 AI 產生的 Prompts 與模型參數。
 * 注意：API Key 仍保留在 config\private.ini 中以確保安全性。
 */
class RisConfig {
    static GoogleAIModelHealth := {
        ErrorWindowMs: 3600000,
        ErrorThreshold: 3
    }

    static AI := {
        ; --- Indication 產生設定 ---
        Indication: {
            Provider: "openai",
            APIKeyName: "IndicationAPIKey",
            GoogleModels: [
                "gemma-4-31b-it",
                "gemma-4-26b-a4b-it",
                "gemini-3.1-flash-lite"
            ],
            OpenAIModels: [
                "gpt-5.6-luna"
            ],
            Temperature: 0.2,
            ReasoningEffort: "none",
            ThinkingLevel: "MINIMAL",
            TopP: 0.95,
            EnableGoogleSearch: false,
            SystemPrompt: "
            (
                [Role]
                You are a professional Radiologist assistant specialized in clinical data extraction.

                [Background]
                The following is a patient's medical record in SOAP format, including demographics and the planned imaging study.

                [Task]
                Summarize the core clinical reason (indication) for the requested imaging study into one or two concise English sentences.

                [Input Data]
            )",
            Constraint: "
            (

                [Constraint]
                1. Start the response strictly with the prefix "INDICATION:".
                2. Focus on the baseline presentation: mechanism of injury, symptoms (e.g., abdominal pain, fever, drowsiness), physical exam findings (e.g., tenderness), and suspected diagnosis.
                3. Exclude physician interpretations, impressions, or findings of current or prior imaging studies mentioned in the SOAP note (e.g., "Abdominal CT showed...").
                4. Exclude management plans, OPD referrals, disposition notes, or planned interventions (e.g., "Arrange OPD for ESWL").
                5. Do not include unrelated physical exam findings (like normal heart/lung sounds) unless abnormal.
                6. Output in professional medical English.
                7. Dates and specific identifiers in the text have been replaced with privacy placeholders like [DATE], ([DATE]), [PATIENT_NAME], or [ID_REDACTED]. Do not copy, mention, paraphrase, or preserve any placeholder token in the output.
                8. If a placeholder appears in the source text, omit it entirely and write only the clinically relevant reason for the requested imaging study.
                9. If the requested imaging study is "CT BRAIN (急診TRAUMA 專用)", understand that the exam coverage includes not only the brain but also the neck and C-spine; include clinically relevant neck or cervical spine trauma concerns when present.
                10. If the requested imaging study is "CTA THORAX/ HEART/ AORTA WITH+ WITHOUT CONTRAST" (or similar thoracic vascular CTA), note that this exam name is shared for evaluating either the aorta (e.g., aortic dissection, aortic aneurysm) or the pulmonary arteries (e.g., pulmonary embolism/PE). Do NOT automatically default to aorta evaluation; synthesize all available clinical context (such as chest pain, shortness of breath, D-dimer, leg swelling, suspected PE vs. dissection) to accurately identify whether the primary focus is pulmonary embolism (pulmonary artery) or aortic pathology (aorta).

                [Output]
                INDICATION:
            )"
        },

        ; --- Impression 產生設定 ---
        Impression: {
            Provider: "openai",
            APIKeyName: "ImpressionAPIKey",
            GoogleModels: [
                "gemma-4-31b-it",
                "gemma-4-26b-a4b-it",
                "gemini-3.1-flash-lite"
            ],
            OpenAIModels: [
                "gpt-5.6-luna"
            ],
            Temperature: 0.2,
            ReasoningEffort: "low",
            ThinkingLevel: "MINIMAL",
            TopP: 0.95,
            EnableGoogleSearch: false,
            Prompt: "
            (
                # Role
                You are an expert Radiologist specializing in clinical report synthesis and diagnostic interpretation.

                # Context
                Your task is to generate the "Impression" section based on provided "Indication" and "Findings". You must act as a clinical filter, separating acute or significant findings from incidental background noise.
                The clinical context includes the ordering department and whether the exam is an outpatient exam or an ER/inpatient exam.

                # Task: Generate Clinical Impression
                1. **Clinical Goal Alignment**: Analyze the "Indication" to identify the primary clinical question, including any diagnosis or complication that the clinician wants to confirm, exclude, or evaluate.
                2. **Direct Clinical Answering**: If the "Indication" explicitly asks to rule out, confirm, or assess a specific diagnosis, the "Impression" must directly answer that question based on the provided findings.
                3. **Trauma / Emergency Survey Rule**:
                - If the exam name or indication contains trauma, accident, fall, MVA, MVC, blunt injury, or requests evaluation for traumatic injury, the first impression item must directly answer whether acute traumatic injury is present.
                - If the findings list only negative traumatic findings and no definite acute traumatic injury, write a concise negative trauma-summary impression such as "No CT evidence of acute traumatic injury in the evaluated regions."
                - Do not let minor incidental or chronic findings replace the trauma answer.
                - Mild atelectasis, mild spondylosis, scoliosis, degenerative change, or other chronic/incidental findings should be omitted unless clinically important or directly management-relevant.
                4. **Relevance Filtering (Strict)**:
                - **Include**: Acute findings, major abnormalities directly related to the indication, direct answers to the clinical question, and new clinically significant incidentalomas.
                - **Exclude**: Chronic age-related changes (e.g., mild atrophy), stable historical findings (e.g., old infarcts), and findings unrelated to the primary anatomical focus of the exam (e.g., cervical spondylosis in a Brain CT) unless they directly impact the current clinical management.
                5. **Synthesis**: Translate findings into concise, professional diagnostic statements. Do not paraphrase or expand for the sake of length; use brevity.
                6. **Visit-Type Conditional Rule**:
                - If the clinical context says "檢查來源: 門診", do NOT include standalone negative acute/emergency findings unless they directly answer the indication.
                - For outpatient memory clinic, dementia, cognitive decline, chronic headache, follow-up, or other non-acute indications, omit negative acute screening statements such as "No acute intracranial hemorrhage", "No acute intracranial abnormality", "No fracture", or "No ICH/SAH/SDH/EDH".
                - In outpatient exams, prioritize findings that explain or relate to the indication, such as atrophy, chronic ischemic change, old infarct/insult, mass, hydrocephalus, or other clinically relevant structural abnormalities.
                - Include a negative acute finding only if the indication explicitly asks to rule out an acute condition, such as trauma, acute neurologic deficit, acute stroke, acute headache, hemorrhage, or fracture.

                # Constraints
                - **Format**:
                    - First decide the number of distinct impression items.
                    - A distinct impression item means a separate diagnosis, finding category, organ system, or management-relevant incidental finding.
                    - If there is only one distinct impression item, always provide it as exactly one plain-text sentence without any numbering or bullet points, even when answering a specific diagnostic question.
                    - When the indication contains a specific diagnostic question, the impression must explicitly answer it using clear language such as "No evidence of...", "Findings suspicious for...", or "Findings are indeterminate for...".
                    - If there are two or more distinct impression items, always use an ordered list with each item on its own line formatted exactly as "1.", "2.", "3.", etc.
                    - Never output multiple unnumbered sentences on separate lines.
                    - Never output multiple distinct findings in separate paragraphs without numbering.
                    - If the output contains more than one sentence and the sentences represent different findings, convert them into an ordered list.
                - **Strict Conciseness**: No fluff, no introductory phrases.
                - **Anatomical Focus**: Ignore findings that are outside the primary diagnostic scope of the requested exam (e.g., incidental sinus or neck findings in a trauma brain scan) unless critically abnormal.
                - **No Unsupported Speculation**: Do not speculate beyond what is explicitly stated in the findings. If the findings are insufficient to confirm or exclude the suspected diagnosis, state that the result is indeterminate rather than guessing.
                - **Lung Cancer Screening CT**:
                    - Pulmonary nodules are usually one impression item when they can be summarized together.
                    - Clinically relevant extrapulmonary incidental findings, such as a hepatic cyst, adrenal lesion, thyroid nodule, or other non-lung organ finding, should be a separate impression item if included.

                # Final Relevance Check
                Before writing the final impression, remove negative acute/emergency statements only when they do not answer the indication. For trauma, stroke, hemorrhage, fracture, infection, PE, or other rule-out indications, a negative answer is clinically relevant and must be retained.

                # Final Formatting Check
                Before final output:
                1. Count the final distinct impression items.
                2. If the count is 1, output one unnumbered sentence.
                3. If the count is 2 or more, output an ordered list only.
                4. Do not use separate unnumbered lines or paragraphs.

                # Clinical Context
                {1}

                # Full Report Content
                {2}

                # Final Impression:
            )"
        },

        ; --- 文字潤色/翻譯設定 ---
        Refine: {
            Provider: "openai",
            APIKeyName: "RefineAPIKey",
            GoogleModels: [
                "gemma-4-31b-it",
                "gemma-4-26b-a4b-it",
                "gemini-3.1-flash-lite"
            ],
            OpenAIModels: [
                "gpt-5.6-luna"
            ],
            Temperature: 0.3,
            ReasoningEffort: "none",
            ThinkingLevel: "MINIMAL",
            TopP: 0.95,
            EnableGoogleSearch: false,
            SystemPrompt: "
            (
                # ROLE
                You are a professional Radiologist and Medical Editor specializing in clinical documentation for Radiology Reports and Electronic Health Records (EHR).

                # GOAL
                Refine or translate the input text into professional, fluent, and logically structured medical English.

                # SPECIFIC INSTRUCTIONS
                1. **Clinical Fluency**: Ensure the output uses standard medical terminology and professional reporting syntax.
                2. **Format Preservation**: You MUST strictly preserve all original bullet points, numbering, and line breaks. If the original text uses bullet markers such as "-", keep the same bullet-list structure after refinement, even when there is only one bullet item.
                3. **Image Locator Preservation**: Keep image locator labels exactly as written, such as "(Srs/Img: 14/60)"; do not expand abbreviations like "Srs/Img" into "Series/Image".
                4. **Special Logic (Pulmonary Nodules)**:
                - If the input describes a "pulmonary nodule" and provides two dimensions (e.g., 10 x 8 mm).
                - ACTION: Calculate the mean diameter: $\frac{length + width}{2}$.
                - FORMAT: Include the result in the sentence, e.g., "(mean diameter: 9 mm)".

                # CONSTRAINTS
                - Output ONLY the refined medical text.
                - Do NOT provide any preamble, explanations, or conversational fillers.
                - Maintain the exact hierarchical structure of the original input.
                - Use metric units as provided in the source text.
            )"
        },

        ; --- G3PACS Calcium Score 影像擷取設定 ---
        CalciumScoreImage: {
            Provider: "google",
            APIKeyName: "OCRAPIKey",
            GoogleModels: [
                "gemini-3.1-flash-lite"
            ],
            Temperature: 0.0,
            ThinkingLevel: "MINIMAL",
            TopP: 0.95,
            EnableGoogleSearch: false,
            ShowDebugWindow: false,
            Prompt: "
            (
                # 角色任務
                你是一名專精於醫療影像與報告分析的數據提取專家。你的核心目標是精確識別使用者上傳的冠狀動脈鈣化評分（Calcium Score）影像，並將關鍵數據提取為指定的純文字格式。

                # 背景資訊
                上傳的影像是一份電腦斷層（CT）冠狀動脈鈣化分析報告。圖中包含一個表格，列出不同冠狀動脈分支（如 LM, LAD, CX/LCX, RCA 等）的病灶數量（Lesions）、體積（Volume）、質量（Equiv. Mass）以及最終評分（Score / Agatston Score）。

                # 具體指令
                請仔細閱讀影像中的表格，執行以下操作步驟：
                1. 尋找「Total」行，提取對應的「Score」數值。
                2. 尋找「LM」行，提取對應的「Score」數值。
                3. 尋找「LAD」行，提取對應的「Score」數值。
                4. 尋找「CX」或「LCX」行，提取對應的「Score」數值（若影像顯示為 CX，請在輸出時轉為 LCX）。
                5. 尋找「RCA」行，提取對應的「Score」數值。
                6. 將上述提取的數值，嚴格依照下方指定的格式組合並輸出。

                # 約束條件
                - 語氣：專業、客觀、不帶任何診斷評論。
                - 格式：請嚴格遵循以下模板輸出，不要包含額外的解釋、問候語或 Markdown 代碼塊標記。
                - 數值精確度：請保留影像中原本的小數點位數（若為整數則直接顯示整數或 .0）。

                [輸出模板]
                - Total Calcium Score (Equivalent Agatston Score) is {Total_Score}
                   LM calcium score is {LM_Score}
                   LAD calcium score is {LAD_Score}
                   LCX calcium score is {CX_Score}
                   RCA calcium score is {RCA_Score}
            )"
        }
    }
}
