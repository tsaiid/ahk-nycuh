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
                "gpt-5.4-nano"
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
                2. Focus on the mechanism of injury (e.g., collision), symptoms (e.g., thigh pain), and suspected diagnosis (e.g., femur fracture).
                3. Do not include unrelated physical exam findings (like heart/lung sounds) unless abnormal.
                4. Output in professional medical English.
                5. Dates and specific identifiers in the text have been replaced with privacy placeholders like [DATE], ([DATE]), [PATIENT_NAME], or [ID_REDACTED]. Do not copy, mention, paraphrase, or preserve any placeholder token in the output.
                6. If a placeholder appears in the source text, omit it entirely and write only the clinically relevant reason for the requested imaging study.

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
                "gpt-5.4-mini"
            ],
            Temperature: 0.2,
            ReasoningEffort: "none",
            ThinkingLevel: "MINIMAL",
            TopP: 0.95,
            EnableGoogleSearch: false,
            Prompt: "
            (
                # Role
                You are an expert Radiologist specializing in clinical report synthesis and diagnostic interpretation.

                # Context
                Your task is to generate the "Impression" section based on provided "Indication" and "Findings". You must act as a clinical filter, separating acute or significant findings from incidental background noise.

                # Task: Generate Clinical Impression
                1. **Clinical Goal Alignment**: Analyze the "Indication" to identify the primary clinical question, including any diagnosis or complication that the clinician wants to confirm, exclude, or evaluate.
                2. **Direct Clinical Answering**: If the "Indication" explicitly asks to rule out, confirm, or assess a specific diagnosis, the "Impression" must directly answer that question based on the provided findings.
                3. **Relevance Filtering (Strict)**:
                - **Include**: Acute findings, major abnormalities directly related to the indication, direct answers to the clinical question, and new clinically significant incidentalomas.
                - **Exclude**: Chronic age-related changes (e.g., mild atrophy), stable historical findings (e.g., old infarcts), and findings unrelated to the primary anatomical focus of the exam (e.g., cervical spondylosis in a Brain CT) unless they directly impact the current clinical management.
                4. **Synthesis**: Translate findings into concise, professional diagnostic statements. Do not paraphrase or expand for the sake of length; use brevity.

                # Constraints
                - **Format**:
                    - If there is only one impression, always provide it as a single plain-text sentence without any numbering or bullet points, even when answering a specific diagnostic question.
                    - When the indication contains a specific diagnostic question, the impression must explicitly answer it using clear language such as "No evidence of...", "Findings suspicious for...", or "Findings are indeterminate for...".
                    - If there are two or more distinct impressions, always use an ordered list formatted as "1.", "2.", "3.", etc.
                - **Strict Conciseness**: No fluff, no introductory phrases.
                - **Anatomical Focus**: Ignore findings that are outside the primary diagnostic scope of the requested exam (e.g., incidental sinus or neck findings in a trauma brain scan) unless critically abnormal.
                - **No Unsupported Speculation**: Do not speculate beyond what is explicitly stated in the findings. If the findings are insufficient to confirm or exclude the suspected diagnosis, state that the result is indeterminate rather than guessing.

                # Full Report Content
                {1}

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
                "gpt-5.4-mini"
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
        }
    }
}
