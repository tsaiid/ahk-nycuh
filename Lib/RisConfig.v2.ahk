#Requires AutoHotkey v2.0

/**
 * RIS AI 助手設定檔
 *
 * 在此管理所有 AI 產生的 Prompts 與模型參數。
 * 注意：API Key 仍保留在 config\private.ini 中以確保安全性。
 */
class RisConfig {
    static AI := {
        ; --- Indication 產生設定 ---
        Indication: {
            Model: "gemma-4-31b-it",
            Temperature: 0.2,
            TopP: 0.95,
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
                5. Note: Dates and specific identifiers in the text have been replaced with placeholders like [DATE] or [PATIENT_NAME] for privacy. Please ignore the placeholders and focus on the clinical findings.

                [Output]
                INDICATION:
            )"
        },

        ; --- Impression 產生設定 ---
        Impression: {
            Model: "gemma-4-31b-it",
            Temperature: 0.2,
            TopP: 0.95,
            Prompt: "
            (
                # Role
                You are an expert Radiologist specializing in clinical report synthesis and diagnostic interpretation.

                # Context
                Your task is to generate the "Impression" section based on provided "Indication" and "Findings". You must act as a clinical filter, separating acute or significant findings from incidental background noise.

                # Task: Generate Clinical Impression
                1. **Clinical Goal Alignment**: Analyze the "Indication" to identify the primary clinical question.
                2. **Relevance Filtering (Strict)**:
                - **Include**: Acute findings, major abnormalities directly related to the indication, and new clinically significant incidentalomas.
                - **Exclude**: Chronic age-related changes (e.g., mild atrophy), stable historical findings (e.g., old infarcts), and findings unrelated to the primary anatomical focus of the exam (e.g., cervical spondylosis in a Brain CT) unless they directly impact the current clinical management.
                3. **Synthesis**: Translate findings into concise, professional diagnostic statements. Do not paraphrase or expand for the sake of length; use brevity.

                # Constraints
                - **Format**: Use a numbered list (1. 2. 3.).
                - **Strict Conciseness**: No fluff, no introductory phrases.
                - **Anatomical Focus**: Ignore findings that are outside the primary diagnostic scope of the requested exam (e.g., incidental sinus or neck findings in a trauma brain scan) unless critically abnormal.
                - **No Inferences**: Do not speculate beyond what is explicitly stated in the findings.

                # Full Report Content
                {1}

                # Final Impression:
            )"
        },

        ; --- 文字潤色/翻譯設定 ---
        Refine: {
            Model: "gemma-4-31b-it",
            Temperature: 0.3,
            TopP: 0.95,
            SystemPrompt: "Refine or translate the input text into professional, fluent medical English for use in radiology reports or clinical records. Ensure logical flow and standard medical terminology. You MUST strictly preserve all original bullet points and line breaks. Provide the refined text directly without any explanation or conversational fillers."
        }
    }
}
