---
name: ahk-hotstring-manager
description: Manage and add new AutoHotkey v2 hotstrings in a structured radiology reporting project. Use when a user requests to add a new abbreviation, hotstring, or reporting template.
---

# AHK Hotstring Manager

This skill provides a structured workflow for adding new hotstrings to the `ahk-nycuh` project, ensuring consistency and preventing conflicts across different radiology sub-specialties.

## Analysis Workflow

When asked to add a new hotstring (e.g., `abc` -> `full text`):

1. **Conflict Search**:
   - Use `grep_search` with the pattern `:?abc::` to check if the hotstring already exists in any file.
   - Use `grep_search` with the pattern `full text` to see if there are similar existing hotstrings.
   - **Ambiguity Check**: If the abbreviation could have different meanings in different contexts (e.g., `cpa` in Neuro vs. Chest), inform the user and suggest context-aware naming or confirm the intended domain.

2. **File Selection**:
   - **General Anatomical Abbreviations**: Use `Hotstrings/abbreviations.v2.ahk`.
   - **Specialty-specific Findings**: Use the corresponding file in `Hotstrings/`, such as:
     - `neuro.v2.ahk` for Brain/Spine.
     - `chest-ct.v2.ahk` for Chest CT.
     - `abdomen-ct.v2.ahk` for Abdomen CT.
   - **Templates/Forms**: For complex multiline templates, use files under `Hotstrings/neuro/`, `Hotstrings/lib/`, etc.

3. **Insertion Strategy**:
   - Open the target file and locate the appropriate section header (e.g., `;; Brain`, `;; Vascular`).
   - Insert the new hotstring in a logical position, preferably grouped with similar terms or alphabetically if applicable.
   - Ensure AHK v2 syntax: `::abbr::replacement` or `{ ... }` blocks for multiline content.

4. **Formatting Standards**:
   - Use `::abbr::replacement` for simple text.
   - Use `{ ... Paste(MyForm) }` for complex or multiline text to avoid indentation issues.
   - Ensure the file encoding remains **UTF-8 with BOM** if it contains Chinese characters.

## Insertion Patterns

### Simple Hotstring
```autohotkey
::cpa::cerebellopontine angle
```

### Multiline Form with Paste
```autohotkey
::myform::
{
    MyForm := "
  (
Line 1
Line 2
  )"
    Paste(MyForm)
}
```

## Quality Control
- **No Overwriting**: Never overwrite an existing hotstring without user confirmation.
- **Ready-to-run**: Ensure all `#Include` dependencies are respected.
- **Double-check**: After adding, confirm the change by reading back the modified section.
