# AGENTS.md

## Validation Rule
- 每次修改 code 後，都必須通過 `powershell -NoProfile -ExecutionPolicy Bypass -File Utilities\compile-check.ps1` 驗證。
- 若驗證失敗，先修正再提交結果，不可略過。

## Line Ending Rule
- Follow `.gitattributes` line endings.
- `*.ahk` and `*.ps1` files must remain CRLF in the working tree.
- After editing AHK or PowerShell files, check `git ls-files --eol` and ensure touched files are not `w/mixed`.
- If needed, normalize touched AHK or PowerShell files to CRLF before validation or commit.

## Commit Rule
- Always use **Traditional Chinese**.
- Commit message follows Conventional Commits.
