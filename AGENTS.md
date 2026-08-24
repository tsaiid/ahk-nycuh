# AGENTS.md

## Validation Rule
- 修改檔案前請勿執行 compile-check 或 test 腳本，僅在完成程式碼修改後再進行驗證。
- 只有修改 `*.ahk` 檔後，才必須通過 `powershell -NoProfile -ExecutionPolicy Bypass -File Utilities\compile-check.ps1` 驗證。
- 只要修改 `*.ahk` 檔，也必須通過 `powershell -NoProfile -ExecutionPolicy Bypass -File Utilities\test.ps1` 單元測試。
- 未修改 `*.ahk` 檔時，使用與變更範圍相符的最小驗證即可，不必執行 `Utilities\compile-check.ps1`。
- 若驗證失敗，先修正再提交結果，不可略過。

## Line Ending Rule
- Follow `.gitattributes` line endings.
- `*.ahk` and `*.ps1` files must remain CRLF in the working tree.
- After editing any `*.ahk` or `*.ps1` file, always run `powershell -NoProfile -ExecutionPolicy Bypass -File Utilities\normalize-line-endings.ps1` before validation.
- `Utilities\compile-check.ps1` does not normalize line endings; after editing `*.ahk` files, run normalize first, then run compile-check before delivery.
- Before delivery, check `git ls-files --eol` and ensure touched AHK or PowerShell files are not `w/mixed`.

## Hotkey Rule
- Any hotkey combination containing `Alt+Shift` or `Ctrl+Shift` must protect IME state, using `WithImeGuard(...)` where available or the equivalent paired IME toggle in standalone scripts.

## Commit Rule
- Always use **Traditional Chinese**.
- Commit message follows Conventional Commits.
