#Requires AutoHotkey v2.0

/**
 * 執行指定的 action，並在前後發送 {Blind}{vkE8} 以防止 Windows 輸入法切換 (Alt+Shift / Ctrl+Shift)
 * @param {Func} action 要執行的回呼函式
 * @returns {*} action 執行的回傳值
 */
WithImeGuard(action) {
    Send("{Blind}{vkE8}")
    try {
        return action.Call()
    } finally {
        Send("{Blind}{vkE8}")
    }
}
