#Requires AutoHotkey v2.0
#SingleInstance Force

#Include <RisConfig.v2>
#Include <RisAIConfigResolver.v2>
#Include <RisAIProviderPolicy.v2>
#Include <RisAIText.v2>
#Include <RisAIPayload.v2>
#Include <RisAIRequestBuilder.v2>
#Include <RisAITransport.v2>
#Include <RisAIDebug.v2>
#Include <RisDialog.v2>
#Include <RisAIDebugGui.v2>
#Include <RisAIService.v2>
#Include <RisAIModelHealth.v2>
#Include <RisAIOrchestration.v2>

; 設定工作列 (Tray) 上的 Icon
if FileExist(A_ScriptDir "\assets\G3PacsTools_icon.png") {
    TraySetIcon(A_ScriptDir "\assets\G3PacsTools_icon.png")
}

; 初始化 NoduleTracker（不覆蓋 Tray 圖示、啟動時不強制彈窗、統一 DPI Awareness 為 -3）
global tracker := NoduleTracker({initTray: false, showGui: false, setDpi: -3})
A_TrayMenu.Add()
A_TrayMenu.Add("顯示 Nodule Tracker`tAlt+Q", (*) => tracker.UpdateGUI())
tracker.InitTrayMenu(A_TrayMenu)

#Include NoduleTracker.ahk
#Include Hotkeys\g3pacs-tools.v2.ahk

RWin:: {
    if !WinExist("ahk_exe G3PACS.exe") {
        G3PacsNotify.Show("找不到 G3PACS 視窗", 2000)
        return
    }
    WinActivate("ahk_exe G3PACS.exe")
}

#^p:: {
    ProcessClose("G3PACS.exe")
}
