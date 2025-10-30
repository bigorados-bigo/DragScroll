
/*
Mouse Scroll v04 (extended)
Original by Mikhail V., 2021
Enhancements: configurable GUI, high-resolution wheel control, process exclusions
*/

#SingleInstance Force
#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.

SetBatchLines, -1

global configFile := A_ScriptDir . "\mouse-scroll.ini"
global running := 0
global passThrough := false
global swap := 0
global horiz := 0
global k := 1.0
global wheelSensitivity := 12.0
global wheelMaxStep := 480
global scanInterval := 20

global wheelBuffer := 0.0
global scrollsTotal := 0
global activationButton := "MButton"
global currentHotkey := ""
global mxLast := 0
global myLast := 0
global processListText := ""
global processBlockSet := {}
global topGuardListText := "OneCommander.exe:60"
global topGuardMap := {}
global topGuardMapCount := 0
global guiBuilt := false
global captureMode := false
global capturePrevDisplay := ""
global captureButtons := ["LButton", "RButton", "MButton", "XButton1", "XButton2"]

autoExec()
return

;------------------------
;  Auto-execute sequence
;------------------------
autoExec()
{
    global scanInterval
    LoadSettings()
    BuildConfigGui()
    ApplySettings()

    Menu, Tray, Add, Settings, ShowConfig
    Menu, Tray, Add
    Menu, Tray, Add, Exit, HandleExit

    SetTimer, ScrollTick, % scanInterval
}

;------------------------
;  Main scrolling loop
;------------------------
ScrollTick:
    global running, horiz, k, mxLast, myLast
    global wheelBuffer, wheelSensitivity, wheelMaxStep
    global swap, scrollsTotal

    if (!running)
        return

    MouseGetPos, mx, my
    if (horiz)
    {
        dy := k * (mx - mxLast)
        mxLast := mx
    }
    else
    {
        dy := k * (my - myLast)
        myLast := my
    }

    if (ShouldSuppressForTopGuard(mx, my))
    {
        wheelBuffer := 0
        return
    }

    wheelBuffer += dy * wheelSensitivity
    delta := Round(wheelBuffer)
    if (delta = 0)
        return

    wheelBuffer -= delta
    if (Abs(delta) > wheelMaxStep)
    {
        clamped := (delta > 0 ? wheelMaxStep : -wheelMaxStep)
        wheelBuffer += (delta - clamped)
        delta := clamped
    }

    eventFlag := horiz ? 0x01000 : 0x0800
    sendVal := swap ? delta : -delta
    DllCall("mouse_event", "UInt", eventFlag, "UInt", 0, "UInt", 0, "Int", sendVal, "Ptr", 0)
    scrollsTotal += Abs(delta)
return

;------------------------
;  Hotkey handlers
;------------------------
ActivationButtonDown:
    global passThrough, running, wheelBuffer, scrollsTotal
    global mxLast, myLast, horiz

    if (ShouldBlockProcess())
    {
        passThrough := true
        Send {%activationButton% Down}
        running := 0
        return
    }

    passThrough := false
    running := 1
    wheelBuffer := 0
    scrollsTotal := 0
    MouseGetPos, mxLast, myLast
return

ActivationButtonUp:
    global passThrough, running, scrollsTotal

    if (passThrough)
    {
        Send {%activationButton% Up}
        passThrough := false
        return
    }

    running := 0
    if (scrollsTotal = 0)
        Send {%activationButton%}
    scrollsTotal := 0
return

;------------------------
;  Settings management
;------------------------
LoadSettings()
{
    global configFile
    global swap, horiz, k, wheelSensitivity, wheelMaxStep
    global activationButton, processListText, scanInterval, topGuardListText

    ; defaults
    swap := 0
    horiz := 0
    k := 1.0
    wheelSensitivity := 12.0
    wheelMaxStep := 480
    activationButton := "MButton"
    processListText := ""
    scanInterval := 20

    if (!FileExist(configFile))
        return

    IniRead, swap, %configFile%, Settings, Swap, %swap%
    IniRead, horiz, %configFile%, Settings, Horizontal, %horiz%
    IniRead, k, %configFile%, Settings, SpeedMultiplier, %k%
    IniRead, wheelSensitivity, %configFile%, Settings, WheelSensitivity, %wheelSensitivity%
    IniRead, wheelMaxStep, %configFile%, Settings, WheelMaxStep, %wheelMaxStep%
    IniRead, activationButton, %configFile%, Settings, ActivationButton, %activationButton%
    IniRead, processListText, %configFile%, Settings, ExcludedProcesses, %processListText%
    IniRead, topGuardListText, %configFile%, Settings, TopGuardZones, %topGuardListText%
    IniRead, scanInterval, %configFile%, Settings, ScanInterval, %scanInterval%
}

SaveSettings()
{
    global configFile
    global swap, horiz, k, wheelSensitivity, wheelMaxStep
    global activationButton, processListText, scanInterval, topGuardListText

    IniWrite, %swap%, %configFile%, Settings, Swap
    IniWrite, %horiz%, %configFile%, Settings, Horizontal
    IniWrite, %k%, %configFile%, Settings, SpeedMultiplier
    IniWrite, %wheelSensitivity%, %configFile%, Settings, WheelSensitivity
    IniWrite, %wheelMaxStep%, %configFile%, Settings, WheelMaxStep
    IniWrite, %activationButton%, %configFile%, Settings, ActivationButton
    IniWrite, %processListText%, %configFile%, Settings, ExcludedProcesses
    IniWrite, %topGuardListText%, %configFile%, Settings, TopGuardZones
    IniWrite, %scanInterval%, %configFile%, Settings, ScanInterval
}

ApplySettings()
{
    global swap, horiz, k, wheelSensitivity, wheelMaxStep
    global activationButton, currentHotkey, scanInterval

    swap := swap ? 1 : 0
    horiz := horiz ? 1 : 0
    k := (k = "" ? 1.0 : k + 0.0)
    if (k = 0)
        k := 1.0

    wheelSensitivity := (wheelSensitivity = "" ? 12.0 : wheelSensitivity + 0.0)
    if (wheelSensitivity = 0)
        wheelSensitivity := 12.0

    wheelMaxStep := Floor(wheelMaxStep)
    if (wheelMaxStep < 120)
        wheelMaxStep := 120

    scanInterval := Floor(scanInterval)
    if (scanInterval < 5)
        scanInterval := 5
    SetTimer, ScrollTick, % scanInterval

    activationButton := NormalizeButton(activationButton)
    if (!IsSupportedActivationButton(activationButton))
    {
        MsgBox, 48, Mouse Scroll, Unsupported activation button.`nResetting to MButton.
        activationButton := "MButton"
    }

    RefreshProcessBlockSet()
    RefreshTopGuardSet()
    RegisterActivationHotkeys()
}

NormalizeButton(button)
{
    static map := {"LBUTTON":"LButton", "RBUTTON":"RButton", "MBUTTON":"MButton", "XBUTTON1":"XButton1", "XBUTTON2":"XButton2"}
    button := Trim(button)
    if (button = "")
        return ""
    StringUpper, upper, button
    if (map.HasKey(upper))
        return map[upper]
    return button
}

IsSupportedActivationButton(button)
{
    static allowed := {"LButton":true, "RButton":true, "MButton":true, "XButton1":true, "XButton2":true}
    return (button != "") && allowed.HasKey(button)
}

RegisterActivationHotkeys()
{
    global activationButton, currentHotkey

    if (currentHotkey != "")
    {
        Hotkey, % currentHotkey, Off
        Hotkey, % currentHotkey . " Up", Off
    }

    currentHotkey := activationButton
    Hotkey, %activationButton%, ActivationButtonDown, On
    Hotkey, %activationButton% Up, ActivationButtonUp, On
}

RefreshProcessBlockSet()
{
    global processListText, processBlockSet
    processBlockSet := {}

    if (processListText = "")
        return

    list := StrReplace(processListText, ",", "`n")
    Loop, Parse, list, `n
    {
        entry := Trim(A_LoopField)
        if (entry = "")
            continue
        StringLower, entry, entry
        processBlockSet[entry] := true
    }
}

RefreshTopGuardSet()
{
    global topGuardListText, topGuardMap, topGuardMapCount
    topGuardMap := {}
    topGuardMapCount := 0

    if (topGuardListText = "")
        return

    list := StrReplace(topGuardListText, ",", "`n")
    Loop, Parse, list, `n
    {
        entry := Trim(A_LoopField)
        if (entry = "")
            continue
        StringSplit, parts, entry, :
        process := Trim(parts1)
        height := Trim(parts2)
        if (process = "" || height = "")
            continue
        height := Abs(Floor(height))
        if (height <= 0)
            continue
        StringLower, processLower, process
        topGuardMap[processLower] := height
        topGuardMapCount++
    }
}

ShouldBlockProcess()
{
    global processBlockSet, processListText
    if (processListText = "")
        return false

    WinGet, procName, ProcessName, A
    if (procName = "")
        return false

    StringLower, procName, procName
    return !!processBlockSet.HasKey(procName)
}

ShouldSuppressForTopGuard(mx, my)
{
    global topGuardMap, topGuardMapCount
    if (!topGuardMapCount)
        return false

    WinGet, procName, ProcessName, A
    if (procName = "")
        return false

    StringLower, procName, procName
    if (!topGuardMap.HasKey(procName))
        return false

    guardHeight := topGuardMap[procName]
    WinGetPos, winX, winY,,, A
    if (winY = "")
        return false

    return (my <= winY + guardHeight)
}

;------------------------
;  GUI logic
;------------------------
BuildConfigGui()
{
    global guiBuilt, swap, horiz, k, wheelSensitivity, wheelMaxStep
    global activationButton, processListText

    if (guiBuilt)
        return

    Gui, Config:New, , Mouse Scroll Settings
    Gui, Config:Margin, 12, 12

    Gui, Config:Add, Text,, Activation button (click "Capture" and press the button):
    Gui, Config:Add, Edit, vGuiActivationDisplay w150 ReadOnly, %activationButton%
    Gui, Config:Add, Button, x+5 gGuiCaptureButton, Capture

    Gui, Config:Add, Checkbox, xm vGuiSwap Checked%swap%, Invert scroll direction
    Gui, Config:Add, Checkbox, vGuiHoriz Checked%horiz%, Horizontal mode (pan)

    Gui, Config:Add, Text,, Speed multiplier:
    Gui, Config:Add, Edit, vGuiK w150, %k%

    Gui, Config:Add, Text,, Wheel sensitivity (delta per pixel):
    Gui, Config:Add, Edit, vGuiWheelSens w150, %wheelSensitivity%

    Gui, Config:Add, Text,, Maximum wheel step:
    Gui, Config:Add, Edit, vGuiWheelMax w150, %wheelMaxStep%

    Gui, Config:Add, Text,, Excluded processes (one per line or comma separated):
    Gui, Config:Add, Edit, vGuiProcessList w260 h70, %processListText%

    Gui, Config:Add, Text,, Top guard zones (process:height px):
    Gui, Config:Add, Edit, vGuiTopGuardList w260 h70, %topGuardListText%

    Gui, Config:Add, Button, xm w80 gGuiApplySettings Default, Apply
    Gui, Config:Add, Button, x+10 w80 gGuiCloseConfig, Close

    guiBuilt := true
}

ShowConfig:
    BuildConfigGui()
    Gui, Config:Show
return

GuiCaptureButton:
    if (captureMode)
        return
    captureMode := true
    GuiControlGet, capturePrevDisplay, Config:, GuiActivationDisplay
    GuiControl, Config:, GuiActivationDisplay, Press button...
    ToolTip, Press desired mouse button or Esc to cancel.
    if (currentHotkey != "")
    {
        Hotkey, %currentHotkey%, Off
        Hotkey, %currentHotkey% Up, Off
    }
    for index, btn in captureButtons
        Hotkey, %btn%, GuiCaptureCommit, On
    Hotkey, Esc, GuiCaptureCancel, On
return

GuiCaptureCommit:
    if (!captureMode)
        return
    button := RegExReplace(A_ThisHotkey, "^[~\*\$\#\+\!\^]+")
    button := NormalizeButton(button)
    GuiControl, Config:, GuiActivationDisplay, %button%
    GuiCaptureReset()
return

GuiCaptureCancel:
    if (!captureMode)
        return
    GuiCaptureReset(true)
return

GuiApplySettings:
    Gui, Config:Submit, NoHide

    global swap, horiz, k, wheelSensitivity, wheelMaxStep
    global activationButton, processListText, topGuardListText

    if (captureMode)
    {
        ShowTempTooltip("Finish capture first.", 1000)
        return
    }

    activationButton := NormalizeButton(GuiActivationDisplay)
    swap := GuiSwap
    horiz := GuiHoriz
    k := GuiK
    wheelSensitivity := GuiWheelSens
    wheelMaxStep := GuiWheelMax
    processListText := GuiProcessList
    topGuardListText := GuiTopGuardList

    GuiControl, Config:, GuiActivationDisplay, %activationButton%

    ApplySettings()
    SaveSettings()
    ShowTempTooltip("Settings applied", 1000)
return

GuiCloseConfig:
    Gui, Config:Hide
    GuiCaptureReset(true)
return

GuiEscape:
    Gui, Config:Hide
    GuiCaptureReset(true)
return

;------------------------
;  Helpers
;------------------------
ShowTempTooltip(text, duration:=1000)
{
    ToolTip, %text%
    SetTimer, __tooltip_clear, % -Abs(duration)
    return

__tooltip_clear:
    SetTimer, __tooltip_clear, Off
    ToolTip
return
}

HandleExit:
    ExitApp
return

GuiCaptureReset(cancel := false)
{
    global captureMode, capturePrevDisplay, captureButtons
    global currentHotkey, activationButton

    if (!captureMode)
    {
        ToolTip
        return
    }

    captureMode := false
    ToolTip

    for index, btn in captureButtons
        Hotkey, %btn%, GuiCaptureCommit, Off
    Hotkey, Esc, GuiCaptureCancel, Off

    if (cancel)
    {
        if (capturePrevDisplay != "")
            GuiControl, Config:, GuiActivationDisplay, %capturePrevDisplay%
    }
    else
    {
        GuiControlGet, tempDisplay, Config:, GuiActivationDisplay
        if (tempDisplay = "")
            GuiControl, Config:, GuiActivationDisplay, %activationButton%
    }

    capturePrevDisplay := ""
    RegisterActivationHotkeys()
}
