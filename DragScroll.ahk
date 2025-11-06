
/*
Mouse Scroll v04 (extended)
Original by Mikhail V., 2021
Enhancements: configurable GUI, high-resolution wheel control, process exclusions
*/

#SingleInstance Force
#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.

SetBatchLines, -1
SetWorkingDir, %A_ScriptDir%

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
global activationButton2 := ""
global currentHotkey2 := ""
global activationHotkeyData := {}
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
global captureSlot := 0
global captureInput := ""
global captureMouseBindings := {}
global activationMovementDetected := false
global activationState := "idle"
global activationNativeDown := false
global activationTriggerData := ""
global activationLastMotionTick := 0
global activationIdleRestoreMs := 220
global activationDragThreshold := 1
global explorerContext := {}
global lastExplorerWindow := 0
global uiAutomation := ""
global debugEnabled := false
global debugLogDefault := A_ScriptDir . "\dragscroll-debug.log"
global debugLogFile := debugLogDefault
global debugHotkeysEnabled := 0
global debugStartEnabled := 0
global debugLogRedirected := false
global debugLogWarned := false

;--- Mouse lock configuration ---
global mouseLockEnabled := true
global mouseLockActive := false
global mouseLockAnchorX := ""
global mouseLockAnchorY := ""
global mouseLockHideCursor := true
global mouseLockCursorHidden := false
global mouseLockBlankCursor := 0
global mouseLockSystemCursorApplied := false
global mouseLockOriginalCursors := {}

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
    ResetDebugLogState()
    ApplySettings()
    ApplyDebugModePreference(true)
    DebugWriteRaw("Startup complete (debugEnabled=" . (debugEnabled ? "true" : "false") . ")")

    Menu, Tray, Add, Settings, ShowConfig
    Menu, Tray, Add, Toggle Debug Logging, TrayToggleDebug
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
    global swap, scrollsTotal, debugEnabled
    global mouseLockEnabled, mouseLockActive, mouseLockAnchorX, mouseLockAnchorY

    if (!running)
        return

    if (!GetCursorPoint(mx, my))
    {
        if (debugEnabled)
            DebugLog("ScrollTick: GetCursorPoint failed; skipping frame")
        return
    }
    lockApplied := (mouseLockEnabled && mouseLockActive)

    if (lockApplied)
    {
        deltaX := mx - mouseLockAnchorX
        deltaY := my - mouseLockAnchorY
        rawDelta := ClampRawDelta(horiz ? deltaX : deltaY)
        dy := k * rawDelta
        if (deltaX != 0 || deltaY != 0)
            DllCall("SetCursorPos", "int", mouseLockAnchorX, "int", mouseLockAnchorY)
        mxLast := mouseLockAnchorX
        myLast := mouseLockAnchorY
    }
    else if (horiz)
    {
        rawDelta := ClampRawDelta(mx - mxLast)
        dy := k * rawDelta
        mxLast := mx
    }
    else
    {
        rawDelta := ClampRawDelta(my - myLast)
        dy := k * rawDelta
        myLast := my
    }

    if (!ProcessActivationDragState(rawDelta, mx, my))
        return

    if (mouseLockEnabled && mouseLockActive)
    {
        mx := mouseLockAnchorX
        my := mouseLockAnchorY
    }

    if (HandleExplorerScroll(dy, mx, my))
        return

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
    global mxLast, myLast, horiz, debugEnabled
    global activationHotkeyData, activationMovementDetected
    global activationState, activationNativeDown, activationTriggerData
    global activationLastMotionTick

    triggerId := GetHotkeyIdentity(A_ThisHotkey)
    triggerData := activationHotkeyData.HasKey(triggerId) ? activationHotkeyData[triggerId] : BuildActivationHotkeyData(triggerId)

    if (debugEnabled)
        DebugLog("ActivationButtonDown trigger=" . (triggerId != "" ? triggerId : "(unknown)"))
    EndMouseLock()

    if (!IsObject(triggerData))
    {
        passThrough := true
        running := 0
        activationState := "idle"
        activationTriggerData := ""
        activationNativeDown := false
        activationMovementDetected := false
        if (debugEnabled)
            DebugLog("Activation hotkey data unavailable; forcing passthrough")
        return
    }

    if (ShouldBlockProcess())
    {
        passThrough := true
        activationNativeDown := SendActivationDown(triggerData)
        running := 0
        activationState := "idle"
        activationTriggerData := ""
        activationMovementDetected := false
        if (debugEnabled)
            DebugLog("Activation blocked for current process; passing through")
        return
    }

    passThrough := false
    running := 1
    wheelBuffer := 0
    scrollsTotal := 0
    activationMovementDetected := false
    activationTriggerData := triggerData
    activationState := "pending"
    activationLastMotionTick := A_TickCount
    activationNativeDown := false
    if (!GetCursorPoint(mxLast, myLast))
    {
        mxLast := 0
        myLast := 0
    }
    TryActivateExplorerMode()
return

ActivationButtonUp:
    global passThrough, running, scrollsTotal, debugEnabled
    global activationHotkeyData, activationMovementDetected
    global activationState, activationNativeDown, activationTriggerData

    triggerId := GetHotkeyIdentity(A_ThisHotkey)
    triggerData := IsObject(activationTriggerData) ? activationTriggerData : (activationHotkeyData.HasKey(triggerId) ? activationHotkeyData[triggerId] : BuildActivationHotkeyData(triggerId))

    if (debugEnabled)
        DebugLog("ActivationButtonUp trigger=" . (triggerId != "" ? triggerId : "(unknown)") . " scrollsTotal=" . scrollsTotal)

    ResetExplorerMode()
    EndMouseLock()

    if (passThrough)
    {
        if (activationNativeDown)
            SendActivationUp(triggerData)
        passThrough := false
        activationMovementDetected := false
        activationState := "idle"
        activationNativeDown := false
        activationTriggerData := ""
        return
    }

    running := 0
    if (activationNativeDown)
        SendActivationUp(triggerData)
    else if (!activationMovementDetected && IsObject(triggerData))
        SendActivationTap(triggerData)
    activationState := "idle"
    activationNativeDown := false
    activationTriggerData := ""
    activationMovementDetected := false
    scrollsTotal := 0
return

;------------------------
;  Settings management
;------------------------
LoadSettings()
{
    global configFile
    global swap, horiz, k, wheelSensitivity, wheelMaxStep
    global activationButton, activationButton2, processListText, scanInterval, topGuardListText
    global debugHotkeysEnabled, debugStartEnabled, mouseLockEnabled, mouseLockHideCursor

    ; defaults
    swap := 0
    horiz := 0
    k := 1.0
    wheelSensitivity := 12.0
    wheelMaxStep := 480
    activationButton := "MButton"
    activationButton2 := ""
    processListText := ""
    scanInterval := 20
    debugHotkeysEnabled := 0
    debugStartEnabled := 1
    mouseLockEnabled := 1
    mouseLockHideCursor := 1

    if (!FileExist(configFile))
        return

    IniRead, swap, %configFile%, Settings, Swap, %swap%
    IniRead, horiz, %configFile%, Settings, Horizontal, %horiz%
    IniRead, k, %configFile%, Settings, SpeedMultiplier, %k%
    IniRead, wheelSensitivity, %configFile%, Settings, WheelSensitivity, %wheelSensitivity%
    IniRead, wheelMaxStep, %configFile%, Settings, WheelMaxStep, %wheelMaxStep%
    IniRead, activationButton, %configFile%, Settings, ActivationButton, %activationButton%
    IniRead, activationButton2, %configFile%, Settings, ActivationButtonSlot2, %activationButton2%
    IniRead, processListText, %configFile%, Settings, ExcludedProcesses, %processListText%
    IniRead, topGuardListText, %configFile%, Settings, TopGuardZones, %topGuardListText%
    IniRead, scanInterval, %configFile%, Settings, ScanInterval, %scanInterval%
    IniRead, debugHotkeysEnabled, %configFile%, Settings, EnableDebugHotkeys, %debugHotkeysEnabled%
    if (ErrorLevel)
        IniRead, debugHotkeysEnabled, %configFile%, Settings, DebugShortcutsEnabled, %debugHotkeysEnabled%
    IniRead, debugStartEnabled, %configFile%, Settings, DebugStartEnabled, %debugStartEnabled%
    IniRead, mouseLockEnabled, %configFile%, Settings, MouseLockEnabled, %mouseLockEnabled%
    IniRead, mouseLockHideCursor, %configFile%, Settings, MouseLockHideCursor, %mouseLockHideCursor%
    debugHotkeysEnabled := debugHotkeysEnabled ? 1 : 0
    debugStartEnabled := debugStartEnabled ? 1 : 0
    mouseLockEnabled := mouseLockEnabled ? 1 : 0
    mouseLockHideCursor := mouseLockHideCursor ? 1 : 0
    if (!debugStartEnabled && debugHotkeysEnabled)
        debugStartEnabled := 1
}

SaveSettings()
{
    global configFile
    global swap, horiz, k, wheelSensitivity, wheelMaxStep
    global activationButton, activationButton2, processListText, scanInterval, topGuardListText
    global debugHotkeysEnabled, debugStartEnabled, mouseLockEnabled, mouseLockHideCursor

    IniWrite, %swap%, %configFile%, Settings, Swap
    IniWrite, %horiz%, %configFile%, Settings, Horizontal
    IniWrite, %k%, %configFile%, Settings, SpeedMultiplier
    IniWrite, %wheelSensitivity%, %configFile%, Settings, WheelSensitivity
    IniWrite, %wheelMaxStep%, %configFile%, Settings, WheelMaxStep
    IniWrite, %activationButton%, %configFile%, Settings, ActivationButton
    IniWrite, %activationButton2%, %configFile%, Settings, ActivationButtonSlot2
    IniWrite, %processListText%, %configFile%, Settings, ExcludedProcesses
    IniWrite, %topGuardListText%, %configFile%, Settings, TopGuardZones
    IniWrite, %scanInterval%, %configFile%, Settings, ScanInterval
    IniWrite, %debugHotkeysEnabled%, %configFile%, Settings, EnableDebugHotkeys
    IniWrite, %debugStartEnabled%, %configFile%, Settings, DebugStartEnabled
    IniWrite, %mouseLockEnabled%, %configFile%, Settings, MouseLockEnabled
    IniWrite, %mouseLockHideCursor%, %configFile%, Settings, MouseLockHideCursor
}

ApplySettings()
{
    global swap, horiz, k, wheelSensitivity, wheelMaxStep
    global activationButton, activationButton2, currentHotkey, currentHotkey2, scanInterval
    global debugHotkeysEnabled, debugStartEnabled, mouseLockEnabled, mouseLockHideCursor
    global mouseLockActive

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

    activationButton := NormalizeActivationHotkey(activationButton)
    activationButton2 := NormalizeActivationHotkey(activationButton2)
    debugHotkeysEnabled := debugHotkeysEnabled ? 1 : 0
    debugStartEnabled := debugStartEnabled ? 1 : 0
    mouseLockEnabled := mouseLockEnabled ? 1 : 0
    mouseLockHideCursor := mouseLockHideCursor ? 1 : 0

    if (activationButton2 = activationButton)
        activationButton2 := ""

    if (!mouseLockEnabled)
        EndMouseLock()
    else if (mouseLockActive)
    {
        if (mouseLockHideCursor)
            HideCursorForLock()
        else
            ShowCursorAfterLock()
    }

    RefreshProcessBlockSet()
    RefreshTopGuardSet()
    RegisterActivationHotkeys()
    RegisterDebugHotkeys()
}

NormalizeActivationHotkey(hotkey)
{
    hotkey := Trim(hotkey)
    if (hotkey = "")
        return ""

    StringUpper, upper, hotkey
    if (upper = "NONE")
        return ""

    human := ConvertHumanReadableHotkey(hotkey)
    if (human != "")
        hotkey := human

    identity := GetHotkeyIdentity(hotkey)
    parts := SplitHotkeyParts(identity)
    if (parts.key = "")
        return ""
    return parts.mods . parts.key
}

ConvertHumanReadableHotkey(spec)
{
    if (!RegExMatch(spec, "i)(ctrl|control|alt|shift|win|windows|meta)"))
        return ""

    mods := ""
    key := ""
    pieces := StrSplit(spec, "+")
    for index, piece in pieces
    {
        part := Trim(piece)
        if (part = "")
            continue
        StringUpper, upper, part
        if (upper = "CTRL" || upper = "CONTROL")
            mods .= "^"
        else if (upper = "ALT")
            mods .= "!"
        else if (upper = "SHIFT")
            mods .= "+"
        else if (upper = "WIN" || upper = "WINDOWS" || upper = "META")
            mods .= "#"
        else
            key := part
    }

    if (key = "")
        return ""

    key := StandardizeKeyName(key)
    return mods . key
}

GetHotkeyIdentity(hotkey)
{
    hotkey := Trim(hotkey)
    if (hotkey = "")
        return ""
    hotkey := RegExReplace(hotkey, "i)\s+up$")
    hotkey := RegExReplace(hotkey, "^[~\*\$]+")
    return hotkey
}

SplitHotkeyParts(hotkey)
{
    result := {mods:"", key:""}
    if (hotkey = "")
        return result

    trimmed := Trim(hotkey)
    if (trimmed = "")
        return result

    index := 1
    len := StrLen(trimmed)
    while (index <= len)
    {
        char := SubStr(trimmed, index, 1)
        if (char ~= "[#!\^+<>")
        {
            result.mods .= char
            index++
        }
        else
            break
    }
    result.key := StandardizeKeyName(SubStr(trimmed, index))
    return result
}

StandardizeKeyName(key)
{
    key := Trim(key)
    if (key = "")
        return ""

    keyLower := key
    StringLower, keyLower, keyLower
            static map := ""
            if (!IsObject(map))
            {
                    map := {}
                    map["lbutton"] := "LButton"
                    map["rbutton"] := "RButton"
                    map["mbutton"] := "MButton"
                    map["xbutton1"] := "XButton1"
                    map["xbutton2"] := "XButton2"
                map["wheelup"] := "WheelUp"
                map["wheeldown"] := "WheelDown"
                    map["wheelleft"] := "WheelLeft"
                    map["wheelright"] := "WheelRight"
                    map["appskey"] := "AppsKey"
                    map["printscreen"] := "PrintScreen"
                    map["lcontrol"] := "LControl"
                    map["rcontrol"] := "RControl"
                    map["lctrl"] := "LControl"
                    map["rctrl"] := "RControl"
                    map["lshift"] := "LShift"
                    map["rshift"] := "RShift"
                    map["lalt"] := "LAlt"
                    map["ralt"] := "RAlt"
                    map["lwin"] := "LWin"
                    map["rwin"] := "RWin"
            }
    if (map.HasKey(keyLower))
        return map[keyLower]

    if (StrLen(key) = 1)
    {
        StringLower, key, key
        return key
    }

    return key
}

EnsureHotkeyRegisterSpec(identity)
{
    if (identity = "")
        return ""
    return (SubStr(identity, 1, 1) = "$") ? identity : ("$" . identity)
}

BuildActivationHotkeyData(hotkey)
{
    identity := NormalizeActivationHotkey(hotkey)
    if (identity = "")
        return ""

    parts := SplitHotkeyParts(identity)
    if (parts.key = "")
        return ""

    data := {}
    data.identity := parts.mods . parts.key
    data.registerSpec := EnsureHotkeyRegisterSpec(data.identity)
    data.isMouse := IsMouseActivationKey(parts.key)
    data.mods := parts.mods
    data.key := parts.key
    data.tap := BuildActivationSendString(parts.mods, parts.key)
    data.down := BuildActivationHoldString(parts.key, "Down")
    data.up := BuildActivationHoldString(parts.key, "Up")
    return data
}

BuildActivationSendString(mods, key)
{
    if (key = "")
        return ""
    keySpec := FormatKeyForSend(key)
    return (mods != "" ? mods : "") . keySpec
}

BuildActivationHoldString(key, event)
{
    if (key = "")
        return ""
    if (!IsHoldCapableKey(key))
        return ""
    keySpec := FormatKeyForSend(key, event)
    return "{Blind}" . keySpec
}

FormatKeyForSend(key, event := "")
{
    if (key = "")
        return ""
    if (event = "")
    {
        if (StrLen(key) = 1 && key ~= "^[a-z0-9]$")
            return key
        return "{" . key . "}"
    }
    return "{" . key . " " . event . "}"
}

IsHoldCapableKey(key)
{
    keyLower := key
    StringLower, keyLower, keyLower
    if (keyLower = "wheelup" || keyLower = "wheeldown" || keyLower = "wheelleft" || keyLower = "wheelright")
        return false
    return true
}

IsMouseActivationKey(key)
{
    static mouseKeys := { "LButton": true, "RButton": true, "MButton": true, "XButton1": true, "XButton2": true, "WheelDown": true, "WheelUp": true, "WheelLeft": true, "WheelRight": true }
    return mouseKeys.HasKey(key)
}

SendActivationEvent(data, mode)
{
    if (!IsObject(data))
        return false

    sendSpec := ""
    if (mode = "tap")
        sendSpec := data.HasKey("tap") ? data.tap : ""
    else if (mode = "down")
        sendSpec := data.HasKey("down") ? data.down : ""
    else if (mode = "up")
        sendSpec := data.HasKey("up") ? data.up : ""

    if (sendSpec = "")
    {
        if (mode = "tap")
            sendSpec := BuildActivationSendString("", data.HasKey("key") ? data.key : "")
        else if (mode = "down")
            sendSpec := BuildActivationHoldString(data.HasKey("key") ? data.key : "", "Down")
        else if (mode = "up")
            sendSpec := BuildActivationHoldString(data.HasKey("key") ? data.key : "", "Up")
    }

    if (sendSpec = "")
        return false

    SendInput, %sendSpec%
    return true
}

SendActivationTap(data)
{
    return SendActivationEvent(data, "tap")
}

SendActivationDown(data)
{
    return SendActivationEvent(data, "down")
}

SendActivationUp(data)
{
    return SendActivationEvent(data, "up")
}

RegisterActivationHotkeys()
{
    global activationButton, activationButton2, currentHotkey, currentHotkey2, activationHotkeyData

    if (currentHotkey != "")
    {
        Hotkey, % currentHotkey, Off
        Hotkey, % currentHotkey . " Up", Off
    }
    if (currentHotkey2 != "")
    {
        Hotkey, % currentHotkey2, Off
        Hotkey, % currentHotkey2 . " Up", Off
    }

    currentHotkey := ""
    currentHotkey2 := ""
    activationHotkeyData := {}

    RegisterSingleActivationHotkey(activationButton, 1)
    if (activationButton2 != "" && activationButton2 != activationButton)
        RegisterSingleActivationHotkey(activationButton2, 2)
}

RegisterSingleActivationHotkey(hotkeySpec, slot)
{
    global activationHotkeyData, currentHotkey, currentHotkey2

    if (hotkeySpec = "")
        return

    data := BuildActivationHotkeyData(hotkeySpec)
    if (!IsObject(data))
        return

    Hotkey, % data.registerSpec, ActivationButtonDown, On
    Hotkey, % data.registerSpec . " Up", ActivationButtonUp, On
    activationHotkeyData[data.identity] := data

    if (slot = 1)
        currentHotkey := data.registerSpec
    else if (slot = 2)
        currentHotkey2 := data.registerSpec
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
    global activationButton, activationButton2, processListText, topGuardListText
    global debugHotkeysEnabled, debugStartEnabled, mouseLockEnabled, mouseLockHideCursor

    if (guiBuilt)
        return

    Gui, Config:New, , Mouse Scroll Settings
    Gui, Config:Margin, 12, 12

    Gui, Config:Add, Text,, Activation button slot 1 (click "Capture" and press the button):
    Gui, Config:Add, Edit, vGuiActivationDisplay1 w150 ReadOnly, %activationButton%
    Gui, Config:Add, Button, x+5 gGuiCaptureButton1, Capture
    Gui, Config:Add, Button, x+5 gGuiClearButton1, Clear

    Gui, Config:Add, Text, xm, Activation button slot 2 (optional):
    Gui, Config:Add, Edit, vGuiActivationDisplay2 w150 ReadOnly, %activationButton2%
    Gui, Config:Add, Button, x+5 gGuiCaptureButton2, Capture
    Gui, Config:Add, Button, x+5 gGuiClearButton2, Clear

    Gui, Config:Add, Checkbox, xm vGuiSwap Checked%swap%, Invert scroll direction
    Gui, Config:Add, Checkbox, vGuiHoriz Checked%horiz%, Horizontal mode (pan)
    Gui, Config:Add, Checkbox, xm vGuiMouseLock Checked%mouseLockEnabled%, Lock mouse cursor while scrolling
    Gui, Config:Add, Checkbox, vGuiMouseLockHide Checked%mouseLockHideCursor%, Hide cursor when active
    Gui, Config:Add, Checkbox, xm vGuiDebugHotkeys Checked%debugHotkeysEnabled%, Enable debug hotkeys (Win+Ctrl+D / Win+Ctrl+L)
    Gui, Config:Add, Checkbox, vGuiDebugStart Checked%debugStartEnabled%, Start with debug logging enabled

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

GuiCaptureButton1:
    BeginGuiCapture(1)
return

GuiCaptureButton2:
    BeginGuiCapture(2)
return

GuiClearButton1:
    ClearActivationSlot(1)
return

GuiClearButton2:
    ClearActivationSlot(2)
return

BeginGuiCapture(slot)
{
    global captureMode, captureSlot, capturePrevDisplay, captureInput
    global currentHotkey, currentHotkey2

    if (captureMode)
        return

    captureMode := true
    captureSlot := slot
    controlName := (slot = 2) ? "GuiActivationDisplay2" : "GuiActivationDisplay1"
    GuiControlGet, capturePrevDisplay, Config:, %controlName%
    GuiControl, Config:, %controlName%, Press key or button...
    ToolTip, Press desired button or key (Esc to cancel).

    if (currentHotkey != "")
    {
        Hotkey, %currentHotkey%, Off
        Hotkey, %currentHotkey% Up, Off
    }
    if (currentHotkey2 != "")
    {
        Hotkey, %currentHotkey2%, Off
        Hotkey, %currentHotkey2% Up, Off
    }

    RegisterCaptureMouseHooks(false)
    RegisterCaptureMouseHooks(true)

    captureInput := InputHook("L1")
    captureInput.KeyOpt("{All}", "E")
    captureInput.OnEnd := Func("GuiCaptureInputFinished")
    captureInput.Start()
}

ClearActivationSlot(slot)
{
    global captureMode, activationButton, activationButton2

    if (captureMode)
    {
        ShowTempTooltip("Finish capture first.", 1000)
        return
    }

    controlName := (slot = 2) ? "GuiActivationDisplay2" : "GuiActivationDisplay1"
    GuiControl, Config:, %controlName%,

    if (slot = 2)
        activationButton2 := ""
    else
        activationButton := ""

    RegisterActivationHotkeys()
    ShowTempTooltip("Activation cleared", 900)
}

RegisterCaptureMouseHooks(enable := true)
{
    global captureMouseBindings
    static buttons := ["LButton", "RButton", "MButton", "XButton1", "XButton2", "WheelUp", "WheelDown", "WheelLeft", "WheelRight"]

    if (!IsObject(captureMouseBindings))
        captureMouseBindings := {}

    for index, btn in buttons
    {
        hotkeySpec := "*" . btn
        if (enable)
        {
            if (captureMouseBindings.HasKey(btn))
                continue
            Hotkey, %hotkeySpec%, GuiCaptureMouseDispatch, On
            captureMouseBindings[btn] := hotkeySpec
        }
        else if (captureMouseBindings.HasKey(btn))
        {
            Hotkey, % captureMouseBindings[btn], GuiCaptureMouseDispatch, Off
            captureMouseBindings.Delete(btn)
        }
    }
}

GuiCaptureMouseDispatch:
    global captureMode
    if (!captureMode)
        return

    button := RegExReplace(A_ThisHotkey, "^\*")

    mods := ""
    if (GetKeyState("Ctrl", "P"))
        mods .= "^"
    if (GetKeyState("Alt", "P"))
        mods .= "!"
    if (GetKeyState("Shift", "P"))
        mods .= "+"
    if (GetKeyState("LWin", "P") || GetKeyState("RWin", "P"))
        mods .= "#"

    hotkey := NormalizeActivationHotkey(mods . button)
    if (hotkey = "")
    {
        ShowTempTooltip("Unsupported input; try again.", 1200)
        return
    }

    ApplyCapturedHotkey(hotkey)
return

GuiCaptureInputFinished(ih)
{
    global captureMode
    if (!captureMode)
        return
    reason := ih.EndReason
    if (reason = "Stopped" || reason = 3)
        return

    endKey := ih.EndKey
    if (endKey = "Escape")
    {
        GuiCaptureReset(true)
        return
    }

    hotkey := ComposeCapturedHotkey(ih)
    if (hotkey = "")
    {
        ShowTempTooltip("Unsupported input; try again.", 1200)
        GuiCaptureReset(true)
        return
    }

    ApplyCapturedHotkey(hotkey)
}

ApplyCapturedHotkey(hotkey)
{
    global captureMode, captureSlot
    if (!captureMode)
        return

    controlName := (captureSlot = 2) ? "GuiActivationDisplay2" : "GuiActivationDisplay1"
    GuiControl, Config:, %controlName%, %hotkey%
    GuiCaptureReset()
}

ComposeCapturedHotkey(ih)
{
    if (!IsObject(ih))
        return ""

    key := ih.EndKey
    if (key = "")
        return ""

    mods := ""
    if (ih.Ctrl)
        mods .= "^"
    if (ih.Alt)
        mods .= "!"
    if (ih.Shift)
        mods .= "+"
    if (ih.Win)
        mods .= "#"

    key := StandardizeKeyName(key)
    keyLower := key
    StringLower, keyLower, keyLower
    if (keyLower = "lcontrol" || keyLower = "rcontrol" || keyLower = "control")
        mods := StrReplace(mods, "^")
    if (keyLower = "lalt" || keyLower = "ralt" || keyLower = "alt")
        mods := StrReplace(mods, "!")
    if (keyLower = "lshift" || keyLower = "rshift" || keyLower = "shift")
        mods := StrReplace(mods, "+")
    if (keyLower = "lwin" || keyLower = "rwin" || keyLower = "win")
        mods := StrReplace(mods, "#")

    return NormalizeActivationHotkey(mods . key)
}

GuiApplySettings:
    Gui, Config:Submit, NoHide

    global swap, horiz, k, wheelSensitivity, wheelMaxStep
    global activationButton, activationButton2, processListText, topGuardListText
    global debugHotkeysEnabled, debugStartEnabled, mouseLockEnabled, mouseLockHideCursor

    if (captureMode)
    {
        ShowTempTooltip("Finish capture first.", 1000)
        return
    }

    activationButton := NormalizeActivationHotkey(GuiActivationDisplay1)
    activationButton2 := NormalizeActivationHotkey(GuiActivationDisplay2)
    swap := GuiSwap
    horiz := GuiHoriz
    debugHotkeysEnabled := GuiDebugHotkeys ? 1 : 0
    debugStartEnabled := GuiDebugStart ? 1 : 0
    mouseLockEnabled := GuiMouseLock ? 1 : 0
    mouseLockHideCursor := GuiMouseLockHide ? 1 : 0
    k := GuiK
    wheelSensitivity := GuiWheelSens
    wheelMaxStep := GuiWheelMax
    processListText := GuiProcessList
    topGuardListText := GuiTopGuardList

    GuiControl, Config:, GuiActivationDisplay1, %activationButton%
    GuiControl, Config:, GuiActivationDisplay2, %activationButton2%

    ApplySettings()
    ApplyDebugModePreference()
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
;  Debug utilities
;------------------------
RegisterDebugHotkeys()
{
    global debugHotkeysEnabled
    state := debugHotkeysEnabled ? "On" : "Off"
    Hotkey, #^d, DebugSnapshotHotkey, %state%
    Hotkey, #^l, DebugDumpExplorerHotkey, %state%
    Hotkey, #^+d, DebugToggleModeHotkey, %state%
}

DebugSnapshotHotkey:
    if (!debugHotkeysEnabled)
        return
    DebugHotkeyAction()
return

DebugDumpExplorerHotkey:
    if (!debugHotkeysEnabled)
        return
    DebugDumpExplorerContext()
return

#^+d::
    ToggleDebugMode()
return

DebugToggleModeHotkey:
    if (!debugHotkeysEnabled)
        return
    ToggleDebugMode()
return

ToggleDebugMode(forceState := "", silent := false)
{
    global debugEnabled, debugLogFile, explorerContext, debugStartEnabled, debugLogDefault

    if (forceState = "")
        newState := !debugEnabled
    else
        newState := !!forceState

    if (debugEnabled = newState)
    {
        if (!silent)
        {
            stateText := debugEnabled ? "enabled" : "disabled"
            ShowTempTooltip("Debug mode " stateText, 1200)
        }
        return
    }

    debugEnabled := newState
    debugStartEnabled := debugEnabled ? 1 : 0
    state := debugEnabled ? "enabled" : "disabled"
    if (debugEnabled)
    {
        ResetDebugLogState()
        FileDelete, %debugLogDefault%
        fallbackPath := GetFallbackLogPath()
        FileDelete, %fallbackPath%
        FileDelete, %debugLogFile%
        DebugWriteRaw("Debug mode enabled")
        DebugLogExplorerContext("Explorer context on enable", explorerContext)
    }
    else
    {
        DebugWriteRaw("Debug mode disabled")
    }
    if (!silent)
        ShowTempTooltip("Debug mode " state, 1200)

    SaveSettings()
}

ApplyDebugModePreference(silent := false)
{
    global debugStartEnabled, debugEnabled
    desired := debugStartEnabled ? true : false
    if (desired = debugEnabled)
        return
    ToggleDebugMode(desired, silent)
}

DebugHotkeyAction()
{
    LogExplorerFocusSnapshot()
}

DebugDumpExplorerContext()
{
    global explorerContext
    if (!IsObject(explorerContext) || !explorerContext.active)
    {
        DebugWriteRaw("Explorer context dump requested but none active")
        ShowTempTooltip("Explorer context: none", 1500)
        return
    }

    summary := BuildExplorerContextSummary(explorerContext)
    DebugWriteRaw("Explorer context dump -> " . summary)
    ShowTempTooltip(summary, 2000)
}

DebugLog(msg)
{
    global debugEnabled
    if (!debugEnabled)
        return
    DebugWriteRaw(msg)
}

DebugWriteRaw(msg)
{
    global debugLogFile, debugLogDefault, debugLogRedirected, debugLogWarned
    FormatTime, stamp, %A_Now%, yyyy-MM-dd HH:mm:ss
    entry := "[" . stamp . "] " . msg . "`r`n"
    outputMessage := "DragScroll: " . msg
    OutputDebug, %outputMessage%
    if (TryAppendDebugEntry(debugLogFile, entry))
        return

    fallbackFile := GetFallbackLogPath()
    if (fallbackFile != debugLogFile)
    {
        if (TryAppendDebugEntry(fallbackFile, entry))
        {
            if (!debugLogRedirected)
            {
                debugLogRedirected := true
                notice := "[" . stamp . "] Logging redirected from " . debugLogFile . " to " . fallbackFile . "`r`n"
                FileAppend, %notice%, %fallbackFile%, UTF-8
                ShowTempTooltip("Debug log redirected to " . fallbackFile, 2500)
            }
            debugLogFile := fallbackFile
            return
        }
    }

    if (!debugLogWarned)
    {
        debugLogWarned := true
        failMsg := "Unable to write debug log.`nPrimary path: " . debugLogDefault . "`nFallback path: " . fallbackFile
        MsgBox, 16, DragScroll, %failMsg%
    }
}

DebugLogExplorerContext(tag, ctx)
{
    if (!IsObject(ctx))
        return
    summary := BuildExplorerContextSummary(ctx)
    DebugLog(tag . ": " . summary)
}

BuildExplorerContextSummary(ctx)
{
    if (!IsObject(ctx))
        return "(null)"
    strategy := ctx.HasKey("strategy") ? ctx.strategy : "?"
    if (strategy = "listview")
    {
        hwnd := ctx.HasKey("listView") ? (ctx.listView + 0) : 0
        scale := ctx.HasKey("pixelScale") ? (ctx.pixelScale + 0.0) : 0.0
        return Format("strategy=listview hwnd=0x{:X} pixelScale={:.2f}", hwnd, scale)
    }
    if (strategy = "uia")
    {
        mode := ctx.HasKey("modeV") ? ctx.modeV : "?"
        vert := ctx.HasKey("verticalScrollable") ? (ctx.verticalScrollable ? "true" : "false") : "?"
        percentVal := ctx.HasKey("lastVerticalPercent") ? (ctx.lastVerticalPercent + 0.0) : ""
        rangeVal := ctx.HasKey("rangePerPixelV") ? (ctx.rangePerPixelV + 0.0) : ""
        percentStr := (percentVal = "") ? "?" : Format("{:.2f}", percentVal)
        rangeStr := (rangeVal = "") ? "?" : Format("{:.3f}", rangeVal)
        return "strategy=uia modeV=" . mode . " vert=" . vert . " percent=" . percentStr . " rangeStep=" . rangeStr
    }
    if (strategy = "scrollbar")
    {
        pos := ctx.HasKey("position") ? ctx.position : "?"
        range := ctx.HasKey("totalUnits") ? ctx.totalUnits : "?"
        units := ctx.HasKey("unitsPerPixel") ? ctx.unitsPerPixel : "?"
        track := ctx.HasKey("trackLength") ? ctx.trackLength : "?"
        folder := ctx.HasKey("folderView") ? (ctx.folderView + 0) : 0
        unitsStr := (units = "?") ? "?" : Format("{:.3f}", units + 0.0)
        trackStr := (track = "?") ? "?" : track
        folderStr := folder ? Format("0x{:X}", folder) : "0"
        return "strategy=scrollbar pos=" . pos . " range=" . range . " unitsPerPixel=" . unitsStr . " track=" . trackStr . " folderView=" . folderStr
    }
    return "(unknown context)"
}

LogExplorerFocusSnapshot()
{
    info := GatherExplorerDebugInfo()
    if (!IsObject(info))
        return

    DebugWriteRaw("Explorer snapshot: " . info.summary)
    for index, line in info.lines
        DebugWriteRaw("  " . line)
    ShowTempTooltip("Explorer snapshot logged", 1200)
}

GatherExplorerDebugInfo()
{
    info := {}
    info.lines := []

    WinGet, hwnd, ID, A
    if (!hwnd)
    {
        info.summary := "No active window"
        return info
    }

    hwnd := hwnd + 0
    title := ""
    WinGetTitle, title, ahk_id %hwnd%
    className := GetWindowClassName(hwnd)
    WinGet, procName, ProcessName, ahk_id %hwnd%
    procName := procName = "" ? "(unknown)" : procName

    procLower := procName
    StringLower, procLower, procName

    focusHwnd := GetWindowFocusHandle(hwnd)
    focusClass := GetWindowClassName(focusHwnd)
    focusText := GetWindowTextSafe(focusHwnd)

    summary := "active=0x" . Format("{:X}", hwnd) . " class=" . className . " proc=" . procName . " title=" . title
    info.summary := summary

    focusLine := "focus=0x" . Format("{:X}", focusHwnd) . " class=" . focusClass
    if (focusText != "")
        focusLine .= " text=" . focusText
    info.lines.Push(focusLine)

    explorerCheck := (procLower = "explorer.exe" && className = "CabinetWClass")
    if (explorerCheck)
    {
        folderView := FindExplorerFolderViewHandle(hwnd)
        folderClass := GetWindowClassName(folderView)
        folderFocused := folderView ? IsWindowDescendant(focusHwnd, folderView) : false
        detailsHandle := FindExplorerDetailsPaneHandle(hwnd)
        detailsVisible := detailsHandle && DllCall("IsWindowVisible", "ptr", detailsHandle)
        detailsClass := GetWindowClassName(detailsHandle)
        detailsLine := "detailsPane=0x" . Format("{:X}", detailsHandle) . " class=" . detailsClass . " visible=" . (detailsVisible ? "true" : "false")
        info.lines.Push("explorer=true folderView=0x" . Format("{:X}", folderView) . " class=" . folderClass . " tabFocused=" . (folderFocused ? "true" : "false"))
        info.lines.Push(detailsLine)

        viewHandles := CollectDescendantsByClass(hwnd, "DirectUIHWND", 8)
        if (viewHandles.Length())
        {
            info.lines.Push("views=" . viewHandles.Length())
            for idx, viewHwnd in viewHandles
            {
                rect := GetWindowRectData(viewHwnd)
                area := (IsObject(rect)) ? rect.width . "x" . rect.height : "?"
                visible := DllCall("IsWindowVisible", "ptr", viewHwnd)
                info.lines.Push("  view[" . idx . "] hwnd=0x" . Format("{:X}", viewHwnd) . " visible=" . (visible ? "true" : "false") . " size=" . area)
            }
        }

        scrollHandles := CollectDescendantsByClass(hwnd, "ScrollBar", 6)
        if (scrollHandles.Length())
        {
            info.lines.Push("scrollbars=" . scrollHandles.Length())
            for idx, sbHwnd in scrollHandles
            {
                rect := GetWindowRectData(sbHwnd)
                size := (IsObject(rect)) ? rect.width . "x" . rect.height : "?"
                orient := IsVerticalScrollBar(sbHwnd) ? "vert" : "horiz"
                visible := DllCall("IsWindowVisible", "ptr", sbHwnd)
                info.lines.Push("  sb[" . idx . "] hwnd=0x" . Format("{:X}", sbHwnd) . " " . orient . " visible=" . (visible ? "true" : "false") . " size=" . size)
            }
        }
    }
    else
    {
        info.lines.Push("explorer=false")
    }

    windows := EnumerateExplorerWindows()
    if (windows.Length())
    {
        info.lines.Push("explorerWindows=" . windows.Length())
        for index, entry in windows
        {
            winLine := "window[" . index . "] hwnd=0x" . Format("{:X}", entry.hwnd) . " visible=" . (entry.visible ? "true" : "false")
            if (entry.active)
                winLine .= " active=true"
            if (entry.title != "")
                winLine .= " title=" . entry.title
            if (entry.path != "")
                winLine .= " path=" . entry.path
            info.lines.Push(winLine)
        }
    }

    return info
}

GetWindowFocusHandle(winHwnd)
{
    if (!winHwnd)
        return 0

    ControlGet, focusHwnd, Hwnd,, Focus, ahk_id %winHwnd%
    if (focusHwnd)
        return focusHwnd + 0

    threadId := DllCall("GetWindowThreadProcessId", "ptr", winHwnd, "uint*", 0)
    currentThread := DllCall("GetCurrentThreadId")
    attached := false
    if (threadId && threadId != currentThread)
    {
        attached := DllCall("AttachThreadInput", "uint", currentThread, "uint", threadId, "int", true)
    }
    focus := DllCall("GetFocus", "ptr")
    if (attached)
        DllCall("AttachThreadInput", "uint", currentThread, "uint", threadId, "int", false)
    return focus ? focus + 0 : 0
}

GetWindowTextSafe(hwnd)
{
    if (!hwnd)
        return ""
    VarSetCapacity(buf, 512, 0)
    len := DllCall("GetWindowText", "ptr", hwnd, "str", buf, "int", 512)
    if (len <= 0)
        return ""
    buf := StrReplace(buf, "`r`n", " ")
    buf := StrReplace(buf, "`n", " ")
    buf := StrReplace(buf, "`r", " ")
    return buf
}

IsWindowDescendant(child, parent)
{
    if (!child || !parent)
        return false
    while (child)
    {
        if (child = parent)
            return true
        child := DllCall("GetParent", "ptr", child, "ptr")
    }
    return false
}

FindExplorerFolderViewHandle(winHwnd, screenX := "", screenY := "")
{
    if (!winHwnd)
        return 0

    views := CollectDescendantsByClass(winHwnd, "DirectUIHWND", 8)
    view := SelectBestWindowFromList(views, false, screenX, screenY)
    if (view)
        return view

    defViews := CollectDescendantsByClass(winHwnd, "SHELLDLL_DefView", 6)
    view := SelectBestWindowFromList(defViews, false, screenX, screenY)
    if (view)
        return view

    return FindDescendantByClass(winHwnd, "DirectUIHWND", 7)
}

ResolveActiveExplorerView(winHwnd, currentView, screenX := "", screenY := "")
{
    if (!winHwnd)
        return currentView
    if (screenX = "" || screenY = "")
        return currentView

    if (currentView && IsWindowTopmostAtPoint(currentView, screenX, screenY))
        return currentView

    view := FindHitTestDescendantByClass(winHwnd, "DirectUIHWND", 8, screenX, screenY)
    if (view)
        return view

    view := FindHitTestDescendantByClass(winHwnd, "SHELLDLL_DefView", 6, screenX, screenY)
    if (view)
        return view

    return currentView
}

FindHitTestDescendantByClass(winHwnd, className, maxDepth, screenX, screenY)
{
    handles := CollectDescendantsByClass(winHwnd, className, maxDepth)
    filtered := []
    for index, hwnd in handles
    {
        if (!DllCall("IsWindow", "ptr", hwnd))
            continue
        if (!DllCall("IsWindowVisible", "ptr", hwnd))
            continue
        rect := GetWindowRectData(hwnd)
        if (!IsObject(rect))
            continue
        if (!RectContainsPointData(rect, screenX, screenY))
            continue
        filtered.Push(hwnd)
    }
    if (filtered.Length() = 0)
        return 0
    return SelectBestWindowFromList(filtered, false, screenX, screenY)
}

FindExplorerDetailsPaneHandle(winHwnd)
{
    if (!winHwnd)
        return 0
    classes := ["AppDetailsPaneHost", "DetailsPaneHost", "DetailsPane", "PreviewPane", "DetailsViewContentHost"]
    for index, className in classes
    {
        handle := FindDescendantByClass(winHwnd, className, 8)
        if (handle)
            return handle
    }
    return 0
}

EnumerateExplorerWindows()
{
    windows := []
    try
        shell := ComObjCreate("Shell.Application")
    catch
        shell := ""
    if (!IsObject(shell))
        return windows

    try
        shellWindows := shell.Windows
    catch
        shellWindows := ""
    if (!IsObject(shellWindows))
        return windows

    count := shellWindows.Count
    Loop, %count%
    {
        try
            win := shellWindows.Item(A_Index - 1)
        catch
            continue
        if (!IsObject(win))
            continue
        entry := {}
        try
            hwnd := win.HWND
        catch
            hwnd := 0
        entry.hwnd := hwnd + 0
        entry.class := GetWindowClassName(entry.hwnd)
    entry.visible := entry.hwnd ? !!DllCall("IsWindowVisible", "ptr", entry.hwnd) : false
    entry.active := entry.hwnd ? (WinActive("ahk_id " . entry.hwnd) ? true : false) : false
        try
            entry.title := win.Document.Title
        catch
            entry.title := ""
        if (entry.title = "")
        {
            try
                entry.title := win.LocationName
            catch
                entry.title := ""
        }
        try
            entry.path := win.Document.Folder.Self.Path
        catch
            entry.path := ""
        windows.Push(entry)
    }
    return windows
}

DebugDumpWindowTree(hwnd, depth := 0, maxDepth := 2, siblingLimit := 12)
{
    global debugEnabled
    if (!debugEnabled || !hwnd || depth > maxDepth)
        return

    static lastDump := {}
    key := hwnd . "-" . maxDepth
    now := A_TickCount
    if (lastDump.HasKey(key) && now - lastDump[key] < 200)
        return
    lastDump[key] := now

    indent := ""
    Loop, %depth%
        indent .= "  "

    className := GetWindowClassName(hwnd)
    DebugLog(indent . Format("depth={} hwnd=0x{:X} class={} ", depth, hwnd + 0, className))

    child := 0
    index := 0
    while (child := DllCall("FindWindowEx", "ptr", hwnd, "ptr", child, "ptr", 0, "ptr", 0))
    {
        DebugDumpWindowTree(child, depth + 1, maxDepth, siblingLimit)
        index++
        if (index >= siblingLimit)
        {
            DebugLog(indent . "  ...")
            break
        }
    }
}

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

GetCursorPoint(ByRef x, ByRef y)
{
    static initialized := false, pt
    if (!initialized)
    {
        VarSetCapacity(pt, 8, 0)
        initialized := true
    }
    if (!DllCall("GetCursorPos", "ptr", &pt))
    {
        x := 0
        y := 0
        return false
    }
    x := NumGet(pt, 0, "int")
    y := NumGet(pt, 4, "int")
    return true
}

ClampRawDelta(delta)
{
    global debugEnabled
    if (Abs(delta) > 4096)
    {
        if (debugEnabled)
            DebugLog("Pointer delta out of range -> " . delta)
        return 0
    }
    return delta
}

ResetDebugLogState()
{
    global debugLogDefault, debugLogFile, debugLogRedirected, debugLogWarned
    debugLogFile := debugLogDefault
    debugLogRedirected := false
    debugLogWarned := false
}

GetFallbackLogPath()
{
    return A_AppData . "\DragScroll\dragscroll-debug.log"
}

TryAppendDebugEntry(path, entry)
{
    if (path = "")
        return false
    SplitPath, path, , dir
    if (dir != "")
        FileCreateDir, %dir%
    FileAppend, %entry%, %path%, UTF-8
    return !ErrorLevel
}

HandleExit:
    global mouseLockBlankCursor, mouseLockOriginalCursors
    EndMouseLock()
    RestoreSystemCursors()
    if (mouseLockBlankCursor)
    {
        DllCall("DestroyCursor", "ptr", mouseLockBlankCursor)
        mouseLockBlankCursor := 0
    }
    for cursorId, original in mouseLockOriginalCursors
    {
        if (original)
            DllCall("DestroyCursor", "ptr", original)
    }
    mouseLockOriginalCursors := {}
    ExitApp
return

TrayToggleDebug:
    ToggleDebugMode()
return

GuiCaptureReset(cancel := false)
{
    global captureMode, capturePrevDisplay, captureSlot, captureInput
    global activationButton, activationButton2

    if (!captureMode)
    {
        ToolTip
        return
    }

    captureMode := false
    ToolTip

    RegisterCaptureMouseHooks(false)

    if (IsObject(captureInput))
    {
        captureInput.OnEnd := ""
        captureInput.Stop()
        captureInput := ""
    }

    controlName := (captureSlot = 2) ? "GuiActivationDisplay2" : "GuiActivationDisplay1"
    fallbackButton := (captureSlot = 2) ? activationButton2 : activationButton

    if (cancel)
    {
        if (capturePrevDisplay != "")
            GuiControl, Config:, %controlName%, %capturePrevDisplay%
    }
    else
    {
        GuiControlGet, tempDisplay, Config:, %controlName%
        if (tempDisplay = "")
            GuiControl, Config:, %controlName%, %fallbackButton%
    }

    capturePrevDisplay := ""
    captureSlot := 0
    RegisterActivationHotkeys()
}

TryActivateExplorerMode()
{
    global explorerContext, horiz, debugEnabled, lastExplorerWindow

    if (horiz)
        return false

    MouseGetPos, mx, my, winHwnd, controlHwnd, 2
    VarSetCapacity(cursorPt, 8, 0)
    DllCall("GetCursorPos", "ptr", &cursorPt)
    screenX := NumGet(cursorPt, 0, "int")
    screenY := NumGet(cursorPt, 4, "int")
    if (!winHwnd)
    {
        WinGet, activeWin, ID, A
        if (DllCall("IsWindow", "ptr", activeWin))
        {
            winHwnd := activeWin
            if (!controlHwnd)
                controlHwnd := GetWindowFocusHandle(winHwnd)
            if (debugEnabled)
                DebugLog(Format("Explorer activation fallback: using active window=0x{:X}", winHwnd + 0))
        }
        else if (DllCall("IsWindow", "ptr", lastExplorerWindow))
        {
            winHwnd := lastExplorerWindow
            if (!controlHwnd)
                controlHwnd := GetWindowFocusHandle(winHwnd)
            if (debugEnabled)
                DebugLog(Format("Explorer activation fallback: using cached window=0x{:X}", winHwnd + 0))
        }
        else
        {
            if (debugEnabled)
                DebugLog("Explorer activation aborted: no window under cursor")
            return false
        }
    }

    WinGet, procName, ProcessName, ahk_id %winHwnd%
    if (procName = "")
    {
        if (debugEnabled)
            DebugLog(Format("Explorer activation aborted: unable to resolve process for window=0x{:X}", winHwnd + 0))
        return false
    }
    procLower := procName
    StringLower, procLower, procName
    winClass := GetWindowClassName(winHwnd)
    if (procLower != "explorer.exe")
    {
        if (debugEnabled)
            DebugLog(Format("Explorer activation skipped: window=0x{:X} class={} proc={}", winHwnd + 0, winClass, procName))
        explorerContext := {}
        return false
    }

    lastExplorerWindow := winHwnd

    reuseLogged := false
    if (IsObject(explorerContext) && explorerContext.active && explorerContext.HasKey("window"))
    {
        ctxWin := explorerContext.window
        if (ctxWin && ctxWin = winHwnd && DllCall("IsWindow", "ptr", ctxWin))
        {
            folderView := explorerContext.HasKey("folderView") ? explorerContext.folderView : 0
            scrollBar := explorerContext.HasKey("scrollBar") ? explorerContext.scrollBar : 0
            folderOk := folderView && DllCall("IsWindow", "ptr", folderView)
            scrollOk := scrollBar ? DllCall("IsWindow", "ptr", scrollBar) : true
            folderVisible := false
            if (folderOk)
                folderVisible := DllCall("IsWindowVisible", "ptr", folderView) ? true : false
            scrollVisible := true
            if (scrollBar)
                scrollVisible := scrollOk ? (DllCall("IsWindowVisible", "ptr", scrollBar) ? true : false) : false
            if (folderOk && folderVisible && scrollOk && scrollVisible)
            {
                scrollText := scrollBar ? (scrollVisible ? "true" : "false") : "(n/a)"
                if (debugEnabled)
                    DebugLog(Format("Explorer context reused for window=0x{:X} viewVisible=true scrollVisible={}", ctxWin + 0, scrollText))
                return true
            }
            if (debugEnabled)
            {
                scrollText := scrollBar ? (scrollVisible ? "true" : "false") : "(n/a)"
                DebugLog(Format("Explorer context reuse aborted (viewOk={} viewVisible={} scrollOk={} scrollVisible={}); rebuilding", folderOk ? "true" : "false", folderVisible ? "true" : "false", scrollOk ? "true" : "false", scrollText))
            }
            reuseLogged := true
        }
        else
        {
            if (debugEnabled)
            {
                valid := (ctxWin && DllCall("IsWindow", "ptr", ctxWin)) ? "true" : "false"
                DebugLog(Format("Explorer context reuse aborted: storedWindow=0x{:X} valid={} currentWindow=0x{:X}", ctxWin + 0, valid, winHwnd + 0))
            }
            reuseLogged := true
        }
    }

    explorerContext := {}

    if (debugEnabled)
    {
        ctrlClass := controlHwnd ? GetWindowClassName(controlHwnd) : "(none)"
        msgPrefix := reuseLogged ? "Explorer activation attempt (rebuild)" : "Explorer activation attempt"
        DebugLog(Format("{} win=0x{:X} class={} control=0x{:X} class={}", msgPrefix, winHwnd + 0, winClass, controlHwnd + 0, ctrlClass))
    }

    targetHwnd := controlHwnd ? controlHwnd : winHwnd
    context := FindExplorerScrollContext(targetHwnd, screenX, screenY, winHwnd)
    if (!IsObject(context))
    {
        if (debugEnabled)
            DebugLog("Explorer activation failed: no context")
        return false
    }

    if (debugEnabled)
        DebugLogExplorerContext("Explorer context created", context)
    explorerContext := context
    return true
}

ResetExplorerMode()
{
    global explorerContext, debugEnabled

    if (IsObject(explorerContext) && explorerContext.strategy = "scrollbar" && explorerContext.scrollMoved)
    {
        if (DllCall("IsWindow", "ptr", explorerContext.scrollBar))
        {
            SendScrollBarThumb(explorerContext, explorerContext.position, 4)
            SendScrollBarCommand(explorerContext, 8)
            if (debugEnabled)
                DebugLog("ScrollBar mode finalized")
        }
    }
    if (debugEnabled && IsObject(explorerContext) && explorerContext.active)
        DebugLog("Explorer context cleared")
    explorerContext := {}
}

ProcessActivationDragState(rawDelta, currentX := "", currentY := "")
{
    global activationState, activationNativeDown, activationTriggerData
    global activationLastMotionTick, activationIdleRestoreMs, activationMovementDetected
    global activationDragThreshold, explorerContext, debugEnabled
    global wheelBuffer

    if (activationState = "idle")
        return true

    moving := Abs(rawDelta) >= activationDragThreshold
    currentTick := A_TickCount
    triggerData := activationTriggerData
    triggerIsMouse := IsObject(triggerData) ? (triggerData.HasKey("isMouse") ? triggerData.isMouse : IsMouseActivationKey(triggerData.key)) : false

    if (activationState = "pending")
    {
        if (moving)
        {
            activationState := "scrolling"
            activationMovementDetected := true
            activationLastMotionTick := currentTick
            if (triggerIsMouse)
                BeginMouseLock(currentX, currentY)
            if (activationNativeDown || !triggerIsMouse)
            {
                SendActivationUp(activationTriggerData)
                activationNativeDown := false
            }
        }
        else
            return false
    }
    else if (activationState = "scrolling")
    {
        if (moving)
        {
            activationLastMotionTick := currentTick
        }
        else if (triggerIsMouse && !activationNativeDown && (currentTick - activationLastMotionTick) >= activationIdleRestoreMs)
        {
            if (SendActivationDown(activationTriggerData))
            {
                activationNativeDown := true
                activationState := "native_hold"
                wheelBuffer := 0
                ResetExplorerMode()
                EndMouseLock()
                if (debugEnabled)
                    DebugLog("Activation returned to native hold")
                return false
            }
            else
                activationLastMotionTick := currentTick
        }
    }
    else if (activationState = "native_hold")
    {
        if (moving)
        {
            if (activationNativeDown)
            {
                SendActivationUp(activationTriggerData)
                activationNativeDown := false
            }
            activationState := "scrolling"
            activationMovementDetected := true
            activationLastMotionTick := currentTick
            if (triggerIsMouse)
                BeginMouseLock(currentX, currentY)
        }
        else
            return false
    }

    return (activationState = "scrolling")
}

;------------------------
;  Mouse lock helpers
;------------------------
BeginMouseLock(anchorX, anchorY)
{
    global mouseLockEnabled, mouseLockActive
    global mouseLockAnchorX, mouseLockAnchorY, debugEnabled
    global mxLast, myLast

    if (!mouseLockEnabled)
    {
        EndMouseLock()
        return false
    }

    if (anchorX = "" || anchorY = "")
        return false

    anchorX := Round(anchorX)
    anchorY := Round(anchorY)

    mouseLockActive := true
    mouseLockAnchorX := anchorX
    mouseLockAnchorY := anchorY
    mxLast := anchorX
    myLast := anchorY
    HideCursorForLock()

    if (debugEnabled)
        DebugLog(Format("Mouse lock engaged at {}, {}", anchorX, anchorY))

    return true
}

EndMouseLock(clearAnchor := true)
{
    global mouseLockActive
    global mouseLockAnchorX, mouseLockAnchorY, debugEnabled

    activeBefore := mouseLockActive
    mouseLockActive := false
    if (clearAnchor)
    {
        mouseLockAnchorX := ""
        mouseLockAnchorY := ""
    }
    ShowCursorAfterLock()

    if (debugEnabled && activeBefore)
        DebugLog("Mouse lock released")

    return activeBefore
}

HideCursorForLock()
{
    global mouseLockCursorHidden, mouseLockHideCursor, debugEnabled

    if (!mouseLockHideCursor)
    {
        ShowCursorAfterLock()
        return
    }

    if (mouseLockCursorHidden)
        return

    if (!ApplyBlankSystemCursors())
    {
        if (debugEnabled)
            DebugLog("Failed to apply blank system cursors; cursor may remain visible")
        return
    }

    mouseLockCursorHidden := true
}

ShowCursorAfterLock()
{
    global mouseLockCursorHidden, debugEnabled

    if (!mouseLockCursorHidden)
        return

    if (!RestoreSystemCursors() && debugEnabled)
        DebugLog("System cursor restore failed; cursor theme may require manual reset")

    mouseLockCursorHidden := false
}

EnsureBlankCursor()
{
    global mouseLockBlankCursor

    if (mouseLockBlankCursor)
        return mouseLockBlankCursor

    static andMask := ""
    static xorMask := ""

    if (andMask = "")
    {
        VarSetCapacity(andMask, 128, 0xFF)
        VarSetCapacity(xorMask, 128, 0)
    }

    cursor := DllCall("CreateCursor", "ptr", 0, "int", 0, "int", 0, "int", 32, "int", 32, "ptr", &andMask, "ptr", &xorMask, "ptr")
    if (cursor)
        mouseLockBlankCursor := cursor

    return mouseLockBlankCursor
}

ApplyBlankSystemCursors()
{
    global mouseLockSystemCursorApplied, mouseLockBlankCursor
    global mouseLockOriginalCursors, debugEnabled

    if (mouseLockSystemCursorApplied)
        return true

    blank := EnsureBlankCursor()
    if (!blank)
        return false

    ids := [32512, 32513, 32514, 32515, 32516, 32642, 32643, 32644, 32645, 32646, 32648, 32649, 32650, 32651]
    success := true
    changed := false

    for index, cursorId in ids
    {
        if (!mouseLockOriginalCursors.HasKey(cursorId))
        {
            orig := DllCall("LoadCursor", "ptr", 0, "ptr", cursorId, "ptr")
            if (orig)
            {
                copyOrig := DllCall("CopyIcon", "ptr", orig, "ptr")
                if (copyOrig)
                    mouseLockOriginalCursors[cursorId] := copyOrig
            }
        }

        if (!mouseLockOriginalCursors.HasKey(cursorId))
            success := false

        copy := DllCall("CopyIcon", "ptr", blank, "ptr")
        if (!copy)
        {
            success := false
            continue
        }

        if (!DllCall("SetSystemCursor", "ptr", copy, "UInt", cursorId))
        {
            success := false
            DllCall("DestroyCursor", "ptr", copy)
        }
        else
            changed := true
    }

    if (!success)
    {
        if (changed)
            DllCall("SystemParametersInfo", "UInt", 0x0057, "UInt", 0, "ptr", 0, "UInt", 0)
        return false
    }

    mouseLockSystemCursorApplied := true
    return true
}

RestoreSystemCursors()
{
    global mouseLockSystemCursorApplied, mouseLockOriginalCursors

    if (!mouseLockSystemCursorApplied)
        return true

    success := true
    changed := false

    for cursorId, original in mouseLockOriginalCursors
    {
        if (!original)
        {
            success := false
            continue
        }

        copy := DllCall("CopyIcon", "ptr", original, "ptr")
        if (!copy)
        {
            success := false
            continue
        }

        if (!DllCall("SetSystemCursor", "ptr", copy, "UInt", cursorId))
        {
            success := false
            DllCall("DestroyCursor", "ptr", copy)
        }
        else
            changed := true
    }

    if (!success)
    {
        if (!DllCall("SystemParametersInfo", "UInt", 0x0057, "UInt", 0, "ptr", 0, "UInt", 0))
            return false

        for cursorId, original in mouseLockOriginalCursors
        {
            if (original)
                DllCall("DestroyCursor", "ptr", original)
        }
        mouseLockOriginalCursors := {}
    }

    mouseLockSystemCursorApplied := false
    return true
}

HandleExplorerScroll(dy, mx, my)
{
    global explorerContext, scrollsTotal, swap, debugEnabled
    maxAttempts := 2
    attempt := 0
    while (attempt < maxAttempts)
    {
        attempt++

        if (!IsObject(explorerContext) || !explorerContext.active)
        {
            if (debugEnabled)
                DebugLog(Format("Explorer context inactive; attempting rebuild (attempt {} of {})", attempt, maxAttempts))
            if (!TryActivateExplorerMode())
            {
                if (debugEnabled)
                    DebugLog("Explorer context rebuild failed")
                return false
            }
            if (!IsObject(explorerContext) || !explorerContext.active)
                continue
        }

        ctx := explorerContext

        if (debugEnabled)
            DebugLog(Format("HandleExplorerScroll attempt={} strategy={} window=0x{:X}", attempt, ctx.HasKey("strategy") ? ctx.strategy : "?", ctx.HasKey("window") ? (ctx.window + 0) : 0))

        if (ctx.strategy = "listview" && ctx.listView)
        {
            if (!DllCall("IsWindow", "ptr", ctx.listView))
            {
                if (debugEnabled)
                    DebugLog("Explorer listview window invalidated; rebuilding context")
                explorerContext := {}
                continue
            }
            if (!DllCall("IsWindowVisible", "ptr", ctx.listView))
            {
                if (debugEnabled)
                    DebugLog("Explorer listview hidden; rebuilding context")
                explorerContext := {}
                continue
            }
            direction := swap ? -1 : 1
            ctx.pixelAccumulator += dy * direction * ctx.pixelScale
            delta := Round(ctx.pixelAccumulator)
            if (delta != 0)
            {
                ctx.pixelAccumulator -= delta
                if (ctx.pixelMaxStep && Abs(delta) > ctx.pixelMaxStep)
                    delta := (delta > 0) ? ctx.pixelMaxStep : -ctx.pixelMaxStep
                DllCall("SendMessage", "ptr", ctx.listView, "uint", 0x1014, "ptr", 0, "int", delta)
                scrollsTotal += Abs(delta)
                if (debugEnabled)
                    DebugLog(Format("ListView scroll delta={} accum={:.2f}", delta, ctx.pixelAccumulator))
            }
            explorerContext := ctx
            return true
        }

        if (ctx.strategy = "scrollbar" && ctx.scrollBar)
        {
            if (!DllCall("IsWindow", "ptr", ctx.scrollBar))
            {
                if (debugEnabled)
                    DebugLog("ScrollBar handle invalidated; rebuilding context")
                explorerContext := {}
                continue
            }
            if (!DllCall("IsWindowVisible", "ptr", ctx.scrollBar))
            {
                if (debugEnabled)
                    DebugLog("ScrollBar hidden; rebuilding context")
                explorerContext := {}
                continue
            }
            if (ctx.HasKey("folderView") && ctx.folderView && DllCall("IsWindow", "ptr", ctx.folderView) && !DllCall("IsWindowVisible", "ptr", ctx.folderView))
            {
                if (debugEnabled)
                    DebugLog("Explorer folder view hidden; rebuilding context")
                explorerContext := {}
                continue
            }

            direction := swap ? -1 : 1
            ctx.pixelAccumulator += dy * direction

            if (ctx.unitsPerPixel <= 0)
                ctx.unitsPerPixel := 1.0

            unitDelta := ctx.pixelAccumulator * ctx.unitsPerPixel
            deltaUnits := Round(unitDelta)
            if (deltaUnits != 0)
            {
                ctx.pixelAccumulator -= deltaUnits / ctx.unitsPerPixel
                newPos := Clamp(ctx.position + deltaUnits, ctx.minPos, ctx.maxPosEff)
                if (newPos != ctx.position)
                {
                    ApplyScrollBarPosition(ctx, newPos)
                    SendScrollBarThumb(ctx, newPos, 5)
                    ctx.position := newPos
                    ctx.scrollMoved := true
                    scrollsTotal += Abs(deltaUnits)
                    if (debugEnabled)
                        DebugLog(Format("ScrollBar track pos={} delta={} accum={:.2f}", newPos, deltaUnits, ctx.pixelAccumulator))
                }
            }

            explorerContext := ctx
            return true
        }

        if (ctx.strategy != "uia")
        {
            if (debugEnabled)
                DebugLog("Explorer context strategy unknown; rebuilding")
            explorerContext := {}
            continue
        }

        if (ctx.verticalScrollable)
        {
            if (ctx.modeV = "range" && !IsObject(ctx.rangePatternV))
            {
                if (IsObject(ctx.verticalBar))
                {
                    refresh := GetRangeValueData(ctx.verticalBar, ctx.verticalRangeLength)
                    if (IsObject(refresh))
                    {
                        ctx.rangePatternV := refresh.pattern
                        ctx.rangeElementV := refresh.element
                        ctx.rangeMinV := refresh.min
                        ctx.rangeMaxV := refresh.max
                        ctx.rangeValueV := refresh.value
                        if (refresh.perPixel != "")
                            ctx.rangePerPixelV := refresh.perPixel
                        if (refresh.threshold > 0)
                            ctx.rangeTriggerV := Max(Abs(ctx.rangePerPixelV) * 0.5, refresh.threshold)
                        if (refresh.minStep > 0)
                            ctx.rangeMinStepV := refresh.minStep
                        if (debugEnabled)
                            DebugLog("Explorer range pattern refreshed from vertical bar")
                    }
                    else
                    {
                        if (debugEnabled)
                            DebugLog("Explorer range pattern unavailable; switching to percent mode")
                        ctx.modeV := "percent"
                    }
                }
                else
                {
                    if (debugEnabled)
                        DebugLog("Explorer vertical bar missing; switching to percent mode")
                    ctx.modeV := "percent"
                }
            }

            if (ctx.modeV = "range" && IsObject(ctx.rangePatternV))
            {
                ctx.rangeAccumulatorV += dy * ctx.rangePerPixelV
                if (Abs(ctx.rangeAccumulatorV) >= ctx.rangeTriggerV)
                {
                    oldValue := ctx.rangeValueV
                    newValue := Clamp(oldValue + ctx.rangeAccumulatorV, ctx.rangeMinV, ctx.rangeMaxV)
                    if (Abs(newValue - oldValue) >= ctx.rangeMinStepV)
                    {
                        setOk := false
                        attemptRange := 0
                        while (!setOk && attemptRange < 2)
                        {
                            attemptRange++
                            try
                            {
                                ctx.rangePatternV.SetValue(newValue)
                                try
                                    ctx.rangeValueV := ctx.rangePatternV.CurrentValue + 0.0
                                catch
                                    ctx.rangeValueV := newValue
                                ctx.rangeAccumulatorV := 0.0
                                scrollsTotal += 1
                                setOk := true
                                if (debugEnabled)
                                    DebugLog(Format("Explorer range scroll applied value={}", ctx.rangeValueV))
                            }
                            catch
                            {
                                if (attemptRange >= 2)
                                    break
                                ctx.rangePatternV := ""
                                if (!IsObject(ctx.verticalBar))
                                    break
                                refresh := GetRangeValueData(ctx.verticalBar, ctx.verticalRangeLength)
                                if (!IsObject(refresh))
                                    break
                                ctx.rangePatternV := refresh.pattern
                                ctx.rangeElementV := refresh.element
                                ctx.rangeMinV := refresh.min
                                ctx.rangeMaxV := refresh.max
                                ctx.rangeValueV := refresh.value
                                if (refresh.perPixel != "")
                                    ctx.rangePerPixelV := refresh.perPixel
                                if (refresh.threshold > 0)
                                    ctx.rangeTriggerV := Max(Abs(ctx.rangePerPixelV) * 0.5, refresh.threshold)
                                if (refresh.minStep > 0)
                                    ctx.rangeMinStepV := refresh.minStep
                                if (debugEnabled)
                                    DebugLog("Explorer range pattern refresh during retry")
                            }
                        }
                        if (!setOk)
                        {
                            if (debugEnabled)
                                DebugLog("Explorer range scroll failed after retries; falling back to percent mode")
                            ctx.modeV := "percent"
                            ctx.rangePatternV := ""
                        }
                    }
                }
            }

            if (!(ctx.modeV = "range" && IsObject(ctx.rangePatternV)))
            {
                ctx.percentAccumulatorV += dy * ctx.percentPerPixelV
                if (Abs(ctx.percentAccumulatorV) >= ctx.percentTriggerV)
                {
                    oldPercent := ctx.lastVerticalPercent
                    newPercent := Clamp(oldPercent + ctx.percentAccumulatorV, ctx.minVerticalPercent, ctx.maxVerticalPercent)
                    if (Abs(newPercent - oldPercent) >= ctx.percentMinStep)
                    {
                        horizParam := ctx.horizontalScrollable ? ctx.lastHorizontalPercent : -1
                        try
                        {
                            ctx.pattern.SetScrollPercent(horizParam, newPercent)
                            try
                                ctx.lastVerticalPercent := ctx.pattern.CurrentVerticalScrollPercent
                            catch
                                ctx.lastVerticalPercent := newPercent
                            ctx.percentAccumulatorV := 0.0
                            scrollsTotal += 1
                            if (debugEnabled)
                                DebugLog(Format("Explorer percent scroll applied percent={:.2f}", ctx.lastVerticalPercent))
                        }
                        catch
                        {
                            if (debugEnabled)
                                DebugLog("Explorer percent scroll failed; rebuilding context")
                            explorerContext := {}
                            continue
                        }
                    }
                }
            }
        }

        explorerContext := ctx
        return true
    }

    if (debugEnabled)
        DebugLog("Explorer scroll handling exhausted attempts without success")
    return false
}

EnsureUIAutomation()
{
    global uiAutomation, debugEnabled
    if (IsObject(uiAutomation))
        return true

    attempts := []
    attempts.Push({progId: "UIAutomationClient.CUIAutomation", iid: ""})
    attempts.Push({progId: "{FF48DBA4-60EF-4201-AA87-54103EEF594E}", iid: ""})
    attempts.Push({progId: "{FF48DBA4-60EF-4201-AA87-54103EEF594E}", iid: "{30CBE57D-D9D0-452A-AB13-7AC5AC4825EE}"})
    for index, attempt in attempts
    {
        try
        {
            if (attempt.iid != "")
                candidate := ComObjCreate(attempt.progId, attempt.iid)
            else
                candidate := ComObjCreate(attempt.progId)
            if (IsObject(candidate))
            {
                uiAutomation := candidate
                if (debugEnabled)
                    DebugLog("UIAutomation created via " . attempt.progId . (attempt.iid != "" ? " (iid=" . attempt.iid . ")" : ""))
                return true
            }
        }
        catch e
        {
            if (debugEnabled)
                DebugLog("UIAutomation creation failed for " . attempt.progId . (attempt.iid != "" ? " (iid=" . attempt.iid . ")" : "") . ": " . e.Message)
        }
    }

    if (!DllCall("GetModuleHandle", "str", "UIAutomationCore.dll"))
    {
        lib := DllCall("LoadLibrary", "str", "UIAutomationCore.dll", "ptr")
        if (debugEnabled)
            DebugLog(Format("UIAutomationCore.dll LoadLibrary result=0x{:X}", lib + 0))
        if (lib)
        {
            for index, attempt in attempts
            {
                try
                {
                    if (attempt.iid != "")
                        candidate := ComObjCreate(attempt.progId, attempt.iid)
                    else
                        candidate := ComObjCreate(attempt.progId)
                    if (IsObject(candidate))
                    {
                        uiAutomation := candidate
                        if (debugEnabled)
                            DebugLog("UIAutomation created after LoadLibrary via " . attempt.progId . (attempt.iid != "" ? " (iid=" . attempt.iid . ")" : ""))
                        return true
                    }
                }
                catch e
                {
                    if (debugEnabled)
                        DebugLog("UIAutomation creation after LoadLibrary failed for " . attempt.progId . (attempt.iid != "" ? " (iid=" . attempt.iid . ")" : "") . ": " . e.Message)
                }
            }
        }
    }

    uiAutomation := ""
    if (debugEnabled)
        DebugLog("UIAutomation unavailable after all attempts")
    return false
}

FindExplorerScrollContext(targetHwnd, screenX, screenY, winHwnd)
{
    global uiAutomation, debugEnabled
    global wheelMaxStep, wheelSensitivity
    static TreeScope_Subtree := 7
    static UIA_ControlTypePropertyId := 30003
    static UIA_ScrollBarControlTypeId := 50014
    static UIA_ThumbControlTypeId := 50027
    static UIA_OrientationPropertyId := 30023
    static Orientation_Horizontal := 1
    static Orientation_Vertical := 2

    if (debugEnabled)
    {
        targetClass := GetWindowClassName(targetHwnd)
        DebugLog(Format("FindExplorerScrollContext target=0x{:X} class={} win=0x{:X}", targetHwnd + 0, targetClass, winHwnd + 0))
    }

    if (targetHwnd && targetHwnd != winHwnd)
    {
        targetClass := GetWindowClassName(targetHwnd)
        if (debugEnabled)
            DebugLog("Primary target class: " . targetClass)
    }

    viewHwnd := ResolveExplorerViewHandle(targetHwnd, winHwnd)
    if (!viewHwnd)
        viewHwnd := FindExplorerFolderViewHandle(winHwnd, screenX, screenY)

    refinedView := ResolveActiveExplorerView(winHwnd, viewHwnd, screenX, screenY)
    if (refinedView && refinedView != viewHwnd)
    {
        if (debugEnabled)
        {
            refinedClass := GetWindowClassName(refinedView)
            DebugLog(Format("Explorer view refined via hit-test hwnd=0x{:X} class={}", refinedView + 0, refinedClass))
        }
        viewHwnd := refinedView
    }

    if (debugEnabled)
    {
        viewClass := GetWindowClassName(viewHwnd)
        DebugLog(Format("Explorer view resolved to hwnd=0x{:X} class={}", viewHwnd + 0, viewClass))
        DebugLogExplorerTabs(winHwnd, viewHwnd)
    }

    scrollBar := ""
    targetClass := GetWindowClassName(targetHwnd)
    if (targetClass = "ScrollBar")
        scrollBar := targetHwnd
    if (!scrollBar && viewHwnd)
    {
        scopedScroll := FindExplorerScrollBarHandle(viewHwnd, 5, screenX, screenY)
        if (scopedScroll)
            scrollBar := scopedScroll
    }
    if (!scrollBar)
        scrollBar := FindExplorerScrollBarHandle(winHwnd, 6, screenX, screenY)

    if (scrollBar)
    {
        ctx := BuildExplorerScrollBarContext(scrollBar, winHwnd)
        if (IsObject(ctx))
        {
            ctx.folderView := viewHwnd
            ctx.window := winHwnd
            if (debugEnabled)
                DebugLogExplorerContext("ScrollBar context built", ctx)
            return ctx
        }
        else if (debugEnabled)
            DebugLog("ScrollBar context failed to build")
    }

    listView := GetExplorerListViewHandle(winHwnd, viewHwnd, screenX, screenY)
    if (listView)
    {
        ctx := {}
        ctx.active := true
        ctx.strategy := "listview"
        ctx.folderView := viewHwnd
        ctx.listView := listView
        ctx.window := winHwnd
        ctx.pixelAccumulator := 0.0
        ctx.pixelScale := (wheelSensitivity != "") ? Max(0.1, wheelSensitivity / 12.0) : 1.0
        ctx.pixelMaxStep := (wheelMaxStep != "") ? Max(10, Floor(wheelMaxStep / 2)) : 400
        if (debugEnabled)
            DebugLog(Format("ListView strategy selected hwnd=0x{:X}", listView + 0))
        return ctx
    }

    if (debugEnabled)
    {
        DebugLog("ListView not found; attempting UIA")
        DebugDumpWindowTree(winHwnd, 0, 2, 20)
    }

    if (!EnsureUIAutomation())
    {
        if (debugEnabled)
            DebugLog("UIAutomation not available")
        return ""
    }

    element := uiAutomation.ElementFromHandle(targetHwnd)
    if (!IsObject(element) && winHwnd && winHwnd != targetHwnd)
        element := uiAutomation.ElementFromHandle(winHwnd)
    info := FindFirstScrollPattern(element, screenX, screenY)
    if (!IsObject(info) && winHwnd && winHwnd != targetHwnd)
    {
        parentElement := uiAutomation.ElementFromHandle(winHwnd)
        info := FindFirstScrollPattern(parentElement, screenX, screenY)
    }
    if (!IsObject(info))
    {
        if (debugEnabled)
            DebugLog("No UIA scroll pattern found")
        return ""
    }

    pattern := info.pattern
    scroller := info.element

    if (debugEnabled)
        DebugLog("UIA scroll pattern acquired")

    try
        vertScrollable := pattern.CurrentVerticallyScrollable
    catch
        vertScrollable := false
    try
        horizScrollable := pattern.CurrentHorizontallyScrollable
    catch
        horizScrollable := false

    if (!vertScrollable && !horizScrollable)
        return ""

    if (!vertScrollable)
    {
        if (debugEnabled)
            DebugLog("UIA reports not vertically scrollable")
        return ""
    }

    try
    {
        baseVert := pattern.CurrentVerticalScrollPercent
        baseHoriz := pattern.CurrentHorizontalScrollPercent
        viewVert := pattern.CurrentVerticalViewSize
        viewHoriz := pattern.CurrentHorizontalViewSize
    }
    catch
        return ""

    if (debugEnabled)
        DebugLog(Format("UIA initial percent vert={:.2f} view={:.2f}", baseVert + 0.0, viewVert + 0.0))

    if (!scroller)
        return ""

    static condScrollBar := ""
    if (!IsObject(condScrollBar))
        condScrollBar := uiAutomation.CreatePropertyCondition(UIA_ControlTypePropertyId, UIA_ScrollBarControlTypeId)

    scrollBars := scroller.FindAll(TreeScope_Subtree, condScrollBar)
    verticalBar := ""
    horizontalBar := ""
    if (IsObject(scrollBars))
    {
        count := scrollBars.Length
        Loop, %count%
        {
            sb := scrollBars.GetElement(A_Index - 1)
            if (!IsObject(sb))
                continue
            try
                orientation := sb.GetCurrentPropertyValue(UIA_OrientationPropertyId)
            catch
                continue
            if (orientation = Orientation_Vertical && !IsObject(verticalBar))
                verticalBar := sb
            else if (orientation = Orientation_Horizontal && !IsObject(horizontalBar))
                horizontalBar := sb
        }
    }

    static condThumb := ""
    if (!IsObject(condThumb))
        condThumb := uiAutomation.CreatePropertyCondition(UIA_ControlTypePropertyId, UIA_ThumbControlTypeId)

    verticalRange := 0.0
    verticalRangeInfo := ""
    if (vertScrollable && IsObject(verticalBar))
    {
        thumb := verticalBar.FindFirst(TreeScope_Subtree, condThumb)
        verticalRange := GetScrollRange(verticalBar, thumb, true)
        verticalRangeInfo := GetRangeValueData(verticalBar, verticalRange)
    }

    horizontalRange := 0.0
    if (horizScrollable && IsObject(horizontalBar))
    {
        thumbH := horizontalBar.FindFirst(TreeScope_Subtree, condThumb)
        horizontalRange := GetScrollRange(horizontalBar, thumbH, false)
    }

    maxVertical := (viewVert > 0 && viewVert < 100) ? 100 - viewVert : 100
    if (maxVertical < 0)
        maxVertical := 0
    maxHorizontal := (viewHoriz > 0 && viewHoriz < 100) ? 100 - viewHoriz : 100
    if (maxHorizontal < 0)
        maxHorizontal := 0

    ctx := {}
    ctx.active := true
    ctx.strategy := "uia"
    ctx.pattern := pattern
    ctx.scroller := scroller
    ctx.folderView := viewHwnd
    ctx.window := winHwnd
    ctx.verticalScrollable := vertScrollable
    ctx.horizontalScrollable := horizScrollable
    ctx.lastVerticalPercent := (baseVert = "") ? 0.0 : baseVert + 0.0
    ctx.lastHorizontalPercent := (baseHoriz = "") ? 0.0 : baseHoriz + 0.0
    ctx.minVerticalPercent := 0.0
    ctx.minHorizontalPercent := 0.0
    ctx.maxVerticalPercent := maxVertical
    ctx.maxHorizontalPercent := maxHorizontal
    ctx.verticalRange := verticalRange
    ctx.horizontalRange := horizontalRange
    ctx.modeV := "percent"
    ctx.rangePatternV := ""
    ctx.rangeElementV := ""
    ctx.rangeMinV := 0.0
    ctx.rangeMaxV := 0.0
    ctx.rangeValueV := 0.0
    ctx.rangePerPixelV := 0.0
    ctx.rangeTriggerV := 0.1
    ctx.rangeMinStepV := 0.1
    ctx.rangeAccumulatorV := 0.0
    ctx.percentPerPixelV := (verticalRange > 1) ? (maxVertical / verticalRange) : (maxVertical / 300)
    ctx.percentPerPixelH := (horizontalRange > 1) ? (maxHorizontal / horizontalRange) : (maxHorizontal / 300)
    if (ctx.percentPerPixelV = "" || ctx.percentPerPixelV = 0)
        ctx.percentPerPixelV := 0.2
    if (ctx.percentPerPixelH = "" || ctx.percentPerPixelH = 0)
        ctx.percentPerPixelH := 0.2
    ctx.percentTriggerV := Max(Abs(ctx.percentPerPixelV) * 0.5, 0.05)
    ctx.percentTriggerH := Max(Abs(ctx.percentPerPixelH) * 0.5, 0.05)
    ctx.percentMinStep := 0.1
    ctx.percentAccumulatorV := 0.0
    ctx.percentAccumulatorH := 0.0
    ctx.verticalBar := verticalBar
    ctx.horizontalBar := horizontalBar
    ctx.verticalRangeLength := verticalRange

    if (IsObject(verticalRangeInfo))
    {
        ctx.modeV := "range"
        ctx.rangePatternV := verticalRangeInfo.pattern
        ctx.rangeElementV := verticalRangeInfo.element
        ctx.rangeMinV := verticalRangeInfo.min
        ctx.rangeMaxV := verticalRangeInfo.max
        ctx.rangeValueV := verticalRangeInfo.value
        ctx.rangePerPixelV := (verticalRangeInfo.perPixel != "") ? verticalRangeInfo.perPixel : ((ctx.rangeMaxV - ctx.rangeMinV) / Max(verticalRange, 1))
        if (ctx.rangePerPixelV = "" || ctx.rangePerPixelV = 0)
            ctx.rangePerPixelV := 0.2
        ctx.rangeTriggerV := Max(Abs(ctx.rangePerPixelV) * 0.5, verticalRangeInfo.threshold)
        ctx.rangeMinStepV := (verticalRangeInfo.minStep > 0) ? verticalRangeInfo.minStep : Max(Abs(ctx.rangePerPixelV), 0.1)
        ctx.rangeAccumulatorV := 0.0
        if (debugEnabled)
            DebugLog("UIA range-value mode configured")
    }
    else if (debugEnabled)
    {
        DebugLog("UIA falling back to percent mode")
    }

    ctx.startMouseX := screenX
    ctx.startMouseY := screenY

    if (debugEnabled)
        DebugLogExplorerContext("UIA context built", ctx)

    return ctx
}

DebugLogExplorerTabs(winHwnd, viewHwnd)
{
    global debugEnabled
    if (!debugEnabled)
        return

    tabs := EnumerateExplorerTabItems(winHwnd)
    DebugWriteRaw("Explorer tabs count=" . tabs.Length())
    for index, tab in tabs
    {
        line := "  tab[" . index . "] hwnd=0x" . Format("{:X}", tab.hwnd) . " text=" . tab.text
        if (tab.selected)
            line .= " selected=true"
        if (tab.visible)
            line .= " visible=true"
        DebugWriteRaw(line)
    }

    if (viewHwnd)
    {
        family := EnumerateAncestorChain(viewHwnd, winHwnd)
        DebugWriteRaw("View ancestor chain: " . family)
    }
}

EnumerateExplorerTabItems(winHwnd)
{
    tabs := []
    if (!winHwnd)
        return tabs

    directHosts := CollectDescendantsByClass(winHwnd, "DirectUIHWND", 4)
    for index, host in directHosts
    {
        childTabs := EnumerateDirectUITabItems(host)
        for idx, tab in childTabs
            tabs.Push(tab)
    }

    return tabs
}

EnumerateDirectUITabItems(hostHwnd)
{
    items := []
    if (!hostHwnd)
        return items

    static uiaTabItem := 50019 ; UIA_TabItemControlTypeId
    static uiaSelectionItemPattern := 10010
    static uiaNameProperty := 30005

    if (!EnsureUIAutomation())
        return items

    try
        element := uiAutomation.ElementFromHandle(hostHwnd)
    catch
        element := ""
    if (!IsObject(element))
        return items

    try
        trueCondition := uiAutomation.CreateTrueCondition()
    catch
        trueCondition := ""
    if (!IsObject(trueCondition))
        return items

    try
        nodeList := element.FindAll(TreeScope_Subtree := 7, trueCondition)
    catch
        nodeList := ""
    if (!IsObject(nodeList))
        return items

    count := nodeList.Length
    Loop, %count%
    {
        node := nodeList.GetElement(A_Index - 1)
        if (!IsObject(node))
            continue
        try
            controlType := node.CurrentControlType
        catch
            controlType := 0
        if (controlType != uiaTabItem)
            continue

        tab := {}
        try
            tab.hwnd := node.CurrentNativeWindowHandle
        catch
            tab.hwnd := 0
        if (tab.hwnd)
            tab.hwnd := tab.hwnd + 0
        try
            tab.text := node.GetCurrentPropertyValue(uiaNameProperty)
        catch
            tab.text := ""
        if (tab.text = "")
        {
            try
                tab.text := node.CurrentName
            catch
                tab.text := ""
        }
        try
            pattern := node.GetCurrentPattern(uiaSelectionItemPattern)
        catch
            pattern := ""
        if (IsObject(pattern))
        {
            try
                tab.selected := pattern.CurrentIsSelected
            catch
                tab.selected := false
        }
        else
            tab.selected := false
        tab.visible := tab.hwnd ? !!DllCall("IsWindowVisible", "ptr", tab.hwnd) : false
        items.Push(tab)
    }

    return items
}

EnumerateAncestorChain(hwnd, stopHwnd := 0)
{
    chain := []
    while (hwnd)
    {
        className := GetWindowClassName(hwnd)
        chain.Push(Format("0x{:X}:{}", hwnd, className))
        if (stopHwnd && hwnd = stopHwnd)
            break
        hwnd := DllCall("GetParent", "ptr", hwnd, "ptr")
    }
    return chain.Length() ? JoinArray(chain, " -> ") : "(none)"
}

JoinArray(arr, delimiter)
{
    if (!IsObject(arr) || !arr.Length())
        return ""
    out := ""
    Loop, % arr.Length()
    {
        if (A_Index > 1)
            out .= delimiter
        out .= arr[A_Index]
    }
    return out
}

FindFirstScrollPattern(element, screenX := "", screenY := "")
{
    global uiAutomation
    static UIA_ScrollPatternId := 10004
    static TreeScope_Subtree := 7
    if (!IsObject(element))
        return ""

    try
        pattern := element.GetCurrentPattern(UIA_ScrollPatternId)
    catch
        pattern := ""
    if (IsObject(pattern))
        return {pattern: pattern, element: element}

    static trueCondition := ""
    if (!IsObject(trueCondition))
        trueCondition := uiAutomation.CreateTrueCondition()

    children := element.FindAll(TreeScope_Subtree, trueCondition)
    if (!IsObject(children))
        return ""
    count := children.Length
    Loop, %count%
    {
        candidate := children.GetElement(A_Index - 1)
        if (!IsObject(candidate))
            continue
        try
            candidatePattern := candidate.GetCurrentPattern(UIA_ScrollPatternId)
        catch
            candidatePattern := ""
        if (IsObject(candidatePattern))
            return {pattern: candidatePattern, element: candidate}
    }

    return ""
}

GetRangeValueData(element, trackLength)
{
    static UIA_RangeValuePatternId := 10003  
    if (!IsObject(element))
        return ""

    info := FindFirstPattern(element, UIA_RangeValuePatternId)
    if (!IsObject(info))
        return ""

    pattern := info.pattern
    patternEl := info.element

    try
    {
        minVal := pattern.CurrentMinimum + 0.0
        maxVal := pattern.CurrentMaximum + 0.0
        curVal := pattern.CurrentValue + 0.0
        small := pattern.CurrentSmallChange + 0.0
        large := pattern.CurrentLargeChange + 0.0
    }
    catch
        return ""

    span := maxVal - minVal
    perPixel := (trackLength > 0 && span != 0) ? (span / trackLength) : ""
    threshold := 0.05
    if (small > 0)
        threshold := Max(threshold, small * 0.25)
    minStep := (small > 0) ? small : (perPixel != "" && perPixel > 0 ? perPixel : 0.1)

    return {pattern: pattern
        , element: patternEl
        , min: minVal
        , max: maxVal
        , value: curVal
        , small: small
        , large: large
        , perPixel: perPixel
        , threshold: threshold
        , minStep: minStep}
}

FindFirstPattern(element, patternId)
{
    global uiAutomation
    static TreeScope_Subtree := 7
    if (!IsObject(element))
        return ""

    try
        pattern := element.GetCurrentPattern(patternId)
    catch
        pattern := ""
    if (IsObject(pattern))
        return {pattern: pattern, element: element}

    static trueCondition := ""
    if (!IsObject(trueCondition))
        trueCondition := uiAutomation.CreateTrueCondition()

    children := element.FindAll(TreeScope_Subtree, trueCondition)
    if (!IsObject(children))
        return ""

    count := children.Length
    Loop, %count%
    {
        candidate := children.GetElement(A_Index - 1)
        if (!IsObject(candidate))
            continue
        try
            pattern := candidate.GetCurrentPattern(patternId)
        catch
            pattern := ""
        if (IsObject(pattern))
            return {pattern: pattern, element: candidate}
    }

    return ""
}

GetScrollRange(scrollBarEl, thumbEl, isVertical)
{
    if (!IsObject(scrollBarEl))
        return 0.0
    try
        rectBar := scrollBarEl.CurrentBoundingRectangle
    catch
        return 0.0
    if (!IsObject(rectBar))
        return 0.0

    barStart := rectBar[isVertical ? 1 : 0]
    barEnd := rectBar[isVertical ? 3 : 2]
    barLength := barEnd - barStart

    if (IsObject(thumbEl))
    {
        try
            rectThumb := thumbEl.CurrentBoundingRectangle
        catch
            rectThumb := ""
        if (IsObject(rectThumb))
        {
            thumbLength := rectThumb[isVertical ? 3 : 2] - rectThumb[isVertical ? 1 : 0]
            range := barLength - thumbLength
            return (range > 1) ? range : barLength
        }
    }
    return barLength
}

GetExplorerListViewHandle(winHwnd, scopeHwnd := 0, screenX := "", screenY := "")
{
    global debugEnabled
    if (!winHwnd)
        return 0

    candidates := []
    if (scopeHwnd)
    {
        listWithinScope := CollectDescendantsByClass(scopeHwnd, "SysListView32", 4)
        for index, handle in listWithinScope
            candidates.Push(handle)
    }

    if (!candidates.Length())
    {
        defViews := CollectDescendantsByClass(winHwnd, "SHELLDLL_DefView", 4)
        for index, view in defViews
        {
            inside := CollectDescendantsByClass(view, "SysListView32", 3)
            for idx, handle in inside
                candidates.Push(handle)
        }
    }

    if (!candidates.Length())
    {
        directUIs := CollectDescendantsByClass(winHwnd, "DirectUIHWND", 6)
        for index, direct in directUIs
        {
            inside := CollectDescendantsByClass(direct, "SysListView32", 3)
            for idx, handle in inside
                candidates.Push(handle)
        }
    }

    if (!candidates.Length())
        candidates := CollectDescendantsByClass(winHwnd, "SysListView32", 6)

    listView := SelectBestWindowFromList(candidates, false, screenX, screenY)
    if (!listView && debugEnabled)
        DebugLog("SysListView32 not found under Explorer window")
    return listView
}

FindDescendantByClass(parentHwnd, targetClass, maxDepth := 5)
{
    global debugEnabled
    if (!parentHwnd || maxDepth < 0)
        return 0

    child := 0
    while (child := DllCall("FindWindowEx", "ptr", parentHwnd, "ptr", child, "ptr", 0, "ptr", 0))
    {
        className := GetWindowClassName(child)
        if (className = targetClass)
        {
            if (debugEnabled)
                DebugLog(Format("FindDescendantByClass found {} at hwnd=0x{:X}", targetClass, child + 0))
            return child
        }
        found := FindDescendantByClass(child, targetClass, maxDepth - 1)
        if (found)
            return found
    }
    return 0
}

GetWindowRectData(hwnd)
{
    if (!hwnd)
        return ""
    VarSetCapacity(rect, 16, 0)
    if (!DllCall("GetWindowRect", "ptr", hwnd, "ptr", &rect))
        return ""
    left := NumGet(rect, 0, "int")
    top := NumGet(rect, 4, "int")
    right := NumGet(rect, 8, "int")
    bottom := NumGet(rect, 12, "int")
    width := right - left
    height := bottom - top
    return {left: left, top: top, right: right, bottom: bottom, width: width, height: height}
}

RectContainsPointData(rect, x, y)
{
    if (!IsObject(rect))
        return false
    return (x >= rect.left && x <= rect.right && y >= rect.top && y <= rect.bottom)
}

IsWindowTopmostAtPoint(hwnd, x, y)
{
    if (!hwnd)
        return false
    x := Round(x)
    y := Round(y)
    point := ((y & 0xFFFFFFFF) << 32) | (x & 0xFFFFFFFF)
    target := DllCall("WindowFromPoint", "int64", point, "ptr")
    if (!target)
        return false
    if (target = hwnd)
        return true
    return IsWindowDescendant(target, hwnd)
}

GetElementBoundingRect(element)
{
    if (!IsObject(element))
        return ""
    try
        rect := element.CurrentBoundingRectangle
    catch
        rect := ""
    return rect
}

RectContainsPoint(rect, x, y)
{
    if (!IsObject(rect))
        return false
    left := rect[0]
    top := rect[1]
    right := rect[2]
    bottom := rect[3]
    return (x >= left && x <= right && y >= top && y <= bottom)
}

RectHasArea(rect)
{
    if (!IsObject(rect))
        return false
    return (rect[2] > rect[0] && rect[3] > rect[1])
}

CollectDescendantsByClass(parentHwnd, targetClass, maxDepth := 6)
{
    result := []
    __CollectDescendantsRecursive(parentHwnd, targetClass, maxDepth, result)
    return result
}

__CollectDescendantsRecursive(parentHwnd, targetClass, depth, result)
{
    if (!parentHwnd || depth < 0)
        return

    child := 0
    while (child := DllCall("FindWindowEx", "ptr", parentHwnd, "ptr", child, "ptr", 0, "ptr", 0))
    {
        className := GetWindowClassName(child)
        if (className = targetClass)
            result.Push(child)
        __CollectDescendantsRecursive(child, targetClass, depth - 1, result)
    }
}

ResolveExplorerViewHandle(targetHwnd, winHwnd)
{
    bridgeCandidate := 0
    while (targetHwnd && targetHwnd != winHwnd)
    {
        className := GetWindowClassName(targetHwnd)
        if (className = "DirectUIHWND" || className = "SHELLDLL_DefView")
            return targetHwnd
        if (!bridgeCandidate && (className = "Windows.UI.Composition.DesktopWindowContentBridge" || className = "Microsoft.UI.Content.DesktopChildSiteBridge"))
            bridgeCandidate := targetHwnd
        targetHwnd := DllCall("GetParent", "ptr", targetHwnd, "ptr")
    }
    return bridgeCandidate
}

SelectBestWindowFromList(handles, preferVertical := false, screenX := "", screenY := "")
{
    if (!IsObject(handles))
        return 0

    best := 0
    bestScore := -1
    hasPoint := (screenX != "" && screenY != "")
    for index, hwnd in handles
    {
        if (!DllCall("IsWindow", "ptr", hwnd))
            continue
        if (!DllCall("IsWindowVisible", "ptr", hwnd))
            continue
        rect := GetWindowRectData(hwnd)
        if (!IsObject(rect))
            continue
        width := rect.width
        height := rect.height
        if (width <= 0 || height <= 0)
            continue
        score := width * height
        if (preferVertical)
        {
            if (height <= width)
                continue
            score := height * 1000 + width
        }
        if (hasPoint)
        {
            if (RectContainsPointData(rect, screenX, screenY))
            {
                if (IsWindowTopmostAtPoint(hwnd, screenX, screenY))
                    score += 1.0e12
                else
                    score -= 1.0e11
            }
            else
            {
                centerX := (rect.left + rect.right) / 2.0
                centerY := (rect.top + rect.bottom) / 2.0
                dx := centerX - screenX
                dy := centerY - screenY
                dist := Sqrt(dx * dx + dy * dy)
                score -= dist * 1000
            }
        }
        if (score > bestScore)
        {
            best := hwnd
            bestScore := score
        }
    }
    if (best)
        return best
    return (IsObject(handles) && handles.Length() >= 1) ? handles[1] : 0
}

FindExplorerScrollBarHandle(parentHwnd, maxDepth := 6, screenX := "", screenY := "")
{
    global debugEnabled
    if (!parentHwnd)
        return 0

    candidates := CollectDescendantsByClass(parentHwnd, "ScrollBar", maxDepth)
    filtered := []
    for index, hwnd in candidates
    {
        if (!IsVerticalScrollBar(hwnd))
            continue
        filtered.Push(hwnd)
    }

    scrollBar := SelectBestWindowFromList(filtered, true, screenX, screenY)
    if (scrollBar && debugEnabled)
        DebugLog(Format("Found vertical ScrollBar hwnd=0x{:X}", scrollBar + 0))
    return scrollBar
}

IsVerticalScrollBar(hwnd)
{
    if (!hwnd)
        return false

    style := DllCall("GetWindowLong", "ptr", hwnd, "int", -16, "uint")
    static SBS_VERT := 0x1
    if (style & SBS_VERT)
        return true

    VarSetCapacity(rect, 16, 0)
    if (DllCall("GetWindowRect", "ptr", hwnd, "ptr", &rect))
    {
        width := NumGet(rect, 8, "int") - NumGet(rect, 0, "int")
        height := NumGet(rect, 12, "int") - NumGet(rect, 4, "int")
        return (height > width)
    }
    return false
}

BuildExplorerScrollBarContext(scrollBar, winHwnd)
{
    global debugEnabled
    if (!DllCall("IsWindow", "ptr", scrollBar))
        return ""

    parent := DllCall("GetParent", "ptr", scrollBar, "ptr")
    info := GetScrollInfoData(scrollBar, 2)
    barType := 2
    if (!IsObject(info) && parent)
    {
        info := GetScrollInfoData(parent, 1)
        barType := 1
    }
    if (!IsObject(info))
    {
        if (debugEnabled)
            DebugLog("GetScrollInfo failed for ScrollBar")
        return ""
    }

    metrics := GetScrollBarMetrics(scrollBar)
    totalUnits := info.max - info.min
    if (info.page > 0 && totalUnits >= info.page)
        totalUnits := totalUnits - info.page + 1
    if (totalUnits < 1)
        totalUnits := (info.max > info.min) ? (info.max - info.min) : 1

    unitsPerPixel := (metrics.track > 0) ? (totalUnits / metrics.track) : (totalUnits / 200.0)
    if (unitsPerPixel <= 0)
        unitsPerPixel := 1.0

    maxPosEff := info.max
    if (info.page > 0)
    {
        adjust := info.page - 1
        if (adjust > 0)
            maxPosEff := info.max - adjust
    }
    if (maxPosEff < info.min)
        maxPosEff := info.min

    ctx := {}
    ctx.active := true
    ctx.strategy := "scrollbar"
    ctx.scrollBar := scrollBar
    ctx.parent := parent ? parent : winHwnd
    ctx.window := winHwnd
    ctx.minPos := info.min
    ctx.maxPos := info.max
    ctx.maxPosEff := maxPosEff
    ctx.position := info.pos
    ctx.barType := barType
    ctx.unitsPerPixel := unitsPerPixel
    ctx.pixelAccumulator := 0.0
    ctx.scrollMoved := false
    ctx.totalUnits := totalUnits
    ctx.trackLength := metrics.track
    ctx.pageSize := info.page
    ctx.lastTrackPos := info.track

    if (debugEnabled)
        DebugLog(Format("ScrollBar context ready hwnd=0x{:X} parent=0x{:X} pos={} range={} track={} unitsPerPixel={:.3f}", scrollBar + 0, ctx.parent + 0, ctx.position, totalUnits, metrics.track, unitsPerPixel))

    return ctx
}

GetScrollInfoData(hwnd, barType)
{
    if (!hwnd)
        return ""

    VarSetCapacity(si, 28, 0)
    NumPut(28, si, 0, "uint")
    NumPut(0x17, si, 4, "uint")  ; SIF_RANGE | SIF_PAGE | SIF_POS | SIF_TRACKPOS
    if (!DllCall("GetScrollInfo", "ptr", hwnd, "int", barType, "ptr", &si))
        return ""

    info := {}
    info.min := NumGet(si, 8, "int")
    info.max := NumGet(si, 12, "int")
    info.page := NumGet(si, 16, "uint")
    info.pos := NumGet(si, 20, "int")
    info.track := NumGet(si, 24, "int")
    return info
}

GetScrollBarMetrics(scrollBar)
{
    VarSetCapacity(rect, 16, 0)
    if (!DllCall("GetWindowRect", "ptr", scrollBar, "ptr", &rect))
        return {height: 0, track: 0}

    height := NumGet(rect, 12, "int") - NumGet(rect, 4, "int")
    arrow := DllCall("GetSystemMetrics", "int", 20)  ; SM_CYVSCROLL
    track := height - (arrow * 2)
    if (track < 1)
        track := height
    return {height: height, track: track}
}

ApplyScrollBarPosition(ctx, pos)
{
    target := (ctx.barType = 2) ? ctx.scrollBar : ctx.parent
    if (!target)
        target := ctx.scrollBar
    if (!target)
        return

    VarSetCapacity(si, 28, 0)
    NumPut(28, si, 0, "uint")
    NumPut(0x4, si, 4, "uint")  ; SIF_POS
    NumPut(pos, si, 20, "int")
    DllCall("SetScrollInfo", "ptr", target, "int", ctx.barType, "ptr", &si, "int", 1)
}

SendScrollBarThumb(ctx, pos, command)
{
    parent := ctx.parent ? ctx.parent : ctx.scrollBar
    if (!parent)
        return
    lParam := (ctx.barType = 2) ? ctx.scrollBar : 0
    newPos := pos & 0xFFFF
    wParam := (newPos << 16) | (command & 0xFFFF)
    DllCall("SendMessage", "ptr", parent, "uint", 0x0115, "ptr", wParam, "ptr", lParam)
}

SendScrollBarCommand(ctx, command)
{
    parent := ctx.parent ? ctx.parent : ctx.scrollBar
    if (!parent)
        return
    lParam := (ctx.barType = 2) ? ctx.scrollBar : 0
    wParam := command & 0xFFFF
    DllCall("SendMessage", "ptr", parent, "uint", 0x0115, "ptr", wParam, "ptr", lParam)
}

GetWindowClassName(hwnd)
{
    if (!hwnd)
        return ""
    VarSetCapacity(classBuf, 256, 0)
    if (DllCall("GetClassName", "ptr", hwnd, "str", classBuf, "int", 256))
        return classBuf
    return ""
}

Clamp(value, minValue, maxValue)
{
    if (value < minValue)
        return minValue
    if (value > maxValue)
        return maxValue
    return value
}
