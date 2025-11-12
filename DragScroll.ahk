ScrollIndicatorUpdateLayered(data)
{
    global scrollIndicatorGuiHwnd, scrollIndicatorCurrentIcon
    global scrollIndicatorCursorWidth, scrollIndicatorCursorHeight
    global scrollIndicatorHotspotX, scrollIndicatorHotspotY
    global debugEnabled

    if (!scrollIndicatorGuiHwnd)
        return false
    if (!IsObject(data) || !data.HasKey("hbm") || !data.hbm)
        return false

    hdc := DllCall("GetDC", "ptr", 0, "ptr")
    if (!hdc)
        return false

    memDC := DllCall("CreateCompatibleDC", "ptr", hdc, "ptr")
    if (!memDC)
    {
        DllCall("ReleaseDC", "ptr", 0, "ptr", hdc)
        return false
    }

    oldBmp := DllCall("SelectObject", "ptr", memDC, "ptr", data.hbm, "ptr")
    if (!oldBmp)
    {
        DllCall("DeleteDC", "ptr", memDC)
        DllCall("ReleaseDC", "ptr", 0, "ptr", hdc)
        return false
    }

    VarSetCapacity(size, 8, 0)
    NumPut(data.width, size, 0, "int")
    NumPut(data.height, size, 4, "int")

    VarSetCapacity(srcPoint, 8, 0)

    VarSetCapacity(blend, 4, 0)
    NumPut(0x00, blend, 0, "uchar")
    NumPut(0x00, blend, 1, "uchar")
    NumPut(255, blend, 2, "uchar")
    NumPut(0x01, blend, 3, "uchar")

    success := DllCall("UpdateLayeredWindow", "ptr", scrollIndicatorGuiHwnd, "ptr", hdc, "ptr", 0, "ptr", &size, "ptr", memDC, "ptr", &srcPoint, "uint", 0, "ptr", &blend, "uint", 2)

    DllCall("SelectObject", "ptr", memDC, "ptr", oldBmp)
    DllCall("DeleteDC", "ptr", memDC)
    DllCall("ReleaseDC", "ptr", 0, "ptr", hdc)

    if (!success)
    {
        if (debugEnabled)
            DebugLog("Scroll indicator: UpdateLayeredWindow failed (" . A_LastError . ")")
        return false
    }

    scrollIndicatorCurrentIcon := data.icon
    scrollIndicatorCursorWidth := data.width
    scrollIndicatorCursorHeight := data.height
    scrollIndicatorHotspotX := data.hotspotX
    scrollIndicatorHotspotY := data.hotspotY

    return true
}

ScrollIndicatorExtractIconBitmap(hCursor, ByRef width, ByRef height, ByRef hotspotX, ByRef hotspotY)
{
    global debugEnabled

    static iconInfoSize := (A_PtrSize = 8) ? 32 : 20
    maskOffset := (A_PtrSize = 8) ? 16 : 12
    colorOffset := maskOffset + A_PtrSize
    bmInfoSize := (A_PtrSize = 8) ? 32 : 24

    VarSetCapacity(iconInfo, iconInfoSize, 0)
    if (!DllCall("GetIconInfo", "ptr", hCursor, "ptr", &iconInfo))
        return 0

    hotspotX := NumGet(iconInfo, 4, "uint")
    hotspotY := NumGet(iconInfo, 8, "uint")
    hbmMask := NumGet(iconInfo, maskOffset, "ptr")
    hbmColor := NumGet(iconInfo, colorOffset, "ptr")

    if (hbmColor)
    {
        VarSetCapacity(bm, bmInfoSize, 0)
        if (DllCall("GetObject", "ptr", hbmColor, "int", bmInfoSize, "ptr", &bm))
        {
            width := NumGet(bm, 4, "int")
            height := Abs(NumGet(bm, 8, "int"))
        }
    }
    if ((!width || !height) && hbmMask)
    {
        VarSetCapacity(bmMask, bmInfoSize, 0)
        if (DllCall("GetObject", "ptr", hbmMask, "int", bmInfoSize, "ptr", &bmMask))
        {
            maskW := NumGet(bmMask, 4, "int")
            maskH := Abs(NumGet(bmMask, 8, "int"))
            if (!width && maskW > 0)
                width := maskW
            if (!height && maskH > 0)
                height := maskH // 2
        }
    }

    if (width <= 0)
        width := 32
    if (height <= 0)
        height := 32

    hbm := ScrollIndicatorDrawIconBitmap(hCursor, width, height)
    if (!hbm && debugEnabled)
        DebugLog("Scroll indicator: DrawIconEx failed; icon will be hidden")

    if (hbmMask)
        DllCall("DeleteObject", "ptr", hbmMask)
    if (hbmColor)
        DllCall("DeleteObject", "ptr", hbmColor)

    return hbm
}

ScrollIndicatorDrawIconBitmap(hIcon, width, height)
{
    if (!hIcon)
        return 0

    hdc := DllCall("GetDC", "ptr", 0, "ptr")
    if (!hdc)
        return 0

    memDC := DllCall("CreateCompatibleDC", "ptr", hdc, "ptr")
    if (!memDC)
    {
        DllCall("ReleaseDC", "ptr", 0, "ptr", hdc)
        return 0
    }

    VarSetCapacity(bi, 40, 0)
    NumPut(40, bi, 0, "uint")
    NumPut(width, bi, 4, "int")
    actualHeight := Abs(height)
    NumPut(-actualHeight, bi, 8, "int")
    NumPut(1, bi, 12, "ushort")
    NumPut(32, bi, 14, "ushort")
    NumPut(0, bi, 16, "uint")
    stride := ((width * 32 + 31) // 32) * 4
    NumPut(stride * actualHeight, bi, 20, "uint")

    bits := 0
    hbm := DllCall("CreateDIBSection", "ptr", hdc, "ptr", &bi, "uint", 0, "ptr*", bits, "ptr", 0, "uint", 0, "ptr")
    if (!hbm)
    {
        DllCall("DeleteDC", "ptr", memDC)
        DllCall("ReleaseDC", "ptr", 0, "ptr", hdc)
        return 0
    }

    oldBmp := DllCall("SelectObject", "ptr", memDC, "ptr", hbm, "ptr")
    if (!oldBmp)
    {
        DllCall("DeleteObject", "ptr", hbm)
        DllCall("DeleteDC", "ptr", memDC)
        DllCall("ReleaseDC", "ptr", 0, "ptr", hdc)
        return 0
    }

    if (bits)
        DllCall("msvcrt\memset", "ptr", bits, "int", 0, "uptr", stride * actualHeight)
    drawSuccess := DllCall("User32.dll\DrawIconEx", "ptr", memDC, "int", 0, "int", 0, "ptr", hIcon, "int", width, "int", height, "uint", 0, "ptr", 0, "uint", 0x0003)
    if (drawSuccess)
        ScrollIndicatorPremultiplyAlpha(bits, width, actualHeight)

    DllCall("SelectObject", "ptr", memDC, "ptr", oldBmp)
    if (!drawSuccess)
    {
        DllCall("DeleteObject", "ptr", hbm)
        hbm := 0
    }
    DllCall("DeleteDC", "ptr", memDC)
    DllCall("ReleaseDC", "ptr", 0, "ptr", hdc)

    return hbm
}

ScrollIndicatorPremultiplyAlpha(bits, width, height)
{
    if (!bits)
        return

    if (width <= 0 || height <= 0)
        return

    total := width * height
    offset := 0
    Loop, %total%
    {
        addr := bits + offset
        alpha := NumGet(addr, 3, "uchar")
        if (alpha = 0)
        {
            NumPut(0, addr, 0, "uint")
        }
        else if (alpha < 255)
        {
            blue := NumGet(addr, 0, "uchar")
            green := NumGet(addr, 1, "uchar")
            red := NumGet(addr, 2, "uchar")
            NumPut((blue * alpha) // 255, addr, 0, "uchar")
            NumPut((green * alpha) // 255, addr, 1, "uchar")
            NumPut((red * alpha) // 255, addr, 2, "uchar")
        }
        offset += 4
    }
}
/*
Mouse Scroll v04 (extended)
Original by Mikhail V., 2021
Enhancements: configurable GUI, high-resolution wheel control, process exclusions
*/

#SingleInstance Force
    global swap, scrollMode, k, wheelSensitivity, wheelMaxStep
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
    global debugHotkeysEnabled, debugStartEnabled, mouseLockEnabled, mouseLockHideCursor
    global multiVerticalBias, multiHorizontalBias, multiActivationThreshold
    global scrollIndicatorEnabled
    global scrollSmoothingEnabled, scrollSmoothingFactor
    global SCROLL_MODE_VERTICAL, SCROLL_MODE_HORIZONTAL, SCROLL_MODE_MULTI
SetBatchLines, -1
SetWorkingDir, %A_ScriptDir%

global configFile := A_ScriptDir . "\mouse-scroll.ini"

    scrollMode := SCROLL_MODE_VERTICAL
global running := 0
global passThrough := false
global swap := 0
global SCROLL_MODE_VERTICAL := 0
global SCROLL_MODE_HORIZONTAL := 1
global SCROLL_MODE_MULTI := 2
global scrollMode := SCROLL_MODE_VERTICAL
global k := 1.0
global wheelSensitivity := 12.0
global wheelMaxStep := 480
global scanInterval := 20
    multiVerticalBias := 1.0
    multiHorizontalBias := 1.0
    multiActivationThreshold := 1.0

global wheelBuffer := 0.0
global wheelBufferHoriz := 0.0
global scrollVelocityY := 0.0
global scrollVelocityX := 0.0
global scrollsTotal := 0
global activationButton := "MButton"
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
global activationTriggerIdentity := ""
global activationTriggerIsHold := false
global activationLastMotionTick := 0
global activationIdleRestoreMs := 220
global activationDragThreshold := 1
global activationPendingStartTick := 0
global activationHoldDecisionDelay := 35
global activationHoldGraceDeadline := 0
global activationHoldActive := false
global activationHoldRepeatTimerActive := false
global activationHoldRepeatDelay := 275
global activationHoldRepeatInterval := 35
global activationHotkeySuspended := false
global explorerContext := {}
global explorerScaleCache := {}
global lastExplorerWindow := 0
global uiAutomation := ""
global debugEnabled := false
global debugLogDefault := A_ScriptDir . "\dragscroll-debug.log"
global debugLogFile := debugLogDefault
global debugHotkeysEnabled := 0
global debugStartEnabled := 0
global debugLogRedirected := false
global debugLogWarned := false

global multiVerticalBias := 1.0
global multiHorizontalBias := 1.0
global multiActivationThreshold := 1.0

global scrollIndicatorEnabled := 0
global scrollIndicatorGuiCreated := false
global scrollIndicatorGuiVisible := false
global scrollIndicatorGuiHwnd := 0
global scrollIndicatorCursorWidth := 32
global scrollIndicatorCursorHeight := 32
global scrollIndicatorHotspotX := 16
global scrollIndicatorHotspotY := 16
global scrollIndicatorLastX := ""
global scrollIndicatorLastY := ""
global scrollIndicatorIconCache := {}
global scrollIndicatorCurrentMode := ""
global scrollIndicatorCurrentIcon := 0

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
global scrollSmoothingEnabled := 1
global scrollSmoothingFactor := 0.65

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
    global running, scrollMode, k, mxLast, myLast
    global wheelBuffer, wheelBufferHoriz, wheelSensitivity, wheelMaxStep
    global scrollVelocityY, scrollVelocityX
    global scrollSmoothingEnabled, scrollSmoothingFactor
    global swap, scrollsTotal, debugEnabled
    global mouseLockEnabled, mouseLockActive, mouseLockAnchorX, mouseLockAnchorY
    global multiVerticalBias, multiHorizontalBias, multiActivationThreshold
    global scrollIndicatorEnabled, scrollIndicatorGuiVisible

    if (!running)
        return

    if (!GetCursorPoint(mx, my))
    {
        if (debugEnabled)
            DebugLog("ScrollTick: GetCursorPoint failed; skipping frame")
        return
    }
    lockApplied := (mouseLockEnabled && mouseLockActive)

    deltaX := 0.0
    deltaY := 0.0

    if (lockApplied)
    {
        deltaX := mx - mouseLockAnchorX
        deltaY := my - mouseLockAnchorY
        if (deltaX != 0 || deltaY != 0)
            DllCall("SetCursorPos", "int", mouseLockAnchorX, "int", mouseLockAnchorY)
        mx := mouseLockAnchorX
        my := mouseLockAnchorY
        mxLast := mouseLockAnchorX
        myLast := mouseLockAnchorY
    }
    else
    {
        deltaX := mx - mxLast
        deltaY := my - myLast
    }

    verticalInput := 0.0
    horizontalInput := 0.0
    rawActivationDelta := 0.0

    if (scrollMode = SCROLL_MODE_VERTICAL)
    {
        rawDelta := ClampRawDelta(deltaY)
        rawActivationDelta := rawDelta
        verticalInput := k * rawDelta
        if (!lockApplied)
            myLast := my
    }
    else if (scrollMode = SCROLL_MODE_HORIZONTAL)
    {
        rawDelta := ClampRawDelta(deltaX)
        rawActivationDelta := rawDelta
        horizontalInput := k * rawDelta
        if (!lockApplied)
            mxLast := mx
    }
    else
    {
        rawVertical := ClampRawDelta(deltaY)
        rawHorizontal := ClampRawDelta(deltaX)
        if (Abs(rawVertical) >= Abs(rawHorizontal))
            rawActivationDelta := rawVertical
        else
            rawActivationDelta := rawHorizontal

        adjVertical := rawVertical * multiVerticalBias
        adjHorizontal := -rawHorizontal * multiHorizontalBias

        if (Abs(adjVertical) >= multiActivationThreshold)
            verticalInput := k * adjVertical
        if (Abs(adjHorizontal) >= multiActivationThreshold)
            horizontalInput := k * adjHorizontal

        if (!lockApplied)
        {
            mxLast := mx
            myLast := my
        }
    }

    if (!ProcessActivationDragState(rawActivationDelta, mx, my))
        return

    if (mouseLockEnabled && mouseLockActive)
    {
        mx := mouseLockAnchorX
        my := mouseLockAnchorY
    }

    smoothingEnabled := (scrollSmoothingEnabled ? true : false)
    if (smoothingEnabled)
    {
        smoothingFactor := scrollSmoothingFactor
        if (smoothingFactor < 0)
            smoothingFactor := 0
        else if (smoothingFactor >= 0.999)
            smoothingFactor := 0.999
        velocityBlend := 1.0 - smoothingFactor
        scrollVelocityY := scrollVelocityY * smoothingFactor + verticalInput * velocityBlend
        scrollVelocityX := scrollVelocityX * smoothingFactor + horizontalInput * velocityBlend
        if (Abs(scrollVelocityY) < 0.0001)
            scrollVelocityY := 0.0
        if (Abs(scrollVelocityX) < 0.0001)
            scrollVelocityX := 0.0
        verticalEffective := scrollVelocityY
        horizontalEffective := scrollVelocityX
    }
    else
    {
        scrollVelocityY := 0.0
        scrollVelocityX := 0.0
        verticalEffective := verticalInput
        horizontalEffective := horizontalInput
    }

    if (verticalEffective != 0)
    {
        if (HandleExplorerScroll(verticalEffective, mx, my))
            verticalEffective := 0
    }

    if (verticalEffective != 0 && ShouldSuppressForTopGuard(mx, my))
    {
        ResetWheelBuffers()
        return
    }

    if (verticalEffective != 0 || (smoothingEnabled && scrollVelocityY != 0))
    {
        wheelBuffer += verticalEffective * wheelSensitivity
        delta := Round(wheelBuffer)
        if (delta != 0)
        {
            wheelBuffer -= delta
            if (Abs(delta) > wheelMaxStep)
            {
                clamped := (delta > 0 ? wheelMaxStep : -wheelMaxStep)
                wheelBuffer += (delta - clamped)
                delta := clamped
            }
            sendVal := swap ? delta : -delta
            DllCall("mouse_event", "UInt", 0x0800, "UInt", 0, "UInt", 0, "Int", sendVal, "Ptr", 0)
            scrollsTotal += Abs(delta)
        }
    }

    if (horizontalEffective != 0 || (smoothingEnabled && scrollVelocityX != 0))
    {
        wheelBufferHoriz += horizontalEffective * wheelSensitivity
        deltaH := Round(wheelBufferHoriz)
        if (deltaH != 0)
        {
            wheelBufferHoriz -= deltaH
            if (Abs(deltaH) > wheelMaxStep)
            {
                clampedH := (deltaH > 0 ? wheelMaxStep : -wheelMaxStep)
                wheelBufferHoriz += (deltaH - clampedH)
                deltaH := clampedH
            }
            sendValH := swap ? deltaH : -deltaH
            DllCall("mouse_event", "UInt", 0x01000, "UInt", 0, "UInt", 0, "Int", sendValH, "Ptr", 0)
            scrollsTotal += Abs(deltaH)
        }
    }

    if (scrollIndicatorEnabled && scrollIndicatorGuiVisible)
        UpdateScrollIndicatorPosition(mx, my)
return

;------------------------
;  Settings management
;------------------------
IniSerializeMultiline(value)
{
    if (value = "")
        return ""
    value := StrReplace(value, "`r`n", "`n")
    value := StrReplace(value, "`r", "`n")
    value := StrReplace(value, "%", "%25")
    return StrReplace(value, "`n", "%0A")
}

IniDeserializeMultiline(value)
{
    if (value = "")
        return ""
    value := StrReplace(value, "%0A", "`n")
    value := StrReplace(value, "%25", "%")
    return StrReplace(value, "`n", "`r`n")
}

IniReadMultiline(filePath, section, key, defaultValue := "")
{
    if (!FileExist(filePath))
        return defaultValue

    FileRead, rawContent, %filePath%
    if (ErrorLevel)
        return defaultValue

    StringLower, sectionLower, section
    foundSection := false
    foundKey := false
    value := ""
    Loop, Parse, rawContent, `n, `r
    {
        line := A_LoopField
        if (line = "")
        {
            if (foundKey)
                break
            else
                continue
        }

        firstChar := SubStr(line, 1, 1)
        if (firstChar = ";")
            continue

        if (firstChar = "[")
        {
            closeIdx := InStr(line, "]")
            if (closeIdx <= 1)
                continue
            sectionName := SubStr(line, 2, closeIdx - 2)
            StringLower, sectionNameLower, sectionName
            if (foundKey)
                break
            foundSection := (sectionNameLower = sectionLower)
            continue
        }

        if (!foundSection)
            continue

        trimmed := LTrim(line)
        keyPrefix := key . "="
        if (!foundKey)
        {
            if (SubStr(trimmed, 1, StrLen(keyPrefix)) = keyPrefix)
            {
                value := SubStr(trimmed, StrLen(keyPrefix) + 1)
                foundKey := true
            }
            continue
        }

        if (InStr(trimmed, "=") || SubStr(trimmed, 1, 1) = "[")
            break

        value .= "`n" . trimmed
    }

    if (!foundKey)
        return defaultValue
    return value
}

LoadSettings()
{
    global configFile
    global swap, scrollMode, k, wheelSensitivity, wheelMaxStep
    global activationButton, activationButton2, processListText, scanInterval, topGuardListText
    global debugHotkeysEnabled, debugStartEnabled, mouseLockEnabled, mouseLockHideCursor
    global multiVerticalBias, multiHorizontalBias, multiActivationThreshold
    global scrollSmoothingEnabled, scrollSmoothingFactor, scrollIndicatorEnabled
    global SCROLL_MODE_VERTICAL, SCROLL_MODE_HORIZONTAL, SCROLL_MODE_MULTI

    swap := 0
    scrollMode := SCROLL_MODE_VERTICAL
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
    multiVerticalBias := 1.0
    multiHorizontalBias := 1.0
    multiActivationThreshold := 1.0
    scrollSmoothingEnabled := 1
    scrollSmoothingFactor := 0.65
    scrollIndicatorEnabled := 0

    if (!FileExist(configFile))
        return

    IniRead, swap, %configFile%, Settings, Swap, %swap%
    IniRead, scrollMode, %configFile%, Settings, ScrollMode, %scrollMode%
    IniRead, k, %configFile%, Settings, SpeedMultiplier, %k%
    IniRead, wheelSensitivity, %configFile%, Settings, WheelSensitivity, %wheelSensitivity%
    IniRead, wheelMaxStep, %configFile%, Settings, WheelMaxStep, %wheelMaxStep%
    IniRead, activationButton, %configFile%, Settings, ActivationButton, %activationButton%
    IniRead, activationButton2, %configFile%, Settings, ActivationButtonSlot2, %activationButton2%
    processListText := IniReadMultiline(configFile, "Settings", "ExcludedProcesses", processListText)
    topGuardListText := IniReadMultiline(configFile, "Settings", "TopGuardZones", topGuardListText)
    IniRead, scanInterval, %configFile%, Settings, ScanInterval, %scanInterval%
    IniRead, debugHotkeysEnabled, %configFile%, Settings, EnableDebugHotkeys, %debugHotkeysEnabled%
    if (ErrorLevel)
        IniRead, debugHotkeysEnabled, %configFile%, Settings, DebugShortcutsEnabled, %debugHotkeysEnabled%
    IniRead, debugStartEnabled, %configFile%, Settings, DebugStartEnabled, %debugStartEnabled%
    IniRead, mouseLockEnabled, %configFile%, Settings, MouseLockEnabled, %mouseLockEnabled%
    IniRead, mouseLockHideCursor, %configFile%, Settings, MouseLockHideCursor, %mouseLockHideCursor%
    IniRead, multiVerticalBias, %configFile%, Settings, MultiVerticalBias, %multiVerticalBias%
    IniRead, multiHorizontalBias, %configFile%, Settings, MultiHorizontalBias, %multiHorizontalBias%
    IniRead, multiActivationThreshold, %configFile%, Settings, MultiActivationThreshold, %multiActivationThreshold%
    IniRead, scrollSmoothingEnabled, %configFile%, Settings, ScrollSmoothingEnabled, %scrollSmoothingEnabled%
    IniRead, scrollSmoothingFactor, %configFile%, Settings, ScrollSmoothingFactor, %scrollSmoothingFactor%
    IniRead, scrollIndicatorEnabled, %configFile%, Settings, ScrollIndicatorEnabled, %scrollIndicatorEnabled%

    scrollMode := scrollMode + 0
    if (scrollMode != SCROLL_MODE_VERTICAL && scrollMode != SCROLL_MODE_HORIZONTAL && scrollMode != SCROLL_MODE_MULTI)
    {
        legacyHoriz := 0
        IniRead, legacyHoriz, %configFile%, Settings, Horizontal, 0
        legacyHoriz := legacyHoriz + 0
        scrollMode := (legacyHoriz = 1) ? SCROLL_MODE_HORIZONTAL : SCROLL_MODE_VERTICAL
    }

    multiVerticalBias := multiVerticalBias + 0.0
    multiHorizontalBias := multiHorizontalBias + 0.0
    multiActivationThreshold := multiActivationThreshold + 0.0

    debugHotkeysEnabled := debugHotkeysEnabled ? 1 : 0
    debugStartEnabled := debugStartEnabled ? 1 : 0
    mouseLockEnabled := mouseLockEnabled ? 1 : 0
    mouseLockHideCursor := mouseLockHideCursor ? 1 : 0
    scrollIndicatorEnabled := scrollIndicatorEnabled ? 1 : 0
    if (!debugStartEnabled && debugHotkeysEnabled)
        debugStartEnabled := 1

    processListText := IniDeserializeMultiline(processListText)
    topGuardListText := IniDeserializeMultiline(topGuardListText)
}

SaveSettings()
{
    global configFile
    global swap, scrollMode, k, wheelSensitivity, wheelMaxStep
    global activationButton, activationButton2, processListText, scanInterval, topGuardListText
    global debugHotkeysEnabled, debugStartEnabled, mouseLockEnabled, mouseLockHideCursor
    global multiVerticalBias, multiHorizontalBias, multiActivationThreshold
    global scrollSmoothingEnabled, scrollSmoothingFactor, scrollIndicatorEnabled
    global SCROLL_MODE_HORIZONTAL, SCROLL_MODE_VERTICAL, SCROLL_MODE_MULTI

    IniWrite, %swap%, %configFile%, Settings, Swap
    IniWrite, %scrollMode%, %configFile%, Settings, ScrollMode
    legacyHoriz := (scrollMode = SCROLL_MODE_HORIZONTAL) ? 1 : 0
    IniWrite, %legacyHoriz%, %configFile%, Settings, Horizontal
    IniWrite, %k%, %configFile%, Settings, SpeedMultiplier
    IniWrite, %wheelSensitivity%, %configFile%, Settings, WheelSensitivity
    IniWrite, %wheelMaxStep%, %configFile%, Settings, WheelMaxStep
    IniWrite, %activationButton%, %configFile%, Settings, ActivationButton
    IniWrite, %activationButton2%, %configFile%, Settings, ActivationButtonSlot2
    IniWrite, %multiVerticalBias%, %configFile%, Settings, MultiVerticalBias
    IniWrite, %multiHorizontalBias%, %configFile%, Settings, MultiHorizontalBias
    IniWrite, %multiActivationThreshold%, %configFile%, Settings, MultiActivationThreshold
    procSerialized := IniSerializeMultiline(processListText)
    guardSerialized := IniSerializeMultiline(topGuardListText)
    IniDelete, %configFile%, Settings, ExcludedProcesses
    IniDelete, %configFile%, Settings, TopGuardZones
    IniWrite, %procSerialized%, %configFile%, Settings, ExcludedProcesses
    IniWrite, %guardSerialized%, %configFile%, Settings, TopGuardZones
    IniWrite, %scanInterval%, %configFile%, Settings, ScanInterval
    IniWrite, %debugHotkeysEnabled%, %configFile%, Settings, EnableDebugHotkeys
    IniWrite, %debugStartEnabled%, %configFile%, Settings, DebugStartEnabled
    IniWrite, %mouseLockEnabled%, %configFile%, Settings, MouseLockEnabled
    IniWrite, %mouseLockHideCursor%, %configFile%, Settings, MouseLockHideCursor
    IniWrite, %scrollSmoothingEnabled%, %configFile%, Settings, ScrollSmoothingEnabled
    IniWrite, %scrollSmoothingFactor%, %configFile%, Settings, ScrollSmoothingFactor
    IniWrite, %scrollIndicatorEnabled%, %configFile%, Settings, ScrollIndicatorEnabled
}

ApplySettings()
{
    global swap, scrollMode, k, wheelSensitivity, wheelMaxStep
    global activationButton, activationButton2, currentHotkey, currentHotkey2, scanInterval
    global debugHotkeysEnabled, debugStartEnabled, mouseLockEnabled, mouseLockHideCursor
    global mouseLockActive
    global multiVerticalBias, multiHorizontalBias, multiActivationThreshold
    global scrollSmoothingEnabled, scrollSmoothingFactor
    global SCROLL_MODE_VERTICAL, SCROLL_MODE_HORIZONTAL, SCROLL_MODE_MULTI

    swap := swap ? 1 : 0
    scrollMode := scrollMode + 0
    if (scrollMode != SCROLL_MODE_VERTICAL && scrollMode != SCROLL_MODE_HORIZONTAL && scrollMode != SCROLL_MODE_MULTI)
        scrollMode := SCROLL_MODE_VERTICAL
    k := (k = "" ? 1.0 : k + 0.0)
    if (k = 0)
        k := 1.0

    wheelSensitivity := (wheelSensitivity = "" ? 12.0 : wheelSensitivity + 0.0)
    if (wheelSensitivity = 0)
        wheelSensitivity := 12.0

    wheelMaxStep := Floor(wheelMaxStep)
    if (wheelMaxStep < 120)
        wheelMaxStep := 120

    multiVerticalBias := (multiVerticalBias = "" ? 1.0 : multiVerticalBias + 0.0)
    if (multiVerticalBias <= 0)
        multiVerticalBias := 0.1
    if (multiVerticalBias > 10)
        multiVerticalBias := 10

    multiHorizontalBias := (multiHorizontalBias = "" ? 1.0 : multiHorizontalBias + 0.0)
    if (multiHorizontalBias <= 0)
        multiHorizontalBias := 0.1
    if (multiHorizontalBias > 10)
        multiHorizontalBias := 10

    multiActivationThreshold := (multiActivationThreshold = "" ? 1.0 : multiActivationThreshold + 0.0)
    if (multiActivationThreshold < 0)
        multiActivationThreshold := 0.0
    if (multiActivationThreshold > 50)
        multiActivationThreshold := 50

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
    scrollSmoothingEnabled := scrollSmoothingEnabled ? 1 : 0
    scrollSmoothingFactor := (scrollSmoothingFactor = "" ? 0.65 : scrollSmoothingFactor + 0.0)
    if (scrollSmoothingFactor < 0)
        scrollSmoothingFactor := 0.0
    if (scrollSmoothingFactor > 0.95)
        scrollSmoothingFactor := 0.95
    scrollIndicatorEnabled := scrollIndicatorEnabled ? 1 : 0

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

    ResetWheelBuffers()
    if (!scrollIndicatorEnabled)
        HideScrollIndicator()
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

SuspendActivationTriggerHotkey(suspend := true)
{
    global activationHotkeySuspended, activationTriggerData

    desired := suspend ? true : false
    if (activationHotkeySuspended = desired)
        return

    if (!IsObject(activationTriggerData))
    {
        if (!desired)
            activationHotkeySuspended := false
        return
    }

    target := activationTriggerData.HasKey("registerSpec") ? activationTriggerData.registerSpec : ""
    if (target = "")
    {
        if (!desired)
            activationHotkeySuspended := false
        return
    }

    state := desired ? "Off" : "On"
    Hotkey, % target, ActivationButtonDown, %state%
    activationHotkeySuspended := desired
}

StartActivationHoldRepeat()
{
    global activationHoldRepeatTimerActive, activationHoldRepeatDelay

    if (activationHoldRepeatTimerActive)
        return

    activationHoldRepeatTimerActive := true
    SetTimer, ActivationHoldRepeatTimer, Off
    SetTimer, ActivationHoldRepeatTimer, % -activationHoldRepeatDelay
}

StopActivationHoldRepeat()
{
    global activationHoldRepeatTimerActive

    if (!activationHoldRepeatTimerActive)
        return

    activationHoldRepeatTimerActive := false
    SetTimer, ActivationHoldRepeatTimer, Off
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

ActivationButtonDown:
    global captureMode, passThrough, activationHotkeyData, activationTriggerData
    global activationState, activationNativeDown, activationMovementDetected
    global activationLastMotionTick, running, wheelBuffer, mxLast, myLast
    global activationTriggerIdentity, debugEnabled
    global activationTriggerIsHold, activationHotkeySuspended
    global activationPendingStartTick, activationHoldGraceDeadline, activationHoldDecisionDelay
    global activationHoldActive

    if (captureMode)
        return

    if (activationState != "idle")
        return

    identity := GetHotkeyIdentity(A_ThisHotkey)
    if (identity = "")
        return

    if (!activationHotkeyData.HasKey(identity))
        return

    data := activationHotkeyData[identity]
    activationTriggerIdentity := identity
    activationTriggerData := data
    activationTriggerIsHold := (!data.isMouse && IsHoldCapableKey(data.key)) ? true : false
    StopActivationHoldRepeat()
    SuspendActivationTriggerHotkey(false)

    if (ShouldBlockProcess())
    {
        passThrough := true
        running := false
        activationState := "idle"
        activationMovementDetected := false
        activationLastMotionTick := 0
        activationPendingStartTick := 0
        activationHoldGraceDeadline := 0
        activationHoldActive := false
        ResetWheelBuffers()
        activationNativeDown := false
        eventDelivered := false
        if (data.isMouse)
        {
            activationNativeDown := SendActivationDown(data) ? true : false
            eventDelivered := activationNativeDown
        }
        else
        {
            if (activationTriggerIsHold)
            {
                activationNativeDown := SendActivationDown(data) ? true : false
                eventDelivered := activationNativeDown
                if (activationNativeDown)
                {
                    activationHoldActive := true
                    activationHoldGraceDeadline := 0
                    StartActivationHoldRepeat()
                }
            }
            else
                eventDelivered := SendActivationTap(data)
        }
        if (!eventDelivered)
            SendActivationTap(data)
        if (debugEnabled)
            DebugLog("Activation pass-through: " . identity)
        if (activationTriggerIsHold)
            SuspendActivationTriggerHotkey(true)
        return
    }

    passThrough := false
    ResetWheelBuffers()
    activationMovementDetected := false
    activationState := "pending"
    activationLastMotionTick := A_TickCount
    activationNativeDown := false
    activationPendingStartTick := A_TickCount
    activationHoldGraceDeadline := activationTriggerIsHold ? (A_TickCount + activationHoldDecisionDelay) : 0
    activationHoldActive := false

    if (GetCursorPoint(mx, my))
    {
        mxLast := mx
        myLast := my
    }

    running := true
    if (debugEnabled)
        DebugLog("Activation start: " . identity)
return

ActivationButtonUp:
    global passThrough, activationTriggerData, activationNativeDown, activationState
    global activationMovementDetected, running, activationTriggerIdentity
    global activationHotkeyData, activationLastMotionTick, wheelBuffer, debugEnabled
    global activationTriggerIsHold, activationHotkeySuspended
    global activationPendingStartTick, activationHoldGraceDeadline, activationHoldActive

    identity := GetHotkeyIdentity(A_ThisHotkey)
    if (identity = "" && activationTriggerIdentity != "")
        identity := activationTriggerIdentity

    HideScrollIndicator()

    triggerData := ""
    if (IsObject(activationTriggerData))
        triggerData := activationTriggerData
    else if (identity != "" && activationHotkeyData.HasKey(identity))
        triggerData := activationHotkeyData[identity]

    if (passThrough)
    {
        if (IsObject(triggerData) && activationNativeDown)
            SendActivationUp(triggerData)
        SuspendActivationTriggerHotkey(false)
        passThrough := false
        activationNativeDown := false
        activationState := "idle"
        activationMovementDetected := false
        activationTriggerData := ""
        activationTriggerIdentity := ""
        activationLastMotionTick := 0
        ResetWheelBuffers()
        activationTriggerIsHold := false
        activationHoldActive := false
        activationHoldGraceDeadline := 0
        activationPendingStartTick := 0
        StopActivationHoldRepeat()
        running := false
        if (debugEnabled)
            DebugLog("Activation pass-through end: " . identity)
        return
    }

    running := false

    if (IsObject(triggerData))
    {
        if (activationState = "pending")
        {
            if (activationNativeDown)
            {
                SendActivationUp(triggerData)
                activationNativeDown := false
                if (!activationTriggerIsHold)
                    SendActivationTap(triggerData)
            }
            else if (!activationTriggerIsHold)
                SendActivationTap(triggerData)
        }
        else if (activationNativeDown)
        {
            SendActivationUp(triggerData)
            activationNativeDown := false
        }
    }
    else
        activationNativeDown := false

    EndMouseLock()
    ResetExplorerMode()

    passThrough := false
    activationState := "idle"
    activationMovementDetected := false
    SuspendActivationTriggerHotkey(false)
    activationTriggerData := ""
    activationTriggerIdentity := ""
    activationLastMotionTick := 0
    ResetWheelBuffers()
    activationNativeDown := false
    activationTriggerIsHold := false
    activationHoldActive := false
    activationHoldGraceDeadline := 0
    activationPendingStartTick := 0
    StopActivationHoldRepeat()

    if (debugEnabled)
        DebugLog("Activation end: " . identity)
return

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
    global guiBuilt, swap, scrollMode, k, wheelSensitivity, wheelMaxStep
    global activationButton, activationButton2, processListText, topGuardListText
    global debugHotkeysEnabled, debugStartEnabled, mouseLockEnabled, mouseLockHideCursor
    global multiVerticalBias, multiHorizontalBias, multiActivationThreshold
    global SCROLL_MODE_VERTICAL, SCROLL_MODE_HORIZONTAL, SCROLL_MODE_MULTI
    global GuiModeVertical, GuiModeHorizontal, GuiModeMulti
    global GuiMultiVertBias, GuiMultiHorizBias, GuiMultiThreshold
    global GuiMultiVertLabel, GuiMultiHorizLabel, GuiMultiThresholdLabel
    global scrollSmoothingEnabled, scrollSmoothingFactor, scrollIndicatorEnabled

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

    modeVertical := (scrollMode = SCROLL_MODE_VERTICAL) ? 1 : 0
    modeHorizontal := (scrollMode = SCROLL_MODE_HORIZONTAL) ? 1 : 0
    modeMulti := (scrollMode = SCROLL_MODE_MULTI) ? 1 : 0

    Gui, Config:Add, Text, xm Section, Scroll mode:
    Gui, Config:Add, Radio, xm+20 vGuiModeVertical gGuiModeChanged Checked%modeVertical% Group, Vertical (up/down)
    Gui, Config:Add, Radio, xm+20 vGuiModeHorizontal gGuiModeChanged Checked%modeHorizontal%, Horizontal (pan)
    Gui, Config:Add, Radio, xm+20 vGuiModeMulti gGuiModeChanged Checked%modeMulti%, Multi-direction (dual axis)

    Gui, Config:Add, Text, xm, Multi-mode tuning:
    Gui, Config:Add, Text, xm+20 vGuiMultiVertLabel, Vertical bias:
    Gui, Config:Add, Edit, x+10 yp w80 vGuiMultiVertBias, %multiVerticalBias%
    Gui, Config:Add, Text, xm+20 vGuiMultiHorizLabel, Horizontal bias:
    Gui, Config:Add, Edit, x+10 yp w80 vGuiMultiHorizBias, %multiHorizontalBias%
    Gui, Config:Add, Text, xm+20 vGuiMultiThresholdLabel, Activation threshold (pixels):
    Gui, Config:Add, Edit, x+10 yp w80 vGuiMultiThreshold, %multiActivationThreshold%

    Gui, Config:Add, Checkbox, xm vGuiMouseLock Checked%mouseLockEnabled%, Lock mouse cursor while scrolling
    Gui, Config:Add, Checkbox, vGuiMouseLockHide Checked%mouseLockHideCursor%, Hide cursor when active
    Gui, Config:Add, Checkbox, xm vGuiScrollIndicatorEnabled Checked%scrollIndicatorEnabled%, Show scroll indicator while active
    Gui, Config:Add, Checkbox, xm vGuiDebugHotkeys Checked%debugHotkeysEnabled%, Enable debug hotkeys (Win+Ctrl+D / Win+Ctrl+L)
    Gui, Config:Add, Checkbox, vGuiDebugStart Checked%debugStartEnabled%, Start with debug logging enabled

    Gui, Config:Add, Text,, Speed multiplier:
    Gui, Config:Add, Edit, vGuiK w150, %k%

    Gui, Config:Add, Text,, Wheel sensitivity (delta per pixel):
    Gui, Config:Add, Edit, vGuiWheelSens w150, %wheelSensitivity%

    Gui, Config:Add, Checkbox, xm vGuiSmoothEnabled Checked%scrollSmoothingEnabled%, Enable scroll smoothing
    Gui, Config:Add, Text,, Smoothing factor (0 - 0.95):
    Gui, Config:Add, Edit, vGuiSmoothFactor w150, %scrollSmoothingFactor%

    Gui, Config:Add, Text,, Maximum wheel step:
    Gui, Config:Add, Edit, vGuiWheelMax w150, %wheelMaxStep%

    Gui, Config:Add, Text,, Excluded processes (one per line or comma separated):
    Gui, Config:Add, Edit, vGuiProcessList w260 h70, %processListText%

    Gui, Config:Add, Text,, Top guard zones (process:height px):
    Gui, Config:Add, Edit, vGuiTopGuardList w260 h70, %topGuardListText%

    Gui, Config:Add, Button, xm w80 gGuiApplySettings Default, Apply
    Gui, Config:Add, Button, x+10 w80 gGuiCloseConfig, Close

    guiBuilt := true
    GuiUpdateScrollModeControls()
}

ShowConfig:
    BuildConfigGui()
    Gui, Config:Show
return

GuiModeChanged:
    global GuiModeVertical, GuiModeHorizontal, GuiModeMulti
    global GuiMultiVertBias, GuiMultiHorizBias, GuiMultiThreshold
    global GuiMultiVertLabel, GuiMultiHorizLabel, GuiMultiThresholdLabel
    Gui, Config:Submit, NoHide
    GuiUpdateScrollModeControls()
return

GuiUpdateScrollModeControls()
{
    GuiControlGet, modeMulti, Config:, GuiModeMulti
    state := modeMulti ? "Enable" : "Disable"
    GuiControl, Config:%state%, GuiMultiVertLabel
    GuiControl, Config:%state%, GuiMultiVertBias
    GuiControl, Config:%state%, GuiMultiHorizLabel
    GuiControl, Config:%state%, GuiMultiHorizBias
    GuiControl, Config:%state%, GuiMultiThresholdLabel
    GuiControl, Config:%state%, GuiMultiThreshold
}

GuiRefreshMultiModeFields()
{
    global multiVerticalBias, multiHorizontalBias, multiActivationThreshold
    GuiControl, Config:, GuiMultiVertBias, %multiVerticalBias%
    GuiControl, Config:, GuiMultiHorizBias, %multiHorizontalBias%
    GuiControl, Config:, GuiMultiThreshold, %multiActivationThreshold%
}

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

    global swap, scrollMode, k, wheelSensitivity, wheelMaxStep
    global activationButton, activationButton2, processListText, topGuardListText
    global debugHotkeysEnabled, debugStartEnabled, mouseLockEnabled, mouseLockHideCursor
    global multiVerticalBias, multiHorizontalBias, multiActivationThreshold
    global SCROLL_MODE_VERTICAL, SCROLL_MODE_HORIZONTAL, SCROLL_MODE_MULTI
    global GuiModeVertical, GuiModeHorizontal, GuiModeMulti
    global GuiMultiVertBias, GuiMultiHorizBias, GuiMultiThreshold
    global GuiMultiVertLabel, GuiMultiHorizLabel, GuiMultiThresholdLabel
    global scrollSmoothingEnabled, scrollSmoothingFactor, scrollIndicatorEnabled

    if (captureMode)
    {
        ShowTempTooltip("Finish capture first.", 1000)
        return
    }

    activationButton := NormalizeActivationHotkey(GuiActivationDisplay1)
    activationButton2 := NormalizeActivationHotkey(GuiActivationDisplay2)
    swap := GuiSwap
    if (GuiModeMulti)
        scrollMode := SCROLL_MODE_MULTI
    else if (GuiModeHorizontal)
        scrollMode := SCROLL_MODE_HORIZONTAL
    else
        scrollMode := SCROLL_MODE_VERTICAL
    debugHotkeysEnabled := GuiDebugHotkeys ? 1 : 0
    debugStartEnabled := GuiDebugStart ? 1 : 0
    mouseLockEnabled := GuiMouseLock ? 1 : 0
    mouseLockHideCursor := GuiMouseLockHide ? 1 : 0
    scrollIndicatorEnabled := GuiScrollIndicatorEnabled ? 1 : 0
    k := GuiK
    wheelSensitivity := GuiWheelSens
    wheelMaxStep := GuiWheelMax
    scrollSmoothingEnabled := GuiSmoothEnabled ? 1 : 0
    scrollSmoothingFactor := GuiSmoothFactor
    processListText := GuiProcessList
    topGuardListText := GuiTopGuardList
    multiVerticalBias := GuiMultiVertBias + 0.0
    multiHorizontalBias := GuiMultiHorizBias + 0.0
    multiActivationThreshold := GuiMultiThreshold + 0.0

    GuiControl, Config:, GuiActivationDisplay1, %activationButton%
    GuiControl, Config:, GuiActivationDisplay2, %activationButton2%

    ApplySettings()
    modeVert := (scrollMode = SCROLL_MODE_VERTICAL) ? 1 : 0
    modeHoriz := (scrollMode = SCROLL_MODE_HORIZONTAL) ? 1 : 0
    modeMulti := (scrollMode = SCROLL_MODE_MULTI) ? 1 : 0
    GuiControl, Config:, GuiModeVertical, %modeVert%
    GuiControl, Config:, GuiModeHorizontal, %modeHoriz%
    GuiControl, Config:, GuiModeMulti, %modeMulti%
    GuiRefreshMultiModeFields()
    GuiUpdateScrollModeControls()
    GuiControl, Config:, GuiSmoothEnabled, %scrollSmoothingEnabled%
    GuiControl, Config:, GuiSmoothFactor, %scrollSmoothingFactor%
    GuiControl, Config:, GuiScrollIndicatorEnabled, %scrollIndicatorEnabled%
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

ExplorerDebugLogScrollMetrics(strategy, data := "")
{
    global debugEnabled
    if (!debugEnabled)
        return

    msg := "Explorer scroll [" . strategy . "]"
    if (IsObject(data))
    {
        first := true
        for key, value in data
        {
            if (first)
            {
                msg .= " "
                first := false
            }
            else
                msg .= " "
            msg .= key . "=" . value
        }
    }
    else if (data != "")
        msg .= " " . data
    DebugLog(msg)
}

ExplorerUpdateObservedUnits(ByRef ctx, unitsDelta, pointerDelta, source := "", atBoundary := false, requestedUnits := "")
{
    if (!IsObject(ctx))
        return
    pointer := pointerDelta + 0.0
    units := unitsDelta + 0.0
    if (pointer = 0 || units = 0)
        return

    requested := ""
    if (requestedUnits != "")
        requested := requestedUnits + 0.0

    if (atBoundary)
    {
        ctx.Delete("observedSampleUnits")
        ctx.Delete("observedSamplePointer")
        ctx.Delete("observedSampleCount")
        ctx.Delete("observedSampleDir")
        return
    }

    if (requested != "")
    {
        mismatch := Abs(units - requested)
        if (mismatch > 0)
        {
            allowed := Max(Abs(requested) * 0.08, (source = "observed_uia") ? 0.2 : 4.0)
            if (mismatch > allowed)
                return
        }
    }

    direction := (pointer > 0) ? 1 : -1
    lastDir := ctx.HasKey("observedSampleDir") ? ctx.observedSampleDir : 0
    if (lastDir != 0 && direction != lastDir)
    {
        ctx.observedSampleUnits := 0.0
        ctx.observedSamplePointer := 0.0
        ctx.observedSampleCount := 0
    }
    ctx.observedSampleDir := direction

    unitsAbs := Abs(units)
    pointerAbs := Abs(pointer)
    ctx.observedSampleUnits := (ctx.HasKey("observedSampleUnits") ? ctx.observedSampleUnits : 0.0) + unitsAbs
    ctx.observedSamplePointer := (ctx.HasKey("observedSamplePointer") ? ctx.observedSamplePointer : 0.0) + pointerAbs
    ctx.observedSampleCount := (ctx.HasKey("observedSampleCount") ? ctx.observedSampleCount : 0) + 1

    sampleUnits := ctx.observedSampleUnits
    samplePointer := ctx.observedSamplePointer

    if (samplePointer < 60 || sampleUnits < 20)
        return

    observed := sampleUnits / samplePointer
    if (observed <= 0)
        observed := 0.0001

    if (source != "")
    {
        ctx.observedSource := source
        ctx.scaleSource := source
        if (ctx.HasKey("rangeBoostPointer") && !InStr(ctx.scaleSource, "+range"))
            ctx.scaleSource .= "+range"
    }

    prior := ctx.HasKey("observedUnitsPerPixel") ? (ctx.observedUnitsPerPixel + 0.0) : 0.0
    if (prior > 0)
    {
        diffRatio := Abs(observed - prior) / prior
        if (diffRatio <= 0.03)
        {
            ctx.observedUnitsTick := A_TickCount
            ctx.observedPointerPx := samplePointer
            ctx.observedPersistent := true
            ctx.unitsPerPixel := prior
            ExplorerStoreScaleCache(ctx, prior, samplePointer)
            ctx.observedSampleUnits := 0.0
            ctx.observedSamplePointer := 0.0
            ctx.observedSampleCount := 0
            return
        }
    }

    weight := 0.15
    if (prior > 0)
        blended := (prior * (1.0 - weight)) + (observed * weight)
    else
        blended := observed
    blended := Max(0.0001, blended)

    ctx.observedUnitsPerPixel := blended
    ctx.observedUnitsTick := A_TickCount
    ctx.observedPointerPx := samplePointer
    ctx.observedPersistent := true
    ctx.unitsPerPixel := blended
    if (source != "")
    {
        ctx.scaleSource := source
        if (ctx.HasKey("rangeBoostPointer") && !InStr(ctx.scaleSource, "+range"))
            ctx.scaleSource .= "+range"
    }
    ctx.scaleSignature := Format("{:.6f}|obs|{:.1f}", blended, samplePointer)
    ExplorerStoreScaleCache(ctx, blended, samplePointer)

    ctx.observedSampleUnits := 0.0
    ctx.observedSamplePointer := 0.0
    ctx.observedSampleCount := 0
}

ExplorerGetScaleCacheKey(ByRef ctx)
{
    if (!IsObject(ctx))
        return ""

    strategy := ctx.HasKey("strategy") ? ctx.strategy : ""
    folder := ctx.HasKey("folderView") ? (ctx.folderView + 0) : 0
    if (!folder && ctx.HasKey("scrollTargetHwnd"))
        folder := ctx.scrollTargetHwnd + 0
    if (!folder || strategy = "")
        return ""

    if (strategy = "scrollbar")
    {
        total := ctx.HasKey("totalUnits") ? Round(ctx.totalUnits + 0.0) : 0
        track := ctx.HasKey("trackLength") ? Round(ctx.trackLength + 0.0) : 0
        page := ctx.HasKey("pageSize") ? Round(ctx.pageSize + 0.0) : 0
        if (total <= 0 || track <= 0)
            return ""
        return Format("scrollbar|0x{:X}|{}|{}|{}", folder, total, track, page)
    }
    else if (strategy = "uia")
    {
        range := ctx.HasKey("maxPosEff") ? Round(ctx.maxPosEff + 0.0) : 0
        view := ctx.HasKey("verticalViewSize") ? Round(ctx.verticalViewSize + 0.0) : 0
        if (range <= 0)
            return ""
        return Format("uia|0x{:X}|{}|{}", folder, range, view)
    }

    return ""
}

ExplorerApplyCachedScale(ByRef ctx)
{
    global explorerScaleCache
    if (!IsObject(ctx))
        return false

    if (!IsObject(explorerScaleCache))
        explorerScaleCache := {}

    key := ExplorerGetScaleCacheKey(ctx)
    if (key = "" || !explorerScaleCache.HasKey(key))
        return false

    entry := explorerScaleCache[key]
    if (!IsObject(entry) || !entry.HasKey("units"))
        return false

    units := entry.units + 0.0
    if (units <= 0)
        return false

    baseSource := (IsObject(entry) && entry.HasKey("source") && entry.source != "") ? entry.source : "observed"
    ctx.cachedSource := baseSource
    ctx.observedUnitsPerPixel := units
    ctx.observedUnitsTick := A_TickCount
    ctx.observedSource := "cache"
    ctx.observedPointerPx := (IsObject(entry) && entry.HasKey("pointer")) ? (entry.pointer + 0.0) : ""
    ctx.observedPersistent := true
    ctx.scaleSource := "cache"
    ctx.unitsPerPixel := units
    ctx.scaleCacheKey := key
    return true
}

ExplorerStoreScaleCache(ByRef ctx, units, pointer)
{
    global explorerScaleCache
    if (!IsObject(ctx))
        return

    key := ExplorerGetScaleCacheKey(ctx)
    if (key = "")
        return

    if (!IsObject(explorerScaleCache))
        explorerScaleCache := {}

    entry := {}
    entry.units := units + 0.0
    entry.pointer := pointer + 0.0
    if (ctx.HasKey("observedSource") && ctx.observedSource = "cache")
        entry.source := ctx.HasKey("cachedSource") ? ctx.cachedSource : "observed"
    else
        entry.source := ctx.HasKey("observedSource") ? ctx.observedSource : ""
    entry.timestamp := A_TickCount
    explorerScaleCache[key] := entry
    ctx.cachedSource := entry.source
    ctx.scaleCacheKey := key
    ExplorerPruneScaleCache()
}

ExplorerComputeRangeBoostUnits(ByRef ctx, currentUnits)
{
    if (!IsObject(ctx) || currentUnits <= 0)
        return currentUnits

    if (!ctx.HasKey("totalUnits"))
        return currentUnits

    totalUnits := ctx.totalUnits + 0.0
    if (totalUnits <= 0 || totalUnits > 1500)
        return currentUnits

    viewPixels := ctx.HasKey("viewPixels") ? (ctx.viewPixels + 0.0) : 0.0
    if (viewPixels <= 0 && ctx.HasKey("trackLength"))
        viewPixels := ctx.trackLength + 0.0
    if (viewPixels <= 0)
        viewPixels := 900.0

    targetPointer := Clamp(viewPixels * 0.35, 180.0, 360.0)
    ctx.rangeBoostPointer := targetPointer
    desiredUnits := totalUnits / targetPointer
    if (desiredUnits <= 0)
        return currentUnits

    minIncrease := currentUnits * 1.03
    if (desiredUnits <= minIncrease)
        return currentUnits

    return desiredUnits
}

ExplorerPruneScaleCache(maxEntries := 20)
{
    global explorerScaleCache
    if (!IsObject(explorerScaleCache))
        return

    count := 0
    for key in explorerScaleCache
        count++
    if (count <= maxEntries)
        return

    oldestKey := ""
    oldestTick := 0
    for key, entry in explorerScaleCache
    {
        tick := (IsObject(entry) && entry.HasKey("timestamp")) ? (entry.timestamp + 0) : 0
        if (oldestKey = "" || tick < oldestTick)
        {
            oldestKey := key
            oldestTick := tick
        }
    }

    if (oldestKey != "")
        explorerScaleCache.Delete(oldestKey)
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
        vert := ctx.HasKey("verticalScrollable") ? (ctx.verticalScrollable ? "true" : "false") : "?"
        target := (ctx.HasKey("scrollTargetHwnd") && ctx.scrollTargetHwnd) ? Format("0x{:X}", ctx.scrollTargetHwnd + 0) : "0"
    wheelBuf := ctx.HasKey("wheelBuffer") ? Format("{:.2f}", ctx.wheelBuffer + 0.0) : "?"
        pending := ctx.HasKey("pendingFocus") ? (ctx.pendingFocus ? "true" : "false") : "?"
        return "strategy=uia vert=" . vert . " target=" . target . " wheelBuf=" . wheelBuf . " focusPending=" . pending
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

ResetWheelBuffers()
{
    global wheelBuffer, wheelBufferHoriz
    global scrollVelocityY, scrollVelocityX
    wheelBuffer := 0.0
    wheelBufferHoriz := 0.0
    scrollVelocityY := 0.0
    scrollVelocityX := 0.0
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
    HideScrollIndicator()
    DestroyScrollIndicatorGui()
    DestroyScrollIndicatorResources()
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
                UpdateExplorerScrollUnitScale(explorerContext)
                explorerContext.pendingFocus := true
                explorerContext.lastFocusTick := 0
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
    EnsureExplorerContextFocus(explorerContext, true, screenX, screenY)
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
    global wheelBuffer, activationTriggerIsHold, activationHotkeySuspended
    global activationPendingStartTick, activationHoldDecisionDelay, activationHoldGraceDeadline, activationHoldActive
    global scrollMode

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
            activationPendingStartTick := 0
            BeginMouseLock(currentX, currentY)
            if (activationNativeDown)
            {
                SendActivationUp(activationTriggerData)
                activationNativeDown := false
                activationHoldActive := false
                activationHoldGraceDeadline := 0
                StopActivationHoldRepeat()
                SuspendActivationTriggerHotkey(false)
            }
            else if (!triggerIsMouse && !activationTriggerIsHold && IsObject(triggerData))
                SendActivationTap(triggerData)

        }
        else
        {
            if (activationTriggerIsHold && !activationNativeDown)
            {
                if (!activationPendingStartTick)
                    activationPendingStartTick := activationLastMotionTick
                if (activationHoldGraceDeadline && currentTick >= activationHoldGraceDeadline && IsObject(triggerData))
                {
                    if (SendActivationDown(triggerData))
                    {
                        activationNativeDown := true
                        activationLastMotionTick := currentTick
                        activationPendingStartTick := 0
                        activationHoldGraceDeadline := 0
                        activationHoldActive := true
                        StartActivationHoldRepeat()
                        SuspendActivationTriggerHotkey(true)
                    }
                    else
                        activationTriggerIsHold := false
                }
            }
            return false
        }
    }
    else if (activationState = "scrolling")
    {
        if (moving)
        {
            activationLastMotionTick := currentTick
        }
        else if (!triggerIsMouse && activationTriggerIsHold && !activationNativeDown && (currentTick - activationLastMotionTick) >= activationIdleRestoreMs)
        {
            if (SendActivationDown(activationTriggerData))
            {
                activationNativeDown := true
                activationState := "native_hold"
                ResetWheelBuffers()
                ResetExplorerMode()
                EndMouseLock()
                activationHoldActive := true
                activationHoldGraceDeadline := 0
                StartActivationHoldRepeat()
                SuspendActivationTriggerHotkey(true)
                if (debugEnabled)
                    DebugLog("Activation returned to native hold")
                HideScrollIndicator()
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
                activationHoldActive := false
                activationHoldGraceDeadline := 0
                StopActivationHoldRepeat()
            }
            activationState := "scrolling"
            activationMovementDetected := true
            activationLastMotionTick := currentTick
            activationPendingStartTick := 0
            BeginMouseLock(currentX, currentY)
            SuspendActivationTriggerHotkey(false)
            StopActivationHoldRepeat()
        }
        else
            return false
    }

    if (activationState = "scrolling")
    {
        if (!EnsureScrollIndicatorVisible(scrollMode, currentX, currentY))
            HideScrollIndicator()
    }
    else
        HideScrollIndicator()

    return (activationState = "scrolling")
}

ActivationHoldRepeatTimer:
    global activationHoldRepeatTimerActive, activationHoldRepeatInterval
    global activationNativeDown, activationTriggerIsHold, activationTriggerData
    global activationHoldActive

    if (!activationHoldRepeatTimerActive)
        return

    if (!activationNativeDown || !activationTriggerIsHold || !activationHoldActive || !IsObject(activationTriggerData))
    {
        activationHoldRepeatTimerActive := false
        SetTimer, ActivationHoldRepeatTimer, Off
        return
    }

    SendActivationDown(activationTriggerData)
    SetTimer, ActivationHoldRepeatTimer, % activationHoldRepeatInterval
return

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

;------------------------
;  Scroll indicator helpers
;------------------------
EnsureScrollIndicatorGui()
{
    global scrollIndicatorGuiCreated, scrollIndicatorGuiVisible
    global scrollIndicatorGuiHwnd
    global debugEnabled

    if (scrollIndicatorGuiCreated)
        return true

    Gui, ScrollIndicator:New, +LastFound +AlwaysOnTop -Caption +ToolWindow +E0x80000 +OwnDialogs -DPIScale
    Gui, ScrollIndicator:Margin, 0, 0
    Gui, ScrollIndicator:+HwndscrollIndicatorGuiHwnd
    Gui, ScrollIndicator:Color, 000000
    Gui, ScrollIndicator:Show, Hide w1 h1
    if (scrollIndicatorGuiHwnd)
    {
        exGet := (A_PtrSize = 8) ? "GetWindowLongPtr" : "GetWindowLong"
        exSet := (A_PtrSize = 8) ? "SetWindowLongPtr" : "SetWindowLong"
        exStyle := DllCall(exGet, "ptr", scrollIndicatorGuiHwnd, "int", -20, "ptr")
        exStyle |= 0x00080000  ; WS_EX_LAYERED
        exStyle |= 0x00000020  ; WS_EX_TRANSPARENT
        exStyle |= 0x08000000  ; WS_EX_NOACTIVATE
        DllCall(exSet, "ptr", scrollIndicatorGuiHwnd, "int", -20, "ptr", exStyle)
    }
    flags := 0x0001 | 0x0002 | 0x0004 | 0x0010 | 0x0200
    if (!DllCall("SetWindowPos", "ptr", scrollIndicatorGuiHwnd, "ptr", -1, "int", 0, "int", 0, "int", 0, "int", 0, "uint", flags))
    {
        if (debugEnabled)
            DebugLog("Scroll indicator: SetWindowPos(HWND_TOPMOST) failed (" . A_LastError . ")")
    }

    scrollIndicatorGuiCreated := true
    scrollIndicatorGuiVisible := false
    if (debugEnabled)
        DebugLog(Format("Scroll indicator GUI created (hwnd=0x{:X})", scrollIndicatorGuiHwnd + 0))
    return true
}

EnsureScrollIndicatorVisible(mode, x := "", y := "")
{
    global scrollIndicatorEnabled, scrollIndicatorCurrentIcon, scrollIndicatorCurrentMode
    global scrollIndicatorCursorWidth, scrollIndicatorCursorHeight
    global scrollIndicatorHotspotX, scrollIndicatorHotspotY
    global scrollIndicatorGuiVisible
    global debugEnabled

    if (!scrollIndicatorEnabled)
        return false

    if (!EnsureScrollIndicatorGui())
        return false

    data := ScrollIndicatorEnsureIcon(mode)
    if (!IsObject(data))
    {
        if (debugEnabled)
            DebugLog("Scroll indicator: cursor data unavailable for mode " . mode)
        return false
    }

    hIcon := data.HasKey("icon") ? data.icon : 0
    if (!hIcon)
    {
        if (debugEnabled)
            DebugLog("Scroll indicator: icon handle missing for mode " . mode)
        return false
    }

    if (!ScrollIndicatorUpdateLayered(data))
        return false

    if (!UpdateScrollIndicatorPosition(x, y, true))
        return false

    scrollIndicatorCurrentMode := mode
    return true
}

UpdateScrollIndicatorPosition(x := "", y := "", forceShow := false)
{
    global scrollIndicatorGuiCreated, scrollIndicatorGuiVisible, scrollIndicatorGuiHwnd
    global scrollIndicatorHotspotX, scrollIndicatorHotspotY
    global scrollIndicatorLastX, scrollIndicatorLastY
    global debugEnabled

    if (!scrollIndicatorGuiCreated)
        return false

    if (x = "" || y = "")
    {
        if (!GetCursorPoint(px, py))
            return false
        x := px
        y := py
    }

    xPos := Round(x - scrollIndicatorHotspotX)
    yPos := Round(y - scrollIndicatorHotspotY)

    if (scrollIndicatorGuiVisible)
    {
        if (scrollIndicatorLastX != xPos || scrollIndicatorLastY != yPos)
        {
            moveFlags := 0x0001 | 0x0004 | 0x0010 | 0x0200
            if (!DllCall("SetWindowPos", "ptr", scrollIndicatorGuiHwnd, "ptr", 0, "int", xPos, "int", yPos, "int", 0, "int", 0, "uint", moveFlags))
            {
                if (debugEnabled)
                    DebugLog("Scroll indicator: SetWindowPos(move) failed (" . A_LastError . ")")
            }
            scrollIndicatorLastX := xPos
            scrollIndicatorLastY := yPos
            if (debugEnabled)
                DebugLog(Format("Scroll indicator moved to {}, {}", xPos, yPos))
        }
    }
    else if (forceShow)
    {
        Gui, ScrollIndicator:Show, NA x%xPos% y%yPos%
        scrollIndicatorGuiVisible := true
        scrollIndicatorLastX := xPos
        scrollIndicatorLastY := yPos
        if (debugEnabled)
        {
            visible := DllCall("IsWindowVisible", "ptr", scrollIndicatorGuiHwnd)
            DebugLog(Format("Scroll indicator shown at {}, {} (visible={} hwnd=0x{:X})", xPos, yPos, visible ? "true" : "false", scrollIndicatorGuiHwnd + 0))
        }
    }
    else
        return false

    return true
}

HideScrollIndicator()
{
    global scrollIndicatorGuiCreated, scrollIndicatorGuiVisible
    global scrollIndicatorCurrentMode, scrollIndicatorCurrentIcon
    global scrollIndicatorLastX, scrollIndicatorLastY

    if (scrollIndicatorGuiCreated && scrollIndicatorGuiVisible)
        Gui, ScrollIndicator:Hide

    stateChanged := scrollIndicatorGuiVisible ? true : false
    scrollIndicatorGuiVisible := false
    scrollIndicatorCurrentMode := ""
    scrollIndicatorCurrentIcon := 0
    scrollIndicatorLastX := ""
    scrollIndicatorLastY := ""
    return stateChanged
}

ScrollIndicatorEnsureIcon(mode)
{
    global scrollIndicatorIconCache, scrollIndicatorCursorWidth, scrollIndicatorCursorHeight
    global scrollIndicatorHotspotX, scrollIndicatorHotspotY, debugEnabled

    key := ScrollIndicatorModeKey(mode)
    if (key = "")
        return ""

    existing := scrollIndicatorIconCache.HasKey(key) ? scrollIndicatorIconCache[key] : ""
    if (IsObject(existing))
        return existing
    if (existing = false)
        return ""

    path := ScrollIndicatorGetCursorPath(key)
    if (path = "")
    {
        if (debugEnabled)
            DebugLog("Scroll indicator cursor missing for mode: " . key)
        scrollIndicatorIconCache[key] := false
        return ""
    }

    flags := 0x0010
    hCursor := DllCall("LoadImage", "ptr", 0, "str", path, "uint", 2, "int", 0, "int", 0, "uint", flags, "ptr")
    if (!hCursor)
    {
        if (debugEnabled && existing != false)
            DebugLog("LoadImage failed for scroll indicator cursor: " . path)
        scrollIndicatorIconCache[key] := false
        return ""
    }

    hIcon := DllCall("CopyImage", "ptr", hCursor, "uint", 1, "int", 0, "int", 0, "uint", 0, "ptr")
    if (!hIcon)
    {
        if (debugEnabled && existing != false)
            DebugLog("CopyImage failed for scroll indicator cursor: " . path)
        DllCall("DestroyCursor", "ptr", hCursor)
        scrollIndicatorIconCache[key] := false
        return ""
    }

    width := scrollIndicatorCursorWidth
    height := scrollIndicatorCursorHeight
    hotspotX := scrollIndicatorHotspotX
    hotspotY := scrollIndicatorHotspotY
    hbm := ScrollIndicatorExtractIconBitmap(hCursor, width, height, hotspotX, hotspotY)
    if (!hbm)
    {
        if (debugEnabled)
            DebugLog("Scroll indicator: failed to extract bitmap for " . path)
        DllCall("DestroyCursor", "ptr", hCursor)
        DllCall("DestroyIcon", "ptr", hIcon)
        scrollIndicatorIconCache[key] := false
        return ""
    }

    data := {icon: hIcon, cursor: hCursor, width: width, height: height, hotspotX: hotspotX, hotspotY: hotspotY, path: path, hbm: hbm}
    scrollIndicatorIconCache[key] := data
    return data
}

ScrollIndicatorModeKey(mode)
{
    global SCROLL_MODE_VERTICAL, SCROLL_MODE_HORIZONTAL, SCROLL_MODE_MULTI

    if (mode = SCROLL_MODE_VERTICAL)
        return "vertical"
    if (mode = SCROLL_MODE_HORIZONTAL)
        return "horizontal"
    if (mode = SCROLL_MODE_MULTI)
        return "multi"
    return ""
}

ScrollIndicatorGetCursorPath(key)
{
    fileName := ""
    if (key = "vertical")
        fileName := "ScrollV.cur"
    else if (key = "horizontal")
        fileName := "ScrollH.cur"
    else if (key = "multi")
        fileName := "ScrollM.cur"

    if (fileName = "")
        return ""

    path := A_ScriptDir . "\" . fileName
    if (!FileExist(path))
        return ""
    return path
}

DestroyScrollIndicatorGui()
{
    global scrollIndicatorGuiCreated, scrollIndicatorGuiVisible, scrollIndicatorGuiHwnd
    global scrollIndicatorLastX, scrollIndicatorLastY
    global debugEnabled

    if (!scrollIndicatorGuiCreated)
        return

    try
    {
        Gui, ScrollIndicator:Destroy
    }
    catch
    {
        if (debugEnabled)
            DebugLog("Scroll indicator: failed to destroy GUI")
    }

    scrollIndicatorGuiCreated := false
    scrollIndicatorGuiVisible := false
    scrollIndicatorGuiHwnd := 0
    scrollIndicatorLastX := ""
    scrollIndicatorLastY := ""
}

DestroyScrollIndicatorResources()
{
    global scrollIndicatorIconCache, scrollIndicatorCurrentMode, scrollIndicatorCurrentIcon
    global scrollIndicatorCursorWidth, scrollIndicatorCursorHeight
    global scrollIndicatorHotspotX, scrollIndicatorHotspotY

    for key, data in scrollIndicatorIconCache
    {
        if (!IsObject(data))
            continue
        if (data.HasKey("hbm") && data.hbm)
            DllCall("DeleteObject", "ptr", data.hbm)
        if (data.HasKey("icon") && data.icon)
            DllCall("DestroyIcon", "ptr", data.icon)
        if (data.HasKey("cursor") && data.cursor)
            DllCall("DestroyCursor", "ptr", data.cursor)
    }

    scrollIndicatorIconCache := {}
    scrollIndicatorCurrentMode := ""
    scrollIndicatorCurrentIcon := 0
    scrollIndicatorCursorWidth := 32
    scrollIndicatorCursorHeight := 32
    scrollIndicatorHotspotX := 16
    scrollIndicatorHotspotY := 16
}

HandleExplorerScroll(dy, mx := "", my := "")
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

        if (ctx.HasKey("pendingFocus") || ctx.HasKey("wheelBuffer"))
            EnsureExplorerContextFocus(ctx, ctx.HasKey("pendingFocus") ? ctx.pendingFocus : false, mx, my)

        if (debugEnabled)
            DebugLog(Format("HandleExplorerScroll attempt={} strategy={} window=0x{:X}", attempt, ctx.HasKey("strategy") ? ctx.strategy : "?", ctx.HasKey("window") ? (ctx.window + 0) : 0))

        scrollResult := HandleExplorerWheelScroll(ctx, dy, mx, my)
        if (scrollResult < 0)
        {
            if (debugEnabled)
                DebugLog("Explorer wheel handler reported failure; rebuilding context")
            explorerContext := {}
            continue
        }

        scrollsTotal += scrollResult
        explorerContext := ctx
        return true
    }

    if (debugEnabled)
        DebugLog("Explorer scroll handling exhausted attempts without success")
    return false
}

HandleExplorerWheelScroll(ctx, dy, mx := "", my := "")
{
    global swap, debugEnabled, wheelSensitivity, wheelMaxStep

    if (!IsObject(ctx))
        return -1

    strategy := ctx.HasKey("strategy") ? ctx.strategy : ""
    if (strategy = "listview")
        return HandleExplorerListViewScroll(ctx, dy)
    if (strategy = "scrollbar")
        return HandleExplorerScrollbarScroll(ctx, dy)
    if (strategy != "uia")
        return -1

    target := GetExplorerScrollTargetHwnd(ctx, mx, my)
    if (target)
        ctx.scrollTargetHwnd := target
    if (TryPromoteExplorerUiaContext(ctx))
        return HandleExplorerScrollbarScroll(ctx, dy)
    return HandleExplorerUiaWheelFallback(ctx, dy)
}

HandleExplorerListViewScroll(ByRef ctx, dy)
{
    global swap, debugEnabled

    if (!ctx.HasKey("listView") || !ctx.listView)
        return -1
    if (!DllCall("IsWindow", "ptr", ctx.listView))
        return -1
    if (!DllCall("IsWindowVisible", "ptr", ctx.listView))
        return -1

    direction := swap ? -1 : 1
    ctx.pixelAccumulator += dy * direction * ctx.pixelScale
    delta := Round(ctx.pixelAccumulator)
    if (delta = 0)
        return 0

    ctx.pixelAccumulator -= delta
    if (ctx.pixelMaxStep && Abs(delta) > ctx.pixelMaxStep)
        delta := (delta > 0) ? ctx.pixelMaxStep : -ctx.pixelMaxStep

    DllCall("SendMessage", "ptr", ctx.listView, "uint", 0x1014, "ptr", 0, "int", delta)
    if (debugEnabled)
    {
        DebugLog(Format("ListView scroll delta={} accum={:.2f}", delta, ctx.pixelAccumulator))
        metrics := {}
        metrics.dy := Format("{:.2f}", dy)
        metrics.pointerPx := Format("{:.2f}", dy * direction)
        metrics.pixelScale := Format("{:.3f}", ctx.pixelScale)
        metrics.applied := delta
        metrics.accumulator := Format("{:.3f}", ctx.pixelAccumulator)
        ExplorerDebugLogScrollMetrics("listview", metrics)
    }
    return Abs(delta)
}

HandleExplorerScrollbarScroll(ByRef ctx, dy)
{
    global swap, debugEnabled

    if (!ctx.HasKey("scrollBar") || !ctx.scrollBar)
        return -1
    if (!DllCall("IsWindow", "ptr", ctx.scrollBar))
        return -1
    if (!DllCall("IsWindowVisible", "ptr", ctx.scrollBar))
        return -1
    if (ctx.HasKey("folderView") && ctx.folderView && DllCall("IsWindow", "ptr", ctx.folderView) && !DllCall("IsWindowVisible", "ptr", ctx.folderView))
        return -1

    priorPos := ctx.position
    UpdateExplorerScrollUnitScale(ctx)

    direction := swap ? -1 : 1
    ctx.pixelAccumulator += dy * direction

    if (ctx.unitsPerPixel <= 0)
        ctx.unitsPerPixel := 1.0

    unitDelta := ctx.pixelAccumulator * ctx.unitsPerPixel
    deltaUnits := Round(unitDelta)
    if (deltaUnits = 0)
        return 0

    priorAccumulator := ctx.pixelAccumulator
    ctx.pixelAccumulator -= deltaUnits / ctx.unitsPerPixel
    newPos := Clamp(ctx.position + deltaUnits, ctx.minPos, ctx.maxPosEff)
    if (newPos != ctx.position)
    {
        ApplyScrollBarPosition(ctx, newPos)
        SendScrollBarThumb(ctx, newPos, 5)
        ctx.position := newPos
        ctx.scrollMoved := true
        pointerDelta := dy * direction
        unitsApplied := newPos - priorPos
        if (pointerDelta != 0 && unitsApplied != 0)
        {
            atEdge := (newPos <= ctx.minPos || newPos >= ctx.maxPosEff)
            ExplorerUpdateObservedUnits(ctx, unitsApplied, pointerDelta, "observed_scrollbar", atEdge, deltaUnits)
        }
        if (debugEnabled)
        {
            DebugLog(Format("ScrollBar track pos={} delta={} accum={:.2f}", newPos, deltaUnits, ctx.pixelAccumulator))
            metrics := {}
            metrics.dy := Format("{:.2f}", dy)
            metrics.pointerPx := Format("{:.2f}", dy * direction)
            metrics.requestedUnits := deltaUnits
            metrics.appliedUnits := unitsApplied
            metrics.unitsPerPixel := Format("{:.5f}", ctx.unitsPerPixel)
            metrics.accumulatorBefore := Format("{:.3f}", priorAccumulator)
            metrics.accumulatorAfter := Format("{:.3f}", ctx.pixelAccumulator)
            metrics.min := ctx.minPos
            metrics.maxEff := ctx.maxPosEff
            ExplorerDebugLogScrollMetrics("scrollbar", metrics)
        }
    }

    return Abs(deltaUnits)
}

HandleExplorerUiaWheelFallback(ByRef ctx, dy)
{
    global swap, debugEnabled

    if (!ctx.HasKey("pattern") || !IsObject(ctx.pattern))
        return -1

    if (!ctx.HasKey("pixelAccumulator"))
        ctx.pixelAccumulator := 0.0

    UpdateExplorerScrollUnitScale(ctx)

    unitsPerPixel := ctx.HasKey("unitsPerPixel") ? (ctx.unitsPerPixel + 0.0) : 0.0
    if (unitsPerPixel <= 0)
        return -1

    direction := swap ? -1 : 1
    ctx.pixelAccumulator += dy * direction

    deltaUnits := ctx.pixelAccumulator * unitsPerPixel
    if (Abs(deltaUnits) < 0.01)
        return 0

    currentPercent := ExplorerGetVerticalScrollPercent(ctx)
    if (currentPercent = "")
    {
        ctx.pixelAccumulator := 0.0
        return -1
    }

    maxPercent := ctx.HasKey("maxPosEff") && ctx.maxPosEff > 0 ? (ctx.maxPosEff + 0.0) : 100.0
    targetPercent := Clamp(currentPercent + deltaUnits, 0.0, maxPercent)
    appliedUnits := targetPercent - currentPercent

    if (Abs(appliedUnits) < 0.01)
    {
        ctx.pixelAccumulator := 0.0
        return 0
    }

    priorAccumulator := ctx.pixelAccumulator
    ctx.pixelAccumulator -= appliedUnits / unitsPerPixel

    try
    {
        ctx.pattern.SetScrollPercent(-1, targetPercent)
        ctx.scrollPercent := targetPercent
        if (debugEnabled)
        {
            DebugLog(Format("UIA precise scroll current={:.3f} target={:.3f} applied={:.3f} pending={:.3f} unitsPerPixel={:.5f}", currentPercent, targetPercent, appliedUnits, ctx.pixelAccumulator * unitsPerPixel, unitsPerPixel))
            actualPercent := ExplorerGetVerticalScrollPercent(ctx)
            if (actualPercent = "")
                actualPercent := targetPercent
            pointerDelta := dy * direction
            percentApplied := actualPercent - currentPercent
            if (pointerDelta != 0 && percentApplied != 0)
            {
                atEdge := (targetPercent <= 0.0 || targetPercent >= maxPercent)
                ExplorerUpdateObservedUnits(ctx, percentApplied, pointerDelta, "observed_uia", atEdge, appliedUnits)
            }
            metrics := {}
            metrics.dy := Format("{:.2f}", dy)
            metrics.pointerPx := Format("{:.2f}", dy * direction)
            metrics.requestedPercent := Format("{:.4f}", appliedUnits)
            metrics.actualPercent := Format("{:.4f}", percentApplied)
            metrics.startPercent := Format("{:.4f}", currentPercent)
            metrics.targetPercent := Format("{:.4f}", targetPercent)
            metrics.unitsPerPixel := Format("{:.5f}", unitsPerPixel)
            metrics.accumulatorBefore := Format("{:.3f}", priorAccumulator)
            metrics.accumulatorAfter := Format("{:.3f}", ctx.pixelAccumulator)
            metrics.pendingUnits := Format("{:.4f}", ctx.pixelAccumulator * unitsPerPixel)
            ExplorerDebugLogScrollMetrics("uia", metrics)
        }
        return Abs(appliedUnits)
    }
    catch e
    {
        ctx.pixelAccumulator += appliedUnits / unitsPerPixel
        if (debugEnabled)
        {
            errMsg := IsObject(e) && e.HasKey("Message") ? e.Message : e
            DebugLog("UIA SetScrollPercent failed: " . errMsg)
        }
        return -1
    }
}

ExplorerGetVerticalScrollPercent(ByRef ctx)
{
    if (!IsObject(ctx) || !ctx.HasKey("pattern") || !IsObject(ctx.pattern))
        return ""
    try
    {
        percent := ctx.pattern.CurrentVerticalScrollPercent
        if (percent = -1)
            percent := 0.0
        ctx.scrollPercent := percent + 0.0
        return percent + 0.0
    }
    catch
        return ""
}


TryPromoteExplorerUiaContext(ByRef ctx)
{
    global debugEnabled
    if (!IsObject(ctx) || !ctx.HasKey("strategy") || ctx.strategy != "uia")
        return false

    src := ctx
    winHwnd := src.HasKey("window") ? src.window : 0

    target := src.HasKey("scrollTargetHwnd") ? src.scrollTargetHwnd : 0
    if (target && DllCall("IsWindow", "ptr", target))
    {
        promoted := BuildExplorerScrollInfoContext(target, winHwnd, src)
        if (IsObject(promoted))
        {
            CopyExplorerUiaMetadata(src, promoted)
            if (!promoted.HasKey("scrollTargetHwnd") || !promoted.scrollTargetHwnd)
                promoted.scrollTargetHwnd := target
            UpdateExplorerScrollUnitScale(promoted, true)
            ctx := promoted
            if (debugEnabled)
                DebugLog("UIA context promoted via scroll target")
            return true
        }
    }

    native := 0
    if (src.HasKey("verticalBar") && IsObject(src.verticalBar))
    {
        try
            native := src.verticalBar.CurrentNativeWindowHandle
        catch
            native := 0
    }
    if (native && DllCall("IsWindow", "ptr", native))
    {
        promoted := BuildExplorerScrollBarContext(native, winHwnd)
        if (IsObject(promoted))
        {
            CopyExplorerUiaMetadata(src, promoted)
            if (!promoted.HasKey("scrollTargetHwnd") || !promoted.scrollTargetHwnd)
                promoted.scrollTargetHwnd := target ? target : native
            UpdateExplorerScrollUnitScale(promoted, true)
            ctx := promoted
            if (debugEnabled)
                DebugLog("UIA context promoted via native scrollbar handle")
            return true
        }
    }

    return false
}

CopyExplorerUiaMetadata(src, ByRef dest)
{
    if (!IsObject(src) || !IsObject(dest))
        return

    if (src.HasKey("folderView"))
        dest.folderView := src.folderView
    if (src.HasKey("scrollTargetHwnd"))
        dest.scrollTargetHwnd := src.scrollTargetHwnd
    if (src.HasKey("focusElement"))
        dest.focusElement := src.focusElement
    if (src.HasKey("pendingFocus"))
        dest.pendingFocus := src.pendingFocus
    if (src.HasKey("lastFocusTick"))
        dest.lastFocusTick := src.lastFocusTick
    if (src.HasKey("startMouseX"))
        dest.startMouseX := src.startMouseX
    if (src.HasKey("startMouseY"))
        dest.startMouseY := src.startMouseY
    if (src.HasKey("wheelBuffer"))
        dest.wheelBuffer := src.wheelBuffer
    if (src.HasKey("pixelAccumulator"))
        dest.pixelAccumulator := src.pixelAccumulator
    if (src.HasKey("unitsPerPixel"))
        dest.unitsPerPixel := src.unitsPerPixel
    if (src.HasKey("scaleSignature"))
        dest.scaleSignature := src.scaleSignature
    if (src.HasKey("scaleSource"))
        dest.scaleSource := src.scaleSource
    if (src.HasKey("scalePageUnits"))
        dest.scalePageUnits := src.scalePageUnits
    if (src.HasKey("scaleTotalUnits"))
        dest.scaleTotalUnits := src.scaleTotalUnits
    if (src.HasKey("viewPixels"))
        dest.viewPixels := src.viewPixels
    if (src.HasKey("totalUnits"))
        dest.totalUnits := src.totalUnits
    if (src.HasKey("pageSize"))
        dest.pageSize := src.pageSize
    if (src.HasKey("verticalViewSize"))
        dest.verticalViewSize := src.verticalViewSize
    if (src.HasKey("minPos"))
        dest.minPos := src.minPos
    if (src.HasKey("maxPos"))
        dest.maxPos := src.maxPos
    if (src.HasKey("maxPosEff"))
        dest.maxPosEff := src.maxPosEff
    if (src.HasKey("scrollPercent"))
        dest.scrollPercent := src.scrollPercent
    if (src.HasKey("uiaSource"))
        dest.uiaSource := src.uiaSource
    dest.active := true
}

EnsureExplorerContextFocus(ByRef ctx, force := false, screenX := "", screenY := "")
{
    global debugEnabled
    if (!IsObject(ctx) || !ctx.HasKey("window"))
        return

    winHwnd := ctx.window
    if (!winHwnd || !DllCall("IsWindow", "ptr", winHwnd))
        return

    now := A_TickCount
    last := ctx.HasKey("lastFocusTick") ? ctx.lastFocusTick : 0
    interval := 250
    if (!force && last && (now - last) < interval)
        return

    ctx.lastFocusTick := now

    if (!WinActive("ahk_id " . winHwnd))
    {
        WinActivate, ahk_id %winHwnd%
    }

    target := GetExplorerScrollTargetHwnd(ctx, screenX, screenY)
    if (target && DllCall("IsWindow", "ptr", target))
        ControlFocus,, ahk_id %target%

    if (ctx.HasKey("focusElement") && IsObject(ctx.focusElement))
    {
        try
            ctx.focusElement.SetFocus()
        catch
        {
            ; ignore focus errors
        }
    }

    focusHandle := GetWindowFocusHandle(winHwnd)
    success := false
    if (focusHandle)
    {
        if (target)
            success := (focusHandle = target) || IsWindowDescendant(focusHandle, target)
        else
            success := true
    }
    ctx.pendingFocus := !success
}

GetExplorerScrollTargetHwnd(ByRef ctx, screenX := "", screenY := "")
{
    if (!IsObject(ctx))
        return 0

    if (ctx.HasKey("scrollTargetHwnd"))
    {
        target := ctx.scrollTargetHwnd
        if (target && DllCall("IsWindow", "ptr", target))
            return target
        ctx.scrollTargetHwnd := 0
    }

    candidates := []
    if (ctx.HasKey("folderView") && ctx.folderView && DllCall("IsWindow", "ptr", ctx.folderView))
        candidates.Push(ctx.folderView)
    listView := GetExplorerListViewHandle(ctx.HasKey("window") ? ctx.window : 0, ctx.HasKey("folderView") ? ctx.folderView : 0, screenX, screenY)
    if (listView && DllCall("IsWindow", "ptr", listView))
        candidates.Push(listView)
    if (ctx.HasKey("window") && ctx.window && DllCall("IsWindow", "ptr", ctx.window))
        candidates.Push(ctx.window)

    for index, handle in candidates
    {
        if (!handle)
            continue
        if (DllCall("IsWindow", "ptr", handle))
        {
            ctx.scrollTargetHwnd := handle
            return handle
        }
    }

    return 0
}

SendExplorerWheelMessage(ctx, target, delta, keyMask, xCoord, yCoord)
{
    if (!target || !DllCall("IsWindow", "ptr", target))
        return false

    wheel := delta & 0xFFFF
    wParam := (keyMask & 0xFFFF) | ((wheel & 0xFFFF) << 16)
    lParam := ((yCoord & 0xFFFF) << 16) | (xCoord & 0xFFFF)

    sent := false
    handles := [target]
    if (ctx.HasKey("window") && ctx.window && ctx.window != target)
        handles.Push(ctx.window)

    for index, hwnd in handles
    {
        if (!hwnd || !DllCall("IsWindow", "ptr", hwnd))
            continue
        DllCall("SendMessage", "ptr", hwnd, "uint", 0x020A, "uptr", wParam, "ptr", lParam)
        sent := true
    }

    return sent
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
            UpdateExplorerScrollUnitScale(ctx, true)
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
        ctx.wheelBuffer := 0.0
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
    if (vertScrollable && IsObject(verticalBar))
    {
        thumb := verticalBar.FindFirst(TreeScope_Subtree, condThumb)
        verticalRange := GetScrollRange(verticalBar, thumb, true)
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
    ctx.verticalBar := verticalBar
    ctx.horizontalBar := horizontalBar
    ctx.verticalRangeLength := verticalRange
    ctx.wheelBuffer := 0.0
    ctx.pixelAccumulator := 0.0
    ctx.unitsPerPixel := 0.0
    try
        nativeHandle := scroller.CurrentNativeWindowHandle
    catch
        nativeHandle := 0
    if (nativeHandle)
        ctx.scrollTargetHwnd := nativeHandle + 0
    else if (viewHwnd && DllCall("IsWindow", "ptr", viewHwnd))
        ctx.scrollTargetHwnd := viewHwnd
    else
        ctx.scrollTargetHwnd := winHwnd
    ctx.focusElement := scroller
    ctx.pendingFocus := true
    ctx.lastFocusTick := 0

    ctx.startMouseX := screenX
    ctx.startMouseY := screenY

    ctx.pageSize := viewVert + 0.0
    ctx.verticalViewSize := viewVert + 0.0
    ctx.totalUnits := maxVertical + 0.0
    ctx.minPos := 0.0
    ctx.maxPos := maxVertical + 0.0
    ctx.maxPosEff := maxVertical + 0.0
    ctx.scrollPercent := baseVert + 0.0

    if (debugEnabled)
        DebugLogExplorerContext("UIA context built", ctx)

    UpdateExplorerScrollUnitScale(ctx, true)

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

InjectExplorerWheel(delta)
{
    static MOUSEEVENTF_WHEEL := 0x0800
    if (delta = 0)
        return true
    DllCall("mouse_event", "UInt", MOUSEEVENTF_WHEEL, "UInt", 0, "UInt", 0, "Int", delta, "Ptr", 0)
    return true
}

DispatchWheelDelta(delta)
{
    if (delta = 0)
        return true
    remaining := Abs(delta)
    sign := (delta > 0) ? 1 : -1
    chunk := 24
    if (chunk <= 0)
        chunk := 24
    while (remaining > 0)
    {
        step := (remaining > chunk) ? chunk : remaining
        if (!InjectExplorerWheel(step * sign))
            return false
        remaining -= step
    }
    return true
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
    ctx.wheelBuffer := 0.0

    if (debugEnabled)
        DebugLog(Format("ScrollBar context ready hwnd=0x{:X} parent=0x{:X} pos={} range={} track={} unitsPerPixel={:.3f}", scrollBar + 0, ctx.parent + 0, ctx.position, totalUnits, metrics.track, unitsPerPixel))

    return ctx
}

BuildExplorerScrollInfoContext(target, winHwnd, src := "")
{
    global debugEnabled
    if (!DllCall("IsWindow", "ptr", target))
        return ""

    info := GetScrollInfoData(target, 1)
    if (!IsObject(info))
    {
        if (debugEnabled)
            DebugLog(Format("GetScrollInfo failed for target=0x{:X}", target + 0))
        return ""
    }

    trackLen := 0.0
    if (IsObject(src) && src.HasKey("verticalRangeLength"))
        trackLen := src.verticalRangeLength + 0.0
    if (trackLen <= 0)
    {
        metrics := GetScrollBarMetrics(target)
        trackLen := metrics.track
    }
    if (trackLen <= 1)
        trackLen := 200.0

    totalUnits := info.max - info.min
    if (info.page > 0 && totalUnits >= info.page)
        totalUnits := totalUnits - info.page + 1
    if (totalUnits < 1)
        totalUnits := (info.max > info.min) ? (info.max - info.min) : 1

    unitsPerPixel := totalUnits / trackLen
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
    ctx.scrollBar := target
    ctx.parent := target
    ctx.window := winHwnd
    ctx.minPos := info.min
    ctx.maxPos := info.max
    ctx.maxPosEff := maxPosEff
    ctx.position := info.pos
    ctx.barType := 1
    ctx.unitsPerPixel := unitsPerPixel
    ctx.pixelAccumulator := 0.0
    ctx.scrollMoved := false
    ctx.totalUnits := totalUnits
    ctx.trackLength := trackLen
    ctx.pageSize := info.page
    ctx.lastTrackPos := info.track
    ctx.wheelBuffer := 0.0

    if (debugEnabled)
        DebugLog(Format("ScrollInfo context ready target=0x{:X} pos={} range={} track={:.1f} unitsPerPixel={:.4f}", target + 0, ctx.position, totalUnits, trackLen, unitsPerPixel))

    return ctx
}

UpdateExplorerScrollUnitScale(ByRef ctx, forceLog := false)
{
    global wheelSensitivity, debugEnabled

    if (!IsObject(ctx) || !ctx.HasKey("strategy"))
        return false

    strategy := ctx.strategy
    if (strategy != "scrollbar" && strategy != "uia")
        return false

    if (!ctx.HasKey("scaleCacheChecked"))
    {
        ctx.scaleCacheChecked := true
        ExplorerApplyCachedScale(ctx)
    }

    viewPixels := ctx.HasKey("viewPixels") ? (ctx.viewPixels + 0.0) : 0.0

    if ((viewPixels <= 0) && ctx.HasKey("folderView") && ctx.folderView)
    {
        rect := GetWindowRectData(ctx.folderView)
        if (IsObject(rect) && rect.height > 0)
            viewPixels := rect.height
    }

    if (viewPixels <= 0 && ctx.HasKey("scrollTargetHwnd") && ctx.scrollTargetHwnd)
    {
        rectTarget := GetWindowRectData(ctx.scrollTargetHwnd)
        if (IsObject(rectTarget) && rectTarget.height > 0)
            viewPixels := rectTarget.height
    }

    if (viewPixels <= 0 && strategy = "uia")
    {
        if (ctx.HasKey("scroller") && IsObject(ctx.scroller))
        {
            rectElem := GetElementBoundingRect(ctx.scroller)
            if (IsObject(rectElem))
                viewPixels := rectElem[3] - rectElem[1]
        }
        if (viewPixels <= 0 && ctx.HasKey("focusElement") && IsObject(ctx.focusElement))
        {
            rectFocus := GetElementBoundingRect(ctx.focusElement)
            if (IsObject(rectFocus))
                viewPixels := rectFocus[3] - rectFocus[1]
        }
    }

    scaleSource := "view"
    if (viewPixels <= 0 && ctx.HasKey("trackLength") && ctx.trackLength > 0)
    {
        viewPixels := ctx.trackLength + 0.0
        scaleSource := "track"
    }
    else if (viewPixels <= 0 && ctx.HasKey("verticalRangeLength") && ctx.verticalRangeLength > 0)
    {
        viewPixels := ctx.verticalRangeLength + 0.0
        scaleSource := "uiaRange"
    }

    if (viewPixels <= 0)
        return false

    ctx.viewPixels := viewPixels

    pageUnits := ctx.HasKey("pageSize") ? (ctx.pageSize + 0.0) : 0.0
    if (pageUnits <= 0 && strategy = "uia" && ctx.HasKey("verticalViewSize"))
        pageUnits := ctx.verticalViewSize + 0.0
    if (pageUnits <= 0 && ctx.HasKey("totalUnits"))
        pageUnits := ctx.totalUnits + 0.0
    if (pageUnits <= 0 && ctx.HasKey("maxPosEff") && ctx.HasKey("minPos"))
        pageUnits := (ctx.maxPosEff + 0.0) - (ctx.minPos + 0.0)
    if (pageUnits <= 0)
        return false

    totalUnits := ctx.HasKey("totalUnits") ? (ctx.totalUnits + 0.0) : 0.0
    if (totalUnits <= 0 && ctx.HasKey("maxPosEff"))
        totalUnits := ctx.maxPosEff + 0.0
    if (totalUnits > 0 && pageUnits > totalUnits)
        pageUnits := totalUnits

    sens := wheelSensitivity
    if (sens = "")
        sens := 12.0
    sens := sens + 0.0
    if (sens <= 0)
        sens := 12.0

    baseScale := pageUnits / viewPixels
    if (baseScale <= 0)
        return false

    baseUnits := baseScale * (sens / 12.0)
    if (baseUnits <= 0)
        baseUnits := baseScale
    baseUnits := Max(0.0001, baseUnits)

    finalUnits := baseUnits
    finalSource := scaleSource
    nowTick := A_TickCount

    if (ctx.HasKey("observedUnitsPerPixel"))
    {
        obsUnits := ctx.observedUnitsPerPixel + 0.0
        obsTick := ctx.HasKey("observedUnitsTick") ? (ctx.observedUnitsTick + 0) : 0
        persistent := ctx.HasKey("observedPersistent") ? ctx.observedPersistent : false
        if (obsUnits > 0)
        {
            if (persistent || !obsTick || (nowTick - obsTick) <= 15000)
            {
                finalUnits := Max(0.0001, obsUnits)
                if (ctx.HasKey("observedSource") && ctx.observedSource != "")
                    finalSource := ctx.observedSource
            }
            else
            {
                ctx.Delete("observedUnitsPerPixel")
                ctx.Delete("observedUnitsTick")
                ctx.Delete("observedSource")
                ctx.Delete("observedPersistent")
            }
        }
    }

    rangeBoostApplied := false
    if (ctx.HasKey("rangeBoostPointer"))
        ctx.Delete("rangeBoostPointer")
    boostedUnits := ExplorerComputeRangeBoostUnits(ctx, finalUnits)
    rangePointerActive := ctx.HasKey("rangeBoostPointer")
    if (boostedUnits > finalUnits)
    {
        finalUnits := boostedUnits
        rangeBoostApplied := true
    }
    else if (rangePointerActive)
        rangeBoostApplied := true

    ctx.unitsPerPixel := finalUnits
    if (rangeBoostApplied)
        finalSource := finalSource . "+range"
    ctx.scaleSource := finalSource
    ctx.scalePageUnits := pageUnits
    ctx.scaleTotalUnits := totalUnits
    ctx.baseUnitsPerPixel := baseUnits

    displaySource := finalSource
    if (finalSource = "cache" && ctx.HasKey("cachedSource") && ctx.cachedSource != "")
        displaySource := "cache[" . ctx.cachedSource . "]"

    signature := Format("{:.6f}|{:.3f}|{:.0f}|{}|{:.6f}", finalUnits, pageUnits, viewPixels, finalSource, baseUnits)
    changed := (!ctx.HasKey("scaleSignature") || ctx.scaleSignature != signature)
    ctx.scaleSignature := signature

    if (debugEnabled && (forceLog || changed))
        DebugLog(Format("Explorer scale recalculated: src={} viewPx={} pageUnits={:.3f} unitsPerPixel={:.6f} baseUnits={:.6f} sens={:.2f}", displaySource, viewPixels, pageUnits, finalUnits, baseUnits, sens))

    return true
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
