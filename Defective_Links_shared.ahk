;--- Variables to set -----------------------------
SetCapsLockState, AlwaysOff

global DoVendorID := 1
global targetMonitor := 3
global TabNumber := 2
global xBoxloc := 1500
global yBoxloc := 350

;--------------------- Hotkeys to change ------------------------
+`:: Tutorial()
`:: Tutorial()
CapsLock:: GetData()
~Xbutton2:: newdata(1)
Mbutton:: newdata(2)
Xbutton1:: newdata(3)

;----------------------------------------------------------------------------

GetData() {
    vendor_id := ""
    Name := ""
    Old_url := ""
    tld := ""
    WinGetActiveTitle, current
    if InStr(current, "xlsx") {
        global OldClip := ClipboardAll
        Clipboard := ""
        Send, {Home}
        Sleep, 100
        Send, ^c
        ClipWait, 2

        vendor_id := Clipboard
        Sleep, 100
        Send, {Right}
        Sleep, 100
        Send, {Right}
        Sleep, 100
        Clipboard := ""

        Send, ^c
        ClipWait, 2

        Name := Clipboard
        Sleep, 100
        Send, {Right}
        Sleep, 100
        Send, {Right}
        Sleep, 100
        Clipboard := ""

        Send, ^c
        ClipWait, 2
        Sleep, 100
        Old_url := Clipboard
        Sleep, 100

        tld := GetTLD(Old_url)

        Clipboard := ""

        Sleep, 100
        Send, {Left}
        Sleep, 100

        Search(OldClip, vendor_id, Name, Old_url, tld)
    } else {
        MsgBox, 5, Error, Please open the Excel file and go to the "new url" field to continue
    }
}

Search(OldClip, vendor_id, Name, Old_url, tld) {
    Sleep, 100
    appurl := "https://app.getvalify.com/admin/vendors"
    Clipboard := appurl

    RunChrome(targetMonitor)
    Sleep, 200

    Sleep, 100
    Send, ^l
    Sleep, 150

    Send, ^v
    Sleep, 100
    Send, {Enter}
    WinWait, Valify - Google Chrome
    Sleep, 150
    TextEnter(vendor_id)
    Sleep, 200

    Send, ^t
    Clipboard := tld
    Sleep, 100
    Send, ^v
    Sleep, 100
    Send, {Enter}
    Sleep, 100

    Send, ^t
    Sleep, 100
    Send, ^l
    Sleep, 100
    Clipboard := Old_url
    ClipWait, 2
    Send, ^v
    Sleep, 100
    Send, {space}
    Sleep, 100
    Send, {Enter}
    Sleep, 100

    Send, ^t
    Sleep, 100
    Send, ^l
    Sleep, 100
    Clipboard := Name
    ClipWait, 2
    Send, ^v
    Sleep, 100
    Send, {space}
    Sleep, 100
    Send, {Enter}
    Sleep, 100
    Send, ^{%TabNumber%}
    Sleep, 100
	Send, {ctrl up}

    Clipboard := vendor_id
}

newdata(mode) {
    if (mode = 1) {
        Clipboard := ""
        Sleep, 100
        Send, {RButton}
        Sleep, 100
        Send, {up 2}{Enter}
        Sleep, 100
        ClipWait, 2
        data := Clipboard
    } else if (mode = 2) {
        Clipboard := ""
        Send, ^l
        Sleep, 100
        Send, ^c
        Sleep, 100
        ClipWait, 2
        data := Clipboard
    } else if (mode = 3) {
        Sleep, 100
        data := "?"
        Sleep, 100
        Clipboard := ""
    }

    if (data = "") {
        MouseGetPos, mouseX, mouseY
        Tooltip, Try again, mouseX, mouseY
        Sleep, 1000
        Tooltip
    } else {
        Sleep, 100
        Send, ^+w
        Sleep, 200
        Inputdata(data)
    }
}

Inputdata(data) {
    Sleep, 200
    WinGetActiveTitle, current
    if InStr(current, "xlsx") {
        Send, %data%
        Sleep, 300
        Send, {Down}
        Sleep, 100
        GetData()
    } else {
        SetTitleMatchMode, 2

        if WinExist("ahk_class XLMAIN") and WinExist("xlsx") {
            WinActivate, ahk_class XLMAIN
            Sleep, 200
            Inputdata(data)
        } else {
            MsgBox, Excel not running. Please enter the sheet to continue.
        }
    }
}

RunChrome(monitorNumber) {
    Run, chrome.exe
    Sleep, 200
    WinGet, hWnd, ID, ahk_class Chrome_WidgetWin_1
    SysGet, MonitorCount, MonitorCount

    WinActivate
    monitorgo := (monitorNumber - 2) * 1922
    WinWait, ahk_class Chrome_WidgetWin_1
    WinRestore, ahk_id %hWnd%
    Sleep, 100
    WinMove, %monitorgo%, 0
    WinMaximize, ahk_id %hWnd%
    Sleep, 100
    WinActivate, ahk_id %hWnd%
}

TextEnter(vendorid) {
    if (DoVendorID = 1) {
        WinGet, hWnd, ID, ahk_class Chrome_WidgetWin_1
        Clipboard := vendorid
        rangeXMin := xBoxloc - 50
        rangeXMax := xBoxloc + 50
        rangeYMin := 347
        rangeYMax := 365
        CoordMode, Mouse, Window
        CoordMode, Caret, Window
        WinWait, ahk_exe chrome.exe,, 5

        Loop {
            MouseGetPos, xpos, ypos
            if (xpos >= rangeXMin and xpos <= rangeXMax) {
                if (ypos >= rangeYMin and ypos <= rangeYMax) {
                    Sleep, 100
                    Send, ^v
                    break
                }
            }
            Sleep, 250
            MouseClick, left, %xBoxloc%, %yBoxloc%,, 5

            if (A_Index = 3) {
                break
            }
        }
    }
}

GetTLD(linky) {
    RegExMatch(StrSplit(linky, "/").3, "[^.]+\.[^.]+$", top)
    domain := "http://" top
    return domain
}

Tutorial() {
    MsgBox, 
    (
    ---------- Defective_links Hotkeys ----------
    To use -
    - Open Excel sheet
    - Put cursor in New Url field
    - Press Search Hotkey

    ------- Important Hotkeys --------

    Esc to end script

    Tutorial Hotkey - tilde

    ----- Default Hotkeys: -----

    Begin Search Hotkey: CapsLock

    Pick URL hotkey: Side mouse button 2 (front)
    (Must be hovering over URL)

    Get current URL: Middle mouse button
    (gets url of site you are currently visiting)

    No Url found: Side mouse button 1 (back)
    (puts ? in the url field)
    )
}

Esc::
    Clipboard := OldClip
    ExitApp

