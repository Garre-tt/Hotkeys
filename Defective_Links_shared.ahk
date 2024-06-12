;------Default Variables------------
SetCapsLockState, AlwaysOff

global DoVendorID := Yes
global targetMonitor := 2
global TabNumber := 2
global xBoxloc := 1500
global yBoxloc := 350




;--------------------- Hotkeys to change ------------------------
`:: Tutorial()
+`:: Settings()
^!s::Settings()
CapsLock:: GetData()
Xbutton2:: newdata(1)
Mbutton:: newdata(2)
Xbutton1:: newdata(3)

;----------------------------------------------------------------------------











GetData() {
    vendor_id := ""
    Name := ""
    Old_url := ""
    tld := ""
    WinGetActiveTitle, current
    if InStr(current, "Excel") {
        global OldClip := ClipboardAll
        Clipboard := ""
        Send, {Home}
        Sleep, 100
        Send, ^c
        ClipWait, 2

        vendor_id := Clipboard
        Sleep, 50
        Send, {Right}
        Sleep, 50
        Send, {Right}
        Sleep, 50
        Clipboard := ""

        Send, ^c
        ClipWait, 2

        Name := Clipboard
        Sleep, 50
        Send, {Right}
        Sleep, 50
        Send, {Right}
        Sleep, 50
        Clipboard := ""

        Send, ^c
        ClipWait, 2
        Sleep, 50
        Old_url := Clipboard
        Sleep, 50

        tld := GetTLD(Old_url)

        Clipboard := ""

        Sleep, 50
        Send, {Left}
        Sleep, 50

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
	WinGetActiveTitle, current
    if InStr(current, "Chrome") {
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
	}
	WinGetActiveTitle, current
    if InStr(current, "Chrome") {
		Send, ^t
		Clipboard := tld
		Sleep, 100
		Send, ^v
		Sleep, 100
		Send, {Enter}
		Sleep, 100
	}
	WinGetActiveTitle, current
    if InStr(current, "Chrome") {
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
	}
	WinGetActiveTitle, current
    if InStr(current, "Chrome") {
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
	}
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
    Sleep, 100
    WinGetActiveTitle, current
    if InStr(current, "excel") {
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
    Sleep, 250
	Loop, 5
	{
		WinGet, hWnd, ID, ahk_class Chrome_WidgetWin_1
		SysGet, MonitorCount, MonitorCount
		WinGetActiveTitle, current
		If Instr(current, "Google - Google Chrome")
		{
			break
		}else{
			sleep 100
		}
	}
    WinActivate, ahk_id %hWnd%
    monitorgo := (monitorNumber - 2) * 1922
    WinWait, ahk_class Chrome_WidgetWin_1
	sleep 50
    WinRestore, ahk_id %hWnd%
    Sleep, 100
    WinMove, %monitorgo%, 0
    WinMaximize, ahk_id %hWnd%
    Sleep, 50
    WinActivate, ahk_id %hWnd%
}

TextEnter(vendorid) {
    if (DoVendorID = Yes) {
        WinGet, hWnd, ID, ahk_class Chrome_WidgetWin_1
        Clipboard := vendorid
        rangeXMin := xBoxloc - 50
        rangeXMax := xBoxloc + 50
        rangeYMin := yBoxloc - 10
        rangeYMax := yBoxloc + 10
        CoordMode, Mouse, Window
        CoordMode, Caret, Window
        WinWait, ahk_exe chrome.exe,, 5
	sleep 200

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
    - Press Search Hotkey (Default Caps Lock)

    ------- Important Hotkeys --------

    	Esc to end script
	
   	Tutorial Hotkey - tilde
	
 	Settings Hotkey - Ctrl + alt + s
			 OR
			  Shift + tilde

    ----- Default Hotkeys: -----

    Begin Search Hotkey: CapsLock

    Pick URL hotkey: Side mouse button 2 (front)
    (Must be hovering over URL)

    Get current URL: Middle mouse button
    (gets url of site you are currently visiting)

    No Url found: Side mouse button 1 (back)
    (puts ? in the url field)

     ------ Navigation Hotkeys -------
     (Only if Chrome is active window)
     1 - Navigate to tab 1
     2 - Navigate to tab 2
     3 - Navigate to tab 3
     4 - Navigate to tab 4
    )
}

Esc::
    ExitApp



#IfWinActive ahk_exe chrome.exe
WinGetActiveTitle, current
If Instr(current, Google)
{
1:: Send, ^1
2::Send, ^2
3:: Send, ^3
4:: Send, ^4
}
return



global TargetMonitorEdit := ""
global TabNumberEdit := ""
global DoVendorIDEdit := ""

Settings() {
    global targetMonitor
    global TabNumber
    global DoVendorID
    global TargetMonitorEdit
    global TabNumberEdit
    global DoVendorIDEdit

    ; Create the GUI
    Gui, Add, Text,, Select which monitor you would like the Chrome window to open on:
    Gui, Add, Edit, vTargetMonitorEdit, %TargetMonitor%
    Gui, Add, Text,, Select which tab you would like to start on (1-4):
    Gui, Add, Edit, vTabNumberEdit, %TabNumber%
    Gui, Add, Text,, Do you want to automatically input Vendor ID?
    Gui, Add, DropDownList, vDoVendorIDedit, Yes||No
    Gui, Add, Button, gSaveSettings, Save
    Gui, Show,, Settings

    Return


    SaveSettings:

        Gui, Submit, NoHide

        targetMonitor := TargetMonitorEdit
        TabNumber := TabNumberEdit
        DoVendorID := DoVendorIDEdit

        Gui, Destroy
        MsgBox, Variables saved:`nMonitor: %targetMonitor%`nTab Number: %TabNumber%`nDo Vendor ID: %DoVendorID%
    Return

    ; Exit the script
    GuiClose:
        Return
    Return
}
