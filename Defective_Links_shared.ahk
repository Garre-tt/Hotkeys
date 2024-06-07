
;--- ----------------Variables to set -----------------------------
; Set which monitor you would like to display the search page on
SetCapsLockState, AlwaysOff

global DoVendorID := 0

global targetMonitor:= 2


;Set which tab to start in -
;1 is Valify app
;2 is old url top level domain
;3 is old url (reccomended to test if it works before finding new one)
;4 is google of vendor name
global TabNumber := 3



;location of text box on valify vendor lookup
;if you want to change, run windows spy and find the location of the box
global xBoxloc := 1500
global yBoxloc := 350


;--------------------- Hotkeys to change ------------------------
;Change the button listed before the :: to change Hotkeys
;All button names listed here
;https://www.autohotkey.com/docs/v1/KeyList.htm


;Hotkey for tooltip
+`:: Tutorial()
`:: Tutorial()
return

;Hotkey to begin search
CapsLock:: GetData()
return


;Hotkey to pick url from search page
~Xbutton2:: newdata(1)
return


;Hotkey to get current url
Mbutton:: newdata(2)
return


;Hotkey to  return ?
Xbutton1:: newdata(3)
return










;----------------------------------------------------------------------------










;---------------------------------------------------------------------------- Take vendor infor
GetData()
{
vendor_id = 
Nmae = 
Old_url =
tld = 
wingetactivetitle, current
if instr(current, "xlsx") 
{

global OldClip:=clipboard

clipboard :=""
Send, {Home}
sleep 0040
Send, ^c
ClipWait

vendor_id:=clipboard

sleep 0050
Send, {Right}
sleep 0060
Send, {Right}
sleep 0040
clipboard =

Send, ^c
ClipWait

Name:= Clipboard

sleep 0050
Send, {Right}
sleep 0050
Send, {Right}
sleep 0040
clipboard =

Send, ^c
ClipWait
sleep 50
Old_url := Clipboard
sleep 0040

tld := GetTLD(old_url)

clipboard= 

sleep 0030
Send, {Left}
sleep 75

Search(OldClip,vendor_id,Name,Old_url,tld)
return
} else{
MsgBox,,error,Please open the excel file and go to the "new url" field to continue, 5
return
}}



return

;---------------------------------------------------------------------------- Search everything in google
Search(OldClip,vendor_id,Name,Old_url,tld)
{
sleep 0080

appurl = https://app.getvalify.com/admin/vendors
clipboard := appurl

RunChrome(targetMonitor)
sleep 0200

sleep 0030
Send, ^l
Sleep 0150

Send, ^v
sleep 0040
send, {enter}
Winwait, Valify - Google Chrome
sleep 0150
Textenter(vendor_id)
sleep 0200

Send, ^t
clipboard := tld
sleep 0060
Send, ^v
sleep 0060
Send, {Enter}
sleep 0060


Send, ^t
sleep 0060
Send, ^l
sleep 0060
clipboard := old_url
ClipWait
Send, ^v
sleep 0050
Send, {space}
sleep 0030
Send, {Enter}
sleep 0090


Send, ^t
sleep 0060
Send, ^l
sleep 0060
clipboard := name
ClipWait
Send, ^v
sleep 0050
Send, {space}
sleep 0030
Send, {Enter}
sleep 0100
Send, {Ctrl down}
sleep 0050
Send, {%TabNumber%}{ctrl up}

clipboard:= vendor_id
return
}

return



;---------------------------------------------------------------------------- Pick new url 
newdata(mode)
{
if (mode = 1)
{

	clipboard=
	sleep 0030
	Send {Rbutton}
	sleep 0090
	Send {up 2}{Enter}
	sleep 0100
	data := clipboard
}

if (mode = 2)
{
	clipboard=
	Send ^l
	sleep 80
	Send ^c
	sleep 50
	Clipwait
	data:= clipboard

}

if (mode = 3)
{
	sleep 50
	data = ?
	sleep 25
	clipboard=
} 

if data = 
{
MouseGetPos, mouseX, MouseY
Tooltip, Try again,mouseX,mouseY
Sleep 1000
Tooltip
} else {
sleep 0100
Send, ^+w
sleep 0200
Inputdata(data)
return
}

}

return
;----------------------------------------------------------------------------



;---------------------------------------------------------------------------- Input data to sheet
Inputdata(data)
{

sleep 0200
wingetactivetitle, current
if instr(current, "xlsx") 
{
	Send, %data%
	sleep 0300
	Send, {Down}
	sleep 0100
	GetData()
} else{
	; Check for an Excel window containing "xlsx" in the title
	SetTitleMatchMode, 2 ; Allows partial match of window titles

	If WinExist("ahk_class XLMAIN") and WinExist("xlsx")
	{
		; Activate the Excel window
		WinActivate, ahk_class XLMAIN
		sleep 0200
		Inputdata(data)
	}
	else
	{
		; Display a message box if no such window is found
		MsgBox, Excel not running. Please enter the sheet to continue.
	}

}




return
}
return


;----------------------------------------------------------------------------
;ctrl hotkeys




;-------------------------------------------------------------------------- Run Chrome and move to correct monitor 


RunChrome(monitorNumber)
{
	
	Run, chrome.exe
	sleep 200
	WinGet, hWnd, ID, ahk_class Chrome_WidgetWin_1
    SysGet, MonitorCount, MonitorCount

	WinActivate
	monitorgo := (monitorNumber - 2) * 1922
	WinWait, ahk_class Chrome_WidgetWin_1
	WinRestore, ahk_id %hWnd%
	sleep 100
	WinMove, %monitorgo% , 0
	WinMaximize, ahk_id %hWnd%
	sleep 0030
	WinActivate, ahk_id %hWnd%
}
return

;--------------------------------------------------------------------------- Set valify app to correct vendor




return

TextEnter(vendorid)
{
	if (DoVendorID = 1)
	{
	WinGet, hWnd, ID, ahk_class Chrome_WidgetWin_1
    clipboard := vendorid
    ; Retrieve the mouse position

	; Define the range of coordinates
	WinGetPos,,,wWidth
	rangeXMin := xBoxloc - 50
	rangeXMax := xBoxloc + 50
	rangeYMin := 347
	rangeYMax := 365
    CoordMode, Mouse, Window
    CoordMode, Caret, Window
    
    ; Wait for the webpage to load
    WinWait, ahk_exe chrome.exe,, 5
    ; If the webpage doesn't load within 5 seconds, the script continues without waiting 

	
Loop
{
	MouseGetPos xpos,ypos
   if xpos between %rangeXMin% and %rangeXMax%
	{
	   if ypos between %rangeYMin% and %rangeYMax%
	   {
	   sleep 100
	   Send, ^v
	   break
	   }
	}
    sleep 250
	
	MouseClick, left, xBoxloc , yBoxloc,,5

	
	if A_Index = 3
	{
		
		break
		
	}
	}
}
}

return



;------------------------------------------------------------------------- Return Top level domain of old url 
;------------------------------------------------------------------------- tld is being returned as a global variable, not sure what to do with it

GetTLD(linky) { ; https://www.autohotkey.com/boards/viewtopic.php?p=434892#p434892
 RegExMatch(StrSplit(linky, "/").3, "[^.]+\.[^.]+$", top)
 domain := "http://" top
 
 return domain
}
return





;---------------------------------------------------------------------------- Tutorial


Tutorial()
{
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

return

Esc::
clipboard:= OldClip
Exitapp
