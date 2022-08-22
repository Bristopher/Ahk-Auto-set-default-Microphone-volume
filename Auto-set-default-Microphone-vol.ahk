; v2.36
/*
Adjusted wording
*/

#Persistent

if !FileExist("Lib\nircmd.exe") {
	MsgBox, Error! make sure you also download and keep the Lib folder in the same folder as this script
}

if (FileExist(A_ScriptDir . "\options.txt")) {
	
}else {
	MsgBox, Options file not found (first time running program most likely)`nEnter properties into options.txt
	
	FileAppend,
(
{0.0.1.00000000}.{77B25CCC-3F98-4A7A-ACDA-6325945A6878}
{0.0.0.00000000}.{691C7A88-58B5-419E-836D-902BCD2C8080}
85
DO NOT EDIT THIS SINGLE LINE OF TEXT





   ____        _   _                 
  / __ \      | | (_)                
 | |  | |_ __ | |_ _  ___  _ __  ___ 
 | |  | | '_ \| __| |/ _ \| '_ \/ __|
 | |__| | |_) | |_| | (_) | | | \__ \
  \____/| .__/ \__|_|\___/|_| |_|___/
        | |                          
        |_|                          
==========**Options for variables**=========
*********EVERYTHING IS CASE SENSITIVE********
Open "Device Manager" -> select your microphone, right click -> "Properties" -> select Property "Device instance path" and use line after second the "\"----------------------------------------------First line usage: {0.0.1.00000000}.{77B25CCC-3F98-4A7A-ACDA-6325945A6878}
; i.e. Device instance path = SWD\MMDEVAPI\{0.0.1.00000000}.{77B25CCC-3F98-4A7A-ACDA-6325945A6878} -> {0.0.1.00000000}.{77B25CCC-3F98-4A7A-ACDA-6325945A6878}
Open "Device Manager" -> select your headphones, right click -> "Properties" -> select Property "Device instance path" and use line after second the "\"----------------------------------------------Second line usage: {0.0.0.00000000}.{691C7A88-58B5-419E-836D-902BCD2C8080}
What volume do you want your mic to be set to for normal usage?---------------------------------------------------------------------------------------------------------------------------------------Third line usage: 85
==============================================


  _   _       _            
 | \ | |     | |           
 |  \| | ___ | |_ ___  ___ 
 | . ` |/ _ \| __/ _ \/ __|
 | |\  | (_) | ||  __/\__ \
 |_| \_|\___/ \__\___||___/
-----------------**NOTES**--------------------
Fourth Line: **This must not be chaged it tell the program this is the end of the options part of this txt**
----------------------------------------------

  _____       __                                  
 |  __ \     / _|                                 
 | |__) |___| |_ ___ _ __ ___ _ __   ___ ___  ___ 
 |  _  // _ \  _/ _ \ '__/ _ \ '_ \ / __/ _ \/ __|
 | | \ \  __/ ||  __/ | |  __/ | | | (_|  __/\__ \
 |_|  \_\___|_| \___|_|  \___|_| |_|\___\___||___/
**References**


My default variables (DON'T EDIT THIS, IT'S A REFERENCE)
-------------------------------------------------------------------------
{0.0.1.00000000}.{77B25CCC-3F98-4A7A-ACDA-6325945A6878}
{0.0.0.00000000}.{691C7A88-58B5-419E-836D-902BCD2C8080}
85
DO NOT EDIT THIS SINGLE LINE OF TEXT
-------------------------------------------------------------------------



), %A_ScriptDir%\options.txt
	

MsgBox Config saved in same directory as this script! Please restart script ;`n`n*

ExitApp
}
gosub Update

Update:
	URLDownloadToFile, https://raw.githubusercontent.com/Bristopher/Ahk-Auto-set-default-Microphone-volume/main/Auto-set-default-Microphone-vol.ahk, update.txt
	FileReadLine, update, update.txt, 1

	if (update = "; v2.36") {
		FileDelete, update.txt
		GoTo NoUpdate
	} else {
		FileReadLine, reason, update.txt, 3
		MsgBox, 4, , A new version of this script has been released!  `n`nPlease press YES to update to the latest version, `nor NO to continue with this version.`n`n`n`nReson for update: %reason%
		
		doUpdate := False
		IfMsgBox, Yes 
			doUpdate := True
		
		if (doUpdate = "True") {
			FileCopy, update.txt, Auto set default Microphone vol v2.1 cleaned up code.ahk, 1
			FileDelete, update.txt
			msgbox, The script will now close.  Please restart it to apply the update!
			ExitApp
			return
		} else {
			msgbox, The script was not updated
			FileDelete, update.txt
			gosub NoUpdate
		}
	
	}



NoUpdate:

; Each array must be initialized before use:
global Array := []

Array[j] := A_LoopField

Array[j, k] := A_LoopReadLine

global ArrayCount := 0
Loop, Read, %A_ScriptDir%\options.txt ; This loop retrieves each line from the file, one at a time.
{
	if InStr(A_LoopReadLine, "DO NOT EDIT THIS SINGLE LINE OF TEXT") {
	break
}

  ArrayCount += 1
  Array[ArrayCount] := A_LoopReadLine
}

Loop % ArrayCount
{
; THIS ARRAY STARTS AT 1, IKR WTF

  element := Array[A_Index]
; TESTER  MsgBox % "Element number " . A_Index . " is " . Array[A_Index]
}


global MicId := Array[1]
global HeadphonesId := Array[2]
global DesiredMicVolume := Array[3]

global MicVolume := 85
global MicVolumeConver := ((65535 * MicVolume) /100)

; Loads user Config


SwapAudioMic(MicId) 
SwapAudioDevice(HeadphonesId)
; 3 options
;	First open = "Sound Control Panel" / "Sound"	|  i.e. open run (win+r) and enter "mmsys.cpl"
;	
;	1)-Name might change over time idk but its easy
; 		You can rename you mic and just feed in the name above like this -> SwapAudioDevice("HeadphonesAt2020").
;
;	2)-This shoulnd't change so I think its preferable (doesn't work for mic tho idk)
;			Or right click and open properties in "Sound Control Panel", open "properties" in "General" tab, click "details tab", 
;			click "Properties" drop down menu and select "Friendly Name", copy it and feed into the function above like this -> SwapAudioMic(MicrophoneAt2020AT2020USB+)
;	3)-
;		Compare id's to id from Device Manager -> select mic/headphones, 
;		right click -> properties -> select property "Device instance path" and use line after second \
;		i.e. SWD\MMDEVAPI\{0.0.1.00000000}.{77B25CCC-3F98-4A7A-ACDA-6325945A6878} -> SwapAudioMic("{0.0.1.00000000}.{77B25CCC-3F98-4A7A-ACDA-6325945A6878}") 

Menu, tray, NoStandard	
Menu, Tray, Icon, Lib\NirCmdfavicon.ico
Menu, Tray, Add, Run MicVolSet for 2m , MicVolSetLoop
Menu, Tray, Add, MicVolSet to MAX, MicVolSetMAX
Menu, Tray, Add, MicVolSet to %MicVolume%, MicVolSet
Menu, Tray, Add, Stop loop, MicVolSetQuit
Menu, Tray, Add, Exit, ExitBut


 ;When Black Ops 3 is opened this loops for every second for 120 times (2min) ahk_class is found using "Window Spy"
WinWait, ahk_class CoDBlackOps
	Run, Lib\nircmd.exe loop 120 1000 setsysvolume %MicVolumeConver% default_record ;Full volume is 65535 so 85% is 65535*0.85
	

SwapAudioMic(device_A_Capture)
{
	global MicId
	global HeadphonesId
	global DesiredMicVolume
	global MicVolume
	global MicVolumeConver
	;Sets system master volume
	

    ; Get device IDs.
    A := VA_GetDevice(MicId), VA_IMMDevice_GetId(A, A_C_id)
	
	if A
    {
		
        ; Get ID of default playback device.
        default := VA_GetDevice("capture")
        VA_IMMDevice_GetId(default, default_id)
        ObjRelease(default)
		
		
        ; If device A isn't default set it, if is is do nothing (or smth like send a message)
		if MicId != default_id
		{
			default_id := A_C_id
			VA_SetDefaultEndpoint(default_id, 0) ;Sets default device
			VA_SetDefaultEndpoint(default_id, 2) ;Sets default communication
			
			Run Lib\nircmd.exe setsysvolume %MicVolumeConver% default_record ;Full volume is 65535 so 85% is 65535*0.85
		} else {
;			msgbox mic already set
		}
		MicVolSet()
    }
    ObjRelease(A)
    if !(A)
        throw Exception("Unknown audio device, try fixing name", -1, MicId)
	
	return
}

SwapAudioDevice(device_A_Playback)
{
	global MicId
	global HeadphonesId
	global DesiredMicVolume
	global MicVolume
	global MicVolumeConver
	
	
    ; Get device IDs.
    A := VA_GetDevice(HeadphonesId), VA_IMMDevice_GetId(A, A_P_id)
    if A
    {
        ; Get ID of default playback device.
        default := VA_GetDevice("playback")
        VA_IMMDevice_GetId(default, default_id)
        ObjRelease(default)
        
        ; If device A isn't default set it, if is is do nothing (or smth like send a message)
		if HeadphonesId != default_id
		{
			default_id := A_P_id
			VA_SetDefaultEndpoint(default_id, 0) ;Sets default device
			VA_SetDefaultEndpoint(default_id, 2) ;Sets default communication
		} else {
;			msgbox mic already set
		}
		
    }
    ObjRelease(A)
    if !(A)
        throw Exception("Unknown audio device, try fixing name", -1, HeadphonesId)
	
	
	
	return
}

MicVolSetLoop() {
	global MicId
	global HeadphonesId
	global DesiredMicVolume
	global MicVolume
	global MicVolumeConver
;	loops for 2 minutes running command every 1 second (2minutes = 120s | 120 loops each 1s)

	Run, Lib\nircmd.exe loop 120 1000 setsysvolume %MicVolumeConver% default_record ;Full volume is 65535 so 85% is 65535*0.85
	
	return
}

MicVolSetMAX() {
	Run, Lib\nircmd.exe setsysvolume 65535 default_record ;Full volume is 65535
return
}

MicVolSet() {
	global MicId
	global HeadphonesId
	global DesiredMicVolume
	global MicVolume
	global MicVolumeConver

	Run Lib\nircmd.exe setsysvolume %MicVolumeConver% default_record

	return
}

MicVolSetQuit() {
	Run, Lib\nircmd.exe killprocess nircmd.exe"
return
}

ExitBut(){
	ExitApp
return
}








; Debug code to display playback and recording defaults device id's
; Compare id's to id from Device Manager -> select mic/headphones, 
;		right click -> properties -> select property "Device instance path" and compare 
;
/*	a1 := VA_GetDevice("capture")
	VA_IMMDevice_GetId(a1, a12)
	a2 := VA_GetDevice("playback")
	VA_IMMDevice_GetId(a2, a23)
	msgbox %a12% %a23%
*/

/* Stray code that doesn't work
	A_Name := VA_GetDeviceName(A)":2"
	msgbox %A_Name%
	
	VA_SetMasterVolume(80, "", %A_Name%)
	SoundSet,80,MICROPHONE:2,VOLUME 
			
	headphone_vol := VA_GetMasterVolume("", device_A_Capture)
	MsgBox, %headphone_vol%
	
	VA_SetMasterVolume(80, device_A_Capture)
*/




/* original 
https://www.autohotkey.com/board/topic/2306-changing-default-audio-device/page-7
https://www.autohotkey.com/board/topic/21984-/
SwapAudioDevice("Speakers", "Digital Output")

SwapAudioDevice(device_A, device_B)
{
    ; Get device IDs.
    A := VA_GetDevice(device_A), VA_IMMDevice_GetId(A, A_id)
    B := VA_GetDevice(device_B), VA_IMMDevice_GetId(B, B_id)
    if A && B
    {
        ; Get ID of default playback device.
        default := VA_GetDevice("playback")
        VA_IMMDevice_GetId(default, default_id)
        ObjRelease(default)
        
        ; If device A is default, set device B; otherwise set device A.
        VA_SetDefaultEndpoint(default_id == A_id ? B : A, 0)
    }
    ObjRelease(B)
    ObjRelease(A)
    if !(A && B)
        throw Exception("Unknown audio device", -1, A ? device_B : device_A)
}
*.
