;===============================================
/*

CSV Buddy Messenger
Written using AutoHotkey_L v1.1.09.03+ (http://ahkscript.org/)
By Jean Lalonde (JnLlnd on AHKScript.org forum)
	
DESCRIPTION

Called from the command line, file shortcuts, the Explorer context menus, etc. to send messages to CSV Buddy in order to launch various actions like:
- set a control content
- activate a command (click a button)
- ...

See RECEIVE_CSVBUDDYMESSENGER function in CSVBuddy.ahk for details.

HISTORY
=======

Version: 0.1 (2022-06-08)
- initial code from QAPmessenger.ahk for Quick Access Popup converted to CSVBuddyMessenger for CSV Buddy

*/ 
;========================================================================================================================
; --- COMPILER DIRECTIVES ---
;========================================================================================================================

; Doc: http://fincs.ahk4.net/Ahk2ExeDirectives.htm
; Note: prefix comma with `

;@Ahk2Exe-SetName CSV Buddy Messenger
;@Ahk2Exe-SetDescription Send messages to CSV Buddy
;@Ahk2Exe-SetVersion 0.1
;@Ahk2Exe-SetOrigFilename CSVBuddyMessenger.exe


;========================================================================================================================
; INITIALIZATION
;========================================================================================================================

#NoEnv
#SingleInstance force
#KeyHistory 0
ListLines, Off

global g_strAppNameText := "CSV Buddy Messenger"
global g_strAppNameFile := "CSVBuddyMessenger"
global g_strAppVersion := "0.1"
global g_strAppVersionBranch := "beta"
global g_strAppVersionLong := "v" . g_strAppVersion . (g_strAppVersionBranch <> "prod" ? " " . g_strAppVersionBranch : "")
global g_strTargetAppTitle := "CSV Buddy ahk_class JeanLalonde.ca"
global g_strTargetAppName := "CSV Buddy"
global g_strCSVBuddyNameFile := "CSVBuddy"

gosub, SetWorkingDirectory
gosub, InitLanguageVariables

global g_strDiagFile := A_WorkingDir . "\" . g_strAppNameFile . "-DIAG.txt"
global g_strIniFile := A_WorkingDir . "\" . g_strCSVBuddyNameFile . ".ini"
global g_blnDiagMode
IniRead, g_blnDiagMode, %g_strIniFile%, Global, MessengerDiagMode, 0

if CSVBuddyIsRunning()
    if CSVBuddySingleInstance()
    {
        ; Use traditional method, not expression
        g_strParam0 = %0% ; number of parameters
        g_strParam1 = %1% ; fisrt parameter, the command name
        g_strParam2 = %2% ; second parameter, the target
        g_strParam3 = %3% ; third parameter, the value (optional)
        
        Diag("g_strParam0", g_strParam0)

        if (g_strParam0 > 0) and StrLen(g_strParam1)
        {
			strParams := ""
			loop, %g_strParam0%
				strParams .= g_strParam%A_Index% . "|"
			strParams := SubStr(strParams, 1, -1) ; remove last delimiter
            Diag("Send_WM_COPYDATA:Param", strParams)
            Diag("Send_WM_COPYDATA:g_strTargetAppTitle", g_strTargetAppTitle)
            intResult := Send_WM_COPYDATA(strParams, g_strTargetAppTitle)
            ; returns FAIL or 0 if an error occurred, 0xFFFF if a CSV Buddy window is open or 1 if success
            Diag("Send_WM_COPYDATA (1=OK)", intResult)
            
            if (intResult = 0xFFFF) ; not implemented in CSV Buddy
                Oops(lMessengerCloseDialog . "`n`n" . lMessengerHelp, g_strTargetAppName)
        }
        else
            Oops(lMessengerErrorNoParam . "`n`n" . lMessengerHelp, g_strAppNameFile)
    }
    else
        Oops(lMessengerSingleInstanceError . "`n`n" . lMessengerHelp, g_strTargetAppName, g_strAppNameText)
else
	Oops(lMessengerErrorNotRunning . "`n`n" . lMessengerHelp, g_strTargetAppName, g_strAppNameText)

return


;-----------------------------------------------------------
InitLanguageVariables:
;-----------------------------------------------------------

lMessengerHelp := "Search for ""Messenger"" on csvbuddy.QuickAccessPopup.com for help."
lMessengerErrorNotRunning := "An error occurred.`n`nMake sure ~1~ is running before sending commands using ~2~."
lMessengerErrorNoParam := "No action parameter detected after ~1~ command."
lMessengerSingleInstanceError := "More than one instance of ~1~ is running.`n`nMake sure only one instance ~1~ is running before sending commands using  ~2~."

return
;-----------------------------------------------------------


;-----------------------------------------------------------
Send_WM_COPYDATA(ByRef strStringToSend, ByRef strTargetScriptTitle) ; ByRef saves a little memory in this case.
; Adapted from AHK documentation (https://autohotkey.com/docs/commands/OnMessage.htm)
; This function sends the specified string to the specified window and returns the reply.
; The reply is 1 if the target window processed the message, or 0 if it ignored it.
;-----------------------------------------------------------
{
    VarSetCapacity(varCopyDataStruct, 3 * A_PtrSize, 0) ; Set up the structure's memory area.
	
    ; First set the structure's cbData member to the size of the string, including its zero terminator:
    intSizeInBytes := (StrLen(strStringToSend) + 1) * (A_IsUnicode ? 2 : 1)
    NumPut(intSizeInBytes, varCopyDataStruct, A_PtrSize) ; OS requires that this be done.
    NumPut(&strStringToSend, varCopyDataStruct, 2 * A_PtrSize) ; Set lpData to point to the string itself.

	strPrevDetectHiddenWindows := A_DetectHiddenWindows
    intPrevTitleMatchMode := A_TitleMatchMode
    DetectHiddenWindows On
    SetTitleMatchMode 2
	
    SendMessage, 0x4a, 0, &varCopyDataStruct, , %strTargetScriptTitle%, , , , 30000 ; 0x4a is WM_COPYDATA. Must use Send not Post.
	
    DetectHiddenWindows %strPrevDetectHiddenWindows% ; Restore original setting for the caller.
    SetTitleMatchMode %intPrevTitleMatchMode% ; Same.
	
    return ErrorLevel ; Return SendMessage's reply back to our caller.
}
;-----------------------------------------------------------


;------------------------------------------------
Oops(strMessage, objVariables*)
;------------------------------------------------
{
	MsgBox, 48, % L("~1~ (~2~)", g_strAppNameText, g_strAppVersionLong), % L(strMessage, objVariables*)
}
; ------------------------------------------------


;------------------------------------------------
L(strMessage, objVariables*)
;------------------------------------------------
{
	Loop
	{
		if InStr(strMessage, "~" . A_Index . "~")
			StringReplace, strMessage, strMessage, ~%A_Index%~, % objVariables[A_Index], A
 		else
			break
	}
	
	return strMessage
}
;------------------------------------------------


;------------------------------------------------------------
CSVBuddyIsRunning()
;------------------------------------------------------------
{
    strPrevDetectHiddenWindows := A_DetectHiddenWindows
    intPrevTitleMatchMode := A_TitleMatchMode
    DetectHiddenWindows, On
    SetTitleMatchMode, 2
	
	SendMessage, 0x2225, , , , %g_strTargetAppTitle%
	intErrorLevel := ErrorLevel
	Diag("CSVBuddyIsRunning:ErrorLevel (1=OK)", intErrorLevel)
    DetectHiddenWindows, %strPrevDetectHiddenWindows%
    SetTitleMatchMode, %intPrevTitleMatchMode%
	Sleep, -1 ; in QAP prevent the cursor to turn to WAIT image for 5 seconds (did not search why) when showing menu from Desktop background
	
    return (intErrorLevel = 1) ; CSV Buddy replies 1 if it runs, else SendMessage returns "FAIL".
}
;------------------------------------------------------------


;------------------------------------------------------------
CSVBuddySingleInstance()
;------------------------------------------------------------
{
    WinGet, intNbInstances, Count, %g_strTargetAppTitle%
	Diag("CSVBuddySingleInstance:intNbInstances (1=OK)", intNbInstances)
    return (intNbInstances = 1) ; only one instance of CSV Buddy is running
}
;------------------------------------------------------------


;-----------------------------------------------------------
SetWorkingDirectory:
;-----------------------------------------------------------

g_blnPortableMode := true ; set this variable for use later during init
SetWorkingDir, %A_ScriptDir% ; do not support alternative settings files, always use value in scriptdir (at worst, default to EN)

return
;-----------------------------------------------------------


;------------------------------------------------
Diag(strName, strData)
;------------------------------------------------
{
	if !(g_blnDiagMode)
		return

	FormatTime, strNow, %A_Now%, yyyyMMdd@HH:mm:ss
	loop
	{
		FileAppend, %strNow%.%A_MSec%`t%strName%`t%strData%`n, %g_strDiagFile%
		if ErrorLevel
			Sleep, 20
	}
	until !ErrorLevel or (A_Index > 50) ; after 1 second (20ms x 50), we have a problem
}
;------------------------------------------------


