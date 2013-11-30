;===============================================
/*
CSV Buddy - Generate Readme Doc
Written using AutoHotkey_L v1.1.09.03+ (http://www.ahkscript.org/)
By JnLlnd on AHK forum
*/ 
;===============================================

#NoEnv
#SingleInstance force
#Include %A_ScriptDir%\CSVBuddy_LANG.ahk
#Include %A_ScriptDir%\CSVBuddy_DOC-LANG.ahk

global strFile := A_ScriptDir . "\readme.md"
global strMD := ""

H(1, L("~1~ (~2~) - Read me", lAppName, lAppVersionLong))
P(lDocDesc450)
P("Written using AutoHotkey_L v1.1.09.03+ (http://www.ahkscript.org)`nBy JnLlnd on [AHK forum](http://www.ahkscript.org/boards/)`nFirst officiel release: 2013-11-30")

H(2, "Links")
B("[Application home](http://code.jeanlalonde.ca/csvbuddy/)")
B("[Download 32-bits / 64-bits](http://code.jeanlalonde.ca/ahk/csvbuddy/csvbuddy.zip) (latest version)")
B("[Description and documentation](http://code.jeanlalonde.ca/ahk/csvbuddy/csvbuddy-doc.html)")

H(2, "History")
H(3, "2013-11-30 v1.0")
B("First official release")
B("Add records to existing data (right-click in the list zone)")
B("Create a new file from scratch (right-click in an empty list zone)")
B("Load the file mentioned as first parameter in the command line")
B("Add validation, confirm before exit and fix various small bugs")
H(3, "2013-11-03 v0.9")
B("Display ""<1"" (instead of ""0"") in status bar when file size is smaller than 0.5 K")
B("Removed CSV Buddy icon from the Tray")
B("Add three test delimited files to the package (see README.txt in the zip file)")
B("Fix default value of blnSkipHelpReadyToEdit in ini file to 0")
H(3, "2013-10-20 v0.8.1")
B("If an .ini file is not found in the program's folder, it is created with default values")
H(3, "2013-10-18 v0.8.0")
B("First release of BETA version")
B("History of ALPHA phase on [BitBucket](https://bitbucket.org/JnLlnd/csvbuddy/) (private repository)")


H(2, lDocCopyrightTitle)
P(lDocCopyrightText)


Sleep, 100
FileDelete, %strFile%
Sleep, 100
FileAppend, %strMD%`r`n, %strFile%, UTF-8
Sleep, 100
run, notepad.exe %strFile%

return





; ------------------ FUNCTIONS ----------------


L(strMessage, objVariables*)
{
	Loop
	{
		if InStr(strMessage, "~" . A_Index . "~")
			StringReplace, strMessage, strMessage, ~%A_Index%~, % objVariables[A_Index]
 		else
			break
	}
	return strMessage
}


B(strMessage)
{
	W("* " . strMessage)
}


P(strMessage)
{
	StringReplace, strMessage, strMessage, `n, %A_Space%%A_Space%`r`n, A
	StringReplace, strMessage, strMessage, `t, &nbsp;&nbsp;&nbsp;&nbsp;, A
	W("`r`n" . strMessage . "`r`n")
}


H(n, strMessage)
{
	s := "`r`n"
	loop, %n%
		s := s . "#"
	W(s . " " . strMessage . "`n")
}


W(strMessage)
{
	strMD := strMD . strMessage . "`r`n"
}
