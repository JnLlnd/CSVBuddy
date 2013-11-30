;===============================================
/*
CSV Buddy - Generate HTML Doc
Written using AutoHotkey_L v1.1.09.03+ (http://www.ahkscript.org/)
By JnLlnd on AHK forum
*/ 
;===============================================

#NoEnv
#SingleInstance force
#Include %A_ScriptDir%\CSVBuddy_LANG.ahk
#Include %A_ScriptDir%\CSVBuddy_DOC-LANG.ahk

global strFile := A_ScriptDir . "\html-doc\csvbuddy-doc.html"
global strHtml := ""

W("<html>")
W("<head>")
W(L("<title>~1~ - ~2~</title>", lAppName, lDocDocumentation))
W("<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"">")
W("<link href=""csvbuddy-doc.css"" rel=""stylesheet"" type=""text/css"">")
W("</head>")
W("<body>")

H(1, L("~1~ (~2~) - ~3~", lAppName, lAppVersionLong, lDocDocumentation))

P(lDocDesc450)

H(2, lDocIntro)

H(4, lDocInstallationTitle)

P(lDocInstallationDetail)
P(lDocInstallationTestFiles)
 
H(4, lDocDescription)

P(lDocDesc2000)

T("CSVBuddy-01.jpg", lTab0Load)
T("CSVBuddy-02.jpg", lTab0Edit)
T("CSVBuddy-03.jpg", lLvEventsEditrowMenu)
T("CSVBuddy-04.jpg", lTab0Save)
T("CSVBuddy-05.jpg", lTab0Export)

H(4, lDocFeatures)

P(lDocFeaturesList)

W("<A name=""help""></a>")
H(2, lDocHelp)

P(lDocHelpIntro)

H(3, lTab0Load . " " . lDocTab)

I("CSVBuddy-tab1.jpg")

Help(L(lTab1HelpFileToLoad, lAppName, lAppName))
Help(lTab1HelpHeader)
Help(lTab1HelpSetHeader)
Help(L(lTab1HelpDelimiter1, lAppName))
Help(L(lTab1HelpEncapsulator1, lAppName))
Help(lTab1HelpMultiline1)
Help(lTab1HelpReadyToEdit)

H(3, lTab0Edit . " " . lDocTab)

I("CSVBuddy-tab2.jpg")

Help(L(lTab2HelpRename, "usually comma", "comma", "usually double-quotes"))
Help(L(lTab2HelpSelect, "usually comma"))
Help(L(lTab2HelpOrder, "usually comma"))

H(3, lTab0Save . " " . lDocTab)

Help(lTab3HelpFileToSave)

I("CSVBuddy-tab3.jpg")

Help(lTab3HelpSaveHeader)
Help(lTab3HelpFieldDelimiter3)
Help(lTab3HelpEncapsulator3)
Help(lTab3HelpSaveMultiline)

H(3, lTab0Export . " " . lDocTab)

I("CSVBuddy-tab4.jpg")

Help(lTab4HelpFileToExport)
Help(lTab4HelpExportFormat)
Help(L(lTab4HelpExportFixed, "usually comma", "16"))
Help(lTab4HelpExportHTML)
Help(lTab4HelpExportXML)
Help(lTab4HelpExportExpress)

H(2, lDocKeyboardHelp)
W(lDocKeyboardHelpDetail)

H(2, lDocAdvancedTopicsTitle)
H(3, lDocIniTitle)
W(lDocIniHelp)
H(3, lDocCommandLineTitle)
W(lDocCommandLineHelp)

H(2, lDocSupportTitle)
P(lDocSupportText)

H(2, lDocCopyrightTitle)
P(lDocCopyrightText)

Loop, 30
	P("&nbsp;")

W("</body>")
W("</html>")

Sleep, 100
FileDelete, %strFile%
Sleep, 100
FileAppend, %strHtml%`r`n, %strFile%, UTF-8
Sleep, 100
run, "C:\Program Files\Internet Explorer\iexplore.exe" %strFile%

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


Help(strMessage, objVariables*)
{
	StringLeft, strTitle, strMessage, % InStr(strMessage, "$") - 1
	StringReplace, strMessage, strMessage, %strTitle%$
	StringReplace, strMessage, strMessage, <, &lt;, A
	StringReplace, strMessage, strMessage, >, &gt;, A
	StringReplace, strMessage, strMessage, % Chr(149), &gt;, A
	H(4, strTitle)
	P(strMessage)
}


P(strMessage)
{
	StringReplace, strMessage, strMessage, `n, <BR />`n, A
	StringReplace, strMessage, strMessage, `t, &nbsp;&nbsp;&nbsp;&nbsp;, A
	W("<P>" . strMessage . "</P>" . "`n")
}


H(n, strMessage)
{
	W("<H" . n . ">" . strMessage . "</H" . n . ">" . "`n")
}


T(strImage, strName)
{
	W("<A HREF=""img/" . strImage . """ TARGET=""_blank"" TITLE=""" . strName . """><IMG SRC=""img/" . strImage . """ BORDER=""1"" WIDTH=""250"" HEIGHT=""175"" ALT=""" . strName . """></A>`n")
}


I(strImage)
{
	W("<P ALIGN=""LEFT""><IMG SRC=""img/" . strImage . """ BORDER=""1""></P>`n")
}


W(strMessage)
{
	strHtml := strHtml . strMessage . "`r`n"
}
