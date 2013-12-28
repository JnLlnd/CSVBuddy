;===============================================
/*
CSV Buddy
Written using AutoHotkey_L v1.1.09.03+ (http://www.ahkscript.org/)
By JnLlnd on AHK forum
This script uses the library ObjCSV v0.3 (https://github.com/JnLlnd/ObjCSV)
*/ 
;===============================================

#KeyHistory 0
#NoEnv
#NoTrayIcon 
#SingleInstance force
ListLines, Off
SetWorkingDir, %A_ScriptDir%

#Include %A_ScriptDir%\..\ObjCSV\lib\ObjCSV.ahk
#Include %A_ScriptDir%\CSVBuddy_LANG.ahk

#MaxMem 4095 ; Variable capacity in megs, default 64
; (see http://auto-hotkey.com/boards/viewtopic.php?f=5&t=115&sid=4e0b3f9f47921441b3b8689138a489b7#p844)
; Tested file size capacity: safe at 100 MB on 32-bits, no limit on 64-bits

; --------------------- COMPILER DIRECTIVES --------------------------

;@Ahk2Exe-SetName CSV Buddy
;@Ahk2Exe-SetDescription Load`, edit`, save and export CSV files
;@Ahk2Exe-SetVersion 1.01
;@Ahk2Exe-SetCopyright Jean Lalonde
;@Ahk2Exe-SetOrigFilename CSVBuddy.exe


; --------------------- GLOBAL AND DEFAULT VALUES --------------------------

strIniFile := A_ScriptDir . "\CSVBuddy.ini"
IfNotExist, %strIniFile%
	FileAppend,
		(LTrim Join`r`n
			[global]
			intDefaultWidth=16
			strTemplateDelimiter=~
			strTextEditorExe=notepad.exe
			blnSkipHelpReadyToEdit=0
			strLatestSkipped=%lAppVersion%
		)
		, %strIniFile%

IniRead, intDefaultWidth, %strIniFile%, global, intDefaultWidth ; used when export to fixed-width format
IniRead, strTemplateDelimiter, %strIniFile%, global, strTemplateDelimiter ; Default ~ (tilde), used when export to HTML and Express formats
IniRead, strTextEditorExe, %strIniFile%, global, strTextEditorExe ; Default notepad.exe
IniRead, blnSkipHelpReadyToEdit, %strIniFile%, global, blnSkipHelpReadyToEdit ; Default 0

intProgressType := -2 ; Status Bar, part 2


; --------------------- GUI1 --------------------------

Gui, 1:New, +Resize, % L(lAppName)

Gui, 1:Font, s12 w700, Verdana
Gui, 1:Add, Text, x10, % L(lAppName)

Gui, 1:Font, s10 w700, Verdana
Gui, 1:Add, Tab2, w950 r4 vtabCSVBuddy gChangedTabCSVBuddy, % L(lTab0List)
Gui, 1:Font

Gui, 1:Tab, 1
Gui, 1:Add, Text,		y+10	x10		vlblCSVFileToLoad w85 right, % L(lTab1CSVFileToLoad)
Gui, 1:Add, Edit,		yp		x100	vstrFileToLoad gChangedFileToLoad disabled
Gui, 1:Add, Button,		yp		x+5		vbtnHelpFileToLoad gButtonHelpFileToLoad, % L(lTab0QuestionMark)
Gui, 1:Add, Button,		yp		x+5		vbtnSelectFileToLoad gButtonSelectFileToLoad w45 default, % L(lTab1Select)
Gui, 1:Add, Text,		y+10	x10 	vlblHeader w85 right, % L(lTab1CSVFileHeader)
Gui, 1:Add, Edit,		yp		x100	vstrFileHeaderEscaped disabled
Gui, 1:Add, Button,		yp		x+5		vbtnHelpHeader gButtonHelpHeader, % L(lTab0QuestionMark)
Gui, 1:Add, Button,		yp		x+5		vbtnPreviewFile gButtonPreviewFile w45 hidden, % L(lTab1PreviewFile)
Gui, 1:Add, Radio,		y+10	x100	vradGetHeader gClickRadGetHeader checked, % L(lTab1Getheaderfromfile)
Gui, 1:Add, Radio,		yp		x+5		vradSetHeader gClickRadSetHeader, % L(lTab1Setheader)
Gui, 1:Add, Button,		yp		x+0		vbtnHelpSetHeader gButtonHelpSetHeader, % L(lTab0QuestionMark)
Gui, 1:Add, Text,		xp		x+27	vlblFieldDelimiter1, % L(lTab1Fielddelimiter)
Gui, 1:Add, Edit,		yp		x+5		vstrFieldDelimiter1 w20 limit1 center, `, ; gChangedFieldDelimiter1 unused
Gui, 1:Add, Button,		yp		x+5		vbtnHelpFieldDelimiter1 gButtonHelpFieldDelimiter1, % L(lTab0QuestionMark)
Gui, 1:Add, Text,		yp		x+27	vlblFieldEncapsulator1, % L(lTab1Fieldencapsulator)
Gui, 1:Add, Edit,		yp		x+5		vstrFieldEncapsulator1 w20 limit1 center, `" ; gChangedFieldEncapsulator1 unused
Gui, 1:Add, Button,		yp		x+5		vbtnHelpEncapsulator1 gButtonHelpEncapsulator1, % L(lTab0QuestionMark)
Gui, 1:Add, Checkbox,	yp		x+27	vblnMultiline1 gChangedMultiline1, % L(lTab1Multilinefields)
Gui, 1:Add, Button,		yp		x+0		vbtnHelpMultiline1 gButtonHelpMultiline1, % L(lTab0QuestionMark)
Gui, 1:Add, Text,		yp		x+5		vlblEndoflineReplacement1 hidden, % L(lTab1EOLreplacement)
Gui, 1:Add, Edit,		yp		x+5		vstrEndoflineReplacement1 w30 center hidden
Gui, 1:Add, Button,		yp		x+5		vbtnLoadFile gButtonLoadFile w45 hidden, % L(lTab1Load)

Gui, 1:Tab, 2
Gui, 1:Add, Text,		y+10	x10		vlblRenameFields w85 right, % L(lTab2Renamefields)
Gui, 1:Add, Edit,		yp		x100	vstrRenameEscaped
Gui, 1:Add, Button,		yp		x+0		vbtnSetRename gButtonSetRename w50, % L(lTab2Rename)
Gui, 1:Add, Button,		yp		x+5		vbtnHelpRename gButtonHelpRename, % L(lTab0QuestionMark)
Gui, 1:Add, Text,		y+10	x10		vlblSelectFields w85 right, % L(lTab2Selectfields)
Gui, 1:Add, Edit,		yp		x100	vstrSelectEscaped
Gui, 1:Add, Button,		yp		x+0		vbtnSetSelect gButtonSetSelect w50, % L(lTab2Select)
Gui, 1:Add, Button,		yp		x+5		vbtnHelpSelect gButtonHelpSelect, % L(lTab0QuestionMark)
Gui, 1:Add, Text,		y+10	x10		vlblOrderFields w85 right, % L(lTab2Orderfields)
Gui, 1:Add, Edit,		yp		x100	vstrOrderEscaped
Gui, 1:Add, Button,		yp		x+0		vbtnSetOrder gButtonSetOrder w50, % L(lTab2Order)
Gui, 1:Add, Button,		yp		x+5		vbtnHelpOrder gButtonHelpOrder, % L(lTab0QuestionMark)

Gui, 1:Tab, 3
Gui, 1:Add, Text,		y+10	x10		vlblCSVFileToSave w85 right, % L(lTab3CSVfiletosave)
Gui, 1:Add, Edit,		yp		x100	vstrFileToSave gChangedFileToSave
Gui, 1:Add, Button,		yp		x+5		vbtnHelpFileToSave gButtonHelpFileToSave, % L(lTab0QuestionMark)
Gui, 1:Add, Button,		yp		x+5		vbtnSelectFileToSave gButtonSelectFileToSave w45 default, % L(lTab3Select)
Gui, 1:Add, Text,		y+10	x100	vlblFieldDelimiter3, % L(lTab3Fielddelimiter)
Gui, 1:Add, Edit,		yp		x200	vstrFieldDelimiter3 gChangedFieldDelimiter3 w20 limit1 center, `, 
Gui, 1:Add, Button,		yp		x+5		vbtnHelpFieldDelimiter3 gButtonHelpFieldDelimiter3, % L(lTab0QuestionMark)
Gui, 1:Add, Text,		y+10	x100	vlblFieldEncapsulator3, % L(lTab3Fieldencapsulator)
Gui, 1:Add, Edit,		yp		x200	vstrFieldEncapsulator3 gChangedFieldEncapsulator3 w20 limit1 center, `"
Gui, 1:Add, Button,		yp		x+5		vbtnHelpEncapsulator3 gButtonHelpEncapsulator3, % L(lTab0QuestionMark)
Gui, 1:Add, Radio,		y100	x300	vradSaveWithHeader checked, % L(lTab3Savewithheader)
Gui, 1:Add, Radio,		y+10	x300	vradSaveNoHeader, % L(lTab3Savewithoutheader)
Gui, 1:Add, Button,		y100	x450	vbtnHelpSaveHeader gButtonHelpSaveHeader, % L(lTab0QuestionMark)
Gui, 1:Add, Radio,		y100	x500	vradSaveMultiline gClickRadSaveMultiline checked, % L(lTab3Savemultiline)
Gui, 1:Add, Radio,		y+10	x500	vradSaveSingleline gClickRadSaveSingleline, % L(lTab3Savesingleline)
Gui, 1:Add, Button,		y100	x620	vbtnHelpMultiline gButtonHelpSaveMultiline, % L(lTab0QuestionMark)
Gui, 1:Add, Text,		y+25	x500	vlblEndoflineReplacement3 hidden, % L(lTab3Endoflinereplacement)
Gui, 1:Add, Edit,		yp		x620	vstrEndoflineReplacement3 hidden w50 center, % chr(182)
Gui, 1:Add, Button,		y105	x+5		vbtnSaveFile gButtonSaveFile w45 hidden, % L(lTab3Save)
Gui, 1:Add, Button,		y137	x+5		vbtnCheckFile hidden w45 gButtonCheckFile, % L(lTab3Check)

Gui, 1:Tab, 4
Gui, 1:Add, Text,		y+10	x10		vlblCSVFileToExport w85 right, % L(lTab4Exportdatatofile)
Gui, 1:Add, Edit,		yp		x100	vstrFileToExport gChangedFileToExport
Gui, 1:Add, Button,		yp		x+5		vbtnHelpFileToExport gButtonHelpFileToExport, % L(lTab0QuestionMark)
Gui, 1:Add, Button,		yp		x+5		vbtnSelectFileToExport gButtonSelectFileToExport w45 default, % L(lTab4Select)
Gui, 1:Add, Text,		y+10	x10		vlblCSVExportFormat w85 right, % L(lTab4Exportformat)
Gui, 1:Add, Radio,		yp		x100	vradFixed gClickRadFixed, % L(lTab4Fixedwidth)
Gui, 1:Add, Radio,		yp		x+15	vradHTML gClickRadHTML, % L(lTab4HTML)
Gui, 1:Add, Radio,		yp		x+15	vradXML gClickRadXML, % L(lTab4XML)
Gui, 1:Add, Radio,		yp		x+15	vradExpress gClickRadExpress, % L(lTab4Express)
Gui, 1:Add, Button,		yp		x+15	vbtnHelpExportFormat gButtonHelpExportFormat, % L(lTab0QuestionMark)
Gui, 1:Add, Button,		yp		x+15	vbtnHelpExportMulti gButtonHelpExportMulti Hidden w125
Gui, 1:Add, Text,		y+10	x10		vlblMultiPurpose w85 right hidden, Hidden Label:
Gui, 1:Add, Edit,		yp		x100	vstrMultiPurpose hidden ; gChangedMultiPurpose unused
Gui, 1:Add, Button,		yp		x+5		vbtnMultiPurpose gButtonMultiPurpose hidden w120
Gui, 1:Add, Button,		y105	x+5		vbtnExportFile gButtonExportFile w45 hidden, % L(lTab4Export)
Gui, 1:Add, Button,		y137	x+5		vbtnCheckExportFile gButtonCheckExportFile w45 hidden, % L(lTab4Check)

Gui, 1:Tab, 5
Gui, 1:Font, s10 w700, Verdana
str32or64 := A_PtrSize  * 8
Gui, 1:Add, Link,		y+10	x25		vlblAboutText1, % L(lTab5Abouttext1, lAppName, lAppVersionLong, str32or64)
Gui, 1:Font, s9 w500, Arial
Gui, 1:Add, Link,		y+20	x25		vlblAboutText2, % L(lTab5Abouttext2)
Gui, 1:Add, Link,		yp		x+150	vlblAboutText3, % L(lTab5Abouttext3)
Gui, 1:Font
Gui, 1:Add, Button,						vbtnCheck4Update gButtonCheck4Update, % L(lTab5Check4Update)

Gui, 1:Tab

Gui, 1:Add, ListView, 	x10 r24 w200 vlvData -ReadOnly NoSort gListViewEvents AltSubmit -LV0x10

Gui, Add, StatusBar
SB_SetParts(200)
SB_SetText(L(lSBEmpty), 1)
if (A_IsCompiled)
	SB_SetIcon(A_ScriptFullPath)
else
	SB_SetIcon("C:\Dropbox\AutoHotkey\CSVBuddy\build\Ico - Visual Pharm\angel.ico")

GuiControl, 1:Focus, btnSelectFileToLoad
GuiControl, 1:+Default, btnSelectFileToLoad
Gui, 1:Show, Autosize

blnButtonCheck4Update := False
Gosub, Check4Update
Gosub, Check4CommandLineParameter

return



ChangedTabCSVBuddy:
Gui, 1:Submit, NoHide
if InStr(tabCSVBuddy, L(lTab0Load))
	GuiControl, 1:+Default, btnSelectFileToLoad
else if InStr(tabCSVBuddy, L(lTab0Edit))
	if LV_GetCount("Column")
		GuiControl, 1:+Default, btnReady
	else
	{
		Oops(lTab0FirstloadaCSVfile)
		GuiControl, 1:Choose, tabCSVBuddy, 1
	}
else if InStr(tabCSVBuddy, L(lTab0Save))
	if LV_GetCount("Column")
		GuiControl, 1:+Default, btnSelectFileToSave
	else
	{
		Oops(lTab0FirstloadaCSVfile)
		GuiControl, 1:Choose, tabCSVBuddy, 1
	}
else if InStr(tabCSVBuddy, L(lTab0Export))
	if LV_GetCount("Column")
		GuiControl, 1:+Default, btnSelectFileToExport
	else
	{
		Oops(lTab0FirstloadaCSVfile)
		GuiControl, 1:Choose, tabCSVBuddy, 1
	}
return




; --------------------- TAB 1 --------------------------


ButtonHelpFileToLoad:
Help(lTab1HelpFileToLoad, lAppName, lAppName)
return



DetectDelimiters:
Gui, 1:Submit, NoHide
strFileHeaderUnEscaped := StrUnEscape(strFileHeaderEscaped)
strCandidates := "`t;,:|~" ; check tab, semi-colon, comma, colon, pipe and tilde
strFieldDelimiterDetected := "," ; comma by default if no delimiter is detected
loop, Parse, strCandidates
	if InStr(strFileHeaderUnEscaped, A_LoopField)
	{
		strFieldDelimiterDetected := A_LoopField
		break
	}
GuiControl, 1:, strFieldDelimiter1, % StrMakeEncodedFieldDelimiter(strFieldDelimiterDetected)
strCandidates := """'~|" ; check double-quote, single-quote, tilde and pipe
strFieldEncapsulatorDetected := """" ; double-quotes by default if no encapsulator is detected
loop, Parse, strCandidates
	if (strFieldDelimiterDetected <> A_LoopField) and (InStr(strFileHeaderUnEscaped, strFieldDelimiterDetected . A_LoopField) or InStr(strFileHeaderUnEscaped, A_LoopField . strFieldDelimiterDetected))
	{
		strFieldEncapsulatorDetected := A_LoopField
		break
	}
GuiControl, 1:, strFieldEncapsulator1, %strFieldEncapsulatorDetected%
return



ButtonSelectFileToLoad:
Gui, 1:Submit, NoHide
Gui, 1:+OwnDialogs 
FileSelectFile, strInputFile, 3, %A_ScriptDir%, % L(lTab1SelectCSVFiletoload)
if !(StrLen(strInputFile))
	return
GuiControl, 1:, strFileToLoad, %strInputFile%
if (radGetHeader) or !(StrLen(strFileHeaderEscaped))
{
	FileReadLine, strCurrentHeader, %strInputFile%, 1
	GuiControl, 1:, strFileHeaderEscaped, % StrEscape(strCurrentHeader)
}
GuiControl, 1:+Default, btnLoadFile
GuiControl, 1:Focus, btnLoadFile
gosub, DetectDelimiters
return


ChangedFileToLoad:
Gui, 1:Submit, NoHide
strFileAttribute := FileExist(strFileToLoad)
if StrLen(strFileAttribute) and !InStr(strFileAttribute, "D")
{
	GuiControl, 1:Show, btnPreviewFile
	GuiControl, 1:Show, btnLoadFile
	GuiControl, 1:, strFileToSave, % NewFileName(strFileToLoad)
	GuiControl, 1:, strFileToExport, % NewFileName(strFileToLoad, "-EXPORT", "txt")
}
else
{
	GuiControl, 1:Hide, btnPreviewFile
	GuiControl, 1:Hide, btnLoadFile
	GuiControl, 1:, strFileToSave
	GuiControl, 1:, strFileToExport
}
return



ButtonHelpHeader:
Help(lTab1HelpHeader)
return



ButtonPreviewFile:
Gui, 1:Submit, NoHide
run, %strTextEditorExe% "%strFileToLoad%"
return



ClickRadGetHeader:
Gui, 1:Submit, NoHide
FileReadLine, strCurrentHeader, %strFileToLoad%, 1
GuiControl, 1:, strFileHeaderEscaped, % StrEscape(strCurrentHeader)
GuiControl, 1:Disable, strFileHeaderEscaped
GuiControl, 1:, lblHeader, % L(lTab1FileCSVHeader)
return



ClickRadSetHeader:
GuiControl, 1:Enable, strFileHeaderEscaped
GuiControl, 1:, lblHeader, % L(lTab1CustomHeader)
GuiControl, 1:Focus, strFileHeaderEscaped
return



ButtonHelpSetHeader:
Help(lTab1HelpSetHeader)
return



/*
ChangedFieldDelimiter1:
return
*/



ButtonHelpFieldDelimiter1:
Help(lTab1HelpDelimiter1, lAppName)
return



/*
ChangedFieldEncapsulator1:
return
*/



ButtonHelpEncapsulator1:
Help(lTab1HelpEncapsulator1, lAppName)
return



ChangedMultiline1:
Gui, 1:Submit, NoHide
if (blnMultiline1)
{
	GuiControl, 1:Show, lblEndoflineReplacement1
	GuiControl, 1:Show, strEndoflineReplacement1
}
else
{
	GuiControl, 1:Hide, lblEndoflineReplacement1
	GuiControl, 1:Hide, strEndoflineReplacement1
}
return



ButtonHelpMultiline1:
Help(lTab1HelpMultiline1)
return


ButtonLoadFile:
Gui, 1:+OwnDialogs
Gui, 1:Submit, NoHide
if !DelimitersOK(1)
	return
if !StrLen(strFileToLoad)
{
	Oops(lTab1FirstusetheSelectbutton)
	return
}
if !StrLen(strFileHeaderEscaped) and (radSetHeader)
{
	MsgBox, 52, % L(lAppName), % L(lTab1CSVHeaderisnotspecified)
	IfMsgBox, No
		return
}
if LV_GetCount("Column")
{
	MsgBox, 36, % L(lAppName), % L(lTab1Replacethecurrentcontentof)
	IfMsgBox, Yes
		gosub, DeleteListviewData
	IfMsgBox, No
	{
		MsgBox, 36, %lAppName%, % L(lTab1DoYouWantToAdd)
		IfMsgBox, No
			return
	}
}
else
	intActualSize := 0

strCurrentHeader := StrUnEscape(strFileHeaderEscaped)
strCurrentFieldDelimiter := StrMakeRealFieldDelimiter(strFieldDelimiter1)
strCurrentVisibleFieldDelimiter := strFieldDelimiter1
strCurrentFieldEncapsulator := strFieldEncapsulator1

FileGetSize, intFileSize, %strFileToLoad%, K
intActualSize := intActualSize + intFileSize

; ObjCSV_CSV2Collection(strFilePath, ByRef strFieldNames [, blnHeader = 1, blnMultiline = 1, intProgressType = 0
;	, strFieldDelimiter = ",", strEncapsulator = """", strEolReplacement = "", strProgressText := ""])
obj := ObjCSV_CSV2Collection(strFileToLoad, strCurrentHeader, radGetHeader, blnMultiline1, intProgressType
	, strCurrentFieldDelimiter, strCurrentFieldEncapsulator, strEndoflineReplacement1, L(lTab1ReadingCSVdata))
if (ErrorLevel)
{
	if (ErrorLevel = 3)
		strError := L(lTab1CSVfilenotloadedNoUnusedRepl)
	else
	{
		strError := L(lTab1CSVfilenotloadedTooLarge, intActualSize)
		if (A_PtrSize = 4) ; 32-bits
			strError := strError . L(lTab1Trythe64bitsversion)
	}
	Oops(strError)
	SB_SetText(lSBEmpty, 1)
	return
}
SB_SetText("", 2)
; ObjCSV_Collection2ListView(objCollection [, strGuiID = "", strListViewID = "", strFieldOrder = "", strFieldDelimiter = ","
;	, strEncapsulator = """", strSortFields = "", strSortOptions = "", intProgressType = 0, strProgressText = ""])
ObjCSV_Collection2ListView(obj, "1", "lvData", strCurrentHeader, strCurrentFieldDelimiter
	, strCurrentFieldEncapsulator, , , intProgressType, L(lTab0LoadingToList))
if (ErrorLevel)
{
	Oops(lTab1CSVfilenotloadedMax200fields)
	SB_SetText(lSBEmpty, 1)
	return
}
SB_SetText(L(lSBRecordsSize, LV_GetCount(), (intActualSize) ? intActualSize : " <1"), 1)
Gosub, UpdateCurrentHeader
if (!blnSkipHelpReadyToEdit)
	Help(lTab1HelpReadyToEdit)
GuiControl, 1:, strFieldDelimiter3, %strCurrentVisibleFieldDelimiter%
GuiControl, 1:, strFieldEncapsulator3, %strCurrentFieldEncapsulator%
obj := ; release object
return


DeleteListviewData:
LV_Delete() ; delete all rows - better performance on large files when we delete rows before columns
loop, % LV_GetCount("Column")
	LV_DeleteCol(1) ; delete all columns
SB_SetText(L(lSBEmpty), 1)
intActualSize := 0
return



; --------------------- TAB 2 --------------------------


ButtonSetRename:
Gui, 1:+OwnDialogs 
Gui, 1:Submit, NoHide
if !LV_GetCount()
{
	Oops(lTab0FirstloadaCSVfile)
	GuiControl, 1:Choose, tabCSVBuddy, 1
	return
}
; ObjCSV_ReturnDSVObjectArray(strCurrentDSVLine, strDelimiter = ",", strEncapsulator = """")
objNewHeader := ObjCSV_ReturnDSVObjectArray(StrUnEscape(strRenameEscaped), strCurrentFieldDelimiter, strCurrentFieldEncapsulator)
intNbFieldNames := objNewHeader.MaxIndex()
intNbColumns := LV_GetCount("Column")
if !StrLen(strRenameEscaped)
{
	MsgBox, 52, %lAppName%, % L(lTab2RenameNoString, strCurrentVisibleFieldDelimiter)
	IfMsgBox, No
		return
}
else if (intNbFieldNames < intNbColumns)
{
	MsgBox, 52, %lAppName%, % L(lTab2RenameLessNames, intNbFieldNames, intNbColumns)
	IfMsgBox, No
		return
}
Loop, % LV_GetCount("Column")
{
	if StrLen(objNewHeader[A_Index])
		LV_ModifyCol(A_Index, "", objNewHeader[A_Index])
	else
		LV_ModifyCol(A_Index, "", "C" . A_Index)
	LV_ModifyCol(A_Index, "AutoHdr")
}
Gosub, UpdateCurrentHeader
objNewHeader := ; release object
return



ButtonHelpRename:
Gui, 1:Submit, NoHide
Help(lTab2HelpRename, strCurrentVisibleFieldDelimiter, strCurrentVisibleFieldDelimiter, strCurrentFieldEncapsulator)
return



ButtonSetSelect:
Gui, 1:Submit, NoHide
if !LV_GetCount()
{
	Oops(lTab0FirstloadaCSVfile)
	GuiControl, 1:Choose, tabCSVBuddy, 1
	return
}
if !StrLen(strSelectEscaped)
{
	Oops(lTab2SelectNoString, strCurrentVisibleFieldDelimiter)
	return
}
; ObjCSV_ReturnDSVObjectArray(strCurrentDSVLine, strDelimiter = ",", strEncapsulator = """")
objCurrentHeader := ObjCSV_ReturnDSVObjectArray(strCurrentHeader, strCurrentFieldDelimiter, strCurrentFieldEncapsulator)
objNewHeader := ObjCSV_ReturnDSVObjectArray(StrUnEscape(strSelectEscaped), strCurrentFieldDelimiter, strCurrentFieldEncapsulator)
intPosPrevious := 0
for intKey, strVal in objNewHeader
{
	intPosThisOne := PositionInArray(strVal, objCurrentHeader)
	if !(intPosThisOne)
	{
		Oops(lTab2SelectFieldMissing, strVal)
		return
	}
	if (intPosThisOne <= intPosPrevious)
	{
		Oops(lTab2SelectBadOrder)
		return
	}
	intPosPrevious := intPosThisOne
}
intMaxCurrent := objCurrentHeader.MaxIndex()
intMaxNew := objNewHeader.MaxIndex()
intIndexCurrent := 1
intIndexNew := 1
intDeleted := 0
Loop
{
	if (objCurrentHeader[intIndexCurrent] = objNewHeader[intIndexNew])
	{
		intIndexCurrent := intIndexCurrent + 1
		intIndexNew := intIndexNew + 1
	}
	else
	{
		LV_DeleteCol(intIndexCurrent - intDeleted)
		intDeleted := intDeleted + 1
		intIndexCurrent := intIndexCurrent + 1
	}
	if (intIndexCurrent > intMaxCurrent)
		break
}
Gosub, UpdateCurrentHeader
objCurrentHeader := ; release object
objNewHeader := ; release object
return



ButtonHelpSelect:
Gui, 1:Submit, NoHide
Help(lTab2HelpSelect, strCurrentVisibleFieldDelimiter)
return



ButtonSetOrder:
Gui, 1:Submit, NoHide
if !StrLen(strOrderEscaped)
{
	Oops(lTab2OrderNoString, strCurrentVisibleFieldDelimiter)
	return
}
if !LV_GetCount()
{
	Oops(lTab0FirstloadaCSVfile)
	GuiControl, 1:Choose, tabCSVBuddy, 1
	return
}
; ObjCSV_ReturnDSVObjectArray(strCurrentDSVLine, strDelimiter = ",", strEncapsulator = """")
objCurrentHeader := ObjCSV_ReturnDSVObjectArray(strCurrentHeader, strCurrentFieldDelimiter, strCurrentFieldEncapsulator)
objNewHeader := ObjCSV_ReturnDSVObjectArray(StrUnEscape(strOrderEscaped), strCurrentFieldDelimiter, strCurrentFieldEncapsulator)
for intKey, strVal in objNewHeader
{
	if !PositionInArray(strVal, objCurrentHeader)
	{
		Oops(lTab2OrderFieldMissing, strVal)
		return
	}
}
LV_Modify(0, "-Select") ; Make sure all rows will be transfered to objNewCollection
; ObjCSV_ListView2Collection([strGuiID = "", strListViewID = "", strFieldOrder = "", strFieldDelimiter = ","
;	, strEncapsulator = """", intProgressType = 0, strProgressText = ""])
objNewCollection := ObjCSV_ListView2Collection("1", "lvData", StrUnEscape(strOrderEscaped), strCurrentFieldDelimiter
	, strCurrentFieldEncapsulator, intProgressType, L(lTab0ReadingFromList))
LV_Delete() ; better performance on large files when we delete rows before columns
loop, % LV_GetCount("Column")
	LV_DeleteCol(1) ; delete all rows
; ObjCSV_Collection2ListView(objCollection [, strGuiID = "", strListViewID = "", strFieldOrder = "", strFieldDelimiter = ","
;	, strEncapsulator = """", strSortFields = "", strSortOptions = "", intProgressType = 0, strProgressText = ""])
ObjCSV_Collection2ListView(objNewCollection, "1", "lvData", StrUnEscape(strOrderEscaped), strCurrentFieldDelimiter
	, strCurrentFieldEncapsulator, , , intProgressType, lTab0LoadingToList)
if (ErrorLevel)
	Oops(lTab2OrderNotLoaded)
Gosub, UpdateCurrentHeader
objNewCollection := ; release object
return



ButtonHelpOrder:
Gui, 1:Submit, NoHide
Help(lTab2HelpOrder, strCurrentVisibleFieldDelimiter)
return




; --------------------- TAB 3 --------------------------


ButtonHelpFileToSave:
Help(lTab3HelpFileToSave)
return



ButtonSelectFileToSave:
Gui, 1:Submit, NoHide
Gui, 1:+OwnDialogs 
FileSelectFile, strOutputFile, 2, %A_ScriptDir%, % L(lTab3SelectCSVFiletosave)
if !(StrLen(strOutputFile))
	return
GuiControl, 1:, strFileToSave, %strOutputFile%
GuiControl, 1:+Default, btnSaveFile
GuiControl, 1:Focus, btnSaveFile
return



ChangedFileToSave:
Gui, 1:Submit, NoHide
SplitPath, strFileToSave, , strOutDir, strOutExtension, strOutNameNoExt
if FileExist(strFileToSave)
	GuiControl, 1:Show, btnCheckFile
else
	GuiControl, 1:Hide, btnCheckFile
if StrLen(strOutNameNoExt)
	GuiControl, 1:Show, btnSaveFile
else
	GuiControl, 1:Hide, btnSaveFile
return



ButtonHelpFieldDelimiter3:
Help(lTab3HelpFieldDelimiter3)
return



ChangedFieldDelimiter3:
strPreviousDelimiter := strFieldDelimiter3
Gui, 1:Submit, NoHide
if (strPreviousDelimiter = strFieldDelimiter3)
	return ; avoid a loop if a field name contains a delimiter at first load time
if StrLen(strFieldDelimiter3)
{
	if !NewDelimiterOrEncapsulatorOK(StrMakeRealFieldDelimiter(strFieldDelimiter3))
	{
		Oops(lTab3BadDelimiter, strFieldDelimiter3)
		GuiControl, 1:, strFieldDelimiter3, %strPreviousDelimiter%
		return
	}
	Gosub, UpdateCurrentHeader
}
return



ChangedFieldEncapsulator3:
strPreviousEncapsulator := strFieldEncapsulator3
Gui, 1:Submit, NoHide
if (strPreviousEncapsulator = strFieldEncapsulator3)
	return ; avoid a loop if a field name contains an encapsulator at first load time
if StrLen(strFieldEncapsulator3)
{
	if !NewDelimiterOrEncapsulatorOK(strFieldEncapsulator3)
	{
		Oops(lTab3BadEncapsulator, strFieldEncapsulator3)
		GuiControl, 1:, strFieldEncapsulator3, %strPreviousEncapsulator%
		return
	}
	Gosub, UpdateCurrentHeader
}
return



ButtonHelpEncapsulator3:
Help(lTab3HelpEncapsulator3)
return



ButtonHelpSaveHeader:
Help(lTab3HelpSaveHeader)
return



ClickRadSaveMultiline:
Gui, 1:Submit, NoHide
GuiControl, 1:Hide, lblEndoflineReplacement3
GuiControl, 1:Hide, strEndoflineReplacement3
return



ClickRadSaveSingleline:
Gui, 1:Submit, NoHide
GuiControl, 1:Show, lblEndoflineReplacement3
GuiControl, 1:Show, strEndoflineReplacement3
return



ButtonHelpSaveMultiline:
Gui, 1:Submit, NoHide
Help(lTab3HelpSaveMultiline)
return



ButtonSaveFile:
Gui, 1:Submit, NoHide
if !DelimitersOK(3)
	return
blnOverwrite := CheckIfFileExistOverwrite(strFileToSave)
if (blnOverwrite < 0)
	return
if !CheckOneRow()
	return
; ObjCSV_ListView2Collection([strGuiID = "", strListViewID = "", strFieldOrder = "", strFieldDelimiter = ","
;	, strEncapsulator = """", intProgressType = 0, strProgressText = ""])
obj := ObjCSV_ListView2Collection("1", "lvData", , , , intProgressType, L(lTab0ReadingFromList))
if (radSaveMultiline)
	strEolReplacement := ""
else
	strEolReplacement := strEndoflineReplacement3
strRealFieldDelimiter3 := StrMakeRealFieldDelimiter(strFieldDelimiter3)
; ObjCSV_Collection2CSV(objCollection, strFilePath [, blnHeader = 0
;	, strFieldOrder = "", intProgressType = 0, blnOverwrite = 0
;	, strFieldDelimiter = ",", strEncapsulator = """", strEolReplacement = "", strProgressText = ""])
ObjCSV_Collection2CSV(obj, strFileToSave, radSaveWithHeader
	, GetListViewHeader(strRealFieldDelimiter3, strFieldEncapsulator3), intProgressType, blnOverwrite
	, strRealFieldDelimiter3, strFieldEncapsulator3, strEolReplacement, L(lTab3SavingCSV))
if FileExist(strFileToSave)
{
	GuiControl, 1:Show, btnCheckFile
	GuiControl, 1:+Default, btnCheckFile
	GuiControl, 1:Focus, btnCheckFile
}
obj := ; release object
return



ButtonCheckFile:
Gui, 1:Submit, NoHide
run, %strTextEditorExe% "%strFileToSave%"
return




; --------------------- TAB 4 --------------------------


ButtonHelpFileToExport:
Help(lTab4HelpFileToExport)
return



ButtonSelectFileToExport:
Gui, 1:Submit, NoHide
Gui, 1:+OwnDialogs 
FileSelectFile, strOutputFile, 2, %A_ScriptDir%, % L(lTab4Selectexportfile)
if !(StrLen(strOutputFile))
	return
GuiControl, 1:, strFileToExport, %strOutputFile%
GuiControl, 1:+Default, btnExportFile
GuiControl, 1:Focus, btnExportFile
return



ChangedFileToExport:
Gui, 1:Submit, NoHide
SplitPath, strFileToExport, , strOutDir, strOutExtension, strOutNameNoExt
if FileExist(strFileToExport)
	GuiControl, 1:Show, btnCheckExportFile
else
	GuiControl, 1:Hide, btnCheckExportFile
if StrLen(strOutNameNoExt)
	GuiControl, 1:Show, btnExportFile
else
	GuiControl, 1:Hide, btnExportFile
return



ClickRadFixed:
Gui, 1:Submit, NoHide
if !DelimitersOK(3)
{
	GuiControl, , radFixed, 0
	return
}
GuiControl, 1:Show, btnHelpExportMulti
GuiControl, 1:, btnHelpExportMulti, % L(lTab4FixedwidthExportHelp)
GuiControl, 1:Show, lblMultiPurpose
GuiControl, 1:, lblMultiPurpose, % L(lTab4Fieldswidth)
GuiControl, 1:Show, strMultiPurpose
GuiControl, 1:, strMultiPurpose
GuiControl, 1:Show, btnMultiPurpose
GuiControl, 1:, btnMultiPurpose, % L(lTab4Changedefaultwidth)
GuiControl, 1:, strFileToExport, % NewFileName(strFileToLoad, "-EXPORT", "txt")
; ObjCSV_ReturnDSVObjectArray(strCurrentDSVLine, strDelimiter = ",", strEncapsulator = """")
objCurrentHeader := ObjCSV_ReturnDSVObjectArray(strCurrentHeader, strCurrentFieldDelimiter, strCurrentFieldEncapsulator)
; strCurrentFieldDelimiter (strFieldDelimiter1) et strCurrentFieldEncapsulator (strFieldEncapsulator1) pour la lecture de strCurrentHeader
strRealFieldDelimiter3 := StrMakeRealFieldDelimiter(strFieldDelimiter3)
strMultiPurpose := ""
Loop, % objCurrentHeader.MaxIndex()
{
	; ObjCSV_Format4CSV(strF4C [, strFieldDelimiter = ",", strEncapsulator = """"])
	strFormat4Csv := ObjCSV_Format4CSV(objCurrentHeader[A_Index], strRealFieldDelimiter3, strFieldEncapsulator3)
	strMultiPurpose := strMultiPurpose . strFormat4Csv . strRealFieldDelimiter3 . intDefaultWidth . strRealFieldDelimiter3
	; strFieldDelimiter3 and strFieldEncapsulator3 for file writing
}
StringTrimRight, strMultiPurpose, strMultiPurpose, 1 ; remove extra delimiter
GuiControl, 1:, strMultiPurpose, % StrEscape(strMultiPurpose)
return



ClickRadHTML:
Gui, 1:Submit, NoHide
GuiControl, 1:Show, btnHelpExportMulti
GuiControl, 1:, btnHelpExportMulti, % L(lTab4HTMLExportHelp)
GuiControl, 1:Show, lblMultiPurpose
GuiControl, 1:, lblMultiPurpose, % L(lTab4HTMLtemplate)
GuiControl, 1:Show, strMultiPurpose
GuiControl, 1:, strMultiPurpose
GuiControl, 1:Show, btnMultiPurpose
GuiControl, 1:, btnMultiPurpose, % L(lTab4SelectHTMLtemplate)
GuiControl, 1:, strFileToExport, % NewFileName(strFileToLoad, "-EXPORT", "html")
return



ClickRadXML:
Gui, 1:Submit, NoHide
GuiControl, 1:Show, btnHelpExportMulti
GuiControl, 1:, btnHelpExportMulti, % L(lTab4XMLExportHelp)
GuiControl, 1:Hide, lblMultiPurpose
GuiControl, 1:Hide, strMultiPurpose
GuiControl, 1:, strMultiPurpose
GuiControl, 1:Hide, btnMultiPurpose
GuiControl, 1:, strFileToExport, % NewFileName(strFileToLoad, "-EXPORT", "xml")
return



ClickRadExpress:
Gui, 1:Submit, NoHide
if !DelimitersOK(3)
{
	GuiControl, , radExpress, 0
	return
}
GuiControl, 1:Show, btnHelpExportMulti
GuiControl, 1:, btnHelpExportMulti, % L(lTab4ExpressExportHelp)
GuiControl, 1:Show, lblMultiPurpose
GuiControl, 1:, lblMultiPurpose, % L(lTab4Expresstemplate)
GuiControl, 1:Show, strMultiPurpose
GuiControl, 1:, strMultiPurpose
GuiControl, 1:Hide, btnMultiPurpose
GuiControl, 1:, strFileToExport, % NewFileName(strFileToLoad, "-EXPORT", "txt")
; ObjCSV_ReturnDSVObjectArray(strCurrentDSVLine, strDelimiter = ",", strEncapsulator = """")
objCurrentHeader := ObjCSV_ReturnDSVObjectArray(strCurrentHeader, strCurrentFieldDelimiter, strCurrentFieldEncapsulator)
; strCurrentFieldDelimiter and strCurrentFieldEncapsulator for file reading of strCurrentHeader
strRealFieldDelimiter3 := StrMakeRealFieldDelimiter(strFieldDelimiter3)
strMultiPurpose := ""
Loop, % objCurrentHeader.MaxIndex()
{
	; ObjCSV_Format4CSV(strF4C [, strFieldDelimiter = ",", strEncapsulator = """"])
	strFormat4Csv := ObjCSV_Format4CSV(objCurrentHeader[A_Index], strRealFieldDelimiter3, strFieldEncapsulator3)
	strMultiPurpose := strMultiPurpose . strTemplateDelimiter . strFormat4Csv . strTemplateDelimiter . A_Space
	; strFieldDelimiter3 and strFieldEncapsulator3 for file writing
}
StringTrimRight, strMultiPurpose, strMultiPurpose, 1 ; remove extra delimiter
GuiControl, 1:, strMultiPurpose, % StrEscape(strMultiPurpose)
return



ButtonHelpExportFormat:
Help(lTab4HelpExportFormat)
return



ButtonHelpExportMulti:
if (radFixed)
	Help(lTab4HelpExportFixed, strFieldDelimiter3, intDefaultWidth)
else if (radHTML)
	Help(lTab4HelpExportHTML)
else if (radXML)
	Help(lTab4HelpExportXML)
else if (radExpress)
	Help(lTab4HelpExportExpress)
return



/*
ChangedMultiPurpose:
return
*/



ButtonMultiPurpose:
Gui, 1:Submit, NoHide
Gui, 1:+OwnDialogs 
if (radFixed)
{
	InputBox, intNewDefaultWidth, % L(lTab4MultiFixedInputTitle, lAppName, lAppVersionLong)
		, % L(lTab4MultiFixedInputPrompt), , , 120, , , , , %intDefaultWidth%
	if !ErrorLevel
		if (intNewDefaultWidth > 0)
			intDefaultWidth := intNewDefaultWidth
		else
			Oops(L(lTab4MultiFixedGreaterZero))
	Gosub, ClickRadFixed
}
else if (radHTML)
{
	FileSelectFile, strHtmlTemplateFile, 3, %A_ScriptDir%, % L(lTab4MultiHTMLSelectHTMLtemplate)
	if !(StrLen(strHtmlTemplateFile))
		return
	GuiControl, 1:, strMultiPurpose, %strHtmlTemplateFile%
}
; else if (radXML)
; else if (radExpress)
return



ButtonExportFile:
Gui, 1:Submit, NoHide
if (radFixed or radExpress)
	blnOverwrite := CheckIfFileExistOverwrite(strFileToExport)
else
	blnOverwrite := true
if (blnOverwrite < 0)
	return
if !CheckOneRow()
	return
if (radFixed)
	if StrLen(strMultiPurpose)
		Gosub, ExportFixed
	else
	{
		Oops(lTab4ExportFixedNoString, strFieldDelimiter3)
		return
	}
else if (radHTML)
	if StrLen(strMultiPurpose)
		Gosub, ExportHTML
	else
	{
		Oops(lTab4ExportHTMLNoString)
		return
	}
else if (radXML)
	Gosub, ExportXML
else if (radExpress)
	if StrLen(strMultiPurpose)
		Gosub, ExportExpress
	else
	{
		Oops(lTab4ExportExpressNoString)
		return
	}
else
	Oops(lTab4ExportSelectFormat)
return



ButtonCheckExportFile:
Gui, 1:Submit, NoHide
if InStr(strFileToExport, ".htm")
	run, "%strFileToExport%"
else
	run, %strTextEditorExe% "%strFileToExport%"
return




; --------------------- TAB 5 --------------------------


ButtonCheck4Update:
blnButtonCheck4Update := True
Gosub, Check4Update
return

; --------------------- LISTVIEW EVENTS --------------------------


ListViewEvents:
if (A_GuiEvent = "ColClick")
{
	intColNumber := A_EventInfo
	Menu, SortMenu, Add, % L(lLvEventsSortalphabetical), MenuSortText
	Menu, SortMenu, Add, % L(lLvEventsSortnumericInteger), MenuSortInteger
	Menu, SortMenu, Add, % L(lLvEventsSortnumericFloat), MenuSortFloat
	Menu, SortMenu, Add, % L(lLvEventsSortdescalphabetical), MenuSortDescText
	Menu, SortMenu, Add, % L(lLvEventsSortdescnumericInteger), MenuSortDescInteger
	Menu, SortMenu, Add, % L(lLvEventsSortdescnumericFloat), MenuSortDescFloat
	Menu, SortMenu, Show
}
if (A_GuiEvent = "DoubleClick")
{
	intRowNumber := A_EventInfo
	Gosub, MenuEditRow
}
SB_SetText(L(lLvEventsrecordsselected, LV_GetCount("Selected")), 2)
return



MenuSortText:
MenuSortInteger:
MenuSortFloat:
MenuSortDescText:
MenuSortDescInteger:
MenuSortDescFloat:
StringReplace, strOption, A_ThisLabel, Menu,
StringReplace, strOption, strOption, Sort, % "Sort "
StringReplace, strOption, strOption, Sort Desc, % "SortDesc "
LV_ModifyCol(intColNumber, strOption)
if InStr(strOption, "Text")
	LV_ModifyCol(intColNumber, "Left")
Menu, SortMenu, Delete
return



GuiContextMenu: ; Launched in response to a right-click or press of the Apps key.
if A_GuiControl <> lvData  ; Display the menu only for clicks inside the ListView.
    return
Menu, SelectMenu, Add
Menu, SelectMenu, DeleteAll ; to avoid ghost lines at the end when menu is re-created
if !LV_GetCount("Column")
	Menu, SelectMenu, Add, % L(lLvEventsCreateNewFile), MenuCreateNewFile
else if !LV_GetCount("")
{
	Menu, SelectMenu, Add, % L(lLvEventsAddrowMenu), MenuAddRow
	Menu, SelectMenu, Add, % L(lLvEventsCreateNewFile), MenuCreateNewFile
}
else
{
	intRowNumber := A_EventInfo
	Menu, SelectMenu, Add, % L(lLvEventsSelectAll), MenuSelectAll
	Menu, SelectMenu, Add, % L(lLvEventsDeselectAll), MenuSelectNone
	Menu, SelectMenu, Add, % L(lLvEventsReverseSelection), MenuSelectReverse
	Menu, SelectMenu, Add
	Menu, SelectMenu, Add, % L(lLvEventsEditrowMenu), MenuEditRow
	Menu, SelectMenu, Add, % L(lLvEventsAddrowMenu), MenuAddRow
	Menu, SelectMenu, Add, % L(lLvEventsDeleteRowMenu), MenuDeleteRow
}
; Show the menu at the provided coordinates, A_GuiX and A_GuiY.  These should be used
; because they provide correct coordinates even if the user pressed the Apps key:
Menu, SelectMenu, Show, %A_GuiX%, %A_GuiY%
return



MenuCreateNewFile:
Gui, 1:+OwnDialogs 
Gui, 1:Submit, NoHide
if !StrLen(strFileHeaderEscaped)
{
	Oops(L(lTab1NewFileInstructions, StrEscape(StrMakeRealFieldDelimiter(strFieldDelimiter1))))
	GuiControl, , radSetHeader, 1
	gosub, ClickRadSetHeader
}
else
{
	if LV_GetCount("Column")
	{
		MsgBox, 36, %lAppName%, % L(lTab1Replacethecurrentcontentof)
		IfMsgBox, Yes
			gosub, DeleteListviewData
		IfMsgBox, No
			return
	}
	else
		intActualSize := 0
	
	strCurrentHeader := StrUnEscape(strFileHeaderEscaped)
	strCurrentFieldDelimiter := StrMakeRealFieldDelimiter(strFieldDelimiter1)
	strCurrentVisibleFieldDelimiter := strFieldDelimiter1
	strCurrentFieldEncapsulator := strFieldEncapsulator1

	GuiControl, , strFileToLoad ; empty the file to load field
	
	objHeader := ObjCSV_ReturnDSVObjectArray(strCurrentHeader, strCurrentFieldDelimiter, strCurrentFieldEncapsulator)
	if objHeader.MaxIndex() > 200 ; ListView cannot display more that 200 columns
		Oops(lTab1CSVfilenotcreatedMax200fields)
	for intIndex, strFieldName in objHeader
		LV_InsertCol(intIndex, "", Trim(strFieldName))
	gosub, MenuAddRow
	gosub, UpdateCurrentHeader
	loop, % LV_GetCount("Column")
		LV_ModifyCol(A_Index, "AutoHdr")
}
return



MenuSelectAll:
GuiControl, Focus, lvData
LV_Modify(0, "Select")
Menu, SelectMenu, Delete
return



MenuSelectNone:
GuiControl, Focus, lvData
LV_Modify(0, "-Select")
Menu, SelectMenu, Delete
return



MenuSelectReverse:
GuiControl, Focus, lvData
Loop, % LV_GetCount()
	if IsRowSelected(A_Index)
		LV_Modify(A_Index, "-Select")
	else
		LV_Modify(A_Index, "Select")
Menu, SelectMenu, Delete
return



IsRowSelected(intRow)
{
	intNextSelectedRow := LV_GetNext(intRow - 1)
	return (intNextSelectedRow = intRow)
}



MenuAddRow:
MenuEditRow:
Gui, 1:Submit, NoHide
if (A_ThisLabel = "MenuAddRow")
{
	LV_Modify(0, "-Select")
	LV_Insert(0xFFFF, "Select Focus") ; add at the end of the list
	intRowNumber := LV_GetNext()
	LV_Modify(intRowNumber, "Vis")
	strSaveRecordButton := "ButtonSaveRecordAddRow"
	strCancelButton := "ButtonCancelAddRow"
}
else
{
	strSaveRecordButton := "ButtonSaveRecord"
	strCancelButton := "ButtonCancel"
}

if (intRowNumber = 0)
	intRowNumber := 1
intGui1WinID := WinExist("A")
Gui, 2:New, +Resize , % L(lLvEventsEditrow, lAppName)
Gui, 2:+Owner1
Gui, 1:Default
SysGet, intMonWork, MonitorWorkArea 
intColWidth := 380
intEditWidth := intColWidth - 20
intMaxNbCol := Floor(intMonWorkRight / intColWidth)
intX := 10
intY := 5
intCol := 1
loop, % LV_GetCount("Column")
{
	if ((intY + 100) > intMonWorkBottom)
	{
		if (intCol = 1)
		{
			Gui, 2:Add, Button, y%intY% x10 vbtnSaveRecord g%strSaveRecordButton%, % L(lLvEventsSave)
			Gui, 2:Add, Button, yp x+5 vbtnCancel g%strCancelButton%, % L(lLvEventsCancel)
		}
		if (intCol = intMaxNbCol)
		{
			intYLabel := intY
			Gui, 2:Add, Text, y%intYLabel% x%intX% vstrLabelMissing, % L(lLvEventsFieldsMissing)
			break
		}
		intCol := intCol + 1
		intX := intX + intColWidth
		intY := 5
	}
	intYLabel := intY
	intYEdit := intY + 15
	LV_GetText(strColHeader, 0, A_Index)
	LV_GetText(strColData, intRowNumber, A_Index)
	Gui, 2:Add, Text, y%intYLabel% x%intX% vstrLabel%A_Index%, %strColHeader%
	Gui, 2:Add, Edit, y%intYEdit% x%intX% w%intEditWidth% vstrEdit%A_Index% +HwndstrEditHandle, %strColData%
	ShrinkEditControl(strEditHandle, 2, "2")
	GuiControlGet, intPosEdit, 2:Pos, %strEditHandle%
	intY := intY + intPosEditH + 19
	intNbFieldsOnScreen := A_Index ; incremented at each occurence of the loop
}
if (intCol = 1) ; duplicate of lines above in the loop, but much simpler that way
{
	Gui, 2:Add, Button, y%intY% x10 vbtnSaveRecord g%strSaveRecordButton%, % L(lLvEventsSave)
	Gui, 2:Add, Button, yp x+5 vbtnCancel g%strCancelButton%, % L(lLvEventsCancel)
}
Gui, 2:Show, AutoSize Center
Gui, 1:+Disabled
return



MenuDeleteRow:
intPrevNbRows := LV_GetCount()
intRowNumber := 0 ; scan each selected row of the ListView
GuiControl, -Redraw, lvData ; stop drawing the ListView during delete
loop, % LV_GetCount("Selected")
{
	intRowNumber := LV_GetNext(intRowNumber) ; get next selected row number
	LV_Delete(intRowNumber)
	intRowNumber := intRowNumber - 1 ; continue searching from the row before the deleted row
}
GuiControl, +Redraw, lvData ; redraw the ListView
intNewNbRows := LV_GetCount()
intActualSize := Round(intActualSize * intNewNbRows / intPrevNbRows)
if (intNewNbRows)
	SB_SetText(L(lSBRecordsSize, intNewNbRows, (intActualSize) ? intActualSize : " <1"))
else
	SB_SetText(L(lSBEmpty), 1)
return




; --------------------- GUI1  --------------------------


GuiSize: ; Expand or shrink the ListView in response to the user's resizing of the window.
if A_EventInfo = 1  ; The window has been minimized.  No action needed.
    return
; Otherwise, the window has been resized or maximized. Resize the controls to match.
GuiControl, 1:Move, tabCSVBuddy, % "W" . (A_GuiWidth - 20)

GuiControl, 1:Move, strFileToLoad, % "W" . (A_GuiWidth - 200)
GuiControl, 1:Move, btnHelpFileToLoad, % "X" . (A_GuiWidth - 90)
GuiControl, 1:Move, btnSelectFileToLoad, % "X" . (A_GuiWidth - 65)
GuiControl, 1:Move, strFileHeaderEscaped, % "W" . (A_GuiWidth - 200)
GuiControl, 1:Move, btnHelpHeader, % "X" . (A_GuiWidth - 90)
GuiControl, 1:Move, btnPreviewFile, % "X" . (A_GuiWidth - 65)
GuiControl, 1:Move, btnLoadFile, % "X" . (A_GuiWidth - 65)

GuiControl, 1:Move, strRenameEscaped, % "W" . (A_GuiWidth - 205)
GuiControl, 1:Move, btnSetRename, % "X" . (A_GuiWidth - 95)
GuiControl, 1:Move, btnHelpRename, % "X" . (A_GuiWidth - 40)
GuiControl, 1:Move, strSelectEscaped, % "W" . (A_GuiWidth - 205)
GuiControl, 1:Move, btnSetSelect, % "X" . (A_GuiWidth - 95)
GuiControl, 1:Move, btnHelpSelect, % "X" . (A_GuiWidth - 40)
GuiControl, 1:Move, strOrderEscaped, % "W" . (A_GuiWidth - 205)
GuiControl, 1:Move, btnSetOrder, % "X" . (A_GuiWidth - 95)
GuiControl, 1:Move, btnHelpOrder, % "X" . (A_GuiWidth - 40)

GuiControl, 1:Move, strFileToSave, % "W" . (A_GuiWidth - 200)
GuiControl, 1:Move, btnHelpFileToSave, % "X" . (A_GuiWidth - 90)
GuiControl, 1:Move, btnSelectFileToSave, % "X" . (A_GuiWidth - 65)
GuiControl, 1:Move, btnSaveFile, % "X" . (A_GuiWidth - 65)
GuiControl, 1:Move, btnCheckFile, % "X" . (A_GuiWidth - 65)

GuiControl, 1:Move, strFileToExport, % "W" . (A_GuiWidth - 200)
GuiControl, 1:Move, btnHelpFileToExport, % "X" . (A_GuiWidth - 90)
GuiControl, 1:Move, btnSelectFileToExport, % "X" . (A_GuiWidth - 65)
GuiControl, 1:Move, btnExportFile, % "X" . (A_GuiWidth - 65)
GuiControl, 1:Move, btnCheckExportFile, % "X" . (A_GuiWidth - 65)
GuiControl, 1:Move, strMultiPurpose, % "W" . (A_GuiWidth - 305)
GuiControl, 1:Move, btnMultiPurpose, % "X" . (A_GuiWidth - 190)

GuiControl, 1:Move, lvData, % "W" . (A_GuiWidth - 20) . " H" . (A_GuiHeight - 215)

return



GuiEscape:
GuiClose:
Gui, 1:+OwnDialogs 
MsgBox, 292, %lAppName%, % L(lQuitApp, lAppName) ; task modal
IfMsgBox, Yes
	ExitApp
return




; --------------------- GUI2  --------------------------


ButtonSaveRecordAddRow:
Gui, 1:Default
intRowNumber := LV_GetCount()
intActualSize := Round(intActualSize + (intActualSize / intRowNumber))
SB_SetText(L(lSBRecordsSize, LV_GetCount(), (intActualSize) ? intActualSize : " <1"))
gosub, ButtonSaveRecord
return



ButtonSaveRecord:
Gui, 2:Submit
Gui, 1:Default
loop, % LV_GetCount("Column")
	LV_Modify(intRowNumber, "Col" . A_Index, strEdit%A_Index%)
Goto, 2GuiClose
return



ButtonCancelAddRow:
Gui, 1:Default
LV_Delete(LV_GetCount()) ; OK because added row is last in the list
gosub, ButtonCancel
return



ButtonCancel:
2GuiClose:
2GuiEscape:
Gui, 1:-Disabled
Gui, 2:Destroy
WinActivate, ahk_id %intGui1WinID%
return



2GuiSize: ; Expand or shrink the ListView in response to the user's resizing of the window.
if A_EventInfo = 1  ; The window has been minimized.  No action needed.
    return
GuiControl, 2:Move, btnSaveRecord, % "X" . (A_GuiWidth - 100)
GuiControl, 2:Move, btnCancel, % "X" . (A_GuiWidth - 50)
if intCol > 1 ; The window has been minimized.  No action needed.
    return
intWidthSize := A_GuiWidth - 20
Loop, %intNbFieldsOnScreen%
{
 	GuiControl, 2:Move, strEdit%A_Index%, % "W" . intWidthSize
}
return




; --------------------- OTHER PROCEDURES --------------------------


UpdateCurrentHeader:
Gui, 1:Submit, NoHide
strCurrentHeader := GetListViewHeader(strCurrentFieldDelimiter, strCurrentFieldEncapsulator)
strCurrentHeaderEscaped := StrEscape(strCurrentHeader)
GuiControl, 1:, strRenameEscaped, %strCurrentHeaderEscaped%
GuiControl, 1:, strSelectEscaped, %strCurrentHeaderEscaped%
GuiControl, 1:, strOrderEscaped, %strCurrentHeaderEscaped%
if (radFixed)
	Gosub, ClickRadFixed
else if (radExpress)
	Gosub, ClickRadExpress
return



ExportFixed:
if !DelimitersOK(3)
	return
; ObjCSV_ListView2Collection([strGuiID = "", strListViewID = "", strFieldOrder = "", strFieldDelimiter = ","
;	, strEncapsulator = """", intProgressType = 0, strProgressText = ""])
obj := ObjCSV_ListView2Collection("1", "lvData", , , , intProgressType, L(lTab0ReadingFromList))
; strFieldDelimiter3 et strFieldEncapsulator3 pour l'ecriture de l'entete seulement
strRealFieldDelimiter3 := StrMakeRealFieldDelimiter(strFieldDelimiter3)
; ObjCSV_ReturnDSVObjectArray(strCurrentDSVLine, strDelimiter = ",", strEncapsulator = """")
objFieldsArray := ObjCSV_ReturnDSVObjectArray(StrUnEscape(strMultiPurpose), strRealFieldDelimiter3, strFieldEncapsulator3)
strFieldsName := ""
strFieldsWidth := ""
loop, % objFieldsArray.MaxIndex() / 2
{
	strThisName := objFieldsArray[(A_Index * 2) - 1]
	; ObjCSV_Format4CSV(strF4C [, strFieldDelimiter = ",", strEncapsulator = """"])
	strFormat4Csv := ObjCSV_Format4CSV(strThisName, strRealFieldDelimiter3, strFieldEncapsulator3)
	strFieldsName := strFieldsName . strFormat4Csv . strRealFieldDelimiter3
	intThisWidth := objFieldsArray[(A_Index * 2)]
	if intThisWidth is integer
		strFieldsWidth := strFieldsWidth . intThisWidth . strRealFieldDelimiter3
	else
	{
		Oops(lExportFixedMustBeInteger, intThisWidth, A_Index, strThisName)
		return
	}
}
StringTrimRight, strFieldsName, strFieldsName, 1 ; remove extra delimiter
StringTrimRight, strFieldsWidth, strFieldsWidth, 1 ; remove extra delimiter
if (radSaveMultiline)
	strEolReplacement := ""
else
	strEolReplacement := strEndoflineReplacement
; ObjCSV_Collection2Fixed(objCollection, strFilePath, strWidth [, blnHeader = 0, strFieldOrder = "", intProgressType = 0, blnOverwrite = 0
;	, strFieldDelimiter = ",", strEncapsulator = """", strEolReplacement = "", strProgressText = ""])
ObjCSV_Collection2Fixed(obj, strFileToExport, strFieldsWidth, radSaveWithHeader, strFieldsName, intProgressType, blnOverwrite
	, strRealFieldDelimiter3, strFieldEncapsulator3, strEolReplacement, L(lExportSaving))
if FileExist(strFileToExport)
{
	GuiControl, 1:Show, btnCheckExportFile
	GuiControl, 1:+Default, btnCheckExportFile
	GuiControl, 1:Focus, btnCheckExportFile
}
obj := ; release object
return



ExportHTML:
; ObjCSV_ListView2Collection([strGuiID = "", strListViewID = "", strFieldOrder = "", strFieldDelimiter = ","
;	, strEncapsulator = """", intProgressType = 0, strProgressText = ""])
obj := ObjCSV_ListView2Collection("1", "lvData", , , , intProgressType, L(lTab0ReadingFromList))
; ObjCSV_Collection2HTML(objCollection, strFilePath, strTemplateFile [, strTemplateEncapsulator = ~
;	, intProgressType = 0, blnOverwrite = 0, strProgressText = ""])
ObjCSV_Collection2HTML(obj, strFileToExport, strMultiPurpose, strTemplateDelimiter
	, intProgressType, blnOverwrite, L(lExportSaving))
if (ErrorLevel)
	if (ErrorLevel > 3)
		Oops(lExportHTMLError)
	else
		Oops(lExportUnknownerror)
if FileExist(strFileToExport)
{
	GuiControl, 1:Show, btnCheckExportFile
	GuiControl, 1:+Default, btnCheckExportFile
	GuiControl, 1:Focus, btnCheckExportFile
}
obj := ; release object
return



ExportXML:
; ObjCSV_ListView2Collection([strGuiID = "", strListViewID = "", strFieldOrder = "", strFieldDelimiter = ","
;	, strEncapsulator = """", intProgressType = 0, strProgressText = ""])
obj := ObjCSV_ListView2Collection("1", "lvData", , , , intProgressType, L(lTab0ReadingFromList))
; ObjCSV_Collection2XML(objCollection, strFilePath [, intProgressType = 0, blnOverwrite = 0, strProgressText = ""])
ObjCSV_Collection2XML(obj, strFileToExport, intProgressType, blnOverwrite, L(lExportSaving))
if (ErrorLevel)
	Oops(lExportUnknownerror)
if FileExist(strFileToExport)
{
	GuiControl, 1:Show, btnCheckExportFile
	GuiControl, 1:+Default, btnCheckExportFile
	GuiControl, 1:Focus, btnCheckExportFile
}
obj := ; release object
return


ExportExpress:
; ObjCSV_ListView2Collection([strGuiID = "", strListViewID = "", strFieldOrder = "", strFieldDelimiter = ","
;	, strEncapsulator = """", intProgressType = 0, strProgressText = ""])
obj := ObjCSV_ListView2Collection("1", "lvData", , , , intProgressType, L(lTab0ReadingFromList))
SplitPath, strFileToExport, , strOutDir
strExpressTemplateTempFile := strOutDir . "\" . GUID() . ".TMP"
strExpressTemplate := strTemplateDelimiter . "ROWS" . strTemplateDelimiter
strExpressTemplate := strExpressTemplate  . strMultiPurpose
strExpressTemplate := strExpressTemplate  . strTemplateDelimiter . "/ROWS" . strTemplateDelimiter
strExpressTemplate := StrUnEscape(strExpressTemplate)
FileAppend, %strExpressTemplate%, %strExpressTemplateTempFile%
if FileExist(strExpressTemplateTempFile)
{
	; ObjCSV_Collection2HTML(objCollection, strFilePath, strTemplateFile [, strTemplateEncapsulator = ~
	;	, intProgressType = 0, blnOverwrite = 0, strProgressText = ""])
	ObjCSV_Collection2HTML(obj, strFileToExport, strExpressTemplateTempFile, strTemplateDelimiter
		, intProgressType, blnOverwrite, L(lExportSaving))
	FileDelete, %strExpressTemplateTempFile%
	if FileExist(strFileToExport)
	{
		GuiControl, 1:Show, btnCheckExportFile
		GuiControl, 1:+Default, btnCheckExportFile
		GuiControl, 1:Focus, btnCheckExportFile
	}
}
else
	Oops(lExportErrorTempFile)
obj := ; release object
return


Check4Update:
Gui, 1:+OwnDialogs 
IniRead, strLatestSkipped, %strIniFile%, global, strLatestSkipped, 0.0
strLatestVersion := Url2Var("https://raw.github.com/JnLlnd/CSVBuddy/master/latest-version.txt")

if RegExMatch(strCurrentVersion, "(alpha|beta)")
	or (FirstVsSecondIs(strLatestSkipped, strLatestVersion) >= 0 and (!blnButtonCheck4Update))
	return

if FirstVsSecondIs(strLatestVersion, lAppVersion) = 1
{
	Gui, +OwnDialogs
	SetTimer, ChangeButtonNames4Update, 50

	MsgBox, 3, % l(lTab5UpdateTitle, lAppName), % l(lTab5UpdatePrompt, lAppName, lAppVersion, strLatestVersion), 30
	IfMsgBox, Yes
		Run, http://code.jeanlalonde.ca/csvbuddy/
	IfMsgBox, No
		IniWrite, %strLatestVersion%, %strIniFile%, global, strLatestSkipped
	IfMsgBox, Cancel ; Remind me
		IniWrite, 0.0, %strIniFile%, Global, LatestVersionSkipped
	IfMsgBox, TIMEOUT ; Remind me
		IniWrite, 0.0, %strIniFile%, Global, LatestVersionSkipped
}
else if (blnButtonCheck4Update)
{
	MsgBox, 4, % l(lTab5UpdateTitle, lAppName), % l(lTab5UpdateYouHaveLatest, lAppVersion, lAppName)
	IfMsgBox, Yes
		Run, http://code.jeanlalonde.ca/csvbuddy/
}
return


ChangeButtonNames4Update: 
IfWinNotExist, % l(lTab5UpdateTitle, lAppName)
    return  ; Keep waiting.
SetTimer, ChangeButtonNames4Update, Off 
WinActivate
ControlSetText, Button3, %lTab5ButtonRemind%
return


Check4CommandLineParameter:
if 0 > 0 
; The left side of a non-expression if-statement is always the name of a variable = %0% command line parameter
{
	strParam = %1%
	if FileExist(strParam)
	{
		GuiControl, 1:, strFileToLoad, %strParam%
		GuiControl, 1:, blnMultiline1, 1 ; always process as multi-line for safety
		FileReadLine, strCurrentHeader, %strParam%, 1
		GuiControl, 1:, strFileHeaderEscaped, % StrEscape(strCurrentHeader)
		gosub, DetectDelimiters
		gosub, ButtonLoadFile
	}
	else
		Oops(L(lTab1CommandLineFileNotFound, strParam))
}
return



; --------------------- FUNCTIONS --------------------------


CheckOneRow()
{
	if (LV_GetCount("Selected") = 1)
	{
		Gui, 1:+OwnDialogs 
		SetTimer, ChangeButtonNamesOneRwo, 50
		MsgBox, 35, % L(lLvEventsOnerecordselectedTitle, lAppName), % L(lLvEventsOnerecordselectedMessage)
		IfMsgBox, No
		{
			GuiControl, Focus, lvData
			LV_Modify(0, "-Select")
		}
		IfMsgBox, Cancel
			return False
	}
	return True
}



ChangeButtonNamesOneRwo: 
IfWinNotExist, % L(lLvEventsOnerecordselectedTitle, lAppName)
    return  ; Keep waiting.
SetTimer, ChangeButtonNamesOneRwo, Off 
WinActivate 
ControlSetText, Button1, %lTab3ThisRecord%
ControlSetText, Button2, %lTab3AllRecord%
return



GetListViewHeader(strRealFieldDelimiter, strFieldEncapsulator)
{
	strHeader := ""
	Loop, % LV_GetCount("Column")
	{
		LV_GetText(strColumnHeader, 0, A_Index)
		; ObjCSV_Format4CSV(strF4C [, strFieldDelimiter = ",", strEncapsulator = """"])
		strHeader := strHeader . ObjCSV_Format4CSV(strColumnHeader, strRealFieldDelimiter, strFieldEncapsulator) . strRealFieldDelimiter
	}
	StringTrimRight, strHeader, strHeader, 1 ; remove extra delimiter
	return strHeader
}



StrEscape(strEscaped)
{
	strReplacement := Chr(131) . Chr(161) . Chr(134) ; temporary string replacement for escape char
	StringReplace, strEscaped, strEscaped, ``, %strReplacement%, All
	StringReplace, strEscaped, strEscaped, `t, ``t, All ; Tab (HT)
	StringReplace, strEscaped, strEscaped, `n, ``n, All ; Linefeed (LF)
	StringReplace, strEscaped, strEscaped, `r, ``r, All ; Carriage return (CR)
	StringReplace, strEscaped, strEscaped, `f, ``f, All ; Form feed (FF)
	StringReplace, strEscaped, strEscaped, %strReplacement%, ````, All
	return strEscaped
}



StrUnEscape(strUnEscaped)
{
	strReplacement := Chr(131) . Chr(161) . Chr(134) ; temporary string replacement for escape char
	StringReplace, strUnEscaped, strUnEscaped, ````, %strReplacement%, All
	StringReplace, strUnEscaped, strUnEscaped, ``t, `t, All ; Tab (HT)
	StringReplace, strUnEscaped, strUnEscaped, ``n, `n, All ; Linefeed (LF)
	StringReplace, strUnEscaped, strUnEscaped, ``r, `r, All ; Carriage return (CR)
	StringReplace, strUnEscaped, strUnEscaped, ``f, `f, All ; Form feed (FF)
	StringReplace, strUnEscaped, strUnEscaped, %strReplacement%, ``, All
	return strUnEscaped
}



StrMakeRealFieldDelimiter(strConverted)
{
	StringReplace, strConverted, strConverted, t, `t, All ; Tab (HT)
	StringReplace, strConverted, strConverted, n, `n, All ; Linefeed (LF)
	StringReplace, strConverted, strConverted, r, `r, All ; Carriage return (CR)
	StringReplace, strConverted, strConverted, f, `f, All ; Form feed (FF)
	return strConverted
}



StrMakeEncodedFieldDelimiter(strConverted)
{
	StringReplace, strConverted, strConverted, `t, t, All ; Tab (HT)
	StringReplace, strConverted, strConverted, `n, n, All ; Linefeed (LF)
	StringReplace, strConverted, strConverted, `r, r, All ; Carriage return (CR)
	StringReplace, strConverted, strConverted, `f, f, All ; Form feed (FF)
	return strConverted
}



ShrinkEditControl(strEditHandle, intMaxRows, strGuiName)
{
	EM_GETLINECOUNT = 0xBA
	SendMessage, %EM_GETLINECOUNT%,,,, AHk_id %strEditHandle%
	intNbRows := ErrorLevel
	if (intNbRows > intMaxRows)
	{
		GuiControlGet, intPosEdit, %strGuiName%:Pos, %strEditHandle%
		intEditMargin := 8 ; top & bottom margin of the Edit control (regardless of the nb of rows)
		intOriginalHeight := intPosEditH
		intHeightOneRow := Round((intOriginalHeight - intEditMargin) / intNbRows)
		intNewHeight := (intHeightOneRow * intMaxRows) + intEditMargin
		; MsgBox, % "intNbRows: " . intNbRows . "`nintOriginalHeight: " . intOriginalHeight 
		;	. "`nintHeightOneRow: " . intHeightOneRow . "`nintNewHeight: " . intNewHeight
		GuiControl, %strGuiName%:Move, %strEditHandle%, h%intNewHeight%
	}
}



CheckIfFileExistOverwrite(strFileName)
{
	if !StrLen(strFileName)
		return -1
	if !FileExist(strFileName)
		return True
	else
	{
		Gui, 1:+OwnDialogs 
		MsgBox, 35, % L(lFuncIfFileExistTitle, lAppName), % L(lFuncIfFileExistMessage, strFileName)
		IfMsgBox, Yes
			return True
		IfMsgBox, No
			return False
		IfMsgBox, Cancel
			return -1
	}
}



NewFileName(strExistingFile, strNote := "", strExtension := "")
{
	SplitPath, strExistingFile, , strOutDir, strOutExtension, strOutNameNoExt
	if !StrLen(strExtension)
		strExtension := strOutExtension
	strNewName := strOutDir . "\" . strOutNameNoExt . strNote . "." . strExtension
	if FileExist(strNewName)
		loop
		{
			strNewName := strOutDir . "\" . strOutNameNoExt . strNote . " (" . A_Index  . ")." . strExtension
			if !FileExist(strNewName)
				break
		}
	return strNewName
}



GUID()         ; 32 hex digits = 128-bit Globally Unique ID
; Source: Laszlo in http://www.autohotkey.com/board/topic/5362-more-secure-random-numbers/
{
   format = %A_FormatInteger%       ; save original integer format
   SetFormat Integer, Hex           ; for converting bytes to hex
   VarSetCapacity(A,16)
   DllCall("rpcrt4\UuidCreate","Str",A)
   Address := &A
   Loop 16
   {
      x := 256 + *Address           ; get byte in hex, set 17th bit
      StringTrimLeft x, x, 3        ; remove 0x1
      h = %x%%h%                    ; in memory: LS byte first
      Address++
   }
   SetFormat Integer, %format%      ; restore original format
   Return h
}



DelimitersOK(intTab)
{
	if (strFieldDelimiter%intTab% = strFieldEncapsulator%intTab%) or (strFieldDelimiter%intTab% = "") or (strFieldEncapsulator%intTab% = "")
	{
		Oops(lFuncDelimitersOK, intTab)
		GuiControl, 1:Choose, tabCSVBuddy, %intTab%
		if (strFieldDelimiter%intTab% = strFieldEncapsulator%intTab%) or (strFieldDelimiter%intTab% = "")
			GuiControl, Focus, strFieldDelimiter%intTab%
		else
			GuiControl, Focus, strFieldEncapsulator%intTab%
		return false
	}
	else
		return true
}



NewDelimiterOrEncapsulatorOK(strChecked)
{
	global strCurrentHeader
	global strCurrentFieldDelimiter
	global strCurrentFieldEncapsulator
	; ObjCSV_ReturnDSVObjectArray(strCurrentDSVLine, strDelimiter = ",", strEncapsulator = """")
	objCurrentHeader := ObjCSV_ReturnDSVObjectArray(strCurrentHeader, strCurrentFieldDelimiter, strCurrentFieldEncapsulator)
	Loop, % objCurrentHeader.MaxIndex()
		If InStr(objCurrentHeader[A_Index], strChecked)
			return false
	return true
}



PositionInArray(strChecked, objArray)
{
	Loop, % objArray.MaxIndex()
		If (objArray[A_Index] = strChecked)
			return A_Index
	return 0
}



Help(strMessage, objVariables*)
{
	Gui, 1:+OwnDialogs 
	StringLeft, strTitle, strMessage, % InStr(strMessage, "$") - 1
	StringReplace, strMessage, strMessage, %strTitle%$
	MsgBox, 0, % L(lFuncHelpTitle, lAppName, lAppVersionLong, strTitle), % L(strMessage, objVariables*)
}



Oops(strMessage, objVariables*)
{
	Gui, 1:+OwnDialogs
	MsgBox, 48, % L(lFuncOopsTitle, lAppName, lAppVersionLong), % L(strMessage, objVariables*)
}



L(strMessage, objVariables*)
{
	Loop
	{
		if InStr(strMessage, "~" . A_Index . "~")
			StringReplace, strMessage, strMessage, ~%A_Index%~, % objVariables[A_Index], All
 		else
			break
	}
	return strMessage
}



Url2Var(strUrl)
{
	ComObjError(False)
	objWebRequest := ComObjCreate("WinHttp.WinHttpRequest.5.1")
	objWebRequest.Open("GET", strUrl)
	objWebRequest.Send()

	Return objWebRequest.ResponseText()
}



FirstVsSecondIs(strFirstVersion, strSecondVersion)
{
	StringSplit, arrFirstVersion, strFirstVersion, `.
	StringSplit, arrSecondVersion, strSecondVersion, `.
	if (arrFirstVersion0 > arrSecondVersion0)
		intLoop := arrFirstVersion0
	else
		intLoop := arrSecondVersion0

	Loop %intLoop%
		if (arrFirstVersion%A_index% > arrSecondVersion%A_index%)
			return 1 ; greater
		else if (arrFirstVersion%A_index% < arrSecondVersion%A_index%)
			return -1 ; smaller
		
	return 0 ; equal
}


