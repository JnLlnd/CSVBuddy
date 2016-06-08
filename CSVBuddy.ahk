;===============================================
/*
CSV Buddy
Written using AutoHotkey_L v1.1.09.03+ (http://www.ahkscript.org/)
By JnLlnd on AHK forum
This script uses the library ObjCSV v0.5.1 (https://github.com/JnLlnd/ObjCSV)

Copyright 2013-2016 Jean Lalonde
--------------------------------
Licensed under the Apache License, Version 2.0 (the "License");
you may not use this file except in compliance with the License.
You may obtain a copy of the License at

    http://www.apache.org/licenses/LICENSE-2.0

Unless required by applicable law or agreed to in writing, software
distributed under the License is distributed on an "AS IS" BASIS,
WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
See the License for the specific language governing permissions and
limitations under the License.

Version history
---------------

2016-06-08 v1.3.2
- Fix bug introduced in v1.2.9.1 preventing from saving manual record edits in some circumstances
- Automatic file encoding detection is now restricted to UTF-8 or UTF-16 encoded files (no BOM)
- Read the new value DefaultFileEncoding= (under [global]) in CSVBuddy.ini to set the default file encoding (possible values "ANSI", "UTF-8", "UTF-16", "UTF-8-RAW", "UTF-16-RAW" or "CPnnn")

2016-05-31 v1.3.1
- Change licence to Apache 2.0 (see Copyright above)

2016-05-23 v1.3
- Select file encoding when loading, saving or exporting files
- Automatic detection of loaded file encoding (this is not possible for files without byte order mark (BOM))
- Encoding supported: ANSI (default), UTF-8 (Unicode 8-bit), UTF-16 (Unicode 16-bit), UTF-8-RAW (no BOM), UTF-16-RAW (no BOM) or custom codepage CPnnnn
- Set custom codepage load and save values in CSVBuddy.ini file
 
Other changes:
- when editing a record, zoom button to edit long strings in a larger window
- deselect all rows after global search is cancelled
- add status bar right section to display hint about list menu
- add help message when right-clicking on column headers
- when there is not enough space on screen to edit all fields in a file, support editing the visible fields without loosing the content of the missing fields
- remove global search and replace (needing to be reworked)

2016-05-18 v1.2.9.1 BETA
- when editing a record, zoom button to edit long strings in a large window
- when there is not enough space on screen to edit all fields in a file, support editing the visible fields without loosing the content of the missing fields

2016-05-15 v1.2.9 BETA
- Select file encoding when loading, saving or exporting files
- Automatic detection of loaded file encoding (this is not possible for files without byte order mark (BOM))
- Encoding supported: ANSI (default), UTF-8 (Unicode 8-bit), UTF-16 (Unicode 16-bit), UTF-8-RAW (no BOM), UTF-16-RAW (no BOM) or custom codepage CPnnnn
- Set custom codepage values in CSVBuddy.ini file
- Other changes:
  - deselect all rows after global search is cancelled
  - add status bar right section to display hint about list menu
  - add help message when right-clicking on column headers
  - remove global search and replace (needing to be reworked)

2014-08-31 v1.2.3 (bug fix)
- fix bug when saving or exporting file with a column sort indicator

2014-03-17 v1.2.2 (bug fix)
- after a column sort, fix names errors in column headers
- by safety, remove sorting column indicator before any action in edit column tab 

2014-03-09 v1.2.1 (bug fix)
- ini variables missing when ini file already existed making grid black text on black background
- bug with multiple instances

2014-03-07 v1.2
- search and replace by column, replacement case sensitive or not
- confirm each replacement or replace all
- during search or replace, select and highlight the current row when displaying the record found
- option in ini file to display or not a grid around cells in list zone
- options in ini file to choose background and text colors in list zone
- up or down arrow to indicate which field is the current sort key
- allow multiple instances of the app to run simultaneously
- import CSV files created by XL that include equal sign before the opening field encasulator

2013-12-30 v1.1
- filter by column: click on a header to retain only rows with the keyword appearing in this column
- global filtering: right-click in the list zone to retain only rows with the keyword appearing in any column
- search by column: find the next row having the keyword in this column and open it in row edit window
- global search: find the next row having the keyword in any column and open it in row edit window
- in edit row window search result, highlight the field containing the searched keyword
- added stop and next buttons to edit row window when search in progress
- added reload original file to the column menu and the list context menu
- display the current edited record number in edit row title bar
- add blnSkipConfirmQuit option in ini file to skip the quit confirm prompt, default to false
- use ObjCSV library v0.4 for better file system error handling

2013-11-30 v1.0
- First official release
- Add records to existing data (right-click in the list zone)
- Create a new file from scratch (right-click in an empty list zone)
- Load the file mentioned as first parameter in the command line
- Add validation, confirm before exit and fix various small bugs

2013-11-03 v0.9
- Display "<1" (instead of "0") in status bar when file size is smaller than 0.5 K
- Removed CSV Buddy icon from the Tray
- Add three test delimited files to the package (see README.txt in the zip file)
- Fix default value of blnSkipHelpReadyToEdit in ini file to 0

2013-10-20 v0.8.1
- If an .ini file is not found in the program's folder, it is created with default values

2013-10-18 v0.8.0
- First release of BETA version
- History of ALPHA phase on BitBucket (private repository)


*/ 
;===============================================

#KeyHistory 0
#NoEnv
#NoTrayIcon 
#SingleInstance off
;@Ahk2Exe-IgnoreBegin
; Piece of code for developement phase only - won't be compiled
; COMPILER EXCLUSION NOT WORKING FOR THIS LINE: #SingleInstance force
; / Piece of code for developement phase only - won't be compiled
;@Ahk2Exe-IgnoreEnd
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
;@Ahk2Exe-SetVersion 1.3.2
;@Ahk2Exe-SetCopyright Jean Lalonde
;@Ahk2Exe-SetOrigFilename CSVBuddy.exe


; --------------------- GLOBAL AND DEFAULT VALUES --------------------------

global intCurrentSortColumn

strListBackgroundColor := "D0D0D0"
strListTextColor := "000000"
blnListGrid := 1
intDefaultWidth := 16
strTemplateDelimiter := "~"
strTextEditorExe := "notepad.exe"

strIniFile := A_ScriptDir . "\CSVBuddy.ini"
IfNotExist, %strIniFile%
	FileAppend,
		(LTrim Join`r`n
			[global]
			ListBackgroundColor=%strListBackgroundColor%
			ListTextColor=%strListTextColor%
			ListGrid=%blnListGrid%
			DefaultWidth=%intDefaultWidth%
			TemplateDelimiter=%strTemplateDelimiter%
			TextEditorExe=%strTextEditorExe%
			SkipHelpReadyToEdit=0
			SkipConfirmQuit=0
			Startups=1
			CodePageLoad=1252
			CodePageSave=1252
		)
		, %strIniFile%

IniRead, strListBackgroundColor, %strIniFile%, global, ListBackgroundColor, %strListBackgroundColor%
IniRead, strListTextColor, %strIniFile%, global, ListTextColor, %strListTextColor%
IniRead, blnListGrid, %strIniFile%, global, ListGrid, %blnListGrid%
IniRead, intDefaultWidth, %strIniFile%, global, DefaultWidth, %intDefaultWidth% ; used when export to fixed-width format
IniRead, strTemplateDelimiter, %strIniFile%, global, TemplateDelimiter, %strTemplateDelimiter% ; Default ~ (tilde), used when export to HTML and Express formats
IniRead, strTextEditorExe, %strIniFile%, global, TextEditorExe, %strTextEditorExe% ; Default notepad.exe
IniRead, blnSkipHelpReadyToEdit, %strIniFile%, global, SkipHelpReadyToEdit ; Default 0
IniRead, blnSkipConfirmQuit, %strIniFile%, global, SkipConfirmQuit ; Default 0
IniRead, strLatestSkipped, %strIniFile%, global, LatestVersionSkipped, 0.0
IniRead, intStartups, %strIniFile%, Global, Startups, 1
IniRead, blnDonator, %strIniFile%, Global, Donator, 0 ; Please, be fair. Don't cheat with this.
IniRead, strCodePageLoad, %strIniFile%, Global, CodePageLoad, 1252 ; default ANSI Latin 1, Western European (Windows)
IniRead, strCodePageSave, %strIniFile%, Global, CodePageSave, 1252 ; default ANSI Latin 1, Western European (Windows)

IniRead, strDefaultFileEncoding, %strIniFile%, Global, DefaultFileEncoding, %A_Space% ; default file encoding (ANSI, UTF-8, UTF-16, UTF-8-RAW, UTF-16-RAW or CPnnn)
if !StrLen(strDefaultFileEncoding)
	strDefaultFileEncoding := "Detect"
strDefaultFileEncoding := StrReplace(L(lFileEncodings, strCodePageLoad), strDefaultFileEncoding . "|", strDefaultFileEncoding . "||")

intProgressType := -2 ; Status Bar, part 2


; --------------------- GUI1 --------------------------

Gui, 1:New, +Resize, % L(lAppName)

Gui, 1:Font, s12 w700, Verdana
Gui, 1:Add, Text, x10, % L(lAppName)

Gui, 1:Font, s10 w700, Verdana
Gui, 1:Add, Tab2, w950 r4 vtabCSVBuddy gChangedTabCSVBuddy, % L(lTab0List)
Gui, 1:Font

Gui, 1:Tab, 1
Gui, 1:Add, Text,			y+10	x10		vlblCSVFileToLoad w85 right, % L(lTab1CSVFileToLoad)
Gui, 1:Add, Edit,			yp		x100	vstrFileToLoad gChangedFileToLoad disabled
Gui, 1:Add, Button,			yp		x+5		vbtnHelpFileToLoad gButtonHelpFileToLoad, % L(lTab0QuestionMark)
Gui, 1:Add, Button,			yp		x+5		vbtnSelectFileToLoad gButtonSelectFileToLoad w45 default, % L(lTab1Select)
Gui, 1:Add, Text,			y+10	x10 	vlblHeader w85 right, % L(lTab1CSVFileHeader)
Gui, 1:Add, Edit,			yp		x100	vstrFileHeaderEscaped disabled
Gui, 1:Add, Button,			yp		x+5		vbtnHelpHeader gButtonHelpHeader, % L(lTab0QuestionMark)
Gui, 1:Add, Button,			yp		x+5		vbtnPreviewFile gButtonPreviewFile w45 hidden, % L(lTab1PreviewFile)
Gui, 1:Add, Radio,			y+10	x20		vradGetHeader gClickRadGetHeader checked, % L(lTab1Getheaderfromfile)
Gui, 1:Add, Radio,			yp		x+5		vradSetHeader gClickRadSetHeader, % L(lTab1Setheader)
Gui, 1:Add, Button,			yp		x+0		vbtnHelpSetHeader gButtonHelpSetHeader, % L(lTab0QuestionMark)
Gui, 1:Add, Text,			xp		x+20	vlblFieldDelimiter1, % L(lTab1Fielddelimiter)
Gui, 1:Add, Edit,			yp		x+5		vstrFieldDelimiter1 w20 limit1 center, `, ; gChangedFieldDelimiter1 unused
Gui, 1:Add, Button,			yp		x+5		vbtnHelpFieldDelimiter1 gButtonHelpFieldDelimiter1, % L(lTab0QuestionMark)
Gui, 1:Add, Text,			yp		x+20	vlblFieldEncapsulator1, % L(lTab1Fieldencapsulator)
Gui, 1:Add, Edit,			yp		x+5		vstrFieldEncapsulator1 w20 limit1 center, `" ; gChangedFieldEncapsulator1 unused
Gui, 1:Add, Button,			yp		x+5		vbtnHelpEncapsulator1 gButtonHelpEncapsulator1, % L(lTab0QuestionMark)
Gui, 1:Add, Checkbox,		yp		x+20	vblnMultiline1 gChangedMultiline1, % L(lTab1Multilinefields)
Gui, 1:Add, Button,			yp		x+0		vbtnHelpMultiline1 gButtonHelpMultiline1, % L(lTab0QuestionMark)
Gui, 1:Add, Text,			yp		x+5		vlblEndoflineReplacement1 hidden, % L(lTab1EOLreplacement)
Gui, 1:Add, Edit,			yp		x+5		vstrEndoflineReplacement1 w30 center hidden
Gui, 1:Add, DropDownList,	yp		x+20	vstrFileEncoding1 w85, %strDefaultFileEncoding%
Gui, 1:Add, Button,			yp		x+7		vbtnHelpFileEncoding1 gButtonHelpFileEncoding1, % L(lTab0QuestionMark)
Gui, 1:Add, Button,			yp		x+5		vbtnLoadFile gButtonLoadFile w45 hidden, % L(lTab1Load)

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
Gui, 1:Add, Text,			y+10	x10		vlblCSVFileToSave w85 right, % L(lTab3CSVfiletosave)
Gui, 1:Add, Edit,			yp		x100	vstrFileToSave gChangedFileToSave
Gui, 1:Add, Button,			yp		x+5		vbtnHelpFileToSave gButtonHelpFileToSave, % L(lTab0QuestionMark)
Gui, 1:Add, Button,			yp		x+5		vbtnSelectFileToSave gButtonSelectFileToSave w45 default, % L(lTab3Select)
Gui, 1:Add, Text,			y+10	x100	vlblFieldDelimiter3, % L(lTab3Fielddelimiter)
Gui, 1:Add, Edit,			yp		x200	vstrFieldDelimiter3 gChangedFieldDelimiter3 w20 limit1 center, `, 
Gui, 1:Add, Button,			yp		x+5		vbtnHelpFieldDelimiter3 gButtonHelpFieldDelimiter3, % L(lTab0QuestionMark)
Gui, 1:Add, Text,			y+10	x100	vlblFieldEncapsulator3, % L(lTab3Fieldencapsulator)
Gui, 1:Add, Edit,			yp		x200	vstrFieldEncapsulator3 gChangedFieldEncapsulator3 w20 limit1 center, `"
Gui, 1:Add, Button,			yp		x+5		vbtnHelpEncapsulator3 gButtonHelpEncapsulator3, % L(lTab0QuestionMark)
Gui, 1:Add, Radio,			y100	x300	vradSaveWithHeader checked, % L(lTab3Savewithheader)
Gui, 1:Add, Radio,			y+10	x300	vradSaveNoHeader, % L(lTab3Savewithoutheader)
Gui, 1:Add, Button,			y100	x450	vbtnHelpSaveHeader gButtonHelpSaveHeader, % L(lTab0QuestionMark)
Gui, 1:Add, Radio,			y100	x500	vradSaveMultiline gClickRadSaveMultiline checked, % L(lTab3Savemultiline)
Gui, 1:Add, Radio,			y+10	x500	vradSaveSingleline gClickRadSaveSingleline, % L(lTab3Savesingleline)
Gui, 1:Add, Button,			y100	x620	vbtnHelpMultiline gButtonHelpSaveMultiline, % L(lTab0QuestionMark)
Gui, 1:Add, DropDownList,	yp		x+20	vstrFileEncoding3 w85, % L(lFileEncodings, strCodePageSave)
Gui, 1:Add, Button,			yp		x+7		vbtnHelpFileEncoding3 gButtonHelpFileEncoding3, % L(lTab0QuestionMark)
Gui, 1:Add, Text,			y+25	x500	vlblEndoflineReplacement3 hidden, % L(lTab3Endoflinereplacement)
Gui, 1:Add, Edit,			yp		x620	vstrEndoflineReplacement3 hidden w50 center, % chr(182)
Gui, 1:Add, Button,			y105	x+5		vbtnSaveFile gButtonSaveFile w45 hidden, % L(lTab3Save)
Gui, 1:Add, Button,			y137	x+5		vbtnCheckFile hidden w45 gButtonCheckFile, % L(lTab3Check)

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
Gui, 1:Add, Button,		yp		x+20		vbtnDonate gButtonDonate, % L(lTab5Donate)

Gui, 1:Tab

Gui, 1:Add, ListView, 	x10 r24 w955 vlvData -ReadOnly NoSort gListViewEvents AltSubmit -LV0x10

Gui, Add, StatusBar
SB_SetParts(200, 300)
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


ButtonHelpFileEncoding1:
ButtonHelpFileEncoding3:
Help(lTab1HelpFileEncoding . (A_ThisLabel = "ButtonHelpFileEncoding3" ? "`n`n" . lTab1HelpFileEncodingExport : "")
	, (A_ThisLabel = "ButtonHelpFileEncoding1" ? lTab1HelpFileEncodingLoad : lTab3HelpFileEncodingSave))
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

intCurrentSortColumn := 0 ; indicate that no field header has the ^ or v sort indicator
strCurrentHeader := StrUnEscape(strFileHeaderEscaped)
strCurrentFieldDelimiter := StrMakeRealFieldDelimiter(strFieldDelimiter1)
strCurrentVisibleFieldDelimiter := strFieldDelimiter1
strCurrentFieldEncapsulator := strFieldEncapsulator1
strCurrentFileEncodingLoad := (strFileEncoding1 = "Detect" ? "" : strFileEncoding1)

FileGetSize, intFileSize, %strFileToLoad%, K
intActualSize := intActualSize + intFileSize

; ObjCSV_CSV2Collection(strFilePath, ByRef strFieldNames [, blnHeader = 1, blnMultiline = 1, intProgressType = 0
;	, strFieldDelimiter = ",", strEncapsulator = """", strEolReplacement = "", strProgressText := ""])
obj := ObjCSV_CSV2Collection(strFileToLoad, strCurrentHeader, radGetHeader, blnMultiline1, intProgressType
	, strCurrentFieldDelimiter, strCurrentFieldEncapsulator, strEndoflineReplacement1, L(lTab1ReadingCSVdata)
	, strCurrentFileEncodingLoad)
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
	SB_SetText("", 3)
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
	SB_SetText("", 3)
	return
}
GuiControl, % (blnListGrid ? "+Grid" : "") . " Background" . strListBackgroundColor . " c" . strListTextColor, lvData

SB_SetText(L(lSBRecordsSize, LV_GetCount(), (intActualSize) ? intActualSize : " <1"), 1)
SB_SetText(lSBHelp, 3)
Gosub, UpdateCurrentHeader
if (!blnSkipHelpReadyToEdit)
	Help(lTab1HelpReadyToEdit)
GuiControl, 1:, strFieldDelimiter3, %strCurrentVisibleFieldDelimiter%
GuiControl, 1:, strFieldEncapsulator3, %strCurrentFieldEncapsulator%
GuiControl, 1:ChooseString, strFileEncoding1, %strCurrentFileEncodingLoad%
GuiControl, 1:ChooseString, strFileEncoding3, %strCurrentFileEncodingLoad%
blnFilterActive := false
obj := ; release object
return


DeleteListviewData:
LV_Delete() ; delete all rows - better performance on large files when we delete rows before columns
loop, % LV_GetCount("Column")
	LV_DeleteCol(1) ; delete all columns
SB_SetText(L(lSBEmpty), 1)
SB_SetText("", 3)
intActualSize := 0
return



; --------------------- TAB 2 --------------------------


ButtonSetRename:
Gui, 1:+OwnDialogs 
Gui, 1:Submit, NoHide

GoSub, RemoveSorting

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

GoSub, RemoveSorting

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

GoSub, RemoveSorting

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
blnOverwrite := CheckIfFileExistOverwrite(strFileToSave, True)
if (blnOverwrite < 0)
	return
if !CheckOneRow()
	return

GoSub, RemoveSorting

; ObjCSV_ListView2Collection([strGuiID = "", strListViewID = "", strFieldOrder = "", strFieldDelimiter = ","
;	, strEncapsulator = """", intProgressType = 0, strProgressText = ""])
obj := ObjCSV_ListView2Collection("1", "lvData", , , , intProgressType, L(lTab0ReadingFromList))
if (radSaveMultiline)
	strEolReplacement := ""
else
	strEolReplacement := strEndoflineReplacement3
strRealFieldDelimiter3 := StrMakeRealFieldDelimiter(strFieldDelimiter3)
strCurrentFileEncodingSave := (strFileEncoding3 = "Detect" ? "" : strFileEncoding3)
; ObjCSV_Collection2CSV(objCollection, strFilePath [, blnHeader = 0
;	, strFieldOrder = "", intProgressType = 0, blnOverwrite = 0
;	, strFieldDelimiter = ",", strEncapsulator = """", strEolReplacement = "", strProgressText = ""])
ObjCSV_Collection2CSV(obj, strFileToSave, radSaveWithHeader
	, GetListViewHeader(strRealFieldDelimiter3, strFieldEncapsulator3), intProgressType, blnOverwrite
	, strRealFieldDelimiter3, strFieldEncapsulator3, strEolReplacement, L(lTab3SavingCSV), strCurrentFileEncodingSave)
if (ErrorLevel)
	if (ErrorLevel = 1)
		Oops(lExportSystemError, A_LastError)
	else
		Oops(lExportUnknownError)
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
		, % L(lTab4MultiFixedInputPrompt), , , 150, , , , , %intDefaultWidth%
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
blnOverwrite := CheckIfFileExistOverwrite(strFileToExport, radFixed)
if (blnOverwrite < 0)
	return

strCurrentFileEncodingSave := (strFileEncoding3 = "Detect" ? "" : strFileEncoding3)

GoSub, RemoveSorting

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



ButtonDonate:
Run, https://www.paypal.com/cgi-bin/webscr?cmd=_s-xclick&hosted_button_id=8UWKXDR5ZQNJW
return




; --------------------- LISTVIEW EVENTS --------------------------


ListViewEvents:
if (A_GuiEvent = "ColClick")
{
	intColNumber := A_EventInfo
	Menu, ColumnMenu, Add
	Menu, ColumnMenu, Delete
	Menu, ColumnMenu, Add, % L(lLvEventsSortalphabetical), MenuColumnText
	Menu, ColumnMenu, Add, % L(lLvEventsSortnumericInteger), MenuColumnInteger
	Menu, ColumnMenu, Add, % L(lLvEventsSortnumericFloat), MenuColumnFloat
	Menu, ColumnMenu, Add
	Menu, ColumnMenu, Add, % L(lLvEventsSortdescalphabetical), MenuColumnDescText
	Menu, ColumnMenu, Add, % L(lLvEventsSortdescnumericInteger), MenuColumnDescInteger
	Menu, ColumnMenu, Add, % L(lLvEventsSortdescnumericFloat), MenuColumnDescFloat
	Menu, ColumnMenu, Add
	Menu, ColumnMenu, Add, % L(lLvEventsFilterColumn), MenuFilter
	Menu, ColumnMenu, Add, % ((lLvEventsFilterReload)), MenuFilterReload
	Menu, ColumnMenu, Add
	Menu, ColumnMenu, Add, % L(lLvEventsSearchColumn), MenuSearch
	Menu, ColumnMenu, Add, % L(lLvEventsReplaceColumn), MenuReplace
	Menu, ColumnMenu, % blnFilterActive ? "Enable" : "Disable", %lLvEventsFilterReload%
	Menu, ColumnMenu, Show
}
if (A_GuiEvent = "DoubleClick")
{
	intRowNumber := A_EventInfo
	Gosub, MenuEditRow
}
SB_SetText(L(lLvEventsrecordsselected, LV_GetCount("Selected")), 2)
return



MenuColumnText:
MenuColumnInteger:
MenuColumnFloat:
MenuColumnDescText:
MenuColumnDescInteger:
MenuColumnDescFloat:
StringReplace, strOption, A_ThisLabel, Menu,
StringReplace, strOption, strOption, Column, % "Sort "
StringReplace, strOption, strOption, Sort Desc, % "SortDesc "

LV_GetText(strColHeader, 0, intCurrentSortColumn)
LV_ModifyCol(intCurrentSortColumn, , SubStr(strColHeader, 3))

intCurrentSortColumn := intColNumber
LV_ModifyCol(intCurrentSortColumn, strOption)
if InStr(strOption, "Text")
	LV_ModifyCol(intColNumber, "Left")
LV_GetText(strColHeader, 0, intCurrentSortColumn)
LV_ModifyCol(intCurrentSortColumn, , (InStr(strOption, "Desc") ? "v" : "^") . " " . strColHeader)
Menu, ColumnMenu, Delete
return



GuiContextMenu: ; Launched in response to a right-click or press of the Apps key.
if (A_GuiControl <> "lvData")
	; Right-click outside of the ListView
    return
if (LV_GetCount("") and A_GuiY < 239)
{
	; Right-click in the ListView header row
	Oops(lLvEventsContextInHeader)
    return
}
intColNumber := 0 ; tell MenuFilter and MenuSearch that search is global
Menu, ContextMenu, Add
Menu, ContextMenu, DeleteAll ; to avoid ghost lines at the end when menu is re-created
if !LV_GetCount("Column")
	Menu, ContextMenu, Add, % L(lLvEventsCreateNewFile), MenuCreateNewFile
else if !LV_GetCount("")
{
	Menu, ContextMenu, Add, % L(lLvEventsAddrowMenu), MenuAddRow
	Menu, ContextMenu, Add, % L(lLvEventsCreateNewFile), MenuCreateNewFile
	Menu, ContextMenu, Add, % ((lLvEventsFilterReload)), MenuFilterReload
	Menu, ContextMenu, % blnFilterActive ? "Enable" : "Disable", %lLvEventsFilterReload%
}
else
{
	intRowNumber := A_EventInfo
	Menu, ContextMenu, Add, % L(lLvEventsSelectAll), MenuSelectAll
	Menu, ContextMenu, Add, % L(lLvEventsDeselectAll), MenuSelectNone
	Menu, ContextMenu, Add, % L(lLvEventsReverseSelection), MenuSelectReverse
	Menu, ContextMenu, Add
	Menu, ContextMenu, Add, % L(lLvEventsEditrowMenu), MenuEditRow
	Menu, ContextMenu, Add, % L(lLvEventsAddrowMenu), MenuAddRow
	Menu, ContextMenu, Add, % L(lLvEventsDeleteRowMenu), MenuDeleteRow
	Menu, ContextMenu, Add
	Menu, ContextMenu, Add, % ((lLvEventsFilterGlobal)), MenuFilter
	Menu, ContextMenu, Add, % ((lLvEventsFilterReload)), MenuFilterReload
	Menu, ContextMenu, % blnFilterActive ? "Enable" : "Disable", %lLvEventsFilterReload%
	Menu, ContextMenu, Add
	Menu, ContextMenu, Add, % ((lLvEventsSearchGlobal)), MenuSearch
	; REMOVED in v1.3 Menu, ContextMenu, Add, % ((lLvEventsReplaceGlobal)), MenuReplace
}
; Show the menu at the provided coordinates, A_GuiX and A_GuiY.  These should be used
; because they provide correct coordinates even if the user pressed the Apps key:
Menu, ContextMenu, Show, %A_GuiX%, %A_GuiY%
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
	blnFilterActive := false
}
return



MenuSelectAll:
GuiControl, Focus, lvData
LV_Modify(0, "Select")
Menu, ContextMenu, Delete
return



MenuSelectNone:
GuiControl, Focus, lvData
LV_Modify(0, "-Select")
Menu, ContextMenu, Delete
return



MenuSelectReverse:
GuiControl, Focus, lvData
Loop, % LV_GetCount()
	if IsRowSelected(A_Index)
		LV_Modify(A_Index, "-Select")
	else
		LV_Modify(A_Index, "Select")
Menu, ContextMenu, Delete
return



IsRowSelected(intRow)
{
	intNextSelectedRow := LV_GetNext(intRow - 1)
	return (intNextSelectedRow = intRow)
}



MenuAddRow:
MenuEditRow:
SearchShowRecord:
ReplaceShowRecord:
strShowRecordLabel := A_ThisLabel
Gui, 1:Submit, NoHide
if (strShowRecordLabel = "MenuAddRow")
{
	LV_Modify(0, "-Select")
	LV_Insert(0xFFFF, "Select Focus") ; add at the end of the list
	intRowNumber := LV_GetNext()
	LV_Modify(intRowNumber, "Vis")
	strSaveRecordButton := "ButtonSaveRecordAddRow"
	strCancelButton := "ButtonCancelAddRow"
	strGuiTitle := L(lLvEventsAddrow, lAppName)
	strCancelButtonLabel := lLvEventsCancel
}
else
{
	strSaveRecordButton := "ButtonSaveRecord"
	strCancelButton := "ButtonCancel"
	strCancelButtonLabel := (strShowRecordLabel = "MenuEditRow" ? lLvEventsCancel : lLvEventsNext)
	strGuiTitle := L((strShowRecordLabel = "MenuEditRow" ? lLvEventsEditrow : lLvEventsSearchEditrow), lAppName, intRowNumber, LV_GetCount())
}

if (intRowNumber = 0)
	intRowNumber := 1
intGui1WinID := WinExist("A")
Gui, 2:New, +Resize +Hwndstr2GuiHandle, %strGuiTitle%
Gui, 2:+Owner1 ; Make the main window (Gui #1) the owner of the EditRow window (Gui #2).
Gui, 1:Default
SysGet, intMonWork, MonitorWorkArea 
intColWidth := 380
intEditWidth := intColWidth - 20
intMaxNbCol := Floor(intMonWorkRight / intColWidth)
intX := 10
intY := 5
intCol := 1
strZoomField := ""
loop, % LV_GetCount("Column")
{
	if ((intY + 100) > intMonWorkBottom)
	{
		if (intCol = 1)
			Gosub, DisplayShowRecordsButtons
		if (intCol = intMaxNbCol)
		{
			intYLabel := intY
			Gui, 2:Add, Text, y%intYLabel% x%intX% vstrLabelMissing, % L(lLvEventsFieldsMissing)
			intLastFieldIn2Gui := A_Index - 1
			break
		}
		intCol := intCol + 1
		intX := intX + intColWidth
		intY := 5
	}
	else
		intLastFieldIn2Gui := A_Index
	intYLabel := intY
	intYEdit := intY + 15
	LV_GetText(strColHeader, 0, A_Index)
	if (A_Index = intCurrentSortColumn)
		strColHeader := SubStr(strColHeader, 3)
	Gui, 2:Add, Text, y%intYLabel% x%intX% vstrLabel%A_Index%, %strColHeader%
	LV_GetText(strColData, intRowNumber, A_Index)
	if (strShowRecordLabel = "ReplaceShowRecord" and A_Index = intColNumber)
	{
		strPreviousCaseSense := A_StringCaseSense 
		StringCaseSense, % (blnReplaceCaseSensitive ? "On" : "Off")
		StringReplace, strColData, strColData, %strSearch%, %strReplace%, All
		StringCaseSense, %strPreviousCaseSense%
	}
	Gui, 2:Add, Edit, y%intYEdit% x%intX% w%intEditWidth% vstrEdit%A_Index% +HwndstrEditHandle, %strColData%
	ShrinkEditControl(strEditHandle, 2, "2")
	GuiControlGet, intPosEdit, 2:Pos, %strEditHandle%
	intY := intY + intPosEditH + 19
	intNbFieldsOnScreen := A_Index ; incremented at each occurence of the loop
}
if (intCol = 1) ; duplicate of line above in the loop, but much simpler that way
	Gosub, DisplayShowRecordsButtons

if ((strShowRecordLabel = "SearchShowRecord" or strShowRecordLabel = "ReplaceShowRecord") and intColFound)
{
	GuiControl, 2:Focus, strEdit%intColFound%
	Send, ^a
}
Gui, 2:Show, AutoSize Center
Gui, 1:+Disabled
return



DisplayShowRecordsButtons:
if (strShowRecordLabel = "ReplaceShowRecord")
	Gui, 2:Add, Button, y%intY% x10 vbtnReplaceAll gButtonReplaceAll, % L(lLvEventsReplaceAll)
if (strShowRecordLabel = "SearchShowRecord" or strShowRecordLabel = "ReplaceShowRecord")
	Gui, 2:Add, Button, % (strShowRecordLabel = "ReplaceShowRecord" ? "yp x+5" : "y" . intY . "x10") . " vbtnStopSearch gButtonStopSearch", % L(lLvEventsStop)
Gui, 2:Add, Button, % (strShowRecordLabel = "SearchShowRecord" ? "yp x+5" : "y" . intY . "x60") . " vbtnSaveRecord g" . strSaveRecordButton, % L(lLvEventsSave)
Gui, 2:Add, Button, yp x+5 vbtnCancel g%strCancelButton%, %strCancelButtonLabel%
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
	SB_SetText(L(lSBRecordsSize, intNewNbRows, (intActualSize) ? intActualSize : " <1"), 1)
else
{
	SB_SetText(L(lSBEmpty), 1)
	SB_SetText("", 3)
}
return



MenuFilter:
Gui, 1:+OwnDialogs 
MsgBox, % (4+32+256), % L(lAppName), % L(lLvEventsFilterCaution)
IfMsgBox, No
	return
InputBox, strFilter, % L(lLvEventsFilterInputTitle, lAppName), %lLvEventsFilterInput%, , , 150
if !StrLen(strFilter)
	return
intPrevNbRows := LV_GetCount()
intProgressBatchSize := ProgressBatchSize(intPrevNbRows)
ProgressStart(intProgressType, intPrevNbRows, lLvEventsFilterProgress)
intRowNumber := 0 ; scan each matching row of the ListView
GuiControl, -Redraw, lvData ; stop drawing the ListView during filtering
loop, %intPrevNbRows%
{
	intRowNumber := intRowNumber + 1
	if !Mod(A_Index, intProgressBatchSize)
		ProgressUpdate(intProgressType, A_Index, intPrevNbRows, lLvEventsFilterProgress)
			; update progress bar only every %intProgressBatchSize% records
	if NotMatchingRow(intRowNumber, strFilter, intColNumber, intColFound, blnReplaceCaseSensitive) ; ByRef intColFound not used here
	{
		LV_Delete(intRowNumber)
		intRowNumber := intRowNumber - 1 ; reduce the counter for deleted row
	}
}
GuiControl, +Redraw, lvData ; redraw the ListView
ProgressStop(intProgressType)
blnFilterActive := true

intNewNbRows := LV_GetCount()
LV_Modify(intNewNbRows, "Select Vis")
intActualSize := Round(intActualSize * intNewNbRows / intPrevNbRows)
if (intNewNbRows)
	SB_SetText(L(lSBRecordsSize, intNewNbRows, (intActualSize) ? intActualSize : " <1"), 1)
else
{
	SB_SetText(L(lSBEmpty), 1)
	SB_SetText("", 3)
}
return



MenuFilterReload:
Gui, 1:+OwnDialogs 
MsgBox, % (3+32+256), % L(lAppName), % L(lLvEventsFilterReloadPrompt)
IfMsgBox, Yes
{
	gosub, DeleteListviewData
	gosub, ButtonLoadFile
}
return



MenuSearch:
MenuReplace:
Gui, 1:+OwnDialogs
intSelectedRows := LV_GetCount("Selected")
if (intSelectedRows > 1)
	InputBox, strSearch, % L((A_ThisLabel = "MenuSearch" ? lLvEventsSearchInputTitle : lLvEventsReplaceInputTitle), lAppName), % L(lLvEventsSearchInputSelected, intSelectedRows), , , 150
else
{
	InputBox, strSearch, % L((A_ThisLabel = "MenuSearch" ? lLvEventsSearchInputTitle : lLvEventsReplaceInputTitle), lAppName), %lLvEventsSearchInput%, , , 150
	intSelectedRows := 0
}
if (A_ThisLabel = "MenuReplace")
{
	InputBox, strReplace, % L(lLvEventsReplaceInputTitle, lAppName), %lLvEventsReplaceInput%, , , 150
	MsgBox, 35, % L(lLvEventsReplaceInputTitle, lAppName), %lLvEventsReplaceCaseSensitive%
	IfMsgBox, Yes
		blnReplaceCaseSensitive := True
	IfMsgBox, No
		blnReplaceCaseSensitive := False
	IfMsgBox, Cancel
		strReplace := ""
}
else
	blnReplaceCaseSensitive := False ; required for NotMatchingRow
if !StrLen(strSearch) or (A_ThisLabel = "MenuReplace" and !StrLen(strReplace))
	return
intRowNumber := 0
intLastRow := LV_GetCount()
blnNotFound := true
blnReplaceAll := false
Loop
{
	if (intSelectedRows > 1)
		intRowNumber := LV_GetNext(intRowNumber) ;  returns 0 if no more selected row
	else
		intRowNumber := intRowNumber + 1
	if (!intRowNumber) or (intRowNumber > intLastRow)
		break
	if !NotMatchingRow(intRowNumber, strSearch, intColNumber, intColFound, blnReplaceCaseSensitive) ; ByRef intColFound to highlight the field in edit row window
	{
		blnNotFound := False
		LV_Modify(intRowNumber, "Vis Select")
		if (A_ThisLabel = "MenuSearch")
			Gosub, SearchShowRecord
		else ; MenuReplace
			if (blnReplaceAll)
			{
				LV_GetText(strCell, intRowNumber, intColNumber)
				strPreviousCaseSense := A_StringCaseSense 
				StringCaseSense, % (blnReplaceCaseSensitive ? "On" : "Off")
				StringReplace, strCell, strCell, %strSearch%, %strReplace%, All
				StringCaseSense, %strPreviousCaseSense%
				LV_Modify(intRowNumber, "Col" . intColNumber, strCell)
			}
			else
				Gosub, ReplaceShowRecord
		WinWaitClose, %strGuiTitle%
		LV_Modify(intRowNumber, "-Select")
	}
}
if (blnNotFound)
	Oops(lLvEventsSearchNotFound, strSearch)
return



NotMatchingRow(intRow, strFilter, intColNumber, ByRef intColFound, blnReplaceCaseSensitive)
{
	if (intColNumber) ; column filter/search
	{
		LV_GetText(strCell, intRow, intColNumber)
		if InStr(strCell, strFilter, blnReplaceCaseSensitive)
		{
			intColFound := intColNumber
			return False
		}
	}
	else ; global filter/search
		Loop, % LV_GetCount("Column")
		{
			LV_GetText(strCell, intRow, A_Index)
			if InStr(strCell, strFilter)
			{
				intColFound := A_Index
				return False
			}
		}
	intColFound := 0
	return True
}




; --------------------- GUI1 --------------------------


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
if (blnSkipConfirmQuit)
	ExitApp
MsgBox, 292, %lAppName%, % L(lQuitApp, lAppName) ; task modal
IfMsgBox, Yes
	ExitApp
return




; --------------------- GUI2 --------------------------


ButtonStopSearch:
Gui, 1:Default
intRowNumber := intLastRow ; cancel search by skipping to last row
gosub, ButtonCancel
return



ButtonReplaceAll:
blnReplaceAll := true
gosub, ButtonSaveRecord
return



ButtonSaveRecordAddRow:
Gui, 1:Default
intRowNumber := LV_GetCount()
intActualSize := Round(intActualSize + (intActualSize / intRowNumber))
SB_SetText(L(lSBRecordsSize, LV_GetCount(), (intActualSize) ? intActualSize : " <1"), 1)
gosub, ButtonSaveRecord
return



ButtonSaveRecord:
Gui, 2:Submit
Gui, 1:Default
loop, % LV_GetCount("Column")
	if (A_Index <= intLastFieldIn2Gui)
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
if (A_ThisLabel = "2GuiEscape")
	and (strShowRecordLabel = "SearchShowRecord" or strShowRecordLabel = "ReplaceShowRecord")
	intRowNumber := intLastRow ; cancel search by skipping to last row
WinActivate, ahk_id %intGui1WinID%
LV_Modify(0, "-Select -Focus")
return



2GuiSize: ; Expand or shrink the ListView in response to the user's resizing of the window.
if A_EventInfo = 1  ; The window has been minimized.  No action needed.
    return
GuiControl, 2:Move, btnReplaceAll, % "x" . (A_GuiWidth - 235)
GuiControl, 2:Move, btnStopSearch, % "x" . (A_GuiWidth - 155)
GuiControl, 2:Move, btnSaveRecord, % "x" . (A_GuiWidth - 105)
GuiControl, 2:Move, btnCancel, % "x" . (A_GuiWidth - 55)
if intCol > 1 ; The window has been minimized.  No action needed.
    return
intWidthSize := A_GuiWidth - 20
Loop, %intNbFieldsOnScreen%
{
	GuiControlGet, strFielfHwnd, 2:Hwnd, strEdit%A_Index%
	GuiControl, 2:Move, strEdit%A_Index%, % "w" . A_GuiWidth - (GetNbLinesOfControl(strFielfHwnd) > 1 ? 40 : 20)
}
return



; --------------------- GUI3 --------------------------


ButtonZoom:
StringReplace, strZoomField, A_GuiControl, strEdit
intGui2WinID := WinExist("A")
GuiControlGet, strZoomFieldContent, , %strZoomField%
Gui, 3:New, +Resize -MinimizeBox, % L(lLvEventsZoomTitle, lAppName)
Gui, 3:+Owner2 ; Make the edit window (Gui #2) the owner of the Zoom window (Gui #3).
Gui, 3:Add, Edit, x10 y5 w400 h300 vstrZoomedEdit, %strZoomFieldContent%
Gui, 3:Add, Button, x10 y+20 vbtnZoomSave gButtonZoomSave, %lLvEventsSave%
Gui, 3:Add, Button, x+10 yp vbtnZoomCancel gButtonZoomCancel, %lLvEventsCancel%
Gui, 3:Show, AutoSize Center
Gui, 2:+Disabled
return



ButtonZoomSave:
ButtonZoomCancel:
3GuiClose:
3GuiEscape:
if (A_ThisLabel = "ButtonZoomSave")
{
	Gui, 3:Submit, NoHide
	strNewZoomFieldContent := strZoomedEdit
	GuiControl, 2:, %strZoomField%, %strNewZoomFieldContent%
}
Gui, 2:-Disabled
Gui, 3:Destroy
WinActivate, ahk_id %intGui2WinID%
strZoomField := ""
return



3GuiSize: ; Expand or shrink the ListView in response to the user's resizing of the window.
if A_EventInfo = 1  ; The window has been minimized.  No action needed.
    return
GuiControl, 3:Move, strZoomedEdit, % "x10 y5 w" . (A_GuiWidth - 15) . " h" . (A_GuiHeight - 45)
GuiControl, 3:Move, btnZoomSave, % "x" . (A_GuiWidth - 105) . " y" . (A_GuiHeight - 30)
GuiControl, 3:Move, btnZoomCancel, % "x" . (A_GuiWidth - 55) . " y" . (A_GuiHeight - 30)
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
	, strRealFieldDelimiter3, strFieldEncapsulator3, strEolReplacement, L(lExportSaving), strCurrentFileEncodingSave)
if (ErrorLevel)
	if (ErrorLevel = 1)
		Oops(lExportSystemError, A_LastError)
	else
		Oops(lExportUnknownError)
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
	, intProgressType, blnOverwrite, L(lExportSaving), strCurrentFileEncodingSave)
if (ErrorLevel)
	if (ErrorLevel = 1)
		Oops(lExportSystemError, A_LastError)
	else if (ErrorLevel = 4 or ErrorLevel = 5)
		Oops(lExportHTMLError)
	else
		Oops(lExportParameterError)
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
ObjCSV_Collection2XML(obj, strFileToExport, intProgressType, blnOverwrite, L(lExportSaving), strCurrentFileEncodingSave)
if (ErrorLevel)
	if (ErrorLevel = 1)
		Oops(lExportSystemError, A_LastError)
	else
		Oops(lExportUnknownError)
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
		, intProgressType, blnOverwrite, L(lExportSaving), strCurrentFileEncodingSave)
	if (ErrorLevel)
		if (ErrorLevel = 1)
			Oops(lExportSystemError, A_LastError)
		else if (ErrorLevel = 4 or ErrorLevel = 5)
			Oops(lExportHTMLError)
		else
			Oops(lExportParameterError)
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

if GetKeyState("Shift") and GetKeyState("LWin")
{
	IniWrite, 1, %strIniFile%, Global, Donator ; stop Freeware donation nagging
	MsgBox, 64, %lAppName%, %lDonateThankyou%
}
else if Time2Donate(intStartups, blnDonator)
{
	MsgBox, 36, % l(lDonateTitle, intStartups, lAppName), % l(lDonatePrompt, lAppName, intStartups)
	IfMsgBox, Yes
		Gosub, ButtonDonate
}
IniWrite, % (intStartups + 1), %strIniFile%, Global, Startups

strLatestVersion := Url2Var("https://raw.github.com/JnLlnd/CSVBuddy/master/latest-version.txt")

if RegExMatch(strCurrentVersion, "(alpha|beta)")
	or (FirstVsSecondIs(strLatestSkipped, strLatestVersion) >= 0 and (!blnButtonCheck4Update))
	return

if FirstVsSecondIs(strLatestVersion, lAppVersion) = 1
{
	Gui, 1:+OwnDialogs
	SetTimer, ChangeButtonNames4Update, 50

	MsgBox, 3, % l(lTab5UpdateTitle, lAppName), % l(lTab5UpdatePrompt, lAppName, lAppVersion, strLatestVersion), 30
	IfMsgBox, Yes
		Run, http://code.jeanlalonde.ca/csvbuddy/
	IfMsgBox, No
		IniWrite, %strLatestVersion%, %strIniFile%, global, LatestVersionSkipped
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



Time2Donate(intStartups, blnDonator)
{
	if (intStartups > 200)
		intDivisor := 50
	else if (intStartups > 120)
		intDivisor := 40
	else if (intStartups > 60)
		intDivisor := 30
	else
		intDivisor := 20
		
	return !Mod(intStartups, intDivisor) and !(blnDonator)
}



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



RemoveSorting:
LV_GetText(strColHeader, 0, intCurrentSortColumn)
LV_ModifyCol(intCurrentSortColumn, , SubStr(strColHeader, 3))
intCurrentSortColumn := 0
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
		if (A_Index = intCurrentSortColumn)
			strColumnHeader := SubStr(strColumnHeader, 3)
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



ShrinkEditControl(strThisEditHandle, intMaxRows, strGuiName)
{
	global ; allow dynamic variable name for Button + (using strEdit plus strThisEditHandle in the variable name)
	
	intNbRows := GetNbLinesOfControl(strThisEditHandle)
	if (intNbRows > intMaxRows)
	{
		GuiControlGet, intPosEdit, %strGuiName%:Pos, %strThisEditHandle%
		intEditMargin := 8 ; top & bottom margin of the Edit control (regardless of the nb of rows)
		intOriginalHeight := intPosEditH
		intHeightOneRow := Round((intOriginalHeight - intEditMargin) / intNbRows)
		intNewHeight := (intHeightOneRow * intMaxRows) + intEditMargin
		; MsgBox, % "intNbRows: " . intNbRows . "`nintOriginalHeight: " . intOriginalHeight 
		;	. "`nintHeightOneRow: " . intHeightOneRow . "`nintNewHeight: " . intNewHeight
		GuiControl, %strGuiName%:Move, %strThisEditHandle%, % "x" . intPosEditX + 20 . "h" . intNewHeight ; width - 20 set by 2GuiSize
		Gui, %strGuiName%:Add, Button, x%intPosEditX% y%intPosEditY% w20 gButtonZoom vstrEdit%strThisEditHandle%, %lLvEventsZoomOut%
	}
}



CheckIfFileExistOverwrite(strFileName, blnCanAppend)
{
	if !StrLen(strFileName)
		return -1
	if !FileExist(strFileName)
		return True
	else
	{
		if (blnCanAppend)
		{
			Gui, 1:+OwnDialogs 
			MsgBox, 35, % L(lFuncIfFileExistTitle, lAppName), % L(lFuncIfFileExistMessageAppend, strFileName)
			IfMsgBox, Yes
				return True ; overwrite
			IfMsgBox, No
				return False ; append
			IfMsgBox, Cancel
				return -1 ; cancel
		}
		else
		{
			Gui, 1:+OwnDialogs 
			MsgBox, 35, % L(lFuncIfFileExistTitle, lAppName), % L(lFuncIfFileExistMessage, strFileName)
			IfMsgBox, Yes
				return True ; overwrite
			return -1 ; cancel
		}

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


GetNbLinesOfControl(strThisEditHandle)
{
	EM_GETLINECOUNT = 0xBA
	SendMessage, %EM_GETLINECOUNT%,,,, ahk_id %strThisEditHandle%
	return ErrorLevel
}



