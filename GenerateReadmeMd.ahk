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

H(1, L("~1~ (~2~) - Read me", lAppName,  "v" . lAppVersion))
P(lDocDesc450)
P("Written using AutoHotkey_L v1.1.09.03+ (http://www.ahkscript.org)`nBy Jean Lalonde (JnLlnd on [AHK forum](http://www.autohotkey.com/boards/)`nFirst official release: 2013-11-30")

H(2, "Links")
B("[Application home](http://code.jeanlalonde.ca/csvbuddy/)")
B("[Download 32-bits / 64-bits](http://code.jeanlalonde.ca/ahk/csvbuddy/csvbuddy.zip) (latest version)")
B("[Description and documentation](http://code.jeanlalonde.ca/ahk/csvbuddy/csvbuddy-doc.html)")

H(2, "History")
H(3, "2022-04-15 v2.1.9.4")
B("adjust display of tabs content for HDPI screens")
B("fix bug opening the wrong record editor")
B("increase size of the undo buttons")
H(3, "2022-04-06 v2.1.9.3")
B("Merge fields")
B("add a ""Merge"" command in tab 2 (previously named ""Reuse"") with separate text boxes for 1) fields and format (including existing fields enclosed by merge delimiters, for example ""[FirstName] [LastName]) and 2) new field name")
B("display error message if a merge field has invalid syntax when loading a CSV file")
B("support placeholder ROWNUMBER (enclosed with reuse delimiters) in reuse fields format section, for example ""#[ROWNUMBER]""")
B("update status bar with progress during build merge fields")
B("User interface")
B("redesign the user interface to support font changes in the main window")
B("new font settings in ""Options"" tab for labels (default Microsoft Sans Serif, 11), text input (default Courier New, 10) and list (default Microsoft Sans Serif, 10)")
B("change the order of commands in tab 2 to ""Rename"", ""Order"", ""Select"" and ""Merge""")
B("add Undo buttons for each commands in tab 2 allowing to revert the last change")
B("track changes in list data and alert user for unsaved changes before quiting the application; add a section to status bar to show that changes need to be saved")
B("remove the previous prompt when quitting (and its associated option in tab 6)")
B("stop quitting the app on Escape key")
B("fix bug preventing from search something and replace it with nothing")
B("disable window during loading file, loading to listview, saving to csv or exporting data")

H(3, "2022-03-01 v2.1.9.2")
B("Support multiple reuse in select command")
B("Known limitation: when adding reuse fields, all existing fields must be present in the Select list (abort select command if required)")
H(3, "2022-02-25 v2.1.9.1")
B("Reuse fields allowing, when loading a file or using the Select command, to create an new field based on the content of previous fields in each row")
B("See reuse specifications and examples at https://github.com/JnLlnd/CSVBuddy/issues/53")
B("Configurable reuse opening and closing delimiters in the ""Options"" tab")
B("In this release, reuse fields are not supported in Export and only one reuse field can be set when loading a file or selecting fields (multiple reuse are planned for future beta releases).")
H(3, "2017-12-10 v2.1.6")
B("Fix bug when changing the Fixed with default in Export tab.")
H(3, "2017-07-20 v2.1.5")
B("Fix bug when processing HTML or XML multi-line content, reversing earlier change done to support non-standard CSV files created by XL causing issue (stripping some ""="") in encapsulated fields.")
H(3, "2017-07-20 v2.1.4")
B("Fix bug: show the end-of-line replacement field when loading a file from the command-line (or by double-click a file in Explorer).")
H(3, "2016-12-23 v2.1.3")
B("Fix bug preventing correct detection of current field delimiter when file is loaded (first delimiter detected in this order: tab, semicolon (;), comma (,), colon (:), pipe (|) or tilde (~)).")
H(3, "2016-12-22 v2.1.2")
B("Fix bug when creating header if ""Set header"" is selected and ""Custom header"" is empty.")
B("Now using library ObjCSV v0.5.8")
H(3, "2016-12-20 v2.1.1")
B("Fix bug when ""Set header"" is selected and ""Custom header"" is empty, columns with ""C"" field names generated are now sorted correctly.")
B("Now using library ObjCSV v0.5.7")
H(3, "2016-10-20 v2.1")
B("Stop trimming (removing spaces at the beginning or end) data value read from CSV file (but still trimming field names read from CSV file).")
B("Now using library ObjCSV v0.5.6")
H(3, "2016-09-06 v2.0")
B("New Record editor dialog box with field-by-field edition (support up to 200 fields per row)")
B("New ""Options"" tab to change setting values saved to the CSVBuddy.ini file")
B("New option ""Record editor"" for choice of 1) ""Full screen Editor"" (legacy) or 2) ""Field-by-field Editor"" (new), default is 2")
B("New option ""Encapsulate all values"" to always enclose saved values with the encapsulator character")
B("Display a ""Create"" button on first tab to create a new file based on the Set header data")
B("Remember the last folder where a file was loaded")
B("Code signature with certificate from DigiCert")
B("See history for v1.3.9 and v1.3.9.1 for details and bug fixes")
H(3, "2016-08-31 v1.3.9.1")
B("Unlock the ""CSV file to load"" zone in first tab, allowing to type or paste a file name")
B("Add a ""Create"" button on first tab to create a new file based on the ""CSV file Header"" data")
B("Remember the last folder where a file was loaded and use it as default location when loading another file")
B("Add items to context menu to add and edit rows with field-by-field editor")
B("Fix bug when filter on a column, hitting the Cancel button now cancels the filtering")
B("Fix bug default encoding in firts tab is now ""Detect"" if no default value is saved to the ini file")
B("Apply grid setting and colors to list of field in field-by-field row editor")
B("Fix bug reading values in ini file for ""Skip Ready Prompt"" and ""Skip Quit Prompt""")
B("Fix visual glitch with labels close to left part of tabs, labels were overlaping left vertical line")
H(3, "2016-08-28 v1.3.9")
B("New Record editor dialog box with field-by-field edition")
B("New ""Options"" tab to change setting values saved to the CSVBuddy.ini file")
B("New option ""Record editor"" for choice of 1) ""Full screen Editor"" (legacy) or 2) ""Field-by-field Editor"" (new). Default is 2.")
B("New option ""Encapsulate all values"" to always enclose saved values with the encapsulator character")
B("Help button for Options")
B("Bug fix: now detect the end-of-line character(s) in fields where line-breaks have to be replaced by a replacement string (detected in this order: CRLF, LF or CR). The first end-of-lines character(s) found is used for remaining fields and records.")
B("Now using library ObjCSV v0.5.5")
H(3, "2016-07-23 v1.3.3")
B("If file encoding is not specified (leave encoding at ""Detect"") when loading a file, it is loaded as UTF-8 or UTF-16 if these formats are detected in file header or as ANSI for all other formats (and displayed as such in load and save encoding encoding lists); UTF-8-RAW and UTF-16-RAW formats cannot be auto-detected and must be selected in encoding list to load files in these formats")
B("Add values SreenHeightCorrection and SreenWidthCorrection in CSVBuddy.ini file (enter negative values in pixels to reduce the height or width of edit row dialog box)")
H(3, "2016-06-08 v1.3.2")
B("Fix bug introduced in v1.2.9.1 preventing from saving manual record edits in some circumstances")
B("Automatic file encoding detection is now restricted to UTF-8 or UTF-16 encoded files (no BOM)")
B("Read the new value DefaultFileEncoding= (under [global]) in CSVBuddy.ini to set the default file encoding (possible values ANSI, UTF-8, UTF-16, UTF-8-RAW, UTF-16-RAW or CPnnn)")
H(3, "2016-05-21 v1.3.1")
B("Change licence to Apache 2.0")
H(3, "2016-05-18 v1.3")
B("Select file encoding when loading, saving or exporting files")
B("Encoding supported: ANSI (default), UTF-8 (Unicode 8-bit), UTF-16 (Unicode 16-bit), UTF-8-RAW (no BOM), UTF-16-RAW (no BOM) or custom codepage CPnnnn")
B("Set custom codepage values in CSVBuddy.ini file")
B("Other changes")
B("When editing a record, zoom button to edit long strings in a large window")
B("When there is not enough space on screen to edit all fields in a file, support editing the visible fields without loosing the content of the missing fields")
B("Automatic detection of loaded file encoding (this is not possible for files without byte order mark (BOM))")
B("Deselect all rows after global search is cancelled")
B("Add status bar right section to display hint about list menu")
B("Add help message when right-clicking on column headers")
B("Remove global search and replace (needing to be reworked)")
H(3, "2014-08-31 v1.2.3")
B("Bug fix: when saving or exporting file with a column sort indicator")
H(3, "2014-03-17 v1.2.2")
B("Bug fix: after a column sort, fix names errors in column headers")
B("Bug fix: by safety, remove sorting column indicator before any action in edit columns tabs")
H(3, "2014-03-09 v1.2.1")
B("Bug fix: ini variables missing when ini file already existed making grid black text on black background")
B("Bug fix: bug with multiple instances")
H(3, "2014-03-07 v1.2")
B("Search and replace by column, replacement case sensitive or not")
B("Confirm each replacement or replace all")
B("During search or replace, select and highlight the current row when displaying the record found")
B("Option in ini file to display or not a grid around cells in list zone")
B("Options in ini file to choose background and text colors in list zone")
B("Up or down arrow to indicate which field is the current sort key")
B("Allow multiple instances of the app to run simultaneously")
B("Import CSV files created by XL that include equal sign before the opening field encasulator")
H(3, "2013-12-30 v1.1")
B("Filter by column: click on a header to retain only rows with the keyword appearing in this column")
B("Global filtering: right-click in the list zone to retain only rows with the keyword appearing in any column")
B("Search by column: find the next row having the keyword in this column and open it in row edit window")
B("Global search: find the next row having the keyword in any column and open it in row edit window")
B("In edit row window search result, highlight the field containing the searched keyword")
B("Added stop and next buttons to edit row window when search in progress")
B("Added reload original file to the column menu and the list context menu")
B("Display the current edited record number in edit row title bar")
B("Add blnSkipConfirmQuit option in ini file to skip the quit confirm prompt, default to false")
B("Use ObjCSV library v0.4 for better file system error handling")
H(3, "2013-12-27 v1.0.1")
B("Fix a small bug in the ""Add row"" dialog box where default values were presented when this dialog box should not have default values")
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
