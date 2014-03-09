
# CSV Buddy (v1.2.1) - Read me


CSV Buddy helps you make your CSV files ready to be imported by a variety of software. You can load files with all sort of field delimiters (comma, tad, semi-colon, etc.) and encapsulators (double/single-quotes or any other character). Convert line breaks in data field making your file XL ready. Rename/reorder fields, add or edit records, filter or search, search and replace, save with any delimiters and export to fixed-width, HTML templates or XML formats. Freeware.


Written using AutoHotkey_L v1.1.09.03+ (http://www.ahkscript.org)  
By JnLlnd on [AHK forum](http://www.ahkscript.org/boards/)  
First official release: 2013-11-30


## Links

* [Application home](http://code.jeanlalonde.ca/csvbuddy/)
* [Download 32-bits / 64-bits](http://code.jeanlalonde.ca/ahk/csvbuddy/csvbuddy.zip) (latest version)
* [Description and documentation](http://code.jeanlalonde.ca/ahk/csvbuddy/csvbuddy-doc.html)

## History


### 2014-03-09 v1.2.1

* Bug fix: ini variables missing when ini file already existed making grid black text on black background
* Bug fix: bug with multiple instances

### 2014-03-07 v1.2

* Search and replace by column, replacement case sensitive or not
* Confirm each replacement or replace all
* During search or replace, select and highlight the current row when displaying the record found
* Option in ini file to display or not a grid around cells in list zone
* Options in ini file to choose background and text colors in list zone
* Up or down arrow to indicate which field is the current sort key
* Allow multiple instances of the app to run simultaneously
* Import CSV files created by XL that include equal sign before the opening field encasulator

### 2013-12-30 v1.1

* Filter by column: click on a header to retain only rows with the keyword appearing in this column
* Global filtering: right-click in the list zone to retain only rows with the keyword appearing in any column
* Search by column: find the next row having the keyword in this column and open it in row edit window
* Global search: find the next row having the keyword in any column and open it in row edit window
* In edit row window search result, highlight the field containing the searched keyword
* Added stop and next buttons to edit row window when search in progress
* Added reload original file to the column menu and the list context menu
* Display the current edited record number in edit row title bar
* Add blnSkipConfirmQuit option in ini file to skip the quit confirm prompt, default to false
* Use ObjCSV library v0.4 for better file system error handling

### 2013-12-27 v1.0.1

* Fix a small bug in the "Add row" dialog box where default values were presented when this dialog box should not have default values

### 2013-11-30 v1.0

* First official release
* Add records to existing data (right-click in the list zone)
* Create a new file from scratch (right-click in an empty list zone)
* Load the file mentioned as first parameter in the command line
* Add validation, confirm before exit and fix various small bugs

### 2013-11-03 v0.9

* Display "<1" (instead of "0") in status bar when file size is smaller than 0.5 K
* Removed CSV Buddy icon from the Tray
* Add three test delimited files to the package (see README.txt in the zip file)
* Fix default value of blnSkipHelpReadyToEdit in ini file to 0

### 2013-10-20 v0.8.1

* If an .ini file is not found in the program's folder, it is created with default values

### 2013-10-18 v0.8.0

* First release of BETA version
* History of ALPHA phase on [BitBucket](https://bitbucket.org/JnLlnd/csvbuddy/) (private repository)

## <a name="copyright"></a>Copyright


CSV Buddy - Copyright (C) 2013-2014  Jean Lalonde  
  
This software is provided 'as-is', without any express or implied warranty.  In no event will the authors be held liable for any damages arising from the use of this software.  
  
Permission is granted to anyone to use this software for any purpose,  including commercial applications, and to alter it and redistribute it freely, subject to the following restrictions:  
  
1. The origin of this software must not be misrepresented; you must not claim that you wrote the original software. If you use this software in a product, an acknowledgment in the product documentation would be appreciated but is not required.  
2. Altered source versions must be plainly marked as such, and must not be misrepresented as being the original software.  
3. This notice may not be removed or altered from any source distribution.  
  
Jean Lalonde, <A HREF="mailto:ahk@jeanlalonde.ca">ahk@jeanlalonde.ca</A>


