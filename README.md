
# CSV Buddy (v1.3.3) - Read me


CSV Buddy helps you make your CSV files ready to be imported by a variety of software. Load files with all sort of field delimiters (comma, tad, semi-colon) and encapsulators (double/single-quotes or any other). Convert line breaks in data field (XL ready). Rename/reorder fields, add/edit records, filter or search, search and replace, save with any delimiters and export to fixed-width, HTML templates or XML formats. Unicode ready. Freeware.


Written using AutoHotkey_L v1.1.09.03+ (http://www.ahkscript.org)  
By JnLlnd on [AHK forum](http://www.ahkscript.org/boards/)  
First official release: 2013-11-30


## Links

* [Application home](http://code.jeanlalonde.ca/csvbuddy/)
* [Download 32-bits / 64-bits](http://code.jeanlalonde.ca/ahk/csvbuddy/csvbuddy.zip) (latest version)
* [Description and documentation](http://code.jeanlalonde.ca/ahk/csvbuddy/csvbuddy-doc.html)

## History


### 2016-07-23 v1.3.3

* If file encoding is not specified (leave encoding at "Detect") when loading a file, it is loaded as UTF-8 or UTF-16 if these formats are detected in file header or as ANSI for all other formats (and displayed as such in load and save encoding encoding lists); UTF-8-RAW and UTF-16-RAW formats cannot be auto-detected and must be selected in encoding list to load files in these formats
* Add values SreenHeightCorrection and SreenWidthCorrection in CSVBuddy.ini file (enter negative values in pixels to reduce the height or width of edit row dialog box)

### 2016-06-08 v1.3.2

* Fix bug introduced in v1.2.9.1 preventing from saving manual record edits in some circumstances
* Automatic file encoding detection is now restricted to UTF-8 or UTF-16 encoded files (no BOM)
* Read the new value DefaultFileEncoding= (under [global]) in CSVBuddy.ini to set the default file encoding (possible values ANSI, UTF-8, UTF-16, UTF-8-RAW, UTF-16-RAW or CPnnn)

### 2016-05-21 v1.3.1

* Change licence to Apache 2.0

### 2016-05-18 v1.3

* Select file encoding when loading, saving or exporting files
* Encoding supported: ANSI (default), UTF-8 (Unicode 8-bit), UTF-16 (Unicode 16-bit), UTF-8-RAW (no BOM), UTF-16-RAW (no BOM) or custom codepage CPnnnn
* Set custom codepage values in CSVBuddy.ini file
* Other changes
* When editing a record, zoom button to edit long strings in a large window
* When there is not enough space on screen to edit all fields in a file, support editing the visible fields without loosing the content of the missing fields
* Automatic detection of loaded file encoding (this is not possible for files without byte order mark (BOM))
* Deselect all rows after global search is cancelled
* Add status bar right section to display hint about list menu
* Add help message when right-clicking on column headers
* Remove global search and replace (needing to be reworked)

### 2014-08-31 v1.2.3

* Bug fix: when saving or exporting file with a column sort indicator

### 2014-03-17 v1.2.2

* Bug fix: after a column sort, fix names errors in column headers
* Bug fix: by safety, remove sorting column indicator before any action in edit columns tabs

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

## <a name="copyright"></a>Copyright 2013-2016 Jean Lalonde


Licensed under the Apache License, Version 2.0 (the "License");  
you may not use this file except in compliance with the License. You may obtain a copy of the License at  
  
<A HREF="http://www.apache.org/licenses/LICENSE-2.0">http://www.apache.org/licenses/LICENSE-2.0</A>  
  
Unless required by applicable law or agreed to in writing, software distributed under the License is distributed on an "AS IS" BASIS, WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied. See the License for the specific language governing permissions and limitations under the License.  
  
Jean Lalonde, <A HREF="mailto:ahk@jeanlalonde.ca">ahk@jeanlalonde.ca</A>


