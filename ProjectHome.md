# SVN4MSOffice #

## About ##

Add-ins for Microsoft Office applications that enables integration with Subversion using [TortoiseSVN](http://tortoisesvn.net/)

SVN4MSOffice was tested on Windows XP with MS Office 2002 and TortoiseSVN 1.4.7 and 14.8
Should work with MS Office 2000 and above, not sure about 2007 version

Report your bugs, feature requests, problems, compatibility issues [using this form](https://spreadsheets.google.com/viewform?key=pZZrFZQpaNhqG84CDsYk7OQ&email=true)

http://kosciak.blox.pl/resource/SVN4MSOffice_Word.JPG


## Features ##

  * integration with [TortoiseSVN](http://tortoisesvn.net/)
  * menu and toolbar - visible only if document's parent folder is under version control
  * buttons/menu items for Add, Commit, Update, Revert, Log, Status, Diff, Lock, Unlock
  * asks for Commit before closing versioned document
  * right buttons/menu items enabled if file is or isn't under version control

## Download ##

Check the [Downloads](http://code.google.com/p/svn4msoffice/downloads/list) section


## Installation ##

[TortoiseSVN](http://tortoisesvn.net/) has to be installed for the Add-ins to work!

**MS Word Add-in**

Copy the SVN4MSOffice-0.x.dot file to:
`%HOMEPATH%\%USERNAME%\Application Data\Microsoft\Word\STARTUP`

for example
`C:\Documents and Settings\User\Application Data\Microsoft\Word\STARTUP`

## Changelog ##

**MS Word Add-in v0.3**
  * first public release
  * more code refactoring - removing unnecessary modules (moved code to classes)
  * some new features and bugfixes
  * menu and toolbars are created in the Normal.dot instead of SVN4MSOffice-0.x.dot when needed
  * disabling and enabling toolbar buttons and menu items without the need to save SVN4MSOffice-0.x.dot

**MS Word Add-in v0.2**
  * major code refactoring
  * lots of bugfixes

**MS Word Add-in v0.1**
  * initial version
  * basic integration with TortoiseSVN
  * toolbar and menu creation


## TODO ##

  * tests on other Operating systems and MS Office versions
  * write code documentation
  * **Excel** and maybe Power Point Add-ins
  * auto locking of file when opeining versioned file [?]
  * auto commiting on save [?]
  * ~~find a way to enable/disable toolbar buttons and menu items without the need to save SVN4MSOffice-0.x.dot when closing MS Word~~



If you don't like my Add-in try [MSOfficeSVN](http://code.google.com/p/msofficesvn/)