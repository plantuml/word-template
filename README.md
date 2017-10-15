
# INTRODUCTION
This repository contains the Word Template Plugin for PlantUML.

The Word template allows using PlantUML directly from MS Word 2010/2013 on Windows (32/64 bit) without need to alter document templates or edit VBA macro's. MS Word Version 2007 may work, but is not tested. 

**Also tested in Word 2016, need to copy contents of "Word2016" folder

# INSTALLATION
First time: 
* install the right template version in Word
  * copy the (dotm) file in to `%appdata%\Microsoft\Word\STARTUP`
  * note: .dotm = Word Doc Template (office 2007/2010) with Macro's enabled
* copy Plantuml.jar to `%appdata%` folder
* install GraphViz
  * http://www.graphviz.org/Download_windows.php
    * use installer if you have rights to install applications; this will install graphviz in your program files (x86)
    * use zip for portable installation
      * extract in `%appdata%` 
      * (executable is then in `%appdata%\release\bin\dot.exe`)
  * if alternative portable installation, please set `GRAPHVIZ_DOT`  to location of DOT.EXE
* restart Word. You now should have a PlantUML menu!


# USING
Once installed, a special menu (PlantUML) should be available in Word as tab "PlantUML"

![](https://raw.githubusercontent.com/plantuml/word-template/master/images/menu.png)

Icon | Description
-- | --
P | show paragraph marks
Show PlantUML | reveal (green text) of PlantUML image sources (for editing)
Hide PlantUML | hide source, just show generated pictures (before releasing a document for review/UCC)
UML.1 | Generate current diagram (cursor in green PlantUML definition)
UML.* | Generate all (note: this may take seconds up to a minute for 100+ pictures). Press Ctrl-Break to abort.


# VBA CODE
For convenience, the current [VBA module](https://github.com/plantuml/word-template/tree/master/module) are listed in the current repository:
* GDIHanling
* PlantUML
* PlantumlFTP
* Registry
* ShellUtil

This allows to clearly follow VBA code changes over versions.

