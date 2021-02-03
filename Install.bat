@ECHO OFF
REM Current Folder is %~dp0
REM Install

REG ADD "HKCU\Software\Classes\*\shell\LyxSoft" /f
REG ADD "HKCU\Software\Classes\*\shell\LyxSoft" /v "Icon" /d "%SystemRoot%\System32\SHELL32.dll,316" /f
REG ADD "HKCU\Software\Classes\*\shell\LyxSoft" /v "MUIVerb" /d "LyxSoft Tools" /f
REG ADD "HKCU\Software\Classes\*\shell\LyxSoft" /v "SubCommands" /d "" /f

REG ADD "HKCU\Software\Classes\*\shell\LyxSoft\shell\LyxSoft.PicFileDate" /f
REG ADD "HKCU\Software\Classes\*\shell\LyxSoft\shell\LyxSoft.PicFileDate" /ve /d "Set Picture Date" /f
REG ADD "HKCU\Software\Classes\*\shell\LyxSoft\shell\LyxSoft.PicFileDate" /v "Icon" /d "%SystemRoot%\System32\SHELL32.dll,313" /f

REG ADD "HKCU\Software\Classes\*\shell\LyxSoft\shell\LyxSoft.PicFileDate\Command" /f
REG ADD "HKCU\Software\Classes\*\shell\LyxSoft\shell\LyxSoft.PicFileDate\Command" /ve /d "WScript.exe """%~dp0PicFileDate.vbs""" """%%1""" %%*" /f

REG ADD "HKCU\Software\Classes\*\shell\LyxSoft\shell\LyxSoft.PicFileDateName" /f
REG ADD "HKCU\Software\Classes\*\shell\LyxSoft\shell\LyxSoft.PicFileDateName" /ve /d "Set Picture Name with Date" /f
REG ADD "HKCU\Software\Classes\*\shell\LyxSoft\shell\LyxSoft.PicFileDateName" /v "Icon" /d "%SystemRoot%\System32\SHELL32.dll,313" /f

REG ADD "HKCU\Software\Classes\*\shell\LyxSoft\shell\LyxSoft.PicFileDateName\Command" /f
REG ADD "HKCU\Software\Classes\*\shell\LyxSoft\shell\LyxSoft.PicFileDateName\Command" /ve /d "WScript.exe """%~dp0PicFileDateName.vbs""" """%%1""" %%*" /f

REG ADD "HKCU\Software\Classes\*\shell\LyxSoft\shell\LyxSoft.PicNewDateFolder" /f
REG ADD "HKCU\Software\Classes\*\shell\LyxSoft\shell\LyxSoft.PicNewDateFolder" /ve /d "Create Folder with Date Name" /f
REG ADD "HKCU\Software\Classes\*\shell\LyxSoft\shell\LyxSoft.PicNewDateFolder" /v "Icon" /d "%SystemRoot%\System32\SHELL32.dll,313" /f

REG ADD "HKCU\Software\Classes\*\shell\LyxSoft\shell\LyxSoft.PicNewDateFolder\Command" /f
REG ADD "HKCU\Software\Classes\*\shell\LyxSoft\shell\LyxSoft.PicNewDateFolder\Command" /ve /d "WScript.exe """%~dp0PicNewDateFolder.vbs""" """%%1""" %%*" /f

REG ADD "HKCU\Software\Classes\*\shell\LyxSoft\shell\LyxSoft.PicFolderFilesToYearFolder" /f
REG ADD "HKCU\Software\Classes\*\shell\LyxSoft\shell\LyxSoft.PicFolderFilesToYearFolder" /ve /d "Put File into Folder with Year of the File" /f
REG ADD "HKCU\Software\Classes\*\shell\LyxSoft\shell\LyxSoft.PicFolderFilesToYearFolder" /v "Icon" /d "%SystemRoot%\System32\SHELL32.dll,313" /f

REG ADD "HKCU\Software\Classes\*\shell\LyxSoft\shell\LyxSoft.PicNewDateFolder\Command" /f
REG ADD "HKCU\Software\Classes\*\shell\LyxSoft\shell\LyxSoft.PicNewDateFolder\Command" /ve /d "WScript.exe """%~dp0PicFolderFilesToYearFolder.vbs""" """%%1""" %%*" /f

REG ADD "HKCU\Software\Classes\*\shell\LyxSoft\shell\LyxSoft.PicFileInfos" /f
REG ADD "HKCU\Software\Classes\*\shell\LyxSoft\shell\LyxSoft.PicFileInfos" /ve /d "Show Picture Informations" /f
REG ADD "HKCU\Software\Classes\*\shell\LyxSoft\shell\LyxSoft.PicFileInfos" /v "Icon" /d "%SystemRoot%\System32\SHELL32.dll,313" /f

REG ADD "HKCU\Software\Classes\*\shell\LyxSoft\shell\LyxSoft.PicFileInfos\Command" /f
REG ADD "HKCU\Software\Classes\*\shell\LyxSoft\shell\LyxSoft.PicFileInfos\Command" /ve /d "WScript.exe """%~dp0PicFileInfos.vbs""" """%%1""" %%*" /f

REG ADD "HKCU\Software\Classes\Folder\shell\LyxSoft" /f
REG ADD "HKCU\Software\Classes\Folder\shell\LyxSoft" /v "Icon" /d "%SystemRoot%\System32\SHELL32.dll,316" /f
REG ADD "HKCU\Software\Classes\Folder\shell\LyxSoft" /v "MUIVerb" /d "LyxSoft Tools" /f
REG ADD "HKCU\Software\Classes\Folder\shell\LyxSoft" /v "SubCommands" /d "" /f

REG ADD "HKCU\Software\Classes\Folder\shell\LyxSoft\shell\LyxSoft.PicFileDate" /f
REG ADD "HKCU\Software\Classes\Folder\shell\LyxSoft\shell\LyxSoft.PicFileDate" /ve /d "Set Picture Date" /f
REG ADD "HKCU\Software\Classes\Folder\shell\LyxSoft\shell\LyxSoft.PicFileDate" /v "Icon" /d "%SystemRoot%\System32\SHELL32.dll,313" /f

REG ADD "HKCU\Software\Classes\Folder\shell\LyxSoft\shell\LyxSoft.PicFileDate\Command" /f
REG ADD "HKCU\Software\Classes\Folder\shell\LyxSoft\shell\LyxSoft.PicFileDate\Command" /ve /d "WScript.exe """%~dp0PicFileDate.vbs""" """%%1""" %%*" /f

REG ADD "HKCU\Software\Classes\Folder\shell\LyxSoft\shell\LyxSoft.PicFileDateName" /f
REG ADD "HKCU\Software\Classes\Folder\shell\LyxSoft\shell\LyxSoft.PicFileDateName" /ve /d "Set Picture Name with Date" /f
REG ADD "HKCU\Software\Classes\Folder\shell\LyxSoft\shell\LyxSoft.PicFileDateName" /v "Icon" /d "%SystemRoot%\System32\SHELL32.dll,313" /f

REG ADD "HKCU\Software\Classes\Folder\shell\LyxSoft\shell\LyxSoft.PicFileDateName\Command" /f
REG ADD "HKCU\Software\Classes\Folder\shell\LyxSoft\shell\LyxSoft.PicFileDateName\Command" /ve /d "WScript.exe """%~dp0PicFileDateName.vbs""" """%%1""" %%*" /f

REG ADD "HKCU\Software\Classes\Folder\shell\LyxSoft\shell\LyxSoft.PicFolderFilesToYearFolder" /f
REG ADD "HKCU\Software\Classes\Folder\shell\LyxSoft\shell\LyxSoft.PicFolderFilesToYearFolder" /ve /d "Put File into Folder with Year of the File" /f
REG ADD "HKCU\Software\Classes\Folder\shell\LyxSoft\shell\LyxSoft.PicFolderFilesToYearFolder" /v "Icon" /d "%SystemRoot%\System32\SHELL32.dll,313" /f

REG ADD "HKCU\Software\Classes\Folder\shell\LyxSoft\shell\LyxSoft.PicFolderFilesToYearFolder\Command" /f
REG ADD "HKCU\Software\Classes\Folder\shell\LyxSoft\shell\LyxSoft.PicFolderFilesToYearFolder\Command" /ve /d "WScript.exe """%~dp0PicFolderFilesToYearFolder.vbs""" """%%1""" %%*" /f

CLS
ECHO Install done.
ECHO.
ECHO If you like to uninstall, delete the register key:
ECHO  HKCU\Software\Classes\*\shell\LyxSoft\shell\LyxSoft.PicFile*
ECHO.
ECHO Li Yingxin @ 2019.04
ECHO.
ECHO Press any key to finish the installation.
ECHO.
PAUSE
