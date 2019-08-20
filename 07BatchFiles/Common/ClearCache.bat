@ECHO OFF 
Echo.
Echo ClearCache 1.1 (2004) for Windows 2000 and XP
Echo Removes temporary files FROM ALL PROFILES.
Echo.
Echo Authors: by Cher Aza, Jimbob60, and Mirfster.
Echo.
Echo Temporary files will be removed from the following locations:
Echo.
Echo   1. [Profile]\Local Settings\Temporary Internet Files
Echo   2. [Profile]\Local Settings\History
Echo   3. [Profile]\Local Settings\Temp
Echo   4. [Profile]\Cookies
Echo   5. [Profile]\Recent
Echo   6. [SystemRoot]\Temp
Echo.
rem Echo This batch file will close automatically when finished.
rem Echo.
rem Echo Press Ctrl+C to abort or any other key to continue . . .
rem Echo.
rem PAUSE > NUL
:: Clear local Temporary Internet Files and History
rem Echo.
Echo The following errors are normal:
RD /S /Q "%Userprofile%\Local Settings\Temporary Internet Files"
RD /S /Q "%Userprofile%\Local Settings\History"
:: Clear all other temporary cache files in all profiles
SET SRC1=C:\Documents and Settings
SET SRC2=Local Settings\Temporary Internet Files
SET SRC3=Local Settings\History
SET SRC4=Local Settings\Temp
SET SRC5=Cookies
SET SRC6=Recent
Echo.
Echo About to delete files from Internet Explorer "Temporary Internet files"
Echo This may take a few minutes.  Please wait...
Echo The following error is normal:
FOR /D %%X IN ("%SRC1%\*") DO FOR /D %%Y IN ("%%X\%SRC2%\*.*") DO RMDIR /S /Q "%%Y"
Echo.
Echo About to delete files from Internet Explorer "History"
Echo This may take a few minutes.  Please wait...
Echo The following error is normal:
FOR /D %%X IN ("%SRC1%\*") DO FOR /D %%Y IN ("%%X\%SRC3%\*.*") DO RMDIR /S /Q "%%Y"
FOR /D %%X IN ("%SRC1%\*") DO FOR  %%Y IN ("%%X\%SRC3%\*.*") DO DEL /F /S /Q "%%Y"
Echo About to delete files from "Local settings\temp"
FOR /D %%X IN ("%SRC1%\*") DO FOR /D %%Y IN ("%%X\%SRC4%\*.*") DO RMDIR  /S /Q "%%Y"
FOR /D %%X IN ("%SRC1%\*") DO FOR  %%Y IN ("%%X\%SRC4%\*.*") DO DEL /F /S /Q "%%Y"
Echo About to delete files from "Cookies"
FOR /D %%X IN ("%SRC1%\*") DO FOR  %%Y IN ("%%X\%SRC5%\*.*") DO DEL /F /S /Q "%%Y"
Echo About to delete files from "Recent" i.e. what appears in Start/Documents/My Documents
FOR /D %%X IN ("%SRC1%\*") DO FOR  %%Y IN ("%%X\%SRC6%\*.lnk") DO DEL /F /S /Q "%%Y"
Echo About to delete files from "[SystemRoot]\Temp"
CD /D %SystemRoot%\Temp
DEL /F /Q *.*
RD /S /Q "%SystemRoot%\Temp"
CD /D %SystemRoot%
MD Temp
:: Done!
Echo Done!
rem EXIT