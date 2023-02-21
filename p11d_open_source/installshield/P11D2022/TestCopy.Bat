SET ISDir=.\Test

REM DELTREE /Y %ISDir%\*.*

REM XCOPY .\Release\Release\DISKIMAGES\DISK1\*.* %ISDir%\. /E 

REM DIR %ISDir%


SET ISDir=.\Test

DELTREE /Y %ISDir%\*.*


XCOPY .\Release\Release_Setup\DISKIMAGES\DISK1\*.* %ISDir%\. /E

XCOPY .\Release\Release_Setup\DISKIMAGES\DISK1\setup.exe "j:\ATMC FTP\P11D\P11d2008.exe" 

wscript.exe dlls
XCOPY c:\dlls.htm "j:\TMC FTP\P11D\p11d2008.dlls.htm"


DIR %ISDir%
