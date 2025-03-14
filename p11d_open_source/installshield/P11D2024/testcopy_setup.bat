SET ISDir=.\Test_Setup

DELTREE /Y %ISDir%\*.*

XCOPY .\Release\Release_Setup\DISKIMAGES\DISK1\*.* %ISDir%\. /E 

DIR %ISDir%

