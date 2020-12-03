SET ISDir=I:\IShield\Setups\P11D2007_Setup

DELTREE /Y %ISDir%\*.*

XCOPY .\Test_Setup\*.* %ISDir%\. /E 

DIR %ISDir%

