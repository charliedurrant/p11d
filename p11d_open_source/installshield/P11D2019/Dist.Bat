SET ISDir=I:\IShield\Setups\P11D2008

DELTREE /Y %ISDir%\*.*

XCOPY .\Test\*.* %ISDir%\. /E 

DIR %ISDir%

