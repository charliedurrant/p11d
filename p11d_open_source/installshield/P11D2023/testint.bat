SET ISDir=I:\Tax\Setups\P11D2008

DELTREE /Y %ISDir%\*.*

XCOPY .\Internal\*.* %ISDir%\. /E

XCOPY ".\Test\P11D2007.msi" %ISDir%\.
XCOPY .\Test\Data1.Cab %ISDir%\.
XCOPY .\Test\0x0409.ini %ISDir%\.

DIR %ISDir%

