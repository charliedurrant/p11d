@ECHO OFF

ECHO.
ECHO Copying P11D2008 Installation Documents ...
ECHO.

SET SrcDir=.\Docs

SET DestDir=.\Test\Docs

MD %DestDir%

XCOPY %SrcDir%\*.* %DestDir%\. /E /Y
ECHO Copied installation notes

ECHO.
ECHO Done
ECHO.
