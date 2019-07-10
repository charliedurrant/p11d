REM batch file called from local.BAT
IF X%1X==XX goto error
IF X%2X==XX goto error
IF X%3X==XX goto error
IF exist %1%2\TCS%2.DLL goto skip_check
IF not exist %1%2\TCS%2.OCX goto error
:skip_check
if X%windir%X == XC:\WINNTX goto winnt
:windows
SET WINSYS=C:\WINDOWS\SYSTEM
DELTREE /Y %3\VB5\LIBRARY\%2
goto common
:winnt
SET WINSYS=C:\WINNT\SYSTEM32
RMDIR /S /Q %3\VB5\LIBRARY\%2 > NUL
:common
IF exist %3\VB5\LIBRARY\%2\TCS%2.DLL goto skip_mkdir
IF exist %3\VB5\LIBRARY\%2\TCS%2.OCX goto skip_mkdir
MD %3\VB5
MD %3\VB5\LIBRARY
MD %3\VB5\LIBRARY\%2
:skip_mkdir
@echo Copying TCS%2
XCOPY %1%2\*.* %3\VB5\LIBRARY\%2 /d /s > nul
IF exist %1%2\RELEASE\FILEREAD.DEP XCOPY %1%2\RELEASE\FILEREAD.DEP %WINSYS%  /d > nul
IF exist %1%2\RELEASE\FILEREAD.DLL XCOPY %1%2\RELEASE\FILEREAD.DLL %WINSYS%  /d > nul
IF exist %1%2\RELEASE\FILEREAD.DLL %WINSYS%\REGSvr32 /s %WINSYS%\FILEREAD.DLL
IF exist %1%2\RELEASE\TCS%2.DEP    XCOPY %1%2\RELEASE\TCS%2.DEP %WINSYS%  /d > nul
IF exist %1%2\RELEASE\TCS%2.DLL  goto dllcopyregister
IF exist %1%2\RELEASE\TCS%2.OCX  goto ocxcopyregister
goto finished

:ocxcopyregister
@echo Registering OCX TCS%2
XCOPY %1%2\RELEASE\TCS%2.OCX %WINSYS%  /d > nul
%WINSYS%\REGSvr32 /s %WINSYS%\TCS%2.OCX
goto finished

:dllcopyregister
@echo Registering DLL TCS%2
XCOPY %1%2\RELEASE\TCS%2.DLL %WINSYS%  /d > nul
%WINSYS%\REGSvr32 /s %WINSYS%\TCS%2.DLL
goto finished

:error
@ECHO usage: %0 Dir LibraryName Drive(%1 %2 %3)
:finished


