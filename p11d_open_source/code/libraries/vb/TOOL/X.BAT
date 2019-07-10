@echo off
IF X%1X==XX goto error
IF exist %1.DLL goto ok
IF not exist %1.OCX goto error
:ok
SET WINSYS=C:\WINDOWS\SYSTEM
if X%windir%X == XC:\WINNTX SET WINSYS=C:\WINNT\SYSTEM32
if X%windir%X == Xc:\winntX SET WINSYS=C:\WINNT\SYSTEM32
if X%windir%X == Xc:\WINNTX SET WINSYS=C:\WINNT\SYSTEM32
if X%windir%X == XC:\winntX SET WINSYS=C:\WINNT\SYSTEM32
IF exist %1.DLL copy %1.DLL release > nul
IF exist %1.OCX copy %1.OCX release  > nul
IF exist %1.OCA copy %1.OCA release  > nul
IF exist %1.EXP copy %1.exp release  > nul
IF exist %1.LIB copy %1.lib release  > nul
IF exist %1.DEP copy %1.dep release  > nul
IF exist %1.DLL COPY %1.DLL %WINSYS%  > nul
IF exist %1.DEP COPY %1.DEP %WINSYS%  > nul
IF exist %1.OCX copy %1.OCX %WINSYS%  > nul
IF exist RELEASE\%1.DBG COPY RELEASE\%1.DBG %WINSYS%  > nul
IF exist RELEASE\FILEREAD.DLL COPY RELEASE\FILEREAD.DLL %WINSYS%  > nul
IF exist RELEASE\FILEREAD.DEP COPY RELEASE\FILEREAD.DEP %WINSYS%  > nul
IF exist RELEASE\FILEREAD.DLL %WINSYS%\REGSvr32 /u /s %WINSYS%\FILEREAD.DLL
IF exist RELEASE\FILEREAD.DLL %WINSYS%\REGSvr32 /s %WINSYS%\FILEREAD.DLL
IF exist RELEASE\TCSPROG.DEP COPY RELEASE\TCSPROG.DEP %WINSYS%   > nul
IF exist RELEASE\TCSPROG.OCX COPY RELEASE\TCSPROG.OCX %WINSYS%   > nul
IF exist RELEASE\TCSPROG.OCX %WINSYS%\REGSvr32 /u /s %WINSYS%\TCSPROG.OCX
IF exist RELEASE\TCSPROG.OCX %WINSYS%\REGSvr32 /s %WINSYS%\TCSPROG.OCX
IF exist %1.DLL %WINSYS%\REGSvr32 /u /s %WINSYS%\%1.dll
IF exist %1.OCX %WINSYS%\REGSvr32 /u /s %WINSYS%\%1.OCX
IF exist %1.DLL %WINSYS%\REGSvr32 /s %WINSYS%\%1.dll
IF exist %1.OCX %WINSYS%\REGSvr32 /s %WINSYS%\%1.OCX
@echo     Copied and registered %1
CD ..
goto finished
:error
@ECHO usage: %0 CoreDllName
:finished
