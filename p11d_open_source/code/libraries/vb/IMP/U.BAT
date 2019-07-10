@echo off
if X%windir%X == XC:\WINNTX goto winnt
:windows
SET WINSYS=C:\WINDOWS\SYSTEM
goto common
:winnt
SET WINSYS=C:\WINNT\SYSTEM32
:common
for %%i in (*.*) do copy %WINSYS%\%%i . > nul
@echo finished

