@echo off
IF EXIST \CONFIG.SYS goto errorlocal
R:
CD \VB5\LIBRARY
rem I:\VERCTRL R:\VB5\LIBRARY\VBSTD.VC
FOR %%i in (CORE DB PARSE IMP REP AUTO ALIGN VTEXT EFG TREE SIZE STAT COMBO QBE TWIST DIR UDM) do call copylib R:\VB5\LIBRARY\ %%i C:
goto finish
:errorlocal
@echo Error executing %0.BAT
@echo Can only be used to Copy from Network to local drive
@echo CD int R:\VB5\LIBRARY
pause
:finish

