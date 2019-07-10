@echo off
IF EXIST \CONFIG.SYS goto errorlocal
rem I:\VERCTRL R:\VB5\LIBRARY\VBSTD.VC
IF X%1X == XX goto error
FOR %%i in (CORE DB PARSE IMP REP AUTO ALIGN VTEXT EFG TREE SIZE STAT COMBO QBE TWIST DBFLX DIR) do call copylib R:\VB5\LIBRARY\ %%i %1
goto finish
:errorlocal
@echo Error executing %0.BAT
@echo Can only be used to Copy from Network to local drive
@echo CD int R:\VB5\LIBRARY
pause
goto finish
:error
@echo usage %0 drive (e.g C:)
:finish


