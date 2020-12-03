@ECHO OFF

ECHO.
DIR | FIND "Directory"

SET DistSrc=.
SET DistDest=.\..\MSM

ECHO.
ECHO Distribute non-abatec Merge Modules
ECHO   from: %DistSrc%
ECHO   to:   %DistDest%
ECHO.

CALL Dist2 atec3Party %DistSrc% %DistDest%
CALL Dist2 sqldmo %DistSrc% %DistDest%
REM CALL Dist2 SCRRUN %DistSrc% %DistDest%
REM CALL Dist2 MSSCRIPT %DistSrc% %DistDest%
REM CALL Dist2 VBSCRIPT %DistSrc% %DistDest%
REM CALL Dist2 MSVCRT40 %DistSrc% %DistDest%
REM CALL Dist2 TLBINF32 %DistSrc% %DistDest%
REM CALL Dist2 VB5DB %DistSrc% %DistDest%
REM CALL Dist2 TDBG5 %DistSrc% %DistDest%
REM CALL Dist2 TDBG6 %DistSrc% %DistDest%
REM CALL Dist2 DAO350 %DistSrc% %DistDest%
REM CALL Dist2 HHUPD %DistSrc% %DistDest%

ECHO.
