@ECHO OFF

ECHO.
DIR | FIND "Directory"

SET DistSrc=.
SET DistDest=.\..\MSM

ECHO.
ECHO Distribute abatec Merge Modules
ECHO   from: %DistSrc%
ECHO   to:   %DistDest%
ECHO.

REM CALL Dist2 atecDesk %DistSrc% %DistDest%
REM CALL Dist2 atecDesk2 %DistSrc% %DistDest%
CALL Dist2 atecDesk3 %DistSrc% %DistDest%
REM CALL Dist2 atecWeb %DistSrc% %DistDest%
CALL Dist2 atecWeb2 %DistSrc% %DistDest%
CALL Dist2 atecIDSMail %DistSrc% %DistDest%
CALL Dist2 AORulePacks %DistSrc% %DistDest%
REM CALL Dist2 atecDAO %DistSrc% %DistDest%
REM CALL Dist2 atecADO %DistSrc% %DistDest%
REM CALL Dist2 TCSBASE %DistSrc% %DistDest%
REM CALL Dist2 TCSCCORE %DistSrc% %DistDest%
REM CALL Dist2 TCSCORE %DistSrc% %DistDest%
REM CALL Dist2 TCSDB %DistSrc% %DistDest%
REM CALL Dist2 TCSPARSE %DistSrc% %DistDest%
REM CALL Dist2 TCSREP %DistSrc% %DistDest%
REM CALL Dist2 TCSPROG %DistSrc% %DistDest%
REM CALL Dist2 TCSIMP %DistSrc% %DistDest%
REM CALL Dist2 TCSHOOK %DistSrc% %DistDest%
REM CALL Dist2 TCSQBE %DistSrc% %DistDest%
REM CALL Dist2 TCSSIZE %DistSrc% %DistDest%
REM CALL Dist2 TCSACTRL %DistSrc% %DistDest%
REM CALL Dist2 TCSAUTO %DistSrc% %DistDest%
REM CALL Dist2 TCSSTAT %DistSrc% %DistDest%
REM CALL Dist2 TCSTWIST %DistSrc% %DistDest%
REM CALL Dist2 TCSSPLIT %DistSrc% %DistDest%
REM CALL Dist2 TCSBG %DistSrc% %DistDest%
REM CALL Dist2 TCSRFILE %DistSrc% %DistDest%
REM CALL Dist2 TCSWHERE %DistSrc% %DistDest%
REM CALL Dist2 TCSALLOC %DistSrc% %DistDest%
REM CALL Dist2 TCSCHECK %DistSrc% %DistDest%
REM CALL Dist2 TCSREPW %DistSrc% %DistDest%
REM CALL Dist2 TCSSQL %DistSrc% %DistDest%
REM CALL Dist2 TCSTOOL %DistSrc% %DistDest%

ECHO.

DIR %DistDest%\atec*.MSM
DIR %DistDest%\aorulepacks.MSM

