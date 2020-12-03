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

CALL Dist2 AORulePacks %DistSrc% %DistDest%

ECHO.

DIR %DistDest%\aorulepacks.MSM

