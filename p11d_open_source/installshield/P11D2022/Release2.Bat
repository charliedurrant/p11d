@ECHO OFF
:Retry
CLS
ECHO %1
DIR /-P | FIND "Directory of "
ECHO.
ECHO Release Commands
ECHO.
ECHO.
ECHO 0. [Back]
ECHO.
ECHO 1. Recreate Working Folder                                     (WDirDel.Bat)
ECHO.
ECHO      (Get latest version of files from SourceSafe project)
ECHO.
ECHO 2. Set Working Folder attributes, Delete SourceSafe files      (WDirAttr.Bat)
ECHO 3. Copy and register COM components                            (CopyReg.Bat)
ECHO 4. Increase version number of version executable               (UpVer.Bat)
ECHO.
ECHO      (Build Release)
ECHO.
ECHO 5. Copy built installation image to Test directory             (TestCopy.Bat)
ECHO 6. Copy documentation files to Test directory                  (CopyDocs.Bat)
ECHO 7. Copy miscellaneous files to Test directory                  (CopyMisc.Bat)
REM ECHO 8. Copy Test directory to internal Test directory              (TestInt.Bat)
ECHO 9. Copy Test directories to composite Test directory           (TestCopyEx.Bat)
REM ECHO A. Copy Test directories to composite internal Test directory  (TestIntEx.Bat)
ECHO.
ECHO      (Test Release)
ECHO.
ECHO B. Copy Test directory/directories to Distribution directory   (Dist.Bat)
ECHO.
ECHO.
ECHO Please choose an option...
ECHO.
CHOICE /C:012345679B /N
ECHO.
IF ERRORLEVEL 2 GOTO NotBack
I:
CD \IShield\Batch
Release.Bat
:NotBack
IF ERRORLEVEL  2 SET Bat=WDirDel.Bat
IF ERRORLEVEL  3 SET Bat=WDirAttr.Bat
IF ERRORLEVEL  4 SET Bat=CopyReg.Bat
IF ERRORLEVEL  5 SET Bat=UpVer.Bat
IF ERRORLEVEL  6 SET Bat=TestCopy.Bat
IF ERRORLEVEL  7 SET Bat=CopyDocs.Bat
IF ERRORLEVEL  8 SET Bat=CopyMisc.Bat
IF ERRORLEVEL  9 SET Bat=TestInt.Bat
IF ERRORLEVEL 10 SET Bat=TestCopyEx.Bat
IF ERRORLEVEL 11 SET Bat=TestIntEx.Bat
IF ERRORLEVEL 12 SET Bat=Dist.Bat
IF NOT EXIST .\%Bat% GOTO NA
ECHO.
CALL %Bat%
ECHO.
PAUSE
GOTO Retry
:NA
ECHO.
ECHO Not Applicable.
ECHO.
PAUSE
GOTO Retry
