@ECHO OFF

O:
CD O:\P11D2003

REM I:
REM CD I:\Tax\Setups\P11D2003

SET MSG_SERVER=http://options.uk.deloitte.com/install_log/logmsg.asp
SET DestDir=C:\Program Files\Deloitte
SET WINSYS=C:\WINDOWS\SYSTEM
if X%windir%X == XC:\WINNTX SET WINSYS=C:\WINNT\SYSTEM32
if X%windir%X == Xc:\winntX SET WINSYS=C:\WINNT\SYSTEM32
if X%windir%X == Xc:\WINNTX SET WINSYS=C:\WINNT\SYSTEM32
if X%windir%X == XC:\winntX SET WINSYS=C:\WINNT\SYSTEM32

MD C:\Data
IF EXIST C:\Data\StartTmp.Bat ATTRIB -R C:\Data\StartTmp.Bat
IF EXIST C:\Data\StartTmp.Bat DEL C:\Data\StartTmp.Bat
IF X%WINSYS%X==XC:\WINDOWS\SYSTEMX start.exe /wait msiRemove.Exe /x P11D2003
IF NOT X%WINSYS%X==XC:\WINDOWS\SYSTEMX start /wait msiRemove.Exe /x P11D2003
CALL C:\Data\StartTmp.Bat

logmsg %MSG_SERVER% START_INSTALL_P11D2003

REM GOTO finished

MD "%DestDir%"
SET DestDir=%DestDir%\P11D2003
MD "%DestDir%"
COPY .\P11D2003_Exe.Exe "%DestDir%\."
COPY .\P11D2003_Exe.Lod "%DestDir%\."
IF X%WINSYS%X==XC:\WINDOWS\SYSTEMX     MD "C:\Windows\Start Menu\Programs\Deloitte"
IF NOT X%WINSYS%X==XC:\WINDOWS\SYSTEMX MD "C:\Documents and Settings\All Users\Start Menu\Programs\Deloitte"
IF X%WINSYS%X==XC:\WINDOWS\SYSTEMX     COPY ".\P11D2003.lnk" "C:\Windows\Start Menu\Programs\Deloitte\."
IF NOT X%WINSYS%X==XC:\WINDOWS\SYSTEMX COPY ".\P11D2003.lnk" "C:\Documents and Settings\All Users\Start Menu\Programs\Deloitte\."

logmsg %MSG_SERVER% P11D2003_SUCCESS
"%DestDir%\P11D2003_Exe.Exe"

:finished
REM EXIT
