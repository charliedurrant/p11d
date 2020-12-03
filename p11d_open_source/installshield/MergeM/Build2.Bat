SET ISCmdBldPath="C:\Program Files\InstallShield\Professional - Windows Installer Edition\System\ISCmdBld.exe"
ECHO Building Merge Module %1
ECHO Building Merge Module %1 from %2\%1.ism >> BuildRep.Txt
ECHO . >> BuildRep.Txt
ECHO %ISCmdBldPath% -p "%2\%1\%1.ism" -d "%1" -b "%2\%1" -r "Release" -a "Release" >> BuildRep.Txt
ECHO . >> BuildRep.Txt
%ISCmdBldPath% -p "%2\%1\%1.ism" -d "%1" -b "%2\%1" -r "Release" -a "Release" >> BuildRep.Txt
ECHO . >> BuildRep.Txt
ECHO Completed build of Merge Module %1 from %2\%1.ism >> BuildRep.Txt
ECHO . >> BuildRep.Txt
ECHO . >> BuildRep.Txt
