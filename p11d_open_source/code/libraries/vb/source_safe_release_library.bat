

set SSDIR=j:\projects\sourcesafe\libraries
ss cp $/release/vb/library -Ychdurrant
ss delete %1
ss add %1 -R -C-

REM  unpin
REM
REM ss unpin *.*

REM label, change to the parent folder ie the folder above the project
REM ss label $/Release

REM PIN
REM ss pin -VLverseionnumberofdlll
