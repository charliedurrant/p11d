@ECHO OFF
ECHO Distributing Blank Merge Modules to I:\IShield\MSM.Blk
COPY I:\IShield\MergeM\Blank\Release\Release\DISKIMAGES\DISK1\Blank.Msm I:\IShield\MSM\.
MD I:\IShield\MSM.Blk
I:
CD \IShield\MSM
FOR %%f IN (*.MSM) DO COPY I:\IShield\MSM\Blank.Msm I:\IShield\MSM.Blk\%%f
CD \IShield\MergeM
