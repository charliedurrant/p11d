Attribute VB_Name = "FileIO"
Option Explicit

Public Enum ITEM_STD
  ITEM_FRIENDLY = 1
  ITEM_PATH_AND_FILE
  
End Enum

Public Function FileOKStandard(sPathAndFile As String, sFileType As String) As Boolean
  FileOKStandard = InStr(1, sPathAndFile, sFileType, vbTextCompare) > 0
End Function
Public Function StandardFile(sPathAndFile As String, RF As IReadFile) As Boolean
  
  On Error GoTo StandardFile_END
  
  If RF.ReadFromFile Then GoTo StandardFile_END
  If Not FileExists(sPathAndFile) Then Call Err.Raise(ERR_FILE_NOT_EXIST, "StandardFile", "The file " & sPathAndFile & " does not exist.")
  If IsFileOpen(sPathAndFile, True) Then Call Err.Raise(ERR_FILE_OPEN_EXCLUSIVE, "StandardFile", "The file " & sPathAndFile & " is open exclusively.")
  If Not RF.FileOK(sPathAndFile) Then Call Err.Raise(ERR_FILE_WRONG_TYPE, "StandardFile", "The file does is not the correct type.")
  Set RF.FileRead = New TCSFileread
  Call RF.FileRead.OpenFile(sPathAndFile)
  Call RF.ReadFile
  RF.ReadFromFile = True

StandardFile_END:
  Exit Function
StandardFile_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "StandardFile", "Standard File", "Error standard read " & sPathAndFile)
  Resume StandardFile_ERR
End Function
