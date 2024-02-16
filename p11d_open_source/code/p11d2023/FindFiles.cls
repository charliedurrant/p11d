VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FindFiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_FileMask As String
Private m_IFF As IFindFile

Public bSearchCancelled As Boolean
Public lTotalNoOfFiles As Long

Public Function FindFiles(ByVal sStartDir As String, ByVal sFileMask As String, Optional IFF As IFindFile, Optional bSubDirectories As Boolean) As Long
  
  Dim FS As FileSystemObject
  Dim f As folder
  Dim lFileCount As Long

  
  On Error GoTo FindFiles_ERR
  If Len(sStartDir) = 0 Then sStartDir = p11d32.WorkingDirectory
  If Right$(sStartDir, 1) <> "\" Then sStartDir = sStartDir & "\"
  
  If Not FileExists(sStartDir, True) Then Call Err.Raise(ERR_FILE_NOT_EXIST, "FindFiles", "That Folder Cannot Be Found")
     
  Set m_IFF = IFF
  m_FileMask = sFileMask
  Set FS = New FileSystemObject
  Set f = FS.GetFolder(sStartDir)
  FindFiles = IterateFiles(f, m_FileMask, bSubDirectories)
  
FindFiles_END:
  Exit Function

FindFiles_ERR:
  If Err.Number = ERR_FILE_NOT_EXIST Then
    Call ErrorMessage(ERR_ERROR, Err, "FindFiles", "P11d - Find Files", Err.Description)
    FindFiles = L_FOLDER_INVALID
  End If
  Resume FindFiles_END
  End Function
  
Private Function IterateFiles(f As folder, sFileMask As String, Optional ByVal bSubDirectories As Boolean) As Long
  Dim sDirectory As String, sFile As String
  Dim File As File
  Dim lNoOfFiles As Long
  Dim sExt, sMain As String
  Dim fdr As folder
  Dim sFileExtensionCurrent As String
  Dim lEmployees As Long
  Dim sPAYEref As String
  
  On Error GoTo IterateFiles_ERR
  
  If Not FileExists(f, True) Then Err.Raise ERR_FILE_NOT_EXIST, "IterateFiles", sFile & "does not exist"
  
  Call SetCursor(vbArrowHourglass)
  sDirectory = f.Path
  
  Call m_IFF.Prenotify(sDirectory, CountFiles(sDirectory))

  sFileExtensionCurrent = p11d32.Rates.FileExtensionCurrent
  
  For Each File In f.Files
    If bSearchCancelled Then Exit For
      lTotalNoOfFiles = lTotalNoOfFiles + 1
      sFile = File.Name
      Call m_IFF.Notify(lTotalNoOfFiles)
      Call SplitPath(File.Name, , sMain, sExt)
        If Len(sMain) > 2 And Len(sMain) <= 8 Then
          If StrComp(Right$(File.Name, 6), sFileExtensionCurrent, vbTextCompare) = 0 Then
            If IsP11DDatabase(sDirectory & "\" & File.Name) Then
              Call m_IFF.Process(sDirectory, sFile, lNoOfFiles + 1, lEmployees, sPAYEref)
              lNoOfFiles = lNoOfFiles + 1
            End If
          End If
        End If
     Next
  
  Call m_IFF.PostNotify
  
  If bSubDirectories Then
    For Each fdr In f.SubFolders
      If bSearchCancelled Then Exit For
      lNoOfFiles = lNoOfFiles + IterateFiles(fdr, sFileMask, bSubDirectories)
    Next
  End If
  
IterateFiles_END:
  IterateFiles = lNoOfFiles
  Call ClearCursor
  Exit Function

IterateFiles_ERR:
  Resume IterateFiles_END
  
  End Function

Private Function IsP11DDatabase(FileName As String) As Boolean
  Dim empr As Employer
  Set empr = New Employer
    
  On Error GoTo IsP11dDatabase_Err
      
  IsP11DDatabase = empr.Validate(FileName)
    
  Set empr = Nothing
    
IsP11dDatabase_End:
  Exit Function

IsP11dDatabase_Err:
  Call ErrorMessage(ERR_ERROR, Err, "IsP11dDatabase", "Is P11D File", "Error verifying if file is P11D")
  Resume IsP11dDatabase_End
End Function

