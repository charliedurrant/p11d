VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FileHelper"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'*  Return Disk / Memory sizes in the following
'*  BYTES
'*  MEGABYTES
'*  GIGABYTES
Public Enum MemoryUnit
  BYTES
  MEGABYTES
  GIGABYTES
End Enum
Private Declare Function GetFullPathName Lib "kernel32" Alias "GetFullPathNameA" (ByVal lpFileName As String, ByVal nBufferLength As Long, ByVal lpBuffer As String, ByVal lpFilePart As String) As Long
Private Declare Function WNetGetConnection Lib "mpr.dll" Alias "WNetGetConnectionA" (ByVal lpszLocalName As String, ByVal lpszRemoteName As String, cbRemoteName As Long) As Long
'Public Declare Function WritePrivateProfileSection Lib "kernel32" Alias "WritePrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpString As String, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Private Declare Function FlushPrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As Long, ByVal lpKeyName As Long, ByVal lpString As Long, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function GetProfileString Lib "kernel32" Alias "GetProfileStringA" (ByVal lpAppName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long) As Long

Public Sub SplitPath(sFullPath As String, Optional sDir As Variant, Optional sFile As Variant, Optional sExt As Variant)
  Dim p As Long, q As Long, tmp As Long
  
  On Error GoTo SplitPath_err
  p = InStrRev(sFullPath, "\")
  If p = 0 Then p = InStrRev(sFullPath, ":")
  
  If (Not IsMissing(sDir)) And (p > 0) Then
    sDir = Left$(sFullPath, p)
  End If
  
  q = InStrRev(sFullPath, ".")
  If Not IsMissing(sFile) Then
    p = p + 1
    If q > 0 Then
      tmp = q - p
    Else
      tmp = Len(sFullPath)
    End If
    sFile = Mid$(sFullPath, p, tmp)
  End If
  
  If Not IsMissing(sExt) And (q > 0) Then
    q = Len(sFullPath) - q + 1
    sExt = Right$(sFullPath, q)
  End If
  Exit Sub
  
SplitPath_err:
  Err.Raise Err.Number, ErrorSourceEx(Err, "SplitPath"), Err.Description
End Sub

'##D Makes Directory recursively
'##V Directory Name of Directory
'##RV True if directory created or already present
'##C FNF
Public Function xMkdir(ByVal Directory As String) As Boolean
  Dim DirString As String
  Dim p0 As Long, p1 As Long
  
  On Error GoTo xMkdir_err
  If Left$(Directory, 2) = "\\" Then
    p0 = 3
    p0 = InStr(p0, Directory, "\")
    If p0 = 0 Then Err.Raise ERR_MKDIR, "xMkdir", "Unable to create directory [" & Directory & "]"
    p0 = InStr(p0 + 1, Directory, "\")
    If p0 = 0 Then Err.Raise ERR_MKDIR, "xMkdir", "Unable to create directory [" & Directory & "]"
    p0 = p0 + 1
  Else
    p0 = 4
  End If
  Do
    p1 = InStr(p0, Directory, "\")
    If p1 = 0 Then
      DirString = Mid$(Directory, 1)
    Else
      DirString = Mid$(Directory, 1, p1 - 1)
      p0 = p1 + 1
    End If
    Call MkDir(DirString)
  Loop Until p1 = 0
  xMkdir = True
  
xMkdir_end:
  Exit Function
  
xMkdir_err:
  xMkdir = False
  If Err.Number = 75 Then Resume Next
  Resume xMkdir_end
End Function

'##D Get DOS Path of a Directory/Path
'##RV The DOS path (includes driver letter)
'##C FNF
Public Function GetDOSPath(ByVal DirectoryOnly As String) As String
  GetDOSPath = GetDrivePathEx(DirectoryOnly)
End Function

Public Function GetDOSFullPath(ByVal FullPath As String) As String
  Dim sDir As String, sFile As String, sExt As String
  
  Call SplitPath(FullPath, sDir, sFile, sExt)
  GetDOSFullPath = GetDrivePathEx(sDir) & sFile & sExt
End Function

'##D xKill wraps Kill function - erases file
'##V FullPath Full Path of file to be erased
'##RV False if failed to kill files, no error is raised if no files erased due to no files matching mask
'##C FNF
Public Function xKill(ByVal FullPath As String) As Boolean
  On Error GoTo xKill_err
  Call kill(FullPath)
  xKill = True
  Exit Function
  
xKill_err:
  If Err.Number = 53 Then xKill = True
End Function

'##D Count the number of files in directory, does not test for existence of path
'##V FullPath the path of the directory to be searched
'##RV Number of files
'##C FNF
Public Function CountFiles(ByVal FilePath As String) As Long
  Dim sFile As String
  
  On Error GoTo CountFiles_err
  If FileExists(FilePath, True, False) Then FilePath = FullPath(FilePath)
  sFile = Dir$(FilePath)
  Do While Len(sFile) > 0
    CountFiles = CountFiles + 1
    sFile = Dir
  Loop
  Exit Function
  
CountFiles_err:
  Err.Raise Err.Number, ErrorSourceEx(Err, "CountFiles"), Err.Description
End Function

Public Function FindFiles(ByVal FileDirectory As String, ByVal FileMask As String, Optional ByVal SubDirs As Boolean = False) As StringList
  Dim sFile As String, sFilePath As String
  Dim i As Long
  Dim sFiles As StringList, sDirs As StringList
  
  On Error GoTo FindFiles_err
  If Not FileExists(FileDirectory, True, False) Then Err.Raise ERR_FINDFILES, "FindFiles", "Directory [" & FileDirectory & "] does not exist"
  FileDirectory = FullPath(FileDirectory)
  Set sDirs = New StringList
  Set sFiles = New StringList
  Call sDirs.Add(FileDirectory)
  
process_dirs:
  For i = 1 To sDirs.Count
    FileDirectory = sDirs.Item(i)
    ' files first
    sFile = Dir$(FileDirectory & FileMask)
    Do While Len(sFile) > 0
      sFilePath = FileDirectory & sFile
      If (GetAttr(sFilePath) And vbDirectory) <> vbDirectory Then Call sFiles.Add(sFilePath)
      sFile = Dir$
    Loop
    
    If SubDirs Then
      ' directories next
      sFile = Dir$(FileDirectory & "*.*", vbDirectory)
      Do While Len(sFile) > 0
        If Not ((StrComp(sFile, ".", vbBinaryCompare) = 0) Or (StrComp(sFile, "..", vbBinaryCompare) = 0)) Then
          sFilePath = FileDirectory & sFile
          If (GetAttr(sFilePath) And vbDirectory) = vbDirectory Then Call sDirs.Add(FullPath(sFilePath))
        End If
        sFile = Dir$
      Loop
      Call sDirs.Remove(FileDirectory)
      GoTo process_dirs
    End If
  Next i
  Set FindFiles = sFiles
  Exit Function
  
FindFiles_err:
  Err.Raise Err.Number, ErrorSourceEx(Err, "FindFiles"), Err.Description
End Function

'##D Determines if a database is open
'##C FNF
Public Function IsDatabaseOpen(ByVal FileName As String, Optional ByVal Exclusive As Boolean = False) As Boolean
  IsDatabaseOpen = IsFileOpenEx(FileName, Exclusive)
End Function

'##D Determines if a file is already open
'##V FileName File Name
'##RV True if file is open
'##C FNF
Public Function IsFileOpen(ByVal FileName As String, Optional ByVal Exclusive As Boolean = False) As Boolean
  IsFileOpen = IsFileOpenEx(FileName, Exclusive)
End Function

'##D Copies file to another location
'##LD This function replaces the normal FileCopy in VB, the function implements file copy as a copy and rename
'##LD and takes account of most long filename issues
'##V Source file
'##V Destination file
'##RV True if copy successful
'##C FNF
Public Function FileCopyEx(ByVal Source As String, ByVal Destination As String) As Boolean
  FileCopyEx = FileCopyExN(Source, Destination, False)
End Function

'##D Copies file to another location
'##LD This function replaces the normal FileCopy in VB, the function implements file copy as a copy and rename
'##LD and takes account of most long filename issues.
'##LD Raises an error if Copy Fails
'##V Source file
'##V Destination file
'##C FNF
Public Sub FileCopyEx2(ByVal Source As String, ByVal Destination As String)
  Call FileCopyExN(Source, Destination, True)
End Sub


'##D Gives (full) network resource name
'##V NetResourceName String where result is placed
'##V LocalResourceName Drive letter (1..27) or printer port
'##V ResourceNameOnly See above default, false
'##V ShowErrors Show error message on error
'##LD Given a local resource name ( e.g G:,H: or LPT1: ) and return either the full network
'##LD resource name ( e.g for files \\server\volume\path ) or just the network resource
'##LD name (server or printer)
'##RV 0 if succeeds, error code if failure (see DLLErrors enumeration)
'##C SQF
Public Function GetNetResource(NetResourceName As String, ByVal LocalResourceName As String, Optional ByVal ResourceNameOnly As Boolean = False, Optional ByVal ShowErrors As Boolean = False) As Long
  Dim sRes As String, p As Long
  
  On Error GoTo GetNetResource_err
  NetResourceName = ""
  sRes = String$(TCSBUFSIZ, 0)
  GetNetResource = WNetGetConnection(UCase$(LocalResourceName), sRes, TCSBUFSIZ)
  If GetNetResource = 0 Then
    sRes = RTrimChar(sRes, vbNullChar)
    NetResourceName = sRes
    ' note UNC of type \\xxxx
    If ResourceNameOnly And (Len(NetResourceName) > 2) Then
      p = InStr(3, NetResourceName, "\")
      If p > 0 Then
        NetResourceName = Mid$(NetResourceName, 3, p - 3)
      End If
    End If
    NetResourceName = sRes
  End If
  
GetNetResource_end:
  Exit Function
  
GetNetResource_err:
  If ShowErrors Then Err.Raise Err.Number, ErrorSourceEx(Err, "GetNetResource"), "Network Resource Name", "Unable to get the network resource name for the resource [" & LocalResourceName & "]" & vbCrLf & Err.Description
  Resume GetNetResource_end
End Function

'##D GetCanonicalPathName return full UNC pathname
'##V FilePath File Path
'##C FNF
Public Function GetCanonicalPathName(ByVal FilePath As String) As String
  Dim sBuffer As String, sFileName As String
  Dim uncpath As String
  Dim lBufSiz As Long, retval As Long
      
  On Error GoTo GetCanonicalPathName_err
  lBufSiz = 1024
  sBuffer = String$(lBufSiz, 0)
  sFileName = String$(lBufSiz, 0)
  retval = GetFullPathName(FilePath, lBufSiz, sBuffer, sFileName)
  If retval <> 0 Then
    GetCanonicalPathName = RTrimChar(sBuffer, vbNullChar)
    If Left$(GetCanonicalPathName, 2) <> "\\" Then
      If GetNetResource(uncpath, Left$(GetCanonicalPathName, 2)) = 0 Then
        GetCanonicalPathName = UCase$(uncpath & Right$(GetCanonicalPathName, Len(GetCanonicalPathName) - 2))
      End If
    End If
  End If
  
GetCanonicalPathName_end:
  Exit Function
  
GetCanonicalPathName_err:
  GetCanonicalPathName = ""
  Resume GetCanonicalPathName_end
End Function

'##D Get the free and total disk space of the drive specified
'##V DriveRootPath Root Path of drive
'##V Total Total space on this drive
'##V Free Free space on drive
'##RV true if successful
'##C SQF
Public Function GetDiskSpace(ByVal DriveRootPath As String, dTotal As Double, dFree As Double, Optional ByVal MemUnits As MemoryUnit = BYTES) As Boolean
  Dim dUserFree As Double
  
  GetDiskSpace = GetDiskSpaceEx(DriveRootPath, dTotal, dUserFree, dFree, MemUnits)
End Function


'##D get the Password for a specific date
'##C TCSPF
Public Function GetPassword(Optional ByVal d0 As Date = UNDATED) As Long
  GetPassword = GetPassword_Daily()
End Function


'##D Write an entry to the ini file specified or of the current application ([AppExeName].INI) if none is specified
'##V section         Ini file section name without brackets
'##V Key             Ini file key name
'##V Value           New value to be written
'##V IniFilePath    Path and file name of ini file (Default = [Windows directory]\[APPEXENAME].INI)
'##RV True if success
'##C IFF
Public Function WriteIniEntry(ByVal Section As String, ByVal Key As String, ByVal Value As String, ByVal IniFilePath As String) As Boolean
  Dim l As Long
  
  If Len(IniFilePath) = 0 Then Err.Raise ERR_FILEHELPER, "WriteIniEntry", "No ini file specified"
  l = WritePrivateProfileString(Section, Key, Value, IniFilePath)
  WriteIniEntry = (l <> 0)
End Function

'##D Read an entry from the ini file specified or of the current application ([AppExeName].INI)
'##V Section             section in ini file
'##V Key                 key
'##V Default               default value for the key (Default = "")
'##V inifilepath Path and filename of ini file (Default = [Windows directory]\[APPEXENAME].INI)
'##RV The value for the key in the ini file, or Default if the Key is not present
'##C IFF
Public Function GetIniEntry(ByVal Section As String, ByVal Key As String, ByVal Default As String, ByVal IniFilePath As String) As String
  Dim sBuffer As String, bsize As Long
  Dim retval As Long
  
  If Len(IniFilePath) = 0 Then Err.Raise ERR_FILEHELPER, "GetIniEntry", "No ini file specified"
  bsize = TCSBUFSIZ
  Do
    bsize = bsize * 2
    sBuffer = String$(bsize, 0)
    retval = GetPrivateProfileString(Section, Key, "", sBuffer, bsize, IniFilePath)
  Loop Until (retval = 0) Or (retval <> (bsize - 1))
  If retval = 0 Then
    sBuffer = Default
  Else
    sBuffer = RTrimChar(sBuffer, vbNullChar)
  End If
  GetIniEntry = sBuffer
End Function

'##D Returns an array KeyNames of all keys in a particular Section in an ini file
'##C IFF
Public Function GetIniKeyNames(KeyNames As Variant, ByVal Section As String, ByVal IniFilePath As String) As Long
  Dim TempKeys() As String
  Dim sBuffer As String, bsize As Long
  Dim p0 As Long, p1 As Long, MaxKey As Long, retval As Long
  
  If Len(IniFilePath) = 0 Then Err.Raise ERR_FILEHELPER, "GetIniKeyNames", "No ini file specified"
  bsize = TCSBUFSIZ
  Do
    bsize = bsize * 2
    sBuffer = String$(bsize, 0)
    retval = GetPrivateProfileString(Section, 0&, "", sBuffer, bsize, IniFilePath)
  Loop Until (retval = 0) Or (retval <> (bsize - 2))
  If retval > 0 Then
    MaxKey = 0: p0 = 1
    Do
      p1 = InStr(p0, sBuffer, vbNullChar)
      MaxKey = MaxKey + 1
      ReDim Preserve TempKeys(1 To MaxKey)
      TempKeys(MaxKey) = Mid$(sBuffer, p0, p1 - p0)
      p0 = p1 + 1
    Loop Until Mid$(sBuffer, p0, 1) = vbNullChar
  End If
  KeyNames = TempKeys
  GetIniKeyNames = MaxKey
End Function

'##D Clears ini file entry
'##LD If no Key specified, clears complete Section otherwise clears that Key
'##C IFF
Public Function ClearIniEntry(ByVal Section As String, Optional ByVal Key As Variant, Optional ByVal Value As Variant, Optional ByVal IniFilePath As String) As Boolean
  Dim s As String
  Dim l As Long
  
  If Len(IniFilePath) = 0 Then Err.Raise ERR_FILEHELPER, "ClearIniEntry", "No ini file specified"
  If IsMissing(Key) Then
    l = WritePrivateProfileString(Section, vbNullString, vbNullString, IniFilePath)
  ElseIf IsMissing(Value) Then
    s = CStr(Key)
    l = WritePrivateProfileString(Section, s, vbNullString, IniFilePath)
  End If
  ClearIniEntry = (l <> 0)
End Function

'##D Similar to GetIniEntry, but acts on WIN.INI
'##C IFF
Public Function GetWinIniEntry(ByVal Section As String, ByVal Key As String, Optional ByVal Default As String = "") As String
  Dim sFname As String
  Dim sBuffer As String, bsize As Long
  Dim retval As Long
  
  bsize = TCSBUFSIZ
  Do
    bsize = bsize * 2
    sBuffer = String$(bsize, 0)
    retval = GetProfileString(Section, Key, Default, sBuffer, bsize)
  Loop Until (retval = 0) Or (retval <> (bsize - 2))
  If retval = 0 Then
    sBuffer = ""
  Else
    sBuffer = RTrimChar(sBuffer, vbNullChar)
  End If
  GetWinIniEntry = sBuffer
End Function

'##D Flush the ini buffer
'##LD Ini files are cached in memory this forces the file to write to disk
'##V IniFilePath  Name of ini file to flush default = windows directory\exename.ini
'##RV True if success
'##C IFF
Public Function FlushIniBuffer(ByVal IniFilePath As String) As Boolean
  Dim l As Long
  
  If Len(IniFilePath) = 0 Then Err.Raise ERR_FILEHELPER, "FlushIniBuffer", "No ini file specified"
  l = FlushPrivateProfileString(0, 0, 0, IniFilePath)
  FlushIniBuffer = (l <> 0)
End Function



