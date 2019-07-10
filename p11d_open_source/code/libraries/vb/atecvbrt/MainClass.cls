VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SystemHelper"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Enum COM_TYPE
  WIN32_INPROC = 1
  WIN32_SERVERPROC = 2
  WIN32_ANY = 3
  WIN16_SERVERPROC = 4
  WIN16_ANY = 4
  WIN_ANY = 7
End Enum

Public Enum VER_QUERY_TYPE
  VQT_PRODUCT_VERSION = 1
  [_VQT_FIRST_ITEM] = VQT_PRODUCT_VERSION
  VQT_PRODUCT_NAME
  VQT_COMPANY_NAME
  VQT_FILE_DESCRIPTION
  VQT_FILE_VERSION
  VQT_INTERNAL_NAME
  VQT_LEGAL_COPYRIGHT
  VQT_ORIGINAL_FILE_NAME
  VQT_COMMENTS
  VQT_LEGAL_TRADEMARKS
  VQT_PRIVATE_BUILD
  'insert items here
  VQT_SPECIAL_BUILD
  [_VQT_LAST_ITEM] = VQT_SPECIAL_BUILD
End Enum

Private Const READ_CONTROL As Long = &H20000
Private Const STANDARD_RIGHTS_ALL As Long = &H1F0000
Private Const STANDARD_RIGHTS_EXECUTE As Long = (READ_CONTROL)
Private Const STANDARD_RIGHTS_READ As Long = (READ_CONTROL)
Private Const STANDARD_RIGHTS_REQUIRED As Long = &HF0000
Private Const STANDARD_RIGHTS_WRITE As Long = (READ_CONTROL)

Private Const KEY_CREATE_LINK As Long = &H20
Private Const KEY_CREATE_SUB_KEY As Long = &H4
Private Const KEY_ENUMERATE_SUB_KEYS As Long = &H8
Private Const KEY_EVENT As Long = &H1     '  Event contains key event record
Private Const KEY_NOTIFY As Long = &H10
Private Const KEY_QUERY_VALUE As Long = &H1
Private Const KEY_SET_VALUE As Long = &H2
Private Const SYNCHRONIZE As Long = &H100000

Private Const KEY_READ As Long = ((STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not SYNCHRONIZE))
Private Const KEY_EXECUTE As Long = ((KEY_READ) And (Not SYNCHRONIZE))
Private Const KEY_WRITE As Long = ((STANDARD_RIGHTS_WRITE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY) And (Not SYNCHRONIZE))
Private Const KEY_ALL_ACCESS As Long = ((STANDARD_RIGHTS_ALL Or KEY_QUERY_VALUE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or KEY_CREATE_LINK) And (Not SYNCHRONIZE))

'registry function
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hkey As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hkey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
' Note that if you declare the lpData parameter as String, you must pass it By Value.
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hkey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByVal lpType As Long, ByVal lpStrData As String, lpcbData As Long) As Long

Private Declare Sub SleepW32 Lib "kernel32" Alias "Sleep" (ByVal dwMilliseconds As Long)
Private Const PROCESS_QUERY_INFORMATION As Long = &H400
Private Const STILL_ACTIVE As Long = &H103
Private Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As Long, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function timeGetTime Lib "winmm.dll" () As Long

'##C SORTF
Public Sub CombSortEx(v As Variant, ByVal lmin As Long, ByVal lmax As Long, SortFn As ISortFunction)
  Call SortAny(v, lmin, lmax, SortFn, COMB_SORT)
End Sub

'##C SORTF
Public Sub QSortEx(v As Variant, ByVal lmin As Long, ByVal lmax As Long, SortFn As ISortFunction)
  Call SortAny(v, lmin, lmax, SortFn, QUICK_SORT)
End Sub

'##C SORTF
Public Sub QSortAsc(v As Variant, ByVal lmin As Long, ByVal lmax As Long)
  Dim SortFn As SortAscending
  Set SortFn = New SortAscending
  Call SortAny(v, lmin, lmax, SortFn, QUICK_SORT)
End Sub

'##C SORTF
Public Sub QSortDesc(v As Variant, ByVal lmin As Long, ByVal lmax As Long)
  Dim SortFn As New SortDescending
  Set SortFn = New SortDescending
  Call SortAny(v, lmin, lmax, SortFn, QUICK_SORT)
End Sub

'##D Is the variable a valid array which can be redim preserved and looped through
'##C STF
'##RV True if Array is valid
Public Function IsArrayEx(ByRef v As Variant) As Boolean
  On Error GoTo IsArrayEx_err
  If IsArray(v) Then
    IsArrayEx = (UBound(v) >= LBound(v))
  End If
  Exit Function
  
IsArrayEx_err:
  IsArrayEx = False
End Function

'##C SQF
Public Function VerQuery(sPropertyResult As String, ByVal sPathAndFile As String, Optional ByVal VQT As VER_QUERY_TYPE = VQT_FILE_VERSION) As Boolean
 sPropertyResult = VersionQueryMap(sPathAndFile, VQT)
 VerQuery = (Len(sPropertyResult) > 0)
End Function

'##D Compares Version Numbers of two files sPathAndFile1 and sPathAndFile2
'##C SQF
Public Function VerComp(ByVal sPathAndFile1 As String, ByVal sPathAndFile2 As String) As Long
  VerComp = VersionCompare(sPathAndFile1, sPathAndFile2)
End Function

'##C MF
Public Function InCollection(col As Object, vItem As Variant) As Boolean
  InCollection = InCollectionEx(col, vItem)
End Function

Private Function xSplitRegKey(KeyName As String) As Long
  Dim p As Long, rootstr As String
  
  p = InStr(KeyName, "\")
  If p > 1 Then
    rootstr = Left$(KeyName, p - 1)
    KeyName = Mid$(KeyName, p + 1)
    If StrComp(rootstr, "HKEY_CLASSES_ROOT", vbTextCompare) = 0 Then
      xSplitRegKey = &H80000000
    ElseIf StrComp(rootstr, "HKEY_CURRENT_USER", vbTextCompare) = 0 Then
      xSplitRegKey = &H80000001
    ElseIf StrComp(rootstr, "HKEY_LOCAL_MACHINE", vbTextCompare) = 0 Then
      xSplitRegKey = &H80000002
    ElseIf StrComp(rootstr, "HKEY_USERS", vbTextCompare) = 0 Then
      xSplitRegKey = &H80000003
    ElseIf StrComp(rootstr, "HKEY_CURRENT_CONFIG", vbTextCompare) = 0 Then
      xSplitRegKey = &H80000005
    End If
  End If
End Function

Private Function xRegGetKeyValue(ByVal KeyName As String, ByVal ValueName As String) As String
  Dim keyroot As Long, hkey As Long
  Dim tmplen As Long, retval As Long
  Dim sBuffer As String
  
  ' get keyroot value
  keyroot = xSplitRegKey(KeyName)
  If keyroot <> 0& Then
    If RegOpenKeyEx(keyroot, KeyName, 0&, KEY_READ, hkey) = 0 Then
      tmplen = TCSBUFSIZ
      sBuffer = String$(tmplen, 0)
      retval = RegQueryValueEx(hkey, ValueName, 0&, 0&, sBuffer, tmplen)
      If (retval = 0) And (tmplen > 1) Then
        xRegGetKeyValue = Left$(sBuffer, tmplen - 1)
      Else
        xRegGetKeyValue = ""
      End If
      Call RegCloseKey(hkey)
    End If
  End If
End Function

Private Function ServerExists(server As String) As Boolean
  Dim p As Long
  ServerExists = False
  p = InStr(server, "/")
  If p > 1 Then server = Left$(server, p - 1)
  server = Trim$(server)
  If Len(server) > 0 Then ServerExists = FileExists(server, False, False)
End Function


'##C SQF
Public Function isCOMPresent(ByVal ProgID As String, ByVal ComType As COM_TYPE) As Boolean
  Dim clsid As String, server As String
  
  On Error GoTo isCOMPresent_err:
  isCOMPresent = False
  clsid = GetCLSIDAsString(ProgID)
  If Len(clsid) > 2 Then
    If (ComType And WIN32_INPROC) > 0 Then
      server = xRegGetKeyValue("HKEY_CLASSES_ROOT\CLSID\" & clsid & "\InProcServer32", "")
      If ServerExists(server) Then
        isCOMPresent = True
        GoTo isCOMPresent_end
      End If
    End If
    If (ComType And WIN32_SERVERPROC) > 0 Then
      server = xRegGetKeyValue("HKEY_CLASSES_ROOT\CLSID\" & clsid & "\LocalServer32", "")
      If ServerExists(server) Then
        isCOMPresent = True
        GoTo isCOMPresent_end
      End If
    End If
    If (ComType And WIN16_SERVERPROC) > 0 Then
      server = xRegGetKeyValue("HKEY_CLASSES_ROOT\CLSID\" & clsid & "\LocalServer", "")
      If ServerExists(server) Then
        isCOMPresent = True
        GoTo isCOMPresent_end
      End If
    End If
  End If
  
isCOMPresent_end:
  Exit Function
  
isCOMPresent_err:
  isCOMPresent = False
  Resume isCOMPresent_end
End Function

Public Function IsRunningInIDE() As Boolean
  Dim FileName As String
  Const VB5_EXE As String = "VB5.EXE"
  Const VB6_EXE As String = "VB6.EXE"
  Static retValue As Long
  
  If retValue = 0 Then
    retValue = 1
    FileName = GetModuleName(0)
    If (StrComp(FileName, VB5_EXE, vbTextCompare) = 0) Or (StrComp(FileName, VB6_EXE, vbTextCompare) = 0) Then
      IsRunningInIDE = True
      retValue = -1
    End If
  Else
    IsRunningInIDE = (retValue = -1)
  End If
End Function

'##D Returns Minimum Value in an Sequence of Parameters
'##C NF
Public Function Min(ParamArray vals() As Variant) As Variant
  Dim i As Long
  
  Min = vals(LBound(vals))
  For i = (LBound(vals) + 1) To UBound(vals)
    If vals(i) < Min Then Min = vals(i)
  Next i
End Function

'##D Returns Maximum Value in a Sequence of Parameters
'##C NF
Public Function Max(ParamArray vals() As Variant) As Variant
  Dim i As Long
  
  Max = vals(LBound(vals))
  For i = (LBound(vals) + 1) To UBound(vals)
    If vals(i) > Max Then Max = vals(i)
  Next i
End Function

'##D Truncates a number to an integer
'##d Number to be truncated
'##C NF
Public Function Trunc(ByVal d As Double) As Long
  Trunc = CLng(Int(d))
End Function


'##D WinExec an application and show error messages
'##V ProgramPath Full path of program
'##RV double representing the program's task ID if successful, otherwise it returns zero
'##LD ##D WinExec an application and show error messages
'##LD if no path is specified windows will search the current dir, windows directory, system directory and the application path
'##C MF
Public Function ShellProgram(ByVal ProgramPath As String, Optional ByVal WinStyle As VbAppWinStyle = vbNormalFocus, Optional ByVal Wait As Boolean = False) As Long
  Dim pid As Long, hproc As Long, hErr As Long, ExitCode As Long
  
  On Error GoTo ShellProgram_err
  pid = CLng(Shell(ProgramPath, WinStyle))
  If Wait Then
    hproc = OpenProcess(PROCESS_QUERY_INFORMATION, False, pid)
    If hproc <> 0 Then
      Do
        hErr = GetExitCodeProcess(hproc, ExitCode)
        Call SleepW32(100)
      Loop Until (ExitCode <> STILL_ACTIVE) Or (hErr = 0)
    End If
  End If
  ShellProgram = pid
  
ShellProgram_end:
  Exit Function

ShellProgram_err:
  Err.Raise ERR_SYSHELPER, "ShellProgram", "Failed to execute [" & ProgramPath & "]" & vbCrLf & Err.Description
  Resume ShellProgram_end
End Function

'##C MF
Public Function ShellOpenFile(ByVal FilePath As String, Optional ByVal WinStyle As VbAppWinStyle = vbNormalFocus) As Long
  ShellOpenFile = ShellExecute(0, "open", FilePath, 0, "", WinStyle)
End Function

Public Function GetTicks() As Long
  GetTicks = timeGetTime
End Function

'##D Gets current time
'##RV Current time
'##C SQF
Public Function GetTime() As Long
  GetTime = timeGetTime
End Function

'##D Stop system for specified amount of time
'##C MF
Public Sub Sleep(ByVal Milliseconds As Long)
  Call SleepW32(Milliseconds)
End Sub

Public Function ErrorSource(ByVal ErrObj As ErrObject, ByVal FunctionName As String) As String
  ErrorSource = ErrorSourceEx(ErrObj, FunctionName)
End Function

Public Function IsValueEqual(ByRef v0 As Variant, ByRef v1 As Variant) As Boolean
  If IsNull(v0) Then
    IsValueEqual = IsNull(v1)
  ElseIf IsNull(v1) Then
    IsValueEqual = IsNull(v0)
  Else
    IsValueEqual = (v0 = v1)
  End If
End Function

