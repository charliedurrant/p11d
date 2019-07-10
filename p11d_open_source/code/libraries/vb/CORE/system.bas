Attribute VB_Name = "system"
Option Explicit

Public Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Public Declare Function GetProfileString Lib "kernel32" Alias "GetProfileStringA" (ByVal lpAppName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long) As Long
Public Declare Sub GetSystemInfo Lib "kernel32" (lpSystemInfo As SYSTEM_INFO)
Public Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Public Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Public Declare Function GetWindowsDirectory32 Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Declare Sub GlobalMemoryStatus Lib "kernel32" (lpBuffer As MEMORYSTATUS)

Public Declare Function GetPrivateProfileSection Lib "kernel32" Alias "GetPrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function WritePrivateProfileSection Lib "kernel32" Alias "WritePrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpString As String, ByVal lpFileName As String) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Public Declare Function FlushPrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As Long, ByVal lpKeyName As Long, ByVal lpString As Long, ByVal lpFileName As String) As Long

Public Declare Function ShellExecute Lib "Shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As Long, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public Declare Function GetLastError Lib "kernel32" () As Long
Public Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Public Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Public Declare Function GetTickCount Lib "kernel32" () As Long
Public Declare Function timeGetTime Lib "winmm.dll" () As Long
Public Declare Function GetModuleFileName Lib "kernel32" Alias "GetModuleFileNameA" (ByVal hModule As Long, ByVal lpFileName As String, ByVal nSize As Long) As Long
Public Declare Function GetWindowModuleFileName Lib "kernel32" Alias "GetWindowFileNameA" (ByVal hWnd As Long, ByVal lpFileName As String, ByVal nSize As Long) As Long
Public Declare Function GetFullPathName Lib "kernel32" Alias "GetFullPathNameA" (ByVal lpFileName As String, ByVal nBufferLength As Long, ByVal lpBuffer As String, ByVal lpFilePart As String) As Long
Public Declare Function GetDiskFreeSpaceEx Lib "kernel32" Alias "GetDiskFreeSpaceExA" (ByVal lpDirectoryName As String, lpFreeBytesAvailableToCaller As ULARGE_INTEGER, lpTotalNumberOfBytes As ULARGE_INTEGER, lpTotalNumberOfFreeBytes As ULARGE_INTEGER) As Long
Public Declare Function LockWindowUpdateW32 Lib "user32" Alias "LockWindowUpdate" (ByVal hwndLock As Long) As Long

Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Public Declare Function WNetOpenEnum Lib "mpr.dll" Alias "WNetOpenEnumA" (ByVal dwScope As Long, ByVal dwType As Long, ByVal dwUsage As Long, ByVal lpNetResource As Long, lphEnum As Long) As Long
Public Declare Function WNetCloseEnum Lib "mpr.dll" (ByVal hEnum As Long) As Long
Public Declare Function WNetEnumResource Lib "mpr.dll" Alias "WNetEnumResourceA" (ByVal hEnum As Long, lpcCount As Long, ByVal lpBuffer As Long, lpBufferSize As Long) As Long

Public Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Public Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal dest As Long, ByVal Source As Long, ByVal cbCopy As Long)
Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Declare Sub SleepW32 Lib "kernel32" Alias "Sleep" (ByVal dwMilliseconds As Long)
Public Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long
Public Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Public Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Public Declare Function GlobalSize Lib "kernel32" (ByVal hMem As Long) As Long
Public Declare Function GetComputerNameW32 Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long


Public Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)
Public Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long

' System functions for GetSpecialFolderEx
Public Declare Function SHGetSpecialFolderLocation Lib "Shell32.dll" (ByVal hwndOwner As Long, ByVal nFolder As Long, pidl As ITEMIDLIST) As Long
Public Declare Function SHGetPathFromIDList Lib "Shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Public Type SHITEMID
  cb As Long
  abID As Byte
End Type

Public Type ITEMIDLIST
  mkid As SHITEMID
End Type



'registry function
Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hkey As Long) As Long
Public Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hkey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
' Note that if you declare the lpData parameter as String, you must pass it By Value.
Public Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hkey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByVal lpType As Long, ByVal lpStrData As String, lpcbData As Long) As Long

'Set volume label
Public Declare Function SetVolumeLabel Lib "kernel32" Alias "SetVolumeLabelA" (ByVal lpRootPathName As String, ByVal lpVolumeName As String) As Long
'GetVolInf
Public Declare Function GetVolumeInformation Lib "kernel32" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long

'Network name
Public Declare Function WNetGetUser Lib "mpr.dll" Alias "WNetGetUserA" (ByVal lpName As String, ByVal lpUserName As String, lpnLength As Long) As Long
'get drive/ resource connection
Public Declare Function WNetGetConnection Lib "mpr.dll" Alias "WNetGetConnectionA" (ByVal lpszLocalName As String, ByVal lpszRemoteName As String, cbRemoteName As Long) As Long

'Temporary File names
Public Declare Function GetTempFileName32 Lib "kernel32" Alias "GetTempFileNameA" (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFileName As String) As Long
Public Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Public Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long

' Activate application
Public Declare Function OpenIcon Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Public Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function EnumWindows Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Public Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, ByVal lpdwProcessId As Long) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function ShowWindowAsync Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function AttachThreadInput Lib "user32" (ByVal idAttach As Long, ByVal idAttachTo As Long, ByVal fAttach As Long) As Long
Public Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Public Declare Function GetCurrentThreadId Lib "kernel32" () As Long
Public Const SW_NORMAL As Long = 1
Public Const SW_SHOW As Long = 5

Public Const GWL_HINSTANCE As Long = (-6)
Public Const GW_HWNDFIRST As Long = 0
Public Const GW_HWNDLAST  As Long = 1
Public Const GW_HWNDNEXT  As Long = 2
Public Const GW_HWNDPREV  As Long = 3



'clipboard
' details from winapi
Public Declare Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As Long) As Long
Public Declare Function GetClipboardData Lib "user32" (ByVal wFormat As Long) As Long

' ?not required
Public Declare Function OpenClipboard Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function CloseClipboard Lib "user32" () As Long
Public Declare Function GetClipboardOwner Lib "user32" () As Long
Public Declare Function SetClipboardViewer Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function GetClipboardViewer Lib "user32" () As Long
Public Declare Function ChangeClipboardChain Lib "user32" (ByVal hWnd As Long, ByVal hWndNext As Long) As Long
Public Declare Function RegisterClipboardFormat Lib "user32" Alias "RegisterClipboardFormatA" (ByVal lpString As String) As Long
Public Declare Function CountClipboardFormats Lib "user32" () As Long
Public Declare Function EnumClipboardFormats Lib "user32" (ByVal wFormat As Long) As Long
Public Declare Function GetClipboardFormatName Lib "user32" Alias "GetClipboardFormatNameA" (ByVal wFormat As Long, ByVal lpString As String, ByVal nMaxCount As Long) As Long
Public Declare Function EmptyClipboard Lib "user32" () As Long
Public Declare Function IsClipboardFormatAvailable Lib "user32" (ByVal wFormat As Long) As Long
Public Declare Function GetPriorityClipboardFormat Lib "user32" (lpPriorityList As Long, ByVal nCount As Long) As Long
Public Declare Function GetOpenClipboardWindow Lib "user32" () As Long
Public Declare Function SetCapture32 Lib "user32" Alias "SetCapture" (ByVal hWnd As Long) As Long
Public Declare Function ReleaseCapture32 Lib "user32" Alias "ReleaseCapture" () As Long
Public Declare Function WindowFromDC Lib "user32" (ByVal hdc As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Const HWND_TOPMOST = -1
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOMOVE = &H2
Public Const HWND_NOTOPMOST = -2
Public Const HWND_TOP = 0

Private Const LVM_SETCOLUMNWIDTH As Long = &H1000 + 30
Private Const LVSCW_AUTOSIZE As Long = -1
Private Const LVSCW_AUTOSIZE_USEHEADER As Long = -2

Private Type SYSTEM_INFO
    dwOemId As Long
    dwPageSize As Long
    lpMinimumApplicationAddress As Long
    lpMaximumApplicationAddress As Long
    dwActiveProcessorMask As Long
    dwNumberOfProcessors As Long
    dwProcessorType As Long
    dwAllocationGranularity As Long
    dwReserved As Long
End Type

Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
    '  Maintenance string for PSS usage
End Type

Private Type MEMORYSTATUS
    dwLength As Long
    dwMemoryLoad As Long
    dwTotalPhys As Long
    dwAvailPhys As Long
    dwTotalPageFile As Long
    dwAvailPageFile As Long
    dwTotalVirtual As Long
    dwAvailVirtual As Long
End Type

Public Type NETRESOURCE
    dwScope As Long
    dwType As Long
    dwDisplayType As Long
    dwUsage As Long
    lpLocalName As Long
    lpRemoteName As Long
    lpComment As Long
    lpProvider As Long
End Type

Public Type BrowseInfo
    hwndOwner As Long
    pIDRoot As Long
    pszDisplayName As String
    lpszTitle As String
    ulFlags As Long
    lpfnCallBack As Long
    lParam As String
    lImage As Long
End Type

Public Type ULARGE_INTEGER
  LowLong As Long
  HighLong As Long
End Type

Public Const RESOURCE_CONNECTED As Long = &H1
Public Const RESOURCE_GLOBALNET As Long = &H2
Public Const RESOURCE_REMEMBERED As Long = &H3

Public Const RESOURCETYPE_ANY As Long = &H0
Public Const RESOURCETYPE_DISK As Long = &H1
Public Const RESOURCETYPE_PRINT As Long = &H2
Public Const RESOURCETYPE_UNKNOWN As Long = &HFFFF

Public Const GMEM_DDESHARE As Long = &H2000
Public Const GMEM_DISCARDABLE As Long = &H100
Public Const GMEM_DISCARDED As Long = &H4000
Public Const GMEM_FIXED As Long = &H0
Public Const GMEM_INVALID_HANDLE As Long = &H8000
Public Const GMEM_SHARE As Long = &H2000
Public Const GMEM_MOVEABLE As Long = &H2
Public Const GMEM_ZEROINIT As Long = &H40
Public Const GPTR As Long = (GMEM_FIXED Or GMEM_ZEROINIT)
Public Const GHND As Long = (GMEM_MOVEABLE Or GMEM_ZEROINIT)

Private Const PROCESSOR_INTEL_386 As Long = 386
Private Const PROCESSOR_INTEL_486 As Long = 486
Private Const PROCESSOR_INTEL_PENTIUM As Long = 586

Public Enum OS_TYPE
    OS_UNKNOWN = 1
    OS_WIN95
    OS_WIN98
    OS_WINME
    OS_Nt35
    OS_NT4
    OS_WIN2000
    OS_WINXP
    OS_WIN2003
End Enum

Public Const VER_PLATFORM_WIN32_WINDOWS As Long = 1
Public Const VER_PLATFORM_WIN32_NT As Long = 2


Public Const PROCESS_QUERY_INFORMATION As Long = &H400
Public Const STILL_ACTIVE As Long = &H103

Public Const READ_CONTROL As Long = &H20000
Public Const STANDARD_RIGHTS_ALL As Long = &H1F0000
Public Const STANDARD_RIGHTS_EXECUTE As Long = (READ_CONTROL)
Public Const STANDARD_RIGHTS_READ As Long = (READ_CONTROL)
Public Const STANDARD_RIGHTS_REQUIRED As Long = &HF0000
Public Const STANDARD_RIGHTS_WRITE As Long = (READ_CONTROL)

Public Const KEY_CREATE_LINK As Long = &H20
Public Const KEY_CREATE_SUB_KEY As Long = &H4
Public Const KEY_ENUMERATE_SUB_KEYS As Long = &H8
Public Const KEY_EVENT As Long = &H1     '  Event contains key event record
Public Const KEY_NOTIFY As Long = &H10
Public Const KEY_QUERY_VALUE As Long = &H1
Public Const KEY_SET_VALUE As Long = &H2
Public Const SYNCHRONIZE As Long = &H100000
Public Const KEY_READ As Long = ((STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not SYNCHRONIZE))
Public Const KEY_EXECUTE As Long = ((KEY_READ) And (Not SYNCHRONIZE))
Public Const KEY_WRITE As Long = ((STANDARD_RIGHTS_WRITE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY) And (Not SYNCHRONIZE))
Public Const KEY_ALL_ACCESS As Long = ((STANDARD_RIGHTS_ALL Or KEY_QUERY_VALUE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or KEY_CREATE_LINK) And (Not SYNCHRONIZE))

Public Const BFFM_INITIALIZED As Long = 1
Public Const BFFM_SETSELECTION As Long = WM_USER + 102


' Predefined Clipboard Formats
Public Const CF_TEXT As Long = 1

' "Private" formats don't get GlobalFree()'d
Public Const CF_PRIVATEFIRST As Long = &H200
Public Const CF_PRIVATELAST As Long = &H2FF
'Private m_PrivateClip As Long
Private m_hWnd As Long

Public Const TWO_POW_32 As Double = 4294967296#
Public Const TWO_POW_10 As Double = 1024
Public Const TWO_POW_20 As Double = 1048576#
Public Const TWO_POW_30 As Double = 1073741824#

'Declare Function GetFileVersionInfo Lib "Version.dll" Alias "GetFileVersionInfoA" (ByVal lptstrFilename As String, ByVal dwhandle As Long, ByVal dwlen As Long, lpData As Any) As Long
'Declare Function GetFileVersionInfoSize Lib "Version.dll" Alias "GetFileVersionInfoSizeA" (ByVal lptstrFilename As String, lpdwHandle As Long) As Long
'Declare Function VerQueryValue Lib "Version.dll" Alias "VerQueryValueA" (pBlock As Any, ByVal lpSubBlock As String, lplpBuffer As Any, puLen As Long) As Long

Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, ByVal Source As Long, ByVal Length As Long)
Declare Function lstrcpy Lib "kernel32" Alias "lstrcpyA" (ByVal lpString1 As String, ByVal lpString2 As Long) As Long
Private mWindowActivated As Boolean

Public Const SM_SWAPBUTTON As Long = 23
Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

Public Declare Function timeBeginPeriod Lib "winmm.dll" (ByVal uPeriod As Long) As Long
Public Declare Function timeEndPeriod Lib "winmm.dll" (ByVal uPeriod As Long) As Long
Public Const TIMERR_BASE = 96   '  was 128, changed to match Win 31 Sonic
Public Const TIMERR_NOERROR = (0)  '  no error
Public Const TIMERR_NOCANDO = (TIMERR_BASE + 1) '  request not completed


Public Sub TimerFrequency(ByVal SetInterval As Boolean)
  Dim ret As Long
  
  If SetInterval Then
    ret = timeBeginPeriod(1)
  Else
    ret = timeEndPeriod(1)
  End If
  'If ret <> TIMERR_NOERROR Then Call ECASE_SYS("TimerFrequency error: " & ret, True)
End Sub

Public Function VersionQueryMap(ByVal sPathAndFile As String, Optional ByVal VQT As VER_QUERY_TYPE = VQT_FILE_VERSION) As String
  Dim sProperty As String

  Select Case VQT
    Case VQT_PRODUCT_VERSION
      sProperty = "ProductVersion"
    Case VQT_PRODUCT_NAME
      sProperty = "ProductName"
    Case VQT_COMPANY_NAME
      sProperty = "CompanyName"
    Case VQT_FILE_DESCRIPTION
      sProperty = "FileDescription"
    Case VQT_FILE_VERSION
      sProperty = "FileVersion"
    Case VQT_INTERNAL_NAME
      sProperty = "InternalName"
    Case VQT_LEGAL_COPYRIGHT
      sProperty = "LegalCopyright"
    Case VQT_ORIGINAL_FILE_NAME
      sProperty = "OriginalFilename"
    Case VQT_COMMENTS
      sProperty = "OriginalFilename"
    Case VQT_LEGAL_TRADEMARKS
      sProperty = "LegalTrademarks"
    Case VQT_PRIVATE_BUILD
      sProperty = "PrivateBuild"
    Case VQT_SPECIAL_BUILD
      sProperty = "SpecialBuild"
    Case Else
      Call ECASE_SYS("Unknown Verquery type.")
  End Select
  VersionQueryMap = cVerquery(sPathAndFile, sProperty)
End Function

Public Function VerCompEx(ByVal sPathAndFile1 As String, ByVal sPathAndFile2 As String) As Long
  Dim vcd As Long
  If Not FileExistsEx(sPathAndFile1, False, False) Then Err.Raise ERR_VERCOMPARE, "VerComp", "Cannot retrieve version information. File " & sPathAndFile1 & " does not exist."
  If Not FileExistsEx(sPathAndFile2, False, False) Then Err.Raise ERR_VERCOMPARE, "VerComp", "Cannot retrieve version information. File " & sPathAndFile2 & " does not exist."
  vcd = cVersionCompare(cFileVerquery(sPathAndFile1), cFileVerquery(sPathAndFile2))
  VerCompEx = 0
  If vcd < 0 Then
    VerCompEx = -1
  ElseIf vcd > 0 Then
    VerCompEx = 1
  End If
End Function

' Success = true
Public Function GetWindowsVersion(lMajorVer As Long, lMinorVer As Long, lBuild As Long, PlatformID As OS_TYPE, sCSDVersion As String) As Boolean
  Dim lpVI As OSVERSIONINFO
  
  lpVI.dwOSVersionInfoSize = Len(lpVI)  ' 148
  GetWindowsVersion = GetVersionEx(lpVI) <> 0
  If GetWindowsVersion Then
    lMajorVer = lpVI.dwMajorVersion
    lMinorVer = lpVI.dwMinorVersion
    lBuild = lpVI.dwBuildNumber
    PlatformID = OS_UNKNOWN
    If lpVI.dwPlatformId = VER_PLATFORM_WIN32_WINDOWS Then
      If lMinorVer = 0 Then
        PlatformID = OS_WIN95
      ElseIf lMinorVer = 10 Then
        PlatformID = OS_WIN98
      ElseIf lMinorVer = 90 Then
        PlatformID = OS_WINME
      End If
    ElseIf lpVI.dwPlatformId = VER_PLATFORM_WIN32_NT Then
      If lMajorVer = 3 Then
        PlatformID = OS_Nt35
      ElseIf lMajorVer = 4 Then
        PlatformID = OS_NT4
      ElseIf lMajorVer = 5 Then
        If lMinorVer = 0 Then
          PlatformID = OS_WIN2000
        ElseIf lMinorVer = 1 Then
          PlatformID = OS_WINXP
        ElseIf lMinorVer = 2 Then
          PlatformID = OS_WIN2003
        End If
      End If
      sCSDVersion = RTrimChar(lpVI.szCSDVersion, vbNullChar)
    End If
  End If
End Function

Public Sub GetSysInfo(sProcessor As String)
  Dim lpSysInfo As SYSTEM_INFO
  
  Call GetSystemInfo(lpSysInfo)
  Select Case lpSysInfo.dwProcessorType
    Case PROCESSOR_INTEL_386
      sProcessor = "Intel 386"
    Case PROCESSOR_INTEL_486
      sProcessor = "Intel 486"
    Case PROCESSOR_INTEL_PENTIUM
      sProcessor = "Intel Pentium"
    Case Else
      sProcessor = "Information unavailable"
  End Select
End Sub

Public Function UpdateSys(frmsys As frmEnvir) As Boolean
  Dim l0 As Long, l1 As Long, l2 As Long, l3 As Long
  Dim pid As OS_TYPE, locInfo As LocaleInfo
  Dim s0 As String, s1 As String, ret As Boolean
  Dim d0 As Double, d1 As Double, d2 As Double
  
  On Error GoTo updatesys_err
  Call GetSysInfo(s0)
  frmsys.lblSysInfo(0).Caption = s0
  If GetWindowsVersion(l0, l1, l2, pid, s0) Then
    s1 = l0 & "." & l1 & "." & l2
    If Len(s0) > 0 Then s1 = s1 & " (" & s0 & ")"
    frmsys.lblSysInfo(1).Caption = s1
    Select Case pid
      Case OS_WIN95
        s0 = "Microsoft Windows 95"
      Case OS_WIN98
        s0 = "Microsoft Windows 98"
      Case OS_WINME
        s0 = "Microsoft Windows Me"
      Case OS_Nt35
        s0 = "Microsoft Windows NT 3.5"
      Case OS_NT4
        s0 = "Microsoft Windows NT"
      Case OS_WIN2000
        s0 = "Microsoft Windows 2000"
      Case OS_WINXP
        s0 = "Microsoft Windows XP"
      Case OS_WIN2003
        s0 = "Windows Server 2003 family"
      Case Else
        s0 = "Unknown OS"
    End Select
    frmsys.lblInformation(1) = s0
  End If
  l0 = GetPhysicalMemory(d0, d1, MEGABYTES)
  frmsys.lblSysInfo(2).Caption = Format$(d0, "#,###0.00 Mb ")
  frmsys.lblSysInfo(3).Caption = Format$(d1, "#,###0.00 Mb ")
  frmsys.lblSysInfo(8).Caption = CStr(l0) & "% "
  s0 = UCase$(left$(mHomeDirectory, 3))
  ret = GetDiskSpaceEx(s0, d0, d1, d2, MEGABYTES)
  frmsys.lblInformation(5).Visible = True
  frmsys.lblSysInfo(4).Visible = True
  frmsys.lblSysInfo(6).Visible = True
  If ret Then
    frmsys.lblInformation(5).Caption = "Application drive " & s0
    frmsys.lblSysInfo(4).Caption = Format$(d0, "#,###0.00 Mb ")
    frmsys.lblSysInfo(6).Caption = Format$(d1, "#,###0.00 Mb ")
  Else
    frmsys.lblInformation(5).Caption = "Application drive " & s0
    frmsys.lblSysInfo(4).Caption = "Unavailable"
    frmsys.lblSysInfo(6).Caption = "Unavailable"
  End If
  s1 = UCase$(left$(CurDir$, 3))
  ret = ret And GetDiskSpaceEx(s1, d0, d1, d2, MEGABYTES)
  If StrComp(s1, s0, vbTextCompare) <> 0 And ret Then
    frmsys.lblInformation(6).Visible = True
    frmsys.lblStatic(1).Visible = True
    frmsys.lblStatic(3).Visible = True
    frmsys.lblSysInfo(5).Visible = True
    frmsys.lblSysInfo(7).Visible = True
    frmsys.lblInformation(6).Caption = "Current drive " & s0
    frmsys.lblSysInfo(5).Caption = Format$(d0, "#,###0.00 Mb ")
    frmsys.lblSysInfo(7).Caption = Format$(d1, "#,###0.00 Mb ")
  Else
    frmsys.lblInformation(6).Visible = False
    frmsys.lblStatic(1).Visible = False
    frmsys.lblStatic(3).Visible = False
    frmsys.lblSysInfo(5).Visible = False
    frmsys.lblSysInfo(7).Visible = False
  End If
  Set locInfo = New LocaleInfo
  s0 = "System Locale ID " & locInfo.GetSystemDefaultLcid & vbCrLf
  s0 = s0 & "Country " & locInfo.GetLocaleValue(LOCALE_SYSTEM_DEFAULT, LOCALE_SENGCOUNTRY) & vbCrLf
  s0 = s0 & "Language " & locInfo.GetLocaleValue(LOCALE_SYSTEM_DEFAULT, LOCALE_SENGLANGUAGE) & vbCrLf
  s0 = s0 & "Currency " & locInfo.GetLocaleValue(LOCALE_SYSTEM_DEFAULT, LOCALE_SCURRENCY) & " (" & locInfo.GetLocaleValue(LOCALE_USER_DEFAULT, LOCALE_SINTLSYMBOL) & ")" & vbCrLf
  s0 = s0 & "Short Date " & locInfo.GetLocaleValue(LOCALE_SYSTEM_DEFAULT, LOCALE_SSHORTDATE) & vbCrLf
  s0 = s0 & "Long Date " & locInfo.GetLocaleValue(LOCALE_SYSTEM_DEFAULT, LOCALE_SLONGDATE) & vbCrLf
  frmsys.lblLocaleSys.Caption = s0
  
  s0 = "User Locale ID " & locInfo.GetUserDefaultLcid & vbCrLf
  s0 = s0 & "Country " & locInfo.GetLocaleValue(LOCALE_USER_DEFAULT, LOCALE_SENGCOUNTRY) & vbCrLf
  s0 = s0 & "Language " & locInfo.GetLocaleValue(LOCALE_USER_DEFAULT, LOCALE_SENGLANGUAGE) & vbCrLf
  s0 = s0 & "Currency " & locInfo.GetLocaleValue(LOCALE_USER_DEFAULT, LOCALE_SCURRENCY) & " (" & locInfo.GetLocaleValue(LOCALE_USER_DEFAULT, LOCALE_SINTLSYMBOL) & ")" & vbCrLf
  s0 = s0 & "Short Date " & locInfo.GetLocaleValue(LOCALE_USER_DEFAULT, LOCALE_SSHORTDATE) & vbCrLf
  s0 = s0 & "Long Date " & locInfo.GetLocaleValue(LOCALE_USER_DEFAULT, LOCALE_SLONGDATE) & vbCrLf
  frmsys.lblLocalUser.Caption = s0
  UpdateSys = True
  
updatesys_exit:
  Exit Function
updatesys_err:
  UpdateSys = False
  Resume updatesys_exit
End Function

' A=1
Public Function GetDriveMappings(maps() As String) As Long
  Dim hEnum As Long, hErr As Long
  Dim lCount As Long, lBufSiz As Long
  Dim nr As NETRESOURCE
  Dim netresbuf As Long, i As Long
  Dim Drive As Byte
  Dim sremote As String
  
  hErr = WNetOpenEnum(RESOURCE_CONNECTED, RESOURCETYPE_DISK, 0, 0&, hEnum)
  If hErr <> 0 Then Exit Function
  
  lCount = &HFFFFFFFF
  lBufSiz = LenB(nr) * 256
  netresbuf = GlobalAlloc(GPTR, lBufSiz)
  hErr = WNetEnumResource(hEnum, lCount, netresbuf, lBufSiz)
  If hErr <> 0 Then Exit Function
  
  For i = 0 To (lCount - 1)
    Call CopyMemory(VarPtr(nr), netresbuf + (i * LenB(nr)), LenB(nr))
    If (nr.lpLocalName <> 0) And (nr.lpRemoteName <> 0) Then
      Call CopyMemory(VarPtr(Drive), nr.lpLocalName, 1)
      If Drive <> 0 Then
        Drive = Drive - Asc("A") + 1
        Call CopyMemtoString(sremote, nr.lpRemoteName)
        maps(Drive) = sremote
        GetDriveMappings = GetDriveMappings + 1
      End If
    End If
  Next i
  Call GlobalFree(netresbuf)
  Call WNetCloseEnum(hEnum)
End Function

Public Function GetDrivePathEx(ByVal sDirectoryOnly As String) As String
  Dim i As Long, p As Long
  Dim maxp As Long, maxi As Long
  Dim s As String
  Dim maps(1 To 27) As String
    
  If Len(sDirectoryOnly) > 0 Then
    If (left$(sDirectoryOnly, 2) = "\\") Then
      If GetDriveMappings(maps) > 0 Then
        For i = 1 To 27
          If Len(maps(i)) > 0 Then
            p = InStr(1, sDirectoryOnly, maps(i), vbTextCompare)
            If p > 0 Then
              p = Len(maps(i))
              If p > maxp Then
                maxp = p
                maxi = i
              End If
            End If
          End If
        Next i
        If maxi > 0 Then
          s = Mid$(sDirectoryOnly, maxp + 2)
          's = right$(sPathOnly, Len(sPathOnly) - maxp - 1)
          sDirectoryOnly = Chr(maxi + Asc("A") - 1) & ":\" & s
        End If
      End If
    End If
  End If
  GetDrivePathEx = FullPathEx(sDirectoryOnly)
End Function

Public Function IsUNCPath(ByVal sPath As String) As Boolean
  IsUNCPath = (left$(sPath, 2) = "\\")
End Function
Private Function xSplitRegKey(KeyName As String) As Long
  Dim p As Long, rootstr As String
  
  p = InStr(KeyName, "\")
  If p > 1 Then
    rootstr = left$(KeyName, p - 1)
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
'CAD P11D private to public
Public Function xRegGetKeyValue(ByVal KeyName As String, ByVal ValueName As String) As String
  Dim keyroot As Long, hkey As Long
  Dim tmplen As Long, retval As Long
  Dim sBuffer As String
  
  ' get keyroot value
  keyroot = xSplitRegKey(KeyName)
  If keyroot <> 0& Then
    If RegOpenKeyEx(keyroot, KeyName, 0&, KEY_READ, hkey) = 0 Then
      tmplen = TCSBUFSIZ
      sBuffer = String$(tmplen, vbNullChar)
      retval = RegQueryValueEx(hkey, ValueName, 0&, 0&, sBuffer, tmplen)
      If (retval = 0) And (tmplen > 1) Then
        xRegGetKeyValue = left$(sBuffer, tmplen - 1)
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
  If p > 1 Then server = left$(server, p - 1)
  server = Trim$(server)
  If Len(server) > 0 Then ServerExists = FileExistsEx(server, False, False)
End Function
    
Public Sub SetWindowZOrderEx(ByVal hWnd As Long, ByVal PositionFlag As Long)
 Call SetWindowPos(hWnd, PositionFlag, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
End Sub
      
Public Function GetComponentPathEx(ClassName As String) As String
  Dim clsid As String, server As String
  
  On Error GoTo GetComponentPathEx_err:
  clsid = GetClsidAsString(ClassName)
  If Len(clsid) > 2 Then
    server = xRegGetKeyValue("HKEY_CLASSES_ROOT\CLSID\" & clsid & "\InProcServer32", "")
    If Len(server) > 0 Then
      GetComponentPathEx = server
      Exit Function
    End If
    server = xRegGetKeyValue("HKEY_CLASSES_ROOT\CLSID\" & clsid & "\LocalServer32", "")
    If Len(server) > 0 Then
      GetComponentPathEx = server
      Exit Function
    End If
    server = xRegGetKeyValue("HKEY_CLASSES_ROOT\CLSID\" & clsid & "\LocalServer", "")
    If Len(server) > 0 Then
      GetComponentPathEx = server
      Exit Function
    End If
  End If
  
GetComponentPathEx_end:
  Exit Function
  
GetComponentPathEx_err:
  GetComponentPathEx = ""
  Resume GetComponentPathEx_end
End Function
    
    
Public Function isCOMPresentEx(ClassName As String, ByVal cType As COM_TYPE) As Boolean
  Dim clsid As String, server As String
  
  On Error GoTo isCOMPresentEx_err:
  isCOMPresentEx = False
  clsid = GetClsidAsString(ClassName)
  If Len(clsid) > 2 Then
    If (cType And WIN32_INPROC) > 0 Then
      server = xRegGetKeyValue("HKEY_CLASSES_ROOT\CLSID\" & clsid & "\InProcServer32", "")
      If ServerExists(server) Then
        isCOMPresentEx = True
        GoTo isCOMPresentEx_end
      End If
    End If
    If (cType And WIN32_SERVERPROC) > 0 Then
      server = xRegGetKeyValue("HKEY_CLASSES_ROOT\CLSID\" & clsid & "\LocalServer32", "")
      If ServerExists(server) Then
        isCOMPresentEx = True
        GoTo isCOMPresentEx_end
      End If
    End If
    If (cType And WIN16_SERVERPROC) > 0 Then
      server = xRegGetKeyValue("HKEY_CLASSES_ROOT\CLSID\" & clsid & "\LocalServer", "")
      If ServerExists(server) Then
        isCOMPresentEx = True
        GoTo isCOMPresentEx_end
      End If
    End If
  End If
  
isCOMPresentEx_end:
  Exit Function
  
isCOMPresentEx_err:
  isCOMPresentEx = False
  Resume isCOMPresentEx_end
End Function

Public Function GetTempDirectoryEx() As String
  Dim sBuffer As String, retval As Long
    
  sBuffer = String$(TCSBUFSIZ, vbNullChar)
  retval = GetTempPath(TCSBUFSIZ, sBuffer)
  If retval = 0 Then Exit Function
  GetTempDirectoryEx = left$(sBuffer, retval)
End Function

Public Function GetTempFileNameEx(ByVal DefaultDir As String, ByVal FilePrefix As String) As String
  Dim l As Long
  Dim sBuffer As String
  
  sBuffer = String$(TCSBUFSIZ, vbNullChar)
  If Len(DefaultDir) = 0 Then DefaultDir = GetTempDirectoryEx()
  If Len(DefaultDir) = 0 Then Exit Function
  FilePrefix = left$(FilePrefix & "TMP", 3)
  Do
    ' l between 1 and FFFF-1
    l = Int((Rnd() * 65534) + 1)
    sBuffer = String$(TCSBUFSIZ, vbNullChar)
    If GetTempFileName32(DefaultDir, FilePrefix, l, sBuffer) = 0 Then Exit Function
    sBuffer = RTrimChar(sBuffer, vbNullChar)
  Loop Until Not FileExistsEx(sBuffer, False, False)
  GetTempFileNameEx = sBuffer
End Function

Function UnsignedToDouble(ByVal Value As Long) As Double
  If Value < 0 Then
    UnsignedToDouble = -(&H80000000 - Value)
    UnsignedToDouble = UnsignedToDouble + &H7FFFFFFF
  Else
    UnsignedToDouble = Value
  End If
End Function

Private Function ULIToDouble(ULI As ULARGE_INTEGER) As Double
  ULIToDouble = UnsignedToDouble(ULI.LowLong) + (TWO_POW_32 * UnsignedToDouble(ULI.HighLong))
End Function

Public Function GetPhysicalMemory(dTotalPhysical As Double, dFreePhysical As Double, ByVal nMemUnit As MemoryUnit) As Long
  Dim lpBuffer As MEMORYSTATUS
  Dim retval As Long
  
  lpBuffer.dwLength = LenB(lpBuffer)
  Call GlobalMemoryStatus(lpBuffer)
  dTotalPhysical = lpBuffer.dwTotalPhys
  dFreePhysical = lpBuffer.dwAvailPhys
  Select Case nMemUnit
    Case GIGABYTES
      dTotalPhysical = dTotalPhysical / TWO_POW_30
      dFreePhysical = dFreePhysical / TWO_POW_30
    Case MEGABYTES
      dTotalPhysical = dTotalPhysical / TWO_POW_20
      dFreePhysical = dFreePhysical / TWO_POW_20
  End Select
  dTotalPhysical = RoundDouble(dTotalPhysical, 2, R_NORMAL)
  dFreePhysical = RoundDouble(dFreePhysical, 2, R_NORMAL)
  GetPhysicalMemory = lpBuffer.dwMemoryLoad
End Function

Public Function GetDiskSpaceEx(ByVal RootPath As String, dblTotal As Double, dblFreeToUser As Double, dblFreeOnDisk As Double, ByVal nMemUnit As MemoryUnit) As Boolean
  Dim FreeToUser As ULARGE_INTEGER
  Dim FreeOnDisk As ULARGE_INTEGER
  Dim Total As ULARGE_INTEGER
  Dim retval As Long
  
  On Error GoTo GetDiskSpaceEx_err
  retval = GetDiskFreeSpaceEx(RootPath, FreeToUser, Total, FreeOnDisk)
  If retval > 0 Then
    dblTotal = ULIToDouble(Total)
    dblFreeToUser = ULIToDouble(FreeToUser)
    dblFreeOnDisk = ULIToDouble(FreeOnDisk)
    Select Case nMemUnit
      Case GIGABYTES
        dblTotal = dblTotal / TWO_POW_30
        dblFreeToUser = dblFreeToUser / TWO_POW_30
        dblFreeOnDisk = dblFreeOnDisk / TWO_POW_30
      Case MEGABYTES
        dblTotal = dblTotal / TWO_POW_20
        dblFreeToUser = dblFreeToUser / TWO_POW_20
        dblFreeOnDisk = dblFreeOnDisk / TWO_POW_20
    End Select
    dblTotal = RoundDouble(dblTotal, 2, R_NORMAL)
    dblFreeToUser = RoundDouble(dblFreeToUser, 2, R_NORMAL)
    dblFreeOnDisk = RoundDouble(dblFreeOnDisk, 2, R_NORMAL)
    GetDiskSpaceEx = True
  End If
GetDiskSpaceEx_end:
  Exit Function
  
GetDiskSpaceEx_err:
  GetDiskSpaceEx = False
  Resume GetDiskSpaceEx_end
End Function

Public Sub SplitPathEx(sFullPath As String, Optional sDir As Variant, Optional sFile As Variant, Optional sExt As Variant)
  Dim p As Long, q As Long, tmp As Long
  
  On Error Resume Next
  p = InStrRev(sFullPath, "\")
  If p = 0 Then p = InStrRev(sFullPath, ":")
  
  If (Not IsMissing(sDir)) And (p > 0) Then
    sDir = left$(sFullPath, p)
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
    sExt = right$(sFullPath, q)
  End If
End Sub

Public Function FileCopyExN(ByVal Source As String, ByVal Destination As String, ByVal RaiseErrors As Boolean) As Boolean
  Dim DestDir As String
  Dim DestDirShort As String
  Dim tmpDestination As String, ret As Long
  Dim blen As Long
  
  On Error GoTo FileCopyExN_Err
  Call SplitPathEx(Destination, DestDir)
  If Len(DestDir) = 0 Then
    Destination = FullPathEx(CurDir()) & Destination
    DestDir = FullPathEx(CurDir())
  End If
  
expand_buffer:
  blen = blen + TCSBUFSIZ
  DestDirShort = String$(TCSBUFSIZ, vbNullChar)
  ret = GetShortPathName(DestDir, DestDirShort, TCSBUFSIZ)
  If ret > TCSBUFSIZ Then GoTo expand_buffer
  If (ret = 0) And (Len(DestDir) > 0) Then Err.Raise ERR_FILECOPY, "FileCopyExN", "Unable to get short name for destination " & DestDir
  DestDirShort = RTrimChar(DestDirShort, vbNullChar)
  tmpDestination = GetTempFileNameEx(DestDirShort, "FC")
  Call FileCopy(Source, tmpDestination)
  If Len(DestDirShort) > 0 Then Destination = DestDirShort & Mid$(Destination, Len(DestDir) + 1)
  Call xKillEx(Destination)
  Name tmpDestination As Destination
  FileCopyExN = True
  
FileCopyExN_End:
  Exit Function
  
FileCopyExN_Err:
  If RaiseErrors Then Call PushErrorMessage(Err)
  FileCopyExN = False
  If FileExistsEx(tmpDestination, False, False) Then Call xKillEx(tmpDestination)
  If RaiseErrors Then
    Call PopErrorMessageErr(Err)
    Err.Raise Err.Number, Err.Source, Err.Description
  End If
  Resume FileCopyExN_End
End Function

Public Function SetAnyClipboardDataEx(ByVal ClipFormat As VBRUN.ClipBoardConstants, String1 As String) As Boolean
  Dim hMem As Long, lpMem As Long
  Dim n As Long, hData As Long
  
  n = Len(String1) + 1
  hMem = GlobalAlloc(GHND, n)
  If hMem <> 0 Then
    lpMem = GlobalLock(hMem)
    If lpMem <> 0 Then
      Call CopyStringtoMem(lpMem, String1)
      Call GlobalUnlock(hMem)
      hData = SetClipboardData(ClipFormat, hMem)
    End If
    If hData = 0 Then Call GlobalFree(hMem)
  End If
  SetAnyClipboardDataEx = (hData <> 0)
End Function

Public Function GetAnyClipboardDataEx(ByVal ClipFormat As VBRUN.ClipBoardConstants) As String
  Dim hMem As Long, lpMem As Long, n As Long
  
  If OpenClipboard(0) Then
    If IsClipboardFormatAvailable(ClipFormat) Then
      hMem = GetClipboardData(ClipFormat)
      n = GlobalSize(hMem)
      If n > 0 Then
        lpMem = GlobalLock(hMem)
        If lpMem <> 0 Then
          Call CopyMemtoString(GetAnyClipboardDataEx, lpMem, n)
          Call GlobalUnlock(hMem)
        End If
      End If
    End If
    Call CloseClipboard
  End If
End Function

Public Function GetComputerName_s() As String
  Dim lRetVal As Long, lBufSiz As Long
  
  lBufSiz = TCSBUFSIZ
  GetComputerName_s = String$(TCSBUFSIZ, 0)
  lRetVal = GetComputerNameW32(GetComputerName_s, lBufSiz)
  If lRetVal <> 0 Then
    GetComputerName_s = RTrimChar(GetComputerName_s, vbNullChar)
  Else
    GetComputerName_s = "(UNKNOWN_COMPUTER)"
  End If
End Function


Public Function GetNetUser_s(bShowErrors As Boolean) As String
  Static UserName As String
  Dim lRetVal As Long, lBufSiz As Long
    
  On Error GoTo GetNetUser_s_err
  If Len(UserName) = 0 Then
    lBufSiz = TCSBUFSIZ
    GetNetUser_s = String$(lBufSiz, 0)
    lRetVal = WNetGetUser("", GetNetUser_s, lBufSiz)
    If lRetVal = 0 Then
      GetNetUser_s = RTrimChar(GetNetUser_s, vbNullChar)
      UserName = GetNetUser_s
    Else
      UserName = "(UNKNOWN)"
    End If
  End If
  
GetNetUser_s_end:
  GetNetUser_s = UserName
  Exit Function
  
GetNetUser_s_err:
  UserName = "(ERR_UNKNOWN)"
  If bShowErrors Then
    Call ErrorMessageEx(ERR_ERROR, Err, "GetNetUser", "ERR_GETNETUSER", "", False)
  End If
  Resume GetNetUser_s_end
End Function

'* lpMem points to Len(String1) + 1 characters
Private Sub CopyStringtoMem(ByVal lpMem As Long, String1 As String)
  Dim i As Long, chW As Integer, ch As Byte, n As Long
  
  n = Len(String1)
  For i = 1 To n
    chW = Asc(Mid$(String1, i, 1))
    ch = (chW And &HFF)
    Call CopyMemory(lpMem + i - 1, VarPtr(ch), 1)
  Next i
  ch = 0
  Call CopyMemory(lpMem + n, VarPtr(ch), 1)
End Sub

'* copy null terminated string from lpMem to String1
Private Sub CopyMemtoString(String1 As String, ByVal lpMem As Long, Optional ByVal MaxBytes As Long = &H7FFFFFFF)
  Dim ch As Byte, i As Long
  
  i = 0: String1 = ""
  Do
    Call CopyMemory(VarPtr(ch), lpMem + i, 1)
    If ch = 0 Then Exit Do
    String1 = String1 & Chr$(ch)
    i = i + 1
  Loop Until i >= MaxBytes
End Sub

Public Function AddressOfFunc(ByVal fnAddr As Long) As Long
  AddressOfFunc = fnAddr
End Function

Public Function BrowseCallback(ByVal hWnd As Long, ByVal msg As Long, ByRef lp As Long, ByRef pdata As Long) As Long
  If (msg = BFFM_INITIALIZED) Then
    Call SendMessage(hWnd, BFFM_SETSELECTION, 1, pdata)
  End If
  BrowseCallback = 0
End Function

Public Function IsFileOpenEx(FileName As String, ByVal Exclusive As Boolean) As Boolean
  Dim i As Integer
  
  On Error GoTo IsFileOpenEx_ERR
  If FileExistsEx(FileName, False, False) Then
    i = FreeFile
    If Exclusive Then
      Open FileName For Input Shared As i
    Else
      Open FileName For Input Lock Read Write As i
    End If
    Close i
  End If
  
IsFileOpenEx_END:
  Exit Function
  
IsFileOpenEx_ERR:
  IsFileOpenEx = True
  Resume IsFileOpenEx_END
End Function

Public Function GetIniKeyNamesInt(KeyNames As Variant, ByVal SectionName As String, ByVal IniFilePath As String) As Long
  Dim sBuffer As String, bsize As Long
  Dim p0 As Long, p1 As Long, MaxKey As Long, retval As Long
  
  If Len(IniFilePath) = 0 Then
    IniFilePath = GetWindowsDirectoryEx() & "\" & mAppExeName & ".INI"
  End If
  bsize = TCSBUFSIZ
  Do
    bsize = bsize * 2
    sBuffer = String$(bsize, 0)
    retval = GetPrivateProfileString(SectionName, 0&, "", sBuffer, bsize, IniFilePath)
  Loop Until (retval = 0) Or (retval <> (bsize - 2))
  If retval > 0 Then
    MaxKey = 0: p0 = 1
    Do
      p1 = InStr(p0, sBuffer, vbNullChar)
      MaxKey = MaxKey + 1
      If IsArrayEx2(KeyNames) Then
        ReDim Preserve KeyNames(1 To MaxKey)
      Else
        ReDim KeyNames(1 To MaxKey)
      End If
      KeyNames(MaxKey) = Mid$(sBuffer, p0, p1 - p0)
      p0 = p1 + 1
    Loop Until Mid$(sBuffer, p0, 1) = vbNullChar
  End If
  GetIniKeyNamesInt = MaxKey
End Function

Public Function GetIniEntryEx(ByVal Section As String, ByVal Key As String, ByVal default As String, ByVal IniFilePath As String) As String
  Dim sBuffer As String, bsize As Long
  Dim retval As Long
  
  If Len(IniFilePath) = 0 Then
    IniFilePath = GetWindowsDirectoryEx() & "\" & mAppExeName & ".INI"
  End If
  bsize = TCSBUFSIZ
  Do
    bsize = bsize * 2
    sBuffer = String$(bsize, 0)
    retval = GetPrivateProfileString(Section, Key, "", sBuffer, bsize, IniFilePath)
  Loop Until (retval = 0) Or (retval <> (bsize - 1))
  If retval = 0 Then
    sBuffer = default
  Else
    sBuffer = RTrimChar(sBuffer, vbNullChar)
  End If
  GetIniEntryEx = sBuffer
'  Dim sBuffer As String
'  Dim retval As Long
'
'  If Len(IniFilePath) = 0 Then
'    IniFilePath = GetWindowsDirectoryEx() & "\" & mAppExeName & ".INI"
'  End If
'  sBuffer = String$(TCSBUFSIZ, 0)
'  retval = GetPrivateProfileString(Section, Key, "", sBuffer, TCSBUFSIZ, IniFilePath)
'  If retval = 0 Then
'    sBuffer = Default
'  Else
'    sBuffer = RTrimChar(sBuffer, vbNullChar)
'  End If
'  GetIniEntryEx = sBuffer
End Function

Function GetWindowsDirectoryEx() As String
  Dim sRes As String
  Dim retval As Long
  
  On Error GoTo GetWindowsDirectoryEx_err
  sRes = String$(TCSBUFSIZ, 0)
  retval = GetWindowsDirectory32(sRes, TCSBUFSIZ)
  If retval = 0 Then
    sRes = WINDIR
  Else
    sRes = RTrimChar(sRes, vbNullChar)
  End If
  
GetWindowsDirectoryEx_end:
  GetWindowsDirectoryEx = sRes
  Exit Function
  
GetWindowsDirectoryEx_err:
  Call ErrorMessageEx(ERR_ERROR, Err, "GetWindowsDirectoryEx", "ERR_GETWINDOWSDIRECTORY", "Unable to retrieve Windows directory." & vbCr & "using " & WINDIR, False)
  Resume GetWindowsDirectoryEx_end
End Function
Public Function GetSysDirectoryEx() As String
  Dim sRes As String
  Dim retval As Long
  
  On Error GoTo GetSysDirectoryEx_err
  sRes = String$(TCSBUFSIZ, 0)
  retval = GetSystemDirectory(sRes, TCSBUFSIZ)
  If retval = 0 Then
    sRes = WINDIR
  Else
    sRes = RTrimChar(sRes, vbNullChar)
  End If
  
GetSysDirectoryEx_end:
  GetSysDirectoryEx = sRes
  Exit Function
  
GetSysDirectoryEx_err:
  Call ErrorMessageEx(ERR_ERROR, Err, "GetSysDirectoryEx", "ERR_GETSYSTEMDIRECTORY", "Unable to retrieve Windows system directory.", False)
  Resume GetSysDirectoryEx_end
End Function

Public Function GetModuleName(ByVal hModule As Long, Optional ByVal FullPath As Boolean = False) As String
  Dim FileName As String, FileNameLen As Long, p As Long
  
  FileName = String$(TCSBUFSIZ, 0)
  FileNameLen = GetModuleFileName(hModule, FileName, TCSBUFSIZ)
  FileName = left$(FileName, FileNameLen)
  If Not FullPath Then
    p = InStrRev(FileName, "\")
    If p > 0 Then FileName = Mid$(FileName, p + 1)
  End If
  GetModuleName = FileName
End Function

Public Function GetWindowModuleName(ByVal hWnd As Long, Optional ByVal FullPath As Boolean = False) As String
  Dim FileName As String, FileNameLen As Long, p As Long
  
  FileName = String$(TCSBUFSIZ, 0)
  FileNameLen = GetWindowModuleFileName(hWnd, FileName, TCSBUFSIZ)
  FileName = left$(FileName, FileNameLen)
  If Not FullPath Then
    p = InStrRev(FileName, "\")
    If p > 0 Then FileName = Mid$(FileName, p + 1)
  End If
  GetWindowModuleName = FileName
End Function

Public Function IsRunningInIDEEx() As Boolean
  Static retValue As Long
  Dim FileName As String
  Const VB5_EXE As String = "VB5.EXE"
  Const VB6_EXE As String = "VB6.EXE"
  
  If retValue = 0 Then
    retValue = 1
    FileName = GetModuleName(ghInstance)
    If (StrComp(FileName, VB5_EXE, vbTextCompare) = 0) Or (StrComp(FileName, VB6_EXE, vbTextCompare) = 0) Then
      IsRunningInIDEEx = True
      retValue = -1
    End If
  Else
    IsRunningInIDEEx = (retValue = -1)
  End If
End Function

Public Function xKillEx(ByVal FullPath As String) As Boolean
  On Error GoTo xKillEx_err
  Call Kill(FullPath)
  xKillEx = True
  
xKillEx_end:
  Exit Function
  
xKillEx_err:
  If Err.Number = 53 Then xKillEx = True
  Resume xKillEx_end
End Function

Public Sub AutoWidthListViewEx(ByVal lv As ListView, ByVal IncludeColumnHeaders As Boolean)
  Dim col As Long, lParam As Long
  
  If lv Is Nothing Then Exit Sub
  If IncludeColumnHeaders Then
    lParam = LVSCW_AUTOSIZE_USEHEADER
  Else
    lParam = LVSCW_AUTOSIZE
  End If
  ' Send the message to all the columns
  For col = 0 To lv.ColumnHeaders.Count - 1
   Call SendMessage(lv.hWnd, LVM_SETCOLUMNWIDTH, col, ByVal lParam)
  Next col
End Sub

Public Function LOWORD(ByVal dw As Long) As Long
  LOWORD = (dw And &HFFFF&)
End Function

Public Function ActivateWindowProc(ByVal hWnd As Long, ByVal pidCur As Long) As Long
  Dim pidNext As Long, hPid As Long
  Dim hThreadNext As Long, hInstNext As Long
  Dim AppModuleName As String, AppNextModuleName As String
 
  hThreadNext = GetWindowThreadProcessId(hWnd, VarPtr(pidNext))
  If pidCur <> pidNext Then
    AppModuleName = GetModuleName(ghInstance)
    
    hInstNext = GetWindowLong(hWnd, GWL_HINSTANCE)
    'hPid = OpenProcess(PROCESS_QUERY_INFORMATION, 0, pidNext)
    AppNextModuleName = GetModuleName(hInstNext)
    'If hPid <> 0 Then Call CloseHandle(hPid)
    If StrComp(AppModuleName, AppNextModuleName, vbTextCompare) = 0 Then
      If AttachThreadInput(hThreadNext, GetCurrentThreadId(), 1) = 0 Then Call ECASE_SYS("Unable to Attach Thread input " & AppNextModuleName)
      If SetForegroundWindow(hWnd) = 0 Then Call ECASE_SYS("Unable to activate application " & AppNextModuleName)
      Call AttachThreadInput(hThreadNext, GetCurrentThreadId(), 0)
      mWindowActivated = True
      ActivateWindowProc = 0
      Exit Function
    End If
  End If
  ActivateWindowProc = -1
End Function

Public Sub ActivatePrevInstanceInt()
  mWindowActivated = False
  If vbg Is Nothing Then Err.Raise ERR_ACTIVATEPREV, "ActivatePrevInstance", "CoreSetup must be called before using Core functions"
  Call EnumWindows(AddressOf ActivateWindowProc, GetCurrentProcessId())
  If Not mWindowActivated Then Err.Raise ERR_ACTIVATEPREV, "ActivatePrevInstance", "Unable to find previous instance"
End Sub

Public Function DateSerialEx(ByVal Year As Long, ByVal Month As Long, ByVal Day As Long) As Date
  Dim DayMax As Long

  On Error GoTo DateSerialEx_err
  If (Month < 1) Or (Month > 12) Then Err.Raise ERR_DATESERIAL, "DateSerialEx", "Invalid Month Given In Date"
  If (Year < 100) Or (Year > 9999) Then Err.Raise ERR_DATESERIAL, "DateSerialEx", "Invalid Year Given In Date"
  
  Select Case Month
    Case 4, 6, 9, 11
      DayMax = 30
    Case 2
      If IsLeapYear(Year) Then
        DayMax = 29
      Else
        DayMax = 28
      End If
    Case Else
      DayMax = 31
  End Select
  If (Day < 1) Or (Day > DayMax) Then Err.Raise ERR_DATESERIAL, "DateSerialEx", "Invalid Day Given In Date"
  DateSerialEx = DateSerial(Year, Month, Day)
  Exit Function
  
DateSerialEx_err:
  Err.Raise Err.Number, ErrorSourceEx(Err, "DateSerialEx"), "Invalid Date '" & Day & "/" & Month & "/" & Year & "' (DMY) " & vbCrLf & Err.Description
End Function

Private Function IsLeapYear(ByVal Year As Long) As Boolean
  If (Year Mod 100) = 0 Then Year = Year / 100
  IsLeapYear = (Year Mod 4 = 0)
End Function

Public Function VarTypetoDatatypeEx(ByVal vbType As VbVarType) As DATABASE_FIELD_TYPES
  Select Case vbType
    Case vbCurrency, vbDecimal, vbDouble, vbSingle
      VarTypetoDatatypeEx = TYPE_DOUBLE
    Case vbInteger, vbLong, vbByte
      VarTypetoDatatypeEx = TYPE_LONG
    Case vbDate
      VarTypetoDatatypeEx = TYPE_DATE
    Case vbString
      VarTypetoDatatypeEx = TYPE_STR
    Case vbBoolean
      VarTypetoDatatypeEx = TYPE_BOOL
    Case Else
      Call ECASE_SYS("Unrecognised data Type: " & CStr(vbType))
  End Select
End Function

Public Function GetSpecialFolderEx(ByVal CSIDL As CSIDL_SPECIAL_FOLDERS) As String
  Dim sPath As String
  Dim IDL As ITEMIDLIST
  ' Retrieve info about system folders such as the "Recent Documents" folder.
  ' Info is stored in the IDL structure.
  If SHGetSpecialFolderLocation(frmAbout.hWnd, CSIDL, IDL) = 0 Then
    ' Get the path from the ID list, and return the folder.
    sPath = String$(TCSBUFSIZ, vbNullChar)
    If SHGetPathFromIDList(IDL.mkid.cb, sPath) Then GetSpecialFolderEx = RTrimChar(sPath, vbNullChar)
  End If
End Function


