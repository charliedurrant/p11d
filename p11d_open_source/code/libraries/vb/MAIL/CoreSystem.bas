Attribute VB_Name = "CoreSystem"
Option Explicit
'This Module includes replication of Core functionality
'AM Environent info TCSCore function. Disk Info and Locale settings currently excluded.
Public Function GetEnvironmentInfo(MyMail As Mail) As String
Dim sOutput As String

'ATC Mail environment information
  sOutput = sOutput & "MailApplication: " & MyMail.MailApplication & vbCrLf
  sOutput = sOutput & "MailSystem: " & MyMail.MailSystem & vbCrLf

'CORE environment information, copied from Public Function UpdateSys(frmsys As frmMailTest) As Boolean
  Dim l0 As Long, l1 As Long, l2 As Long, l3 As Long
  Dim pid As OS_TYPE ', locInfo As LocaleInfo
  Dim s0 As String, s1 As String, s2 As String, ret As Boolean
  Dim d0 As Double, d1 As Double, d2 As Double
  
  On Error GoTo GetEnvironmentInfo_err
  Call GetSysInfo(s0)
  sOutput = sOutput & "Processor type: " & s0 & vbCrLf
  'frmsys.lblSysInfo(0).Caption = s0
  If GetWindowsVersion(l0, l1, l2, pid, s0) Then
    s1 = l0 & "." & l1 & "." & l2
    If Len(s0) > 0 Then s1 = s1 & " (" & s0 & ")"
    'sOutput = sOutput & s1 & vbCrLf
    'frmsys.lblSysInfo(1).Caption = s1
    Select Case pid
      Case OS_NT4
        s2 = "Microsoft Windows NT"
      Case OS_WIN95
        s2 = "Microsoft Windows 95"
      Case OS_WIN98
        s2 = "Microsoft Windows 98"
      Case OS_W2000
        s2 = "Microsoft Windows 2000"
      Case Else
        s2 = "Unknown OS"
    End Select
    sOutput = sOutput & s2 & " version: " & s1 & vbCrLf
    'frmsys.lblInformation(1) = s0
  End If
  l0 = GetPhysicalMemory(d0, d1, MEGABYTES)
  sOutput = sOutput & "Total physical memory available: " & Format$(d0, "#,###0.00 Mb ") & vbCrLf
  sOutput = sOutput & "Free physical memory available: " & Format$(d1, "#,###0.00 Mb ") & vbCrLf
  sOutput = sOutput & "Overall memory usage: " & CStr(l0) & "% "
  'frmsys.lblSysInfo(2).Caption = Format$(d0, "#,###0.00 Mb ")
  'frmsys.lblSysInfo(3).Caption = Format$(d1, "#,###0.00 Mb ")
  'frmsys.lblSysInfo(8).Caption = CStr(l0) & "% "
'  s0 = UCase$(Left$(mHomeDirectory, 3))
'  ret = GetDiskSpaceEx(s0, d0, d1, d2, MEGABYTES)
''  frmsys.lblInformation(5).Visible = True
''  frmsys.lblSysInfo(4).Visible = True
''  frmsys.lblSysInfo(6).Visible = True
'  If ret Then
'    sOutput = sOutput & "Application drive " & s0 & vbCrLf
'    sOutput = sOutput & Format$(d0, "#,###0.00 Mb ") & vbCrLf
'    sOutput = sOutput & Format$(d1, "#,###0.00 Mb ") & vbCrLf
''    frmsys.lblInformation(5).Caption = "Application drive " & s0
''    frmsys.lblSysInfo(4).Caption = Format$(d0, "#,###0.00 Mb ")
''    frmsys.lblSysInfo(6).Caption = Format$(d1, "#,###0.00 Mb ")
'  Else
''    sOutput = sOutput & "Application drive " & s0 & vbCrLf
''    sOutput = sOutput & "Unavailable" & vbCrLf
''    sOutput = sOutput & "Unavailable" & vbCrLf
''    frmsys.lblInformation(5).Caption = "Application drive " & s0
''    frmsys.lblSysInfo(4).Caption = "Unavailable"
''    frmsys.lblSysInfo(6).Caption = "Unavailable"
'  End If
'  s1 = UCase$(Left$(CurDir$, 3))
'  ret = ret And GetDiskSpaceEx(s1, d0, d1, d2, MEGABYTES)
'  If StrComp(s1, s0, vbTextCompare) <> 0 And ret Then
''    frmsys.lblInformation(6).Visible = True
''    frmsys.lblStatic(1).Visible = True
''    frmsys.lblStatic(3).Visible = True
''    frmsys.lblSysInfo(5).Visible = True
''    frmsys.lblSysInfo(7).Visible = True
'    sOutput = sOutput & "Current drive " & s0 & vbCrLf
'    sOutput = sOutput & Format$(d0, "#,###0.00 Mb ") & vbCrLf
'    sOutput = sOutput & Format$(d1, "#,###0.00 Mb ") & vbCrLf
''    frmsys.lblInformation(6).Caption = "Current drive " & s0
''    frmsys.lblSysInfo(5).Caption = Format$(d0, "#,###0.00 Mb ")
''    frmsys.lblSysInfo(7).Caption = Format$(d1, "#,###0.00 Mb ")
'  Else
''    frmsys.lblInformation(6).Visible = False
''    frmsys.lblStatic(1).Visible = False
''    frmsys.lblStatic(3).Visible = False
''    frmsys.lblSysInfo(5).Visible = False
''    frmsys.lblSysInfo(7).Visible = False
'  End If
'  Set locInfo = New LocaleInfo
'  s0 = "System Locale ID " & locInfo.GetSystemDefaultLcid & vbCrLf
'  s0 = s0 & "Country " & locInfo.GetLocaleValue(LOCALE_SYSTEM_DEFAULT, LOCALE_SENGCOUNTRY) & vbCrLf
'  s0 = s0 & "Language " & locInfo.GetLocaleValue(LOCALE_SYSTEM_DEFAULT, LOCALE_SENGLANGUAGE) & vbCrLf
'  s0 = s0 & "Currency " & locInfo.GetLocaleValue(LOCALE_SYSTEM_DEFAULT, LOCALE_SCURRENCY) & " (" & locInfo.GetLocaleValue(LOCALE_USER_DEFAULT, LOCALE_SINTLSYMBOL) & ")" & vbCrLf
'  s0 = s0 & "Short Date " & locInfo.GetLocaleValue(LOCALE_SYSTEM_DEFAULT, LOCALE_SSHORTDATE) & vbCrLf
'  s0 = s0 & "Long Date " & locInfo.GetLocaleValue(LOCALE_SYSTEM_DEFAULT, LOCALE_SLONGDATE) & vbCrLf
'  frmsys.lblLocaleSys.Caption = s0
'
'  s0 = "User Locale ID " & locInfo.GetUserDefaultLcid & vbCrLf
'  s0 = s0 & "Country " & locInfo.GetLocaleValue(LOCALE_USER_DEFAULT, LOCALE_SENGCOUNTRY) & vbCrLf
'  s0 = s0 & "Language " & locInfo.GetLocaleValue(LOCALE_USER_DEFAULT, LOCALE_SENGLANGUAGE) & vbCrLf
'  s0 = s0 & "Currency " & locInfo.GetLocaleValue(LOCALE_USER_DEFAULT, LOCALE_SCURRENCY) & " (" & locInfo.GetLocaleValue(LOCALE_USER_DEFAULT, LOCALE_SINTLSYMBOL) & ")" & vbCrLf
'  s0 = s0 & "Short Date " & locInfo.GetLocaleValue(LOCALE_USER_DEFAULT, LOCALE_SSHORTDATE) & vbCrLf
'  s0 = s0 & "Long Date " & locInfo.GetLocaleValue(LOCALE_USER_DEFAULT, LOCALE_SLONGDATE) & vbCrLf
'  frmsys.lblLocalUser.Caption = s0
  GetEnvironmentInfo = sOutput
  
GetEnvironmentInfo_exit:
  Exit Function
GetEnvironmentInfo_err:
  GetEnvironmentInfo = "Error retrieving environment information"
  Resume GetEnvironmentInfo_exit
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

Public Function GetWindowsVersion(lMajorVer As Long, lMinorVer As Long, lBuild As Long, PlatformID As OS_TYPE, sCSDVersion As String) As Boolean
  Dim lpVI As OSVERSIONINFO
  
  lpVI.dwOSVersionInfoSize = Len(lpVI)  ' 148
  GetWindowsVersion = GetVersionEx(lpVI) <> 0
  If GetWindowsVersion Then
    lMajorVer = lpVI.dwMajorVersion
    lMinorVer = lpVI.dwMinorVersion
    lBuild = lpVI.dwBuildNumber
    PlatformID = OS_UNKNOWN
    If lpVI.dwPlatformId = VER_PLATFORM_WIN32_NT Then
      If lMajorVer = 4 Then
        PlatformID = OS_NT4
      Else
        PlatformID = OS_W2000
      End If
    ElseIf lpVI.dwPlatformId = VER_PLATFORM_WIN32_WINDOWS Then
      If lMinorVer = 0 Then
        PlatformID = OS_WIN95
      Else
        PlatformID = OS_WIN98
      End If
    End If
    'AM RTrimChar not available
    sCSDVersion = Replace(lpVI.szCSDVersion, vbNullChar, "")
    'sCSDVersion = RTrimChar(lpVI.szCSDVersion, vbNullChar)
  End If
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

Public Function RoundDouble(ByVal Number As Double, ByVal DecimalPlaces As Long, ByVal rType As ROUND_TYPE) As Double
  Dim d As Double
  Dim TenPow As Double
  
  If (DecimalPlaces >= LOW_POW) And (DecimalPlaces <= HIGH_POW) Then
    TenPow = Powers(DecimalPlaces)
  Else
    TenPow = 10 ^ DecimalPlaces
  End If
  
  Select Case rType
    Case R_NORMAL
      RoundDouble = Int((Number * TenPow) + 0.5) / TenPow
    Case R_UP, R_DOWN
      d = Number * TenPow
      If Int(d) <> d Then
        If rType = R_UP Then d = d + 1
        RoundDouble = Int(d) / TenPow
      Else
        RoundDouble = d
      End If
    Case R_BANKERS
      RoundDouble = CLng(Number * TenPow) / TenPow
  End Select
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

Private Function ULIToDouble(ULI As ULARGE_INTEGER) As Double
  ULIToDouble = UnsignedToDouble(ULI.LowLong) + (TWO_POW_32 * UnsignedToDouble(ULI.HighLong))
End Function

Function UnsignedToDouble(ByVal Value As Long) As Double
  If Value < 0 Then
    UnsignedToDouble = -(&H80000000 - Value)
    UnsignedToDouble = UnsignedToDouble + &H7FFFFFFF
  Else
    UnsignedToDouble = Value
  End If
End Function

Public Sub MathInit()
  Dim i As Long
  For i = LOW_POW To HIGH_POW
    Powers(i) = 10 ^ i
  Next i
End Sub


Public Function FileExists(sFname As String) As Boolean
  Dim Attrs As Long
  On Error GoTo fileexists_err
  
  FileExists = False
  Attrs = GetAttr(sFname)
  If (Attrs And vbNormal) = 0 Then
      FileExists = True
  Else
      FileExists = False
  End If

fileexists_end:
  Exit Function
  
fileexists_err:
  FileExists = False
  Resume fileexists_end
End Function

