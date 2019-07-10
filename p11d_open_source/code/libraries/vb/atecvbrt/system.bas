Attribute VB_Name = "system"
Option Explicit
Private Declare Function WNetOpenEnum Lib "mpr.dll" Alias "WNetOpenEnumA" (ByVal dwScope As Long, ByVal dwType As Long, ByVal dwUsage As Long, ByVal lpNetResource As Long, lphEnum As Long) As Long
Private Declare Function WNetCloseEnum Lib "mpr.dll" (ByVal hEnum As Long) As Long
Private Declare Function WNetEnumResource Lib "mpr.dll" Alias "WNetEnumResourceA" (ByVal hEnum As Long, lpcCount As Long, ByVal lpBuffer As Long, lpBufferSize As Long) As Long

Private Type ULARGE_INTEGER
  LowLong As Long
  HighLong As Long
End Type

Private Const TWO_POW_32 As Double = 4294967296#
Private Const TWO_POW_10 As Double = 1024
Private Const TWO_POW_20 As Double = 1048576#
Private Const TWO_POW_30 As Double = 1073741824#

Private Const LOW_POW As Long = -10
Private Const HIGH_POW As Long = 10
Private Powers(LOW_POW To HIGH_POW) As Double

Public Enum ROUND_TYPE
  R_NORMAL
  R_UP
  R_DOWN
  R_BANKERS
End Enum

Private Type NETRESOURCE
    dwScope As Long
    dwType As Long
    dwDisplayType As Long
    dwUsage As Long
    lpLocalName As Long
    lpRemoteName As Long
    lpComment As Long
    lpProvider As Long
End Type

Private Const RESOURCE_CONNECTED As Long = &H1
Private Const RESOURCE_GLOBALNET As Long = &H2
Private Const RESOURCE_REMEMBERED As Long = &H3

Private Const RESOURCETYPE_ANY As Long = &H0
Private Const RESOURCETYPE_DISK As Long = &H1
Private Const RESOURCETYPE_PRINT As Long = &H2
Private Const RESOURCETYPE_UNKNOWN As Long = &HFFFF

Private Const GMEM_DDESHARE As Long = &H2000
Private Const GMEM_DISCARDABLE As Long = &H100
Private Const GMEM_DISCARDED As Long = &H4000
Private Const GMEM_FIXED As Long = &H0
Private Const GMEM_INVALID_HANDLE As Long = &H8000
Private Const GMEM_SHARE As Long = &H2000
Private Const GMEM_MOVEABLE As Long = &H2
Private Const GMEM_ZEROINIT As Long = &H40
Private Const GPTR As Long = (GMEM_FIXED Or GMEM_ZEROINIT)
Private Const GHND As Long = (GMEM_MOVEABLE Or GMEM_ZEROINIT)

Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal dest As Long, ByVal Source As Long, ByVal cbCopy As Long)
Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
Private Declare Function GetDiskFreeSpaceEx Lib "kernel32" Alias "GetDiskFreeSpaceExA" (ByVal lpDirectoryName As String, lpFreeBytesAvailableToCaller As ULARGE_INTEGER, lpTotalNumberOfBytes As ULARGE_INTEGER, lpTotalNumberOfFreeBytes As ULARGE_INTEGER) As Long

Public Function GetDrivePathEx(ByVal sDirectoryOnly As String) As String
  Dim i As Long, p As Long
  Dim maxp As Long, maxi As Long
  Dim s As String
  Dim maps(1 To 27) As String
    
  On Error GoTo GetDrivePathEx_err
  If Len(sDirectoryOnly) > 0 Then
    If (Left$(sDirectoryOnly, 2) = "\\") Then
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
  GetDrivePathEx = FullPath(sDirectoryOnly)
  Exit Function
  
GetDrivePathEx_err:
  Err.Raise Err.Number, ErrorSourceEx(Err, "GetDrivePathEx"), Err.Description
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
        sremote = CopyLPSTRtoString(nr.lpRemoteName)
        maps(Drive) = sremote
        GetDriveMappings = GetDriveMappings + 1
      End If
    End If
  Next i
  Call GlobalFree(netresbuf)
  Call WNetCloseEnum(hEnum)
End Function


Public Function IsFileOpenEx(FileName As String, ByVal Exclusive As Boolean) As Boolean
  Dim i As Integer
  
  On Error GoTo IsFileOpenEx_ERR
  If FileExists(FileName, False, False) Then
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


Public Function FileCopyExN(ByVal Source As String, ByVal Destination As String, ByVal RaiseErrors As Boolean) As Boolean
  Dim DestDir As String
  Dim DestDirShort As String
  Dim tmpDestination As String, ret As Long
  Dim blen As Long
    
  On Error GoTo FileCopyExN_Err
  Dim fh As New FileHelper
  Call fh.SplitPath(Destination, DestDir)
  If Len(DestDir) = 0 Then
    Destination = FullPath(CurDir()) & Destination
    DestDir = FullPath(CurDir())
  End If
  
expand_buffer:
  blen = blen + TCSBUFSIZ
  DestDirShort = String$(TCSBUFSIZ, 0)
  ret = GetShortPathName(DestDir, DestDirShort, TCSBUFSIZ)
  If ret > TCSBUFSIZ Then GoTo expand_buffer
  If (ret = 0) And (Len(DestDir) > 0) Then Err.Raise ERR_FILECOPY, "FileCopyExN", "Unable to get short name for destination " & DestDir
  DestDirShort = RTrimChar(DestDirShort, vbNullChar)
  tmpDestination = GetTempFilename(DestDirShort, "FC")
  Call FileCopy(Source, tmpDestination)
  If Len(DestDirShort) > 0 Then Destination = DestDirShort & Mid$(Destination, Len(DestDir) + 1)
  Call fh.xKill(Destination)
  Name tmpDestination As Destination
  FileCopyExN = True
  
FileCopyExN_End:
  Exit Function
  
FileCopyExN_Err:
  If RaiseErrors Then
    Dim eStack As New ErrHelper
    Call eStack.Push(Err)
  End If
  FileCopyExN = False
  If FileExists(tmpDestination, False, False) Then Call fh.xKill(tmpDestination)
  If RaiseErrors Then
    Call eStack.Pop(Err)
    Err.Raise Err.Number, ErrorSourceEx(Err, "FileCopyExN"), Err.Description
  End If
  Resume FileCopyExN_End
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

Public Function RoundDouble(ByVal Number As Double, ByVal DecimalPlaces As Long, ByVal rType As ROUND_TYPE) As Double
  Dim d As Double
  Dim TenPow As Double
  
  Call MathInit
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

Public Sub MathInit()
  Dim i As Long
  
  If Powers(0) = 0 Then
    For i = LOW_POW To HIGH_POW
      Powers(i) = 10 ^ i
    Next i
  End If
End Sub



