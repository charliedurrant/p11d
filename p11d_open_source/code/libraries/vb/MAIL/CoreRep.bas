Attribute VB_Name = "CoreRep"
Option Explicit
'This Module includes replication of Core constants

'Get System directory function
Public Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long

'System directory (when not found by GetSystemDirectory)
Public Const WINDIR As String = "C:\WINDOWS"

'Environment info
Public Enum OS_TYPE
  OS_UNKNOWN = 1
  OS_WIN95
  OS_WIN98
  OS_NT4
  OS_W2000
End Enum
Public Const VER_PLATFORM_WIN32_WINDOWS As Long = 1
Public Const VER_PLATFORM_WIN32_NT As Long = 2

Public Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long

Public Declare Sub GetSystemInfo Lib "kernel32" (lpSystemInfo As SYSTEM_INFO)

Public Declare Sub GlobalMemoryStatus Lib "kernel32" (lpBuffer As MEMORYSTATUS)

Public Enum MemoryUnit
  BYTES
  MEGABYTES
  GIGABYTES
End Enum

Public Const LOW_POW As Long = -10
Public Const HIGH_POW As Long = 10
Public Powers(LOW_POW To HIGH_POW) As Double

Public Const TWO_POW_32 As Double = 4294967296#
Public Const TWO_POW_10 As Double = 1024
Public Const TWO_POW_20 As Double = 1048576#
Public Const TWO_POW_30 As Double = 1073741824#

Public Enum ROUND_TYPE
  R_NORMAL
  R_UP
  R_DOWN
  R_BANKERS
End Enum

Public mHomeDirectory As String

Public Declare Function GetDiskFreeSpaceEx Lib "kernel32" Alias "GetDiskFreeSpaceExA" (ByVal lpDirectoryName As String, lpFreeBytesAvailableToCaller As ULARGE_INTEGER, lpTotalNumberOfBytes As ULARGE_INTEGER, lpTotalNumberOfFreeBytes As ULARGE_INTEGER) As Long

Public Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
    '  Maintenance string for PSS usage
End Type

Public Type MEMORYSTATUS
    dwLength As Long
    dwMemoryLoad As Long
    dwTotalPhys As Long
    dwAvailPhys As Long
    dwTotalPageFile As Long
    dwAvailPageFile As Long
    dwTotalVirtual As Long
    dwAvailVirtual As Long
End Type

Public Type ULARGE_INTEGER
  LowLong As Long
  HighLong As Long
End Type

Public Type SYSTEM_INFO
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

Public Const PROCESSOR_INTEL_386 As Long = 386
Public Const PROCESSOR_INTEL_486 As Long = 486
Public Const PROCESSOR_INTEL_PENTIUM As Long = 586
