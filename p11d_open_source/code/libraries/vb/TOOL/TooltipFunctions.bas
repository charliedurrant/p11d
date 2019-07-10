Attribute VB_Name = "TooltipFunctions"
Option Explicit

' API declarations
'-----------------

' Ensures that the common control DLL is loaded
Public Declare Sub InitCommonControls Lib "comctl32.dll" ()
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hwndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Public Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal dwLength As Long)
Public Declare Function GetComctl32Version Lib "comctl32" Alias "DllGetVersion" (pdvi As DLLVersionInfo) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long


' Constants
' ---------

#Const WIN32_IE = &H400
Public Const TOOLTIPS_CLASS = "tooltips_class32"
Public Const WM_NOTIFY As Long = &H4E
Public Const WM_USER = &H400
Public Const GWL_STYLE As Long = (-16)
Public Const LVM_FIRST As Long = &H1
Public Const TV_FIRST As Long = &H1100
Public Const TCM_FIRST As Long = &H13
Public Const TVM_GETTOOLTIPS As Long = (TV_FIRST + 25)  ' Treeview constant
Public Const TB_GETTOOLTIPS = WM_USER + 35              ' Toolbar constant
Public Const RB_GETTOOLTIPS = WM_USER + 17              ' Rebar constants
Public Const TBM_GETTOOLTIPS = WM_USER + 30             ' Trackbar constants
Public Const LVM_GETTOOLTIPS = LVM_FIRST + 78           ' Listview constants
Public Const TCM_GETTOOLTIPS = TCM_FIRST + 45           ' TabControl constants
' Treeview style
Public Const TVS_NOTOOLTIPS = &H80
' Tooltip styles
Public Const TTS_ALWAYSTIP = &H1
Public Const TTS_NOPREFIX = &H2
Public Const TTS_BALLOON = &H40
' Comctl32.dll constants - PlatformID
Public Const DLL_PLATFORM_WINDOWS = &H1  ' built for all Windows versions
Public Const DLL_PLATFORM_NT = &H2       ' built specifically for NT


' Structures
'------------

Public Type DLLVersionInfo
  cbSize As Long
  dwMajorVersion As Long
  dwMinorVersion As Long
  dwBuildNumber As Long
  dwPlatformID As Long
End Type

' The pointer to this structure is specified as the lParam of the WM_NOTIFY message
Public Type NMHDR
  hwndFrom As Long   ' Window handle of control sending message
  idFrom As Long     ' Identifier of control sending message
  code  As Long      ' Notification code
End Type

Public Type POINTAPI
  x As Long
  y As Long
End Type

Public Type RECT
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type

Public Type TOOLINFO
  cbSize As Long
  uFlags As TT_Flags
  hWnd As Long
  uId As Long
  RECT As RECT
  hinst As Long
  lpszText As String   ' Long
#If (WIN32_IE >= &H300) Then
  lParam As Long
#End If
End Type

Public Type TTHITTESTINFO
  hWnd As Long
  pt As POINTAPI
  ti As TOOLINFO
End Type

Public Type NMTTDISPINFO
  hdr As NMHDR
  lpszText As Long
#If UNICODE Then
  szText As String * 160
#Else
  szText As String * 80
#End If
  hinst As Long
  uFlags As Long
#If (WIN32_IE >= &H300) Then
  lParam As Long
#End If
End Type


' Enumerations
' ------------

Public Enum TT_Flags
  TTF_IDISHWND = &H1
  TTF_CENTERTIP = &H2
  TTF_RTLREADING = &H4
  TTF_SUBCLASS = &H10
#If (WIN32_IE >= &H300) Then
  TTF_TRACK = &H20
  TTF_ABSOLUTE = &H80
  TTF_TRANSPARENT = &H100
  TTF_DI_SETITEM = &H8000&        ' valid only on the TTN_NEEDTEXT callback
#End If
End Enum

Public Enum TT_DelayTime
  TTDT_AUTOMATIC = 0
  TTDT_RESHOW = 1
  TTDT_AUTOPOP = 2
  TTDT_INITIAL = 3
End Enum

Public Enum TT_Msgs
  TTM_ACTIVATE = (WM_USER + 1)
  TTM_SETDELAYTIME = (WM_USER + 3)
  TTM_RELAYEVENT = (WM_USER + 7)
  TTM_GETTOOLCOUNT = (WM_USER + 13)
  TTM_WINDOWFROMPOINT = (WM_USER + 16)
    
#If UNICODE Then
  TTM_ADDTOOL = (WM_USER + 50)
  TTM_DELTOOL = (WM_USER + 51)
  TTM_NEWTOOLRECT = (WM_USER + 52)
  TTM_GETTOOLINFO = (WM_USER + 53)
  TTM_SETTOOLINFO = (WM_USER + 54)
  TTM_HITTEST = (WM_USER + 55)
  TTM_GETTEXT = (WM_USER + 56)
  TTM_UPDATETIPTEXT = (WM_USER + 57)
  TTM_ENUMTOOLS = (WM_USER + 58)
  TTM_GETCURRENTTOOL = (WM_USER + 59)
#Else
  TTM_ADDTOOL = (WM_USER + 4)
  TTM_DELTOOL = (WM_USER + 5)
  TTM_NEWTOOLRECT = (WM_USER + 6)
  TTM_GETTOOLINFO = (WM_USER + 8)
  TTM_SETTOOLINFO = (WM_USER + 9)
  TTM_HITTEST = (WM_USER + 10)
  TTM_GETTEXT = (WM_USER + 11)
  TTM_UPDATETIPTEXT = (WM_USER + 12)
  TTM_ENUMTOOLS = (WM_USER + 14)
  TTM_GETCURRENTTOOL = (WM_USER + 15)
#End If

#If (WIN32_IE >= &H300) Then
  TTM_TRACKACTIVATE = (WM_USER + 17)       ' wParam = TRUE/FALSE start end  lparam = LPTOOLINFO
  TTM_TRACKPOSITION = (WM_USER + 18)       ' lParam = dwPos
  TTM_SETTIPBKCOLOR = (WM_USER + 19)
  TTM_SETTIPTEXTCOLOR = (WM_USER + 20)
  TTM_GETDELAYTIME = (WM_USER + 21)
  TTM_GETTIPBKCOLOR = (WM_USER + 22)
  TTM_GETTIPTEXTCOLOR = (WM_USER + 23)
  TTM_SETMAXTIPWIDTH = (WM_USER + 24)
  TTM_GETMAXTIPWIDTH = (WM_USER + 25)
  TTM_SETMARGIN = (WM_USER + 26)           ' lParam = lprc
  TTM_GETMARGIN = (WM_USER + 27)           ' lParam = lprc
  TTM_POP = (WM_USER + 28)
#End If

#If (WIN32_IE >= &H400) Then
  TTM_UPDATE = (WM_USER + 29)
#End If
End Enum

Public Enum TT_Notifications
  TTN_FIRST = -520&   '   (0U-520U)
  TTN_LAST = -549&    '   (0U-549U)
#If UNICODE Then
  TTN_NEEDTEXT = (TTN_FIRST - 10)   ' is now TTN_GETDISPINFO
#Else
  TTN_NEEDTEXT = (TTN_FIRST - 0)
#End If   ' UNICODE
  TTN_SHOW = (TTN_FIRST - 1)
  TTN_POP = (TTN_FIRST - 2)
End Enum


' Functions
' ---------

' Returns the low-order word from the given 32-bit value.
Public Function LOWORD(dwValue As Long) As Integer
  CopyMemory LOWORD, dwValue, 2
End Function

' Returns the larger of the two passed paramaters
Public Function Max(param1 As Long, param2 As Long) As Long
  If param1 > param2 Then Max = param1 Else Max = param2
End Function

Public Function GetStrFromBufferA(szA As String) As String
  If InStr(szA, vbNullChar) Then
    GetStrFromBufferA = Left$(szA, InStr(szA, vbNullChar) - 1)
  Else
    ' If sz had no null char, the Left$ function
    ' above would returnn a zero length string ("")
    GetStrFromBufferA = szA
  End If
End Function


