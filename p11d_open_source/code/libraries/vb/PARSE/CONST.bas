Attribute VB_Name = "CONST"
Option Explicit

Public Const MAX_TOKEN_LEN As Long = 128

'TCS core defined errors
Public Enum TCSPARSE_UDE
  ERR_PARSESETTINGS = TCSPARSER_ERROR + 1
  ERR_PARSETOKEN
  ERR_TOKENPARAMETERS
  ERR_PARSEDUPLICATEITEM
  ERR_PARSEMODE
  ERR_SETPARSESETTINGS
End Enum

'Public Function GetHashValue(s As String) As Long
'  Dim sbyte() As Byte
'  Dim i As Long, tlen As Long
'
'  tlen = (Len(s) * 2) - 1
'  sbyte = s
'  For i = 0 To tlen Step 2
'    GetHashValue = GetHashValue + (sbyte(i) * (i + 1))
'  Next i
'  'GetHashValue = GetHashValue + (sbyte(0) * sbyte(tlen))
'End Function
