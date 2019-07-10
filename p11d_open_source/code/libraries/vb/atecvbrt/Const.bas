Attribute VB_Name = "Const"
Option Explicit

Public Enum TCSCORE_UDE
  ERR_SORT = ABATEC_VBRT
  ERR_GETVALUE
  ERR_INITIALISE
  ERR_VERQUERY
  ERR_MKDIR
  ERR_FINDFILES
  ERR_FILEHELPER
  ERR_FILECOPY
  ERR_SYSHELPER
  ERR_STATIC
  ERR_DATACONV
  ERR_DATESERIAL
  ERR_CONVERTDATE
  ERR_INVALID_TYPE
  ERR_DBHELPER
  ERR_INVALIDOFFSET
  ERR_INIFILECACHE
End Enum

Public Const CONTACTSTR As String = "For help, please contact abatec on (020) 7438 3669"
Public Const INTERNALCONTACTSTR As String = "For help, please contact abatec on 45544"
Public Const LogFileExt As String = ".LOG"
Public Const ErrFileExt As String = ".ERR"

