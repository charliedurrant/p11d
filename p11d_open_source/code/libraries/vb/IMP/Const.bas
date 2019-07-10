Attribute VB_Name = "Const"
Option Explicit

Public Const DATABASE_FIELD_TYPES_COUNT As Long = 5

Public Enum TCSIMP_UDE
  ERR_IMPORT = TCSIMPORT_ERROR + 1
  ERR_PARSETOKEN
  ERR_IMPORTINIT
  ERR_PARSEBOOLEAN
  ERR_IMPORTMASK
  ERR_ADDFAILED
  ERR_NOIMPDEST
  ERR_IMPORTEMPTY
  ERR_IMPORTFIELDSIZE
  ERR_IMPORTWIZARD
  ERR_RESTOREIMPORT
  ERR_FIXUPCOLUMNS
  ERR_ADDPRIMARYFIELD
  ERR_ADDREQUIREDFIELD
  ERR_REQUIREDFIELD
  ERR_INVALID_UPDATETYPE
  ERR_PREVIEW_UPDATES
  ERR_PRIMARY_DEST
  ERR_IMPORTUPDATEONLY
  ERR_UPDATEPRIMARYONLY
  ERR_DEFSTATIC
  ERR_DEFCOPIEDFIELD
  ERR_RSNOTUPDATEABLE
  ERR_NO_DESTRS
  ERR_ADDCONSTRAINT
  ERR_VALIDATEDEST
  ERR_IMPORTFW
  ERR_NOLINK
End Enum
                      
Public Type FlexGridDropInfo
  FromForm As String
  FromFG As String
  FromCol As Long
  FromRow As Long
  ToForm As String
  ToFG As String
  ToCol As Long
  ToRow As Long
End Type

Public Const DEFAULT_TEXT_DELIMITER As String = "."
Public m_ImportWizard As ImportWizard
Public Const FIXED_COLCOUNT As Long = 1
Public Const FIXED_ROWCOUNT As Long = 2
Public Const MISC_FIXED_ROWCOUNT As Long = FIXED_ROWCOUNT
Public Const LINK_FIXED_ROWCOUNT As Long = FIXED_ROWCOUNT

Public mDisableRecalc As Boolean
