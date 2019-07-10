Attribute VB_Name = "Const"
Option Explicit

' core classes for internal use of core functions
Public gCore As CoreClass
Public vbg As VB.Global
Public ghInstance As Long
Public ghThreadID As Long

Public gTCSEventClass As TCSEventClass
Public m_IPreErrorFilter As IErrorFilter
Public m_IPostErrorProcess As IErrorPostProcess

Public m_IDebugMenu As IDebugMenu
Public m_INotifyTCSPassword As IBaseNotify
Public mTCS_InitialiseDefaultWS As Boolean
Public m_FatalError As Boolean
Public m_HelpAboutText As String
Public m_OtherCaption As String
Public inErrMainMsg As Boolean
Public mFormattedErrorStrings As Boolean


Public Const contactstr As String = "For help, please contact abatec on (020) 7303 8102"
Public Const internalcontactstr As String = "For help, please contact abatec on 30543"
Public Const WINDIR As String = "C:\WINDOWS"
Public Const LogFileExt As String = ".LOG"
Public Const ErrFileExt As String = ".ERR"
Public Const NO_TIMEOUT As Long = 16777216
Public Const GLOBAL_LBCOLCHARS As Long = 25

'core
Public mCoreInitCount As Long
Public mCoreShutDownDone  As Boolean

Public mCoreTrace As Boolean
Public gstaticdata As Collection
Public LibraryVersions As Collection

Public mDBTarget As DATABASE_TARGET

Public mAppPath As String
Public mAppExeName As String
Public mAppVersion As String
Public mAppName As String
Public mAppCmdParam As String
Public mTCSCoreVersion As String
Public mSilentError As Boolean
Public mInitDefaultWS  As Boolean
Public mForceErrorTopMost  As Boolean

'Debug menu collection
Public gMenusCollection As Collection
Public gPasswordPrompt As String
Public gPasswordTitle As String

'TCS core defined errors
Public Enum TCSCORE_UDE
  CoreErrorRecursive = TCSCORE_ERROR + 1
  ERR_CONVERTDATE = TCSCORE_ERROR + 4  ' note: see VTEXT if changed
  ERR_NOFILEEXISTS
  ERR_NOVERSIONINFO
  ERR_INITAPP
  ERR_EXITAPP
  ERR_INITCORE
  ERR_CALLSTACKCORRUPT
  ERR_CALLSTACKERROR
  ERR_SYSTEM
  ERR_REPORT_STACK
  ERR_NO_INIT_REPORT
  ERR_REPORT_CANCEL
  ERR_INVALID_PREVIEW_OBJECT
  ERR_INVALID_ZOOM_VALUE
  ERR_FORM_ISNOTHING
  ERR_ADD_COL_FAIL
  ERR_FINDWINDOW
  ERR_ERRORLOGRS
  ERR_INVALIDOFFSET
  ERR_POPCURSOR
  ERR_INVALID_TYPE
  ERR_STARTTIMER
  ERR_ROUND
  ERR_DISPLAYMSG
  ERR_SORT
  ERR_VERCOMPARE
  ERR_MKDIR
  ERR_ACTIVATEPREV
  ERR_CMDPARAM
  ERR_GETVALUE
  ERR_DATESERIAL
  ERR_FINDFILES
  ERR_FILECOPY
  ERR_PRECISION
  
End Enum


'CommonControl styles
Public Const CCS_NODIVIDER As Long = &H40          'Prevents a toolbar from drawing a divider line on the top edge.

'Toolbar styles
Public Const TBSTYLE_FLAT As Long = &H800          'Makes toolbar buttons transparent and flat.

'Messages: Window messages
Public Const WM_PAINT As Long = &HF
Public Const WM_USER As Long = &H400&

'Toolbar messages
Public Const TB_GETSTYLE As Long = (WM_USER + 57)
Public Const TB_SETSTYLE As Long = (WM_USER + 56)
Public Const STATICS_SECTION As String = "STATICDATA"
Public mStaticFileName As String
Public mHomeDirectory As String

Public Sub Main()
  Call MathInit
  mFormattedErrorStrings = False
End Sub
