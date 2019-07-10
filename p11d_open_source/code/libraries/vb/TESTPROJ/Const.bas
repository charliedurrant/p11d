Attribute VB_Name = "Const"
Option Explicit
Public gbForceExit As Boolean
Public gbAllowAppExit As Boolean
Public gTBCtrl As ToolBarControl
Public gbMDILoaded As Boolean

Public Enum ApplicationErrors
  ERR_CORESETUP = TCSCLIENT_ERROR
  ERR_ADDBUTTONIMAGE
  ERR_TESTERROR
  'ERR_NEXTERROR etc...
End Enum
' END --- TEMPLATE CODE

Public gMAPI As Mail
Public RW As ReportWizard

