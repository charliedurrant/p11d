Attribute VB_Name = "Constant"
Option Explicit

' AUTOINSERTCOMMENT BEGIN STDGLOBALS
Public gbForceExit As Boolean
Public gbAllowAppExit As Boolean
' AUTOINSERTCOMMENT END STDGLOBALS

Public Enum ApplicationErrors
  ERR_APPLICATION = TCSCLIENT_ERROR
  'ERR_NEXTERROR = TCSCLIENT_ERROR + 1 etc...
End Enum

