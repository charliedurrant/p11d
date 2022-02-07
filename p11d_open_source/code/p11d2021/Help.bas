Attribute VB_Name = "HelpLink"
Option Explicit


Option Explicit

Public Const HH_DISPLAY_TOPIC = &H0
Public Const HH_DISPLAY_TEXT_POPUP = &HE
Public Const HH_HELP_CONTEXT = &HF
Public Const HH_CLOSE_ALL = &H12

Public Declare Function HtmlHelp Lib "hhctrl.ocx" Alias "HtmlHelpA" _
   (ByVal hwndCaller As Long, _
    ByVal pszFile As String, _
    ByVal uCommand As Long, _
    ByVal dwData As Long) As Long
    




