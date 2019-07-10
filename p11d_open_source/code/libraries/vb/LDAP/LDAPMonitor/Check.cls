VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Check"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements ITimerProc
Public SvrContext As String
Public DebugServerContext As String
Public Username As String
Public Password As String
Public Monfrm As frmLDAPMon

Private Sub ITimerProc_OnTimer(ByVal SysTime As Long)
  Dim LD As LDAPHelper
  
  On Error GoTo ITimerProc_OnTimer_err
  If Len(SvrContext) = 0 Then
    SvrContext = "<tcsau5001.arthurandersen.com;tcsau5003.arthurandersen.com;tcsau5004.arthurandersen.com;tcsau5005.arthurandersen.com;tcsau5007.arthurandersen.com;tcsau5008.arthurandersen.com;>"
  End If
  Set LD = New LDAPHelper
  LD.ServerContext = SvrContext
  If Not LD.Authenticate(Username, Password, ADS_STANDARD, False) Then
    
  End If
  SvrContext = LD.ServerContext
  DebugServerContext = LD.DebugServerContext
  Monfrm.lblDebugContext = LD.DebugServerContext
  Exit Sub
  
ITimerProc_OnTimer_err:
  Monfrm.lblDebugContext = "Error accessing ldap servers: " & Err.Description
End Sub
