Attribute VB_Name = "StartUp"
Option Explicit

Public Sub Main()
  If CoreSetup("", VB.Global) Then
    Set datest.dacon = New DAConnection
    datest.dacon.debugmode = True
  Else
    MsgBox "TCS core failed to initialise"
  End If

End Sub

