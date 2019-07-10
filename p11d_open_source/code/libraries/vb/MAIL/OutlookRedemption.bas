Attribute VB_Name = "OutlookRedemption"
Option Explicit

Global Application As Object
Global Namespace As Object

Public Sub InitOutlookRedemption()

If Application Is Nothing Then
  Exit Sub
End If

Set Application = CreateObject("Outlook.Application")
Set Namespace = Application.GetNamespace("MAPI")

End Sub
