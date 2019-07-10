Attribute VB_Name = "Notify"
Option Explicit

Public Function ProvideFeedback(lCurrent As Long, lMax As Long, sMessage As String)
  If Not gNotify Is Nothing Then
    Call gNotify.Notify(lCurrent, lMax, sMessage)
  End If
End Function
