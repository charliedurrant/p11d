Attribute VB_Name = "MainModule"
Option Explicit


Public Sub Main()
  Dim Message As String, Title As String

  Message = "Version:  " & App.Major & "." & App.Minor & "." & App.Revision

  Title = App.Title

  Call MsgBox(Message, vbOKOnly + vbInformation, Title)

End Sub
