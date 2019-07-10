VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IErrorFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'* Log error messages return true to allow the normal error handler to
'* handle the error
Public Function FilterErrorMessage(UserName As String, DateTime As Date, ErrorNumber As Long, ErrorName As String, ErrorText As String) As Boolean
Attribute FilterErrorMessage.VB_Description = "Log error messages return true to allow the normal error handler to handle the error"

End Function

Public Function ViewAvailable() As Boolean

End Function

Public Function PrintAvailable() As Boolean

End Function

Public Function ViewErrors() As Boolean

End Function

Public Function PrintErrors() As Boolean

End Function

