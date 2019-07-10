VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ImporterWebErrors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private m_Errors As Collection

Private Sub Class_Initialize()
  Set m_Errors = New Collection
End Sub

Public Property Get Count() As Long
  Count = m_Errors.Count
End Property

Public Property Get Item(ByVal vIndex As Variant) As WebError
  On Error GoTo Item_err
  Set Item = m_Errors.Item(vIndex)
  Exit Property
  
Item_err:
  Err.Raise 380, "ImporterWebErrors.Item", "Invalid Error index [" & vIndex & "]" & vbCrLf & Err.Description
End Property

Public Sub AddErr(ByVal Description As String, Optional ByVal ErrObj As ErrObject)
  Dim wErr As WebErrors
  
  If Not ErrObj Is Nothing Then
    Set wErr = New WebErrors
    wErr.Number = ErrObj.Number
    If Len(Description) = 0 Then
      wErr.Description = ErrObj.Description
    Else
      wErr.Description = Description & vbCrLf & ErrObj.Description
    End If
    wErr.Source = ErrObj.Source
    m_Errors.Add wErr
  Else
    Set wErr = New WebErrors
    wErr.Description = Description
    m_Errors.Add wErr
  End If
End Sub

Public Property Get ErrorsToHTML() As String
  Dim qsError As QString
  Dim wErr As WebErrors
  Dim i As Long
  Dim eCount As Long
  Dim errid As String
  
  Set qsError = New QString
  qsError.Append "<TABLE " & HTMLAttrib("id", "errormessage") & HTMLAttrib("class", "errortable") & ">"
  For i = 1 To m_Errors.Count
    Set wErr = m_Errors.Item(i)
    errid = "errorcallstack" & i
    qsError.Append "<TR><TD>"
    qsError.Append "<SPAN " & HTMLAttrib("class", "errordescription") & ">"
    qsError.Append wErr.ErrorDescription
    qsError.Append "</SPAN>"
    qsError.Append "</TD></TR>"
    qsError.Append "<TR><TD><BR></TD></TR>"
    qsError.Append "<TR><TD>"
    qsError.Append "<DIV " & HTMLAttrib("class", "errorcallstackheader") & HTMLAttrib("onclick", "flipdisplay(" & errid & ")") & ">"
    qsError.Append "Show error details"
    qsError.Append "</DIV>"
    qsError.Append "<DIV " & HTMLAttrib("id", errid) & HTMLAttrib("class", "errorcallstack") & ">"
    qsError.Append wErr.ErrorDetails
    qsError.Append "</DIV>"
    qsError.Append "</TD></TR>"
    qsError.Append "<TR><TD><BR></TD></TR>"
  Next i
  qsError.Append "</TABLE>"
  ErrorsToHTML = qsError.bstr
  Set qsError = Nothing
End Property
  
Public Property Get ErrorsToXML() As String
  Dim qsError As QString
  Dim wErr As WebErrors
  Dim i As Long
  Dim eCount As Long
  Dim errid As String
  
  Set qsError = New QString
  qsError.Append "<ERRORS>"
  For i = 1 To m_Errors.Count
    Set wErr = m_Errors.Item(i)
    qsError.Append "<ERROR>"
    qsError.Append "<DESCRIPTION>" & wErr.ErrorDescriptionXML & "</DESCRIPTION>"
    qsError.Append "<DETAILS>" & wErr.ErrorDetailsXML & "</DETAILS>"
    qsError.Append "</ERROR>"
  Next i
  qsError.Append "</ERRORS>"
  ErrorsToXML = qsError.bstr
  Set qsError = Nothing
End Property

