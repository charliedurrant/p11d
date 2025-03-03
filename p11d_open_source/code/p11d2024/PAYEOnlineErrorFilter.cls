VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PayeOnlineErrorFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IErrorFilter
Private m_Filter As IPAYEOnlineOnError

Private m_ErrorCount As Long

Public Property Set Filter(value As IPAYEOnlineOnError)
  Set m_Filter = value
End Property
Public Property Get Filter() As IPAYEOnlineOnError
  Set Filter = m_Filter
End Property
Public Function IErrorFilter_FilterErrorMessage(ByVal Username As String, ByVal DateTime As Date, ByVal ErrorNumber As Long, ByVal ErrorName As String, ByVal ErrorText As String, ByVal SourceFunction As String) As Boolean
  Dim error As PAYEOnlineError
  Dim s() As String
  Dim i As Long
  
  m_ErrorCount = m_ErrorCount + 1
  
  If Not m_Filter Is Nothing Then
    
    Set error = New PAYEOnlineError
    
    error.Number = ErrorNumber
    error.Description = ErrorText
    
    Call GetDelimitedValues(s, ErrorName)
    For i = LBound(s) To UBound(s)
      Select Case i
        Case 1
          error.Employer = s(1)
        Case 2
          error.Employee = s(2)
        Case 3
          error.HMITSection = s(3)
        Case 4
          error.benefit = s(4)
        Case 5
          error.BenefitType = s(5)
      End Select
    Next
    
    Call m_Filter.OnError(m_ErrorCount, error)
  End If
End Function
Public Property Get ErrorCount() As Long
  ErrorCount = m_ErrorCount
End Property
Private Function IErrorFilter_PrintAvailable() As Boolean

End Function

Private Function IErrorFilter_PrintErrors() As Boolean

End Function

Private Function IErrorFilter_ViewAvailable() As Boolean

End Function

Private Function IErrorFilter_ViewErrors() As Boolean

End Function
