Attribute VB_Name = "EmployerList"
Option Explicit
Private m_MAXINDEX As Long
Private m_CURINDEX As Long
Private Const Increment As Long = 20
Private m_Employers() As clsEmployer

Public Function AddEmployer(NewEmployer As clsEmployer) As Long
  On Error GoTo AddEmployer_Err
  Call xSet("AddEmployer")
  m_CURINDEX = m_CURINDEX + 1
  If m_CURINDEX > m_MAXINDEX Then
    m_MAXINDEX = m_MAXINDEX + Increment
    ReDim Preserve m_Employers(1 To m_MAXINDEX) As clsEmployer
  End If
  Set m_Employers(m_CURINDEX) = NewEmployer
  AddEmployer = m_CURINDEX
AddEmployer_End:
  Call xReturn("AddEmployer")
  Exit Function

AddEmployer_Err:
  AddEmployer = -1
  Call ErrorMessage(ERR_ERROR, Err, "AddEmployer", "Adding Employer", "Error adding new employer to employer list")
  Resume AddEmployer_End
End Function


Public Function GetEmployer(index As Long) As clsEmployer
  On Error GoTo GetEmployer_Err
  Call xSet("GetEmployer")
  
  If index > 0 And index <= m_CURINDEX Then
    Set GetEmployer = m_Employers(index)
  End If
  
GetEmployer_End:
  Call xReturn("GetEmployer")
  Exit Function

GetEmployer_Err:
  Call ErrorMessage(ERR_ERROR, Err, "GetEmployer", "ERR_UNDEFINED", "Undefined error.")
  Resume GetEmployer_End
End Function
Public Sub RemoveEmployer(index As Long)
  If (index < 1) Or (index > m_CURINDEX) Then Call Err.Raise(380) 'apf
  Set m_Employers(index) = Nothing
End Sub

Public Sub ClearEmployerList()
  Dim i As Long
  On Error GoTo ClearEmployerList_Err
  Call xSet("ClearEmployerList")
  For i = 1 To m_CURINDEX
    Set m_Employers(i) = Nothing
  Next i
  m_CURINDEX = 0
  
ClearEmployerList_End:
  Call xReturn("ClearEmployerList")
  Exit Sub

ClearEmployerList_Err:
  Call ErrorMessage(ERR_ERROR, Err, "ClearEmployerList", "ERR_UNDEFINED", "Undefined error.")
  Resume ClearEmployerList_End
End Sub



