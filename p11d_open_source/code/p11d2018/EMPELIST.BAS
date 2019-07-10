Attribute VB_Name = "EmployeeList"
Option Explicit
Private m_MAXINDEX As Long
Private m_CURINDEX As Long
Private Const Increment As Long = 500
Private m_Employees() As clsEmployee

Public Function AddEmployee(NewEmployee As clsEmployee) As Long
  On Error GoTo AddEmployee_Err
  Call xSet("AddEmployee")
  m_CURINDEX = m_CURINDEX + 1
  If m_CURINDEX > m_MAXINDEX Then
    m_MAXINDEX = m_MAXINDEX + Increment
    ReDim Preserve m_Employees(1 To m_MAXINDEX) As clsEmployee
  End If
  Set m_Employees(m_CURINDEX) = NewEmployee
  AddEmployee = m_CURINDEX
  
AddEmployee_End:
  Call xReturn("AddEmployee")
  Exit Function

AddEmployee_Err:
  AddEmployee = -1
  Call ErrorMessage(ERR_ERROR, Err, "AddEmployee", "Adding Employee", "Error adding new Employee to Employee list")
  Resume AddEmployee_End
End Function


Public Function GetEmployee(Index As Long) As clsEmployee
  On Error GoTo GetEmployee_Err
'  Call xSet("GetEmployee")
  
  If Index > 0 And Index <= m_CURINDEX Then
    Set GetEmployee = m_Employees(Index)
  End If
  
GetEmployee_End:
'  Call xReturn("GetEmployee")
  Exit Function

GetEmployee_Err:
  Call ErrorMessage(ERR_ERROR, Err, "GetEmployee", "ERR_UNDEFINED", "Undefined error.")
  Resume GetEmployee_End
End Function

Public Function GetEmployeeIndex(ee As clsEmployee) As Long
  Dim i As Long
  On Error GoTo GetEmployeeIndex_Err
  
  GetEmployeeIndex = -1
  For i = 1 To m_CURINDEX
    If m_Employees(i) Is ee Then
      GetEmployeeIndex = i
      Exit For
    End If
  Next i
  
GetEmployeeIndex_End:
  Exit Function

GetEmployeeIndex_Err:
  Resume GetEmployeeIndex_End
End Function


Public Sub ClearEmployeeList()
  Dim i As Long
  On Error GoTo ClearEmployeeList_Err
  
  Call xSet("ClearEmployeeList")
  For i = m_CURINDEX To 1 Step -1
    Set m_Employees(i) = Nothing
  Next i
  m_CURINDEX = 0
  
ClearEmployeeList_End:
  Call xReturn("ClearEmployeeList")
  Exit Sub

ClearEmployeeList_Err:
  Call ErrorMessage(ERR_ERROR, Err, "ClearEmployeeList", "ERR_UNDEFINED", "Undefined error.")
  Resume ClearEmployeeList_End
End Sub


Public Function EmployeeCount() As Long
  On Error GoTo EmployeeCount_Err
  Call xSet("EmployeeCount")
  
  EmployeeCount = m_CURINDEX
  
EmployeeCount_End:
  Call xReturn("EmployeeCount")
  Exit Function

EmployeeCount_Err:
  Call ErrorMessage(ERR_ERROR, Err, "EmployeeCount", "ERR_UNDEFINED", "Undefined error.")
  Resume EmployeeCount_End
End Function

Public Sub RemoveEmployee(ByVal Index As Long)
  If (Index < 1) Or (Index > m_CURINDEX) Then Call Err.Raise(380) 'apf
  Set m_Employees(Index) = Nothing
End Sub

