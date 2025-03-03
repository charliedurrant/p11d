VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SharedVans"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Option Base 0
Implements IBenefitClass

Private m_BenClass As BEN_CLASS
Private m_Parent As Employer
Private m_ReadFromDB As Boolean
Public Vans As ObjectList
Private m_dirty As Boolean
Private m_InvalidFields As Long
Private m_NeedToCalculate As Boolean
Private m_shvan_benefit_van_only As Variant
Private m_shvan_fuel_benefit As Variant
Private m_shvan_value_van_only As Variant
Private m_shvan_fuel_value As Variant


Private Sub Class_Terminate()
 'Debug.Print"Shared vans - terminate"
End Sub

Private Function IBenefitClass_CalculateBody() As Variant

End Function

Private Property Get IBenefitClass_ImageListKey() As String

End Property

Private Property Let IBenefitClass_NeedToCalculate(ByVal RHS As Boolean)
  m_NeedToCalculate = NeedToCalculateHelper(Me, RHS)
End Property

Private Property Get IBenefitClass_NeedToCalculate() As Boolean
  IBenefitClass_NeedToCalculate = m_NeedToCalculate
End Property

Private Property Let IBenefitClass_LinkBen(RHS As Boolean)

End Property

Private Property Get IBenefitClass_LinkBen() As Boolean

End Property

Private Function IBenefitClass_CanBringForward() As Boolean

End Function

Private Function IBenefitClass_Copy(Parent As Object) As IBenefitClass

End Function
Private Sub Class_Initialize()
  Set Vans = New ObjectList
End Sub
Private Property Get IBenefitClass_BenefitClass() As BEN_CLASS
  IBenefitClass_BenefitClass = BC_SHAREDVANs_G
End Property
Private Property Let IBenefitClass_BenefitClass(NewValue As BEN_CLASS)
End Property
Public Function IBenefitClass_Calculate() As Variant
  Dim lTotal  As Long, lTempTotal As Long
  Dim SharedVan As IBenefitClass
  
  Dim benEmployer As IBenefitClass
  Dim benEmployee As IBenefitClass
  Dim i As Long
  
  On Error GoTo SharedVans_Calculate_Error
  
  Call xSet("SharedVans_Calculate")
  
  
  IBenefitClass_Calculate = 0
  m_shvan_benefit_van_only = 0
  m_shvan_fuel_benefit = 0
  m_shvan_fuel_value = 0
  m_shvan_value_van_only = 0
  
  

  For i = 1 To Vans.Count
    Set SharedVan = Vans(i)
    If Not SharedVan Is Nothing Then
      Call SharedVan.Calculate
      m_shvan_benefit_van_only = m_shvan_benefit_van_only + SharedVan.value(shvan_benefit_van_only)
      m_shvan_fuel_benefit = m_shvan_fuel_benefit + SharedVan.value(shvan_fuel_benefit)
      
      m_shvan_value_van_only = m_shvan_value_van_only + SharedVan.value(shvan_Value_Non_OPRA)
      m_shvan_fuel_value = m_shvan_fuel_value + SharedVan.value(shvan_fuel_value)
      
    End If
  Next i
  
  Set benEmployer = p11d32.CurrentEmployer
  
  If benEmployer.value(employer_EmployeesWithSharedVan) > 0 Then
    m_shvan_benefit_van_only = m_shvan_benefit_van_only / (benEmployer.value(employer_EmployeesWithSharedVan))
    m_shvan_fuel_benefit = m_shvan_fuel_benefit / (benEmployer.value(employer_EmployeesWithSharedVan))
    
    m_shvan_value_van_only = m_shvan_value_van_only / (benEmployer.value(employer_EmployeesWithSharedVan))
    m_shvan_fuel_value = m_shvan_fuel_value / (benEmployer.value(employer_EmployeesWithSharedVan))
    
    lTotal = m_shvan_fuel_benefit + m_shvan_benefit_van_only
    IBenefitClass_Calculate = lTotal
  Else
    IBenefitClass_Calculate = 0
  End If
  
SharedVans_Calculate_End:
  xReturn "SharedVans_Calculate"
  Exit Function
SharedVans_Calculate_Error:
  IBenefitClass_Calculate = S_ERROR
  Resume SharedVans_Calculate_End
  Resume
End Function

Private Property Let IBenefitClass_CompanyDefined(ByVal NewValue As Boolean)
  ECASE "SharedVans_CompanyDefined"
End Property

Private Property Get IBenefitClass_CompanyDefined() As Boolean
  IBenefitClass_CompanyDefined = False
End Property
Private Property Let IBenefitClass_Dirty(NewValue As Boolean)
  m_dirty = DirtyHelper(Me, NewValue)
End Property

Private Property Get IBenefitClass_Dirty() As Boolean
  IBenefitClass_Dirty = m_dirty
End Property

Private Property Get IBenefitClass_HasBookMark() As Boolean
  'not used
End Property
Private Sub IBenefitClass_Kill()
  Dim i As Long
  Dim b As IBenefitClass
  For i = 1 To Vans.Count
    Set b = Vans(i)
    If Not b Is Nothing Then
      b.Kill
    End If
  Next i
  Call Vans.RemoveAll
  Set m_Parent = Nothing
End Sub
Private Property Get IBenefitClass_Name() As String
  IBenefitClass_Name = "Vans"
End Property

Public Property Set IBenefitClass_Parent(NewValue As Object)
  Dim ben As IBenefitClass
  
  Set m_Parent = NewValue
End Property

Public Property Get IBenefitClass_Parent() As Object
  Set IBenefitClass_Parent = m_Parent
End Property
Private Property Get IBenefitClass_PrintHeader() As String
  ECASE ("IBenefitClass_PrintHeader")
End Property

Private Function IBenefitClass_PrintWkBody(rep As Reporter) As Boolean

End Function

Private Function IBenefitClass_ReadDB() As Long
  Dim i As Long
  Dim van As IBenefitClass
  Dim rs As Recordset
  
  On Error GoTo SharedVans_ReadDB_err
  
  Call xSet("SharedVans_ReadDB")
  
  If m_ReadFromDB Then GoTo SharedVans_ReadDB_end
  
  Set rs = p11d32.CurrentEmployer.rsBenTables(TBL_SHARED_VANS)
  With rs
    Do While Not .EOF
      Set van = New SharedVan
      van.RSBookMark = .Bookmark
      Set van.Parent = Me
      Call van.ReadDB
      
      Call Vans.Add(van)
      .MoveNext
      i = i + 1
    Loop
  End With
  
  m_ReadFromDB = True
  
SharedVans_ReadDB_end:
  IBenefitClass_ReadDB = i
  Set van = Nothing
  Set rs = Nothing
  xReturn "SharedVans_ReadDB"
  Exit Function
SharedVans_ReadDB_err:
  Call ErrorMessage(ERR_ERROR, Err, "SharedVans_ReadDB", "SharedVans ReadDB", "Error reading the shared vans.")
  Resume SharedVans_ReadDB_end
  Resume
End Function
Private Property Let IBenefitClass_ReadFromDB(ByVal NewValue As Boolean)
  m_ReadFromDB = NewValue
End Property
Private Property Get IBenefitClass_ReadFromDB() As Boolean
  ECASE "SharedVans_ReadFromDB"
End Property
Private Property Let IBenefitClass_RSBookMark(NewValue As String)
  ECASE "SharedVans_sBookMark"
End Property
Private Property Get IBenefitClass_RSBookMark() As String
  ECASE "SharedVans_sBookMark"
End Property
Private Sub IBenefitClass_SetBenItemsInformation()

End Sub
Public Function IBenefitClass_PrintWk(rep As Reporter) As Boolean
End Function
Private Sub IBenefitClass_SetCalcDefaults()
  
End Sub

Private Property Let IBenefitClass_value(ByVal Item As Long, RHS As Variant)
  ECASE "SharedVans_IBenefitClass_Value"
End Property

Private Property Get IBenefitClass_value(ByVal Item As Long) As Variant
  If (Item = shvan_benefit_van_only) Then
    IBenefitClass_value = m_shvan_benefit_van_only
  ElseIf Item = shvan_fuel_benefit Then
    IBenefitClass_value = m_shvan_fuel_benefit
  ElseIf (Item = shvan_Value_Non_OPRA) Then
    IBenefitClass_value = m_shvan_value_van_only
  ElseIf Item = shvan_fuel_value Then
    IBenefitClass_value = m_shvan_fuel_value
  Else
    ECASE "SharedVans_IBenefitClass_Value"
  End If
  
  
End Property

Public Function IBenefitClass_WriteDB() As Boolean
  Dim ben As IBenefitClass, ibc As IBenefitClass
  Dim i As Long
  
  On Error GoTo SharedVans_WriteDB_Err
  Call xSet("SharedVans_WriteDB")
  
  Set ibc = Me

  If ibc.Dirty Or p11d32.BringForward.Yes Then
    For i = 1 To Vans.Count
      Set ben = Vans(i)
      If Not (ben Is Nothing) Then
        If ben.Dirty Or p11d32.BringForward.Yes Then
          Call ben.WriteDB
        End If
      End If
    Next i
    ibc.Dirty = False
  End If
  
  IBenefitClass_WriteDB = True


SharedVans_WriteDB_End:
  Set ibc = Nothing
  Set ben = Nothing
  Call xReturn("SharedVans_WriteDB")
  Exit Function

SharedVans_WriteDB_Err:
  IBenefitClass_WriteDB = False
  Call ErrorMessage(ERR_ERROR, Err, "SharedVans_WriteDB", "SharedVans Write DB", "Error writing the shared vans to the database.")
  Resume SharedVans_WriteDB_End
End Function



Private Property Get IBenefitClass_TABLE() As BENEFIT_TABLES
  IBenefitClass_TABLE = TBL_VANS
End Property


Private Function IBenefitClass_DeleteDB() As Boolean
  Dim ben As IBenefitClass
  Dim i As Long
  
  On Error GoTo SharedVans_DeleteDB_Err
  Call xSet("SharedVans_DeleteDB")
  
  For i = 1 To Vans.Count
    Set ben = Vans(i)
    If Not ben Is Nothing Then
      ben.DeleteDB
    End If
  Next i

SharedVans_DeleteDB_End:
  Call xReturn("SharedVans_DeleteDB")
  Exit Function

SharedVans_DeleteDB_Err:
  Call ErrorMessage(ERR_ERROR, Err, "SharedVans_DeleteDB", "Shared Vans Delete DB", "Error deleting the shared vans from the database.")
  Resume SharedVans_DeleteDB_End
End Function

Private Property Let IBenefitClass_InvalidFields(ByVal NewValue As Long)
  m_InvalidFields = NewValue
End Property

Private Property Get IBenefitClass_InvalidFields() As Long
  Dim ben As New IBenefitClass
  Dim i As Long
  For i = 1 To Vans.Count
    Set ben = Vans(i)
    If Not ben Is Nothing Then
      If ben.InvalidFields > 0 Then
        IBenefitClass_InvalidFields = ben.InvalidFields
        Exit For
      End If
    End If
  Next i
End Property


