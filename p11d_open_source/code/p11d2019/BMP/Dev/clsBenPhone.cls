VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsBenPhone"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Base 0
Implements IBenefitClass
Public Enum PhoneItems
  'Longs
    Pho_Value
    Pho_MadeGood
    Pho_RentValue
    Pho_RentMadeGood
    Pho_CallsMadeGood
    Pho_CallsValue
    Pho_Unavailable
    Pho_ApportionedBen
    Pho_Benefit
    
  'Strings
    Pho_EmployeeReference
    Pho_UDBCode
    Pho_Class
    Pho_item
  'Booleans
    Pho_PrivateUse
    Pho_ReqMakeGood
    Pho_ActMadeGood
    Pho_CapMadeGood
  'Dates
    Pho_availablefrom
    Pho_availableto
  'LASTITEM
    Pho_LASTITEM
End Enum

Private m_parent As Object
Private m_BenefitClass As BenClass
Private m_sbookmark As String
Private m_readfromdb As Boolean
Private m_companydefined  As Boolean
Private m_PhoneItems(0 To Pho_LASTITEM) As Variant
Private m_dirty As Boolean
Private m_dateedit As Date
Private m_InvalidFields As Long

Private Property Let IBenefitClass_BenefitClass(RHS As BenClass)
  m_BenefitClass = RHS
End Property

Private Property Get IBenefitClass_BenefitClass() As BenClass
  IBenefitClass_BenefitClass = m_BenefitClass
End Property

Private Function IBenefitClass_Calculate() As Variant
  Dim l As Long
  On Error GoTo clsBenPhone_Calculate_err
  
  IBenefitClass_Calculate =
  
clsBenPhone_Calculate_end:
  Exit Function

clsBenPhone_Calculate_err:
  IBenefitClass_Calculate = S_ERROR
  'ErrorMessage ERR_ERROR, Err, "clsBenPhone_Calculate", "ERR_OTHER_CALCULATE", "Unable to correctly calculate the benefit on the item " & m_OtherItems(Oth_item) & "."
  Resume clsBenPhone_Calculate_end
End Function

Private Property Let IBenefitClass_CompanyDefined(RHS As Boolean)
  m_companydefined = RHS
End Property

Private Property Get IBenefitClass_CompanyDefined() As Boolean
  IBenefitClass_CompanyDefined = m_companydefined
End Property

Private Property Get IBenefitClass_DateEdit() As Date
 IBenefitClass_DateEdit = m_dateedit
End Property

Private Property Get IBenefitClass_Dirty() As Boolean
  IBenefitClass_Dirty = m_dirty
End Property

Private Function IBenefitClass_GetItem(lItem As Long) As Variant
  IBenefitClass_GetItem = m_OtherItems(lItem)
End Function

Private Sub IBenefitClass_Kill()
  Set m_parent = Nothing
End Sub

Private Sub IBenefitClass_MakeDirty()
  m_dirty = True
End Sub

Private Property Get IBenefitClass_name() As String
  IBenefitClass_name = m_OtherItems(Oth_item)
End Property

Private Property Set IBenefitClass_Parent(RHS As Object)
  Set m_parent = RHS
End Property

Private Property Get IBenefitClass_Parent() As Object
  Set IBenefitClass_Parent = m_parent
End Property

Private Function IBenefitClass_ReadDB() As Long
  Dim phone As IBenefitClass
  Dim rs As Recordset
  Dim s As String
  Dim sIndex As String
  Dim i As Long
  
  On Error GoTo clsBenPhone_ReadDB_err
  xSet "clsBenPhone_ReadDB"
  If m_readfromdb Then GoTo clsBenPhone_ReadDB_end
  Set rs = m_parent.Parent.rsBenTables(TBL_PHONE)
  If Len(m_sbookmark) = 0 Then
    s = m_parent.PersonelNo
    rs.FindFirst ("P_NUM = '" & s & "'")
    If Not rs.NoMatch Then
      i = i + 1
      m_sbookmark = rs.Bookmark
      rs.FindNext ("P_NUM = '" & s & "'")
      Do While Not rs.NoMatch
        i = i + 1
        Set phone = New clsBenPhone
        phone.sBookMark = rs.Bookmark
        phone.CompanyDefined = m_companydefined
        Set phone.Parent = m_parent
        Call m_parent.benefits.Add(phone)
        Set phone = Nothing
        rs.FindNext ("P_NUM = '" & s & "'")
      Loop
    Else
      m_readfromdb = True
    End If
  End If
  If Len(m_sbookmark) > 0 And m_readfromdb = False Then
    rs.Bookmark = m_sbookmark
    
    m_PhoneItems(Pho_EmployeeReference) = "" & rs.Fields("P_NUM")
    m_PhoneItems(Pho_Class) = "" & rs.Fields("Class")
    m_PhoneItems(Pho_RentValue) = IIf(IsNull(rs.Fields("ValueRent")), 0, rs.Fields("ValueRent"))
    m_PhoneItems(Pho_RentMadeGood) = IIf(IsNull(rs.Fields("ContribRent")), 0, rs.Fields("ContribRent"))
    m_PhoneItems(Pho_CallsValue) = IIf(IsNull(rs.Fields("ValueCalls")), 0, rs.Fields("ValueCalls"))
    m_PhoneItems(Pho_CallsMadeGood) = IIf(IsNull(rs.Fields("ContribCalls")), 0, rs.Fields("ContribCalls"))
    m_PhoneItems(Pho_Unavailable) = IIf(IsNull(rs.Fields("Unavailable")), 0, rs.Fields("Unavailable"))
    m_PhoneItems(Pho_UDBCode) = "" & rs.Fields("UDBCode")
    m_PhoneItems(Pho_item) = "" & rs.Fields("Item")
    m_PhoneItems(Pho_PrivateUse) = rs.Fields("PvtUse")
    m_PhoneItems(Pho_ReqMakeGood) = rs.Fields("MakeGood")
    m_PhoneItems(Pho_ActMadeGood) = rs.Fields("MadeGood")
    m_PhoneItems(Pho_CapMadeGood) = rs.Fields("CapMadeGood")
    m_PhoneItems(Pho_availablefrom) = IIf(IsNull(rs.Fields("AvailFrom")), rates.GetItem(taxyearstart), rs.Fields("AvailFrom"))
    m_PhoneItems(Pho_availableto) = IIf(IsNull(rs.Fields("AvailTo")), rates.GetItem(taxyearend), rs.Fields("AvailTo"))
    
    m_readfromdb = True
    
  End If
clsBenPhone_ReadDB_end:
  IBenefitClass_ReadDB = i
  xReturn "clsBenPhone_ReadDB"
  Exit Function
clsBenPhone_ReadDB_err:
   ErrorMessage ERR_ERROR, Err, "clsBenPhone_ReadDB", "ERR_READ_Phone_BENEFITS", "Unable to read in the Phone benefits"
   Resume clsBenPhone_ReadDB_end
End Function

Private Property Let IBenefitClass_ReadFromDB(RHS As Boolean)
  m_readfromdb = RHS
End Property

Private Property Get IBenefitClass_ReadFromDB() As Boolean
  IBenefitClass_ReadFromDB = m_readfromdb
End Property


Private Property Get IBenefitClass_Reference() As String
  IBenefitClass_Reference = m_OtherItems(Oth_EmployeeReference)
End Property

Private Property Let IBenefitClass_sBookMark(RHS As String)
  m_sbookmark = RHS
End Property

Private Property Get IBenefitClass_sBookMark() As String
  IBenefitClass_sBookMark = m_sbookmark
End Property

Private Function IBenefitClass_SetItem(Item As Long, Value As Variant) As Boolean
  m_PhoneItems(Item) = Value
  IBenefitClass_SetItem = True
End Function

Public Function IBenefitClass_DeleteDB() As Boolean
  Dim rs As Recordset
  Dim s As String

  On Error GoTo clsBenPhone_DeleteDB_Err
  Call xSet("clsBenPhone_DeleteDB")
  Set rs = m_parent.Parent.rsBenTables(TBL_PHONE)
  If Len(m_sbookmark) > 0 Then
    rs.Bookmark = m_sbookmark
    rs.Delete
  End If
  IBenefitClass_DeleteDB = True
clsBenPhone_DeleteDB_End:
  Set rs = Nothing
  Call xReturn("clsBenPhone_DeleteDB")
  Exit Function

clsBenPhone_DeleteDB_Err:
  Call ErrorMessage(ERR_ERROR, Err, "clsBenPhone_DeleteDB", "ERR_UNDEFINED", "Undefined error.")
  Resume clsBenPhone_DeleteDB_End
End Function


Public Function IBenefitClass_WriteDB() As Boolean
  Dim rs As Recordset
  Dim s As String

  On Error GoTo clsBenPhone_WriteDB_Err
  Call xSet("clsBenPhone_WriteDB")
  Set rs = m_parent.Parent.rsBenTables(TBL_PHONE)
  
  If Len(m_sbookmark) > 0 Then
    rs.Bookmark = m_sbookmark
    rs.Edit
  Else
    rs.AddNew
    rs.Fields("P_NUM") = m_OtherItems(Oth_EmployeeReference)
  End If
  rs.Fields("Category") = m_OtherItems(Oth_Category)
  rs.Fields("Class") = m_OtherItems(Oth_Class)
  rs.Fields("UDB") = m_companydefined
  rs.Fields("UDBCode") = m_OtherItems(Oth_UDBCode)
  rs.Fields("Item") = m_OtherItems(Oth_item)
  rs.Fields("From") = IIf(m_OtherItems(Oth_availablefrom) = UNDATED, Null, m_OtherItems(Oth_availablefrom))
  rs.Fields("To") = IIf(m_OtherItems(Oth_availableto) = UNDATED, Null, m_OtherItems(Oth_availableto))
  rs.Fields("Value") = m_OtherItems(Oth_Value)
  rs.Fields("Madegood") = m_OtherItems(Oth_MadeGood)
  rs.Update
  If Len(m_sbookmark) = 0 Then
    m_sbookmark = rs.LastModified
  End If
  m_dirty = False
  IBenefitClass_WriteDB = True
clsBenPhone_WriteDB_End:
  Set rs = Nothing
  Call xReturn("clsBenPhone_WriteDB")
  Exit Function
clsBenPhone_WriteDB_Err:
  Call ErrorMessage(ERR_ERROR, Err, "clsBenPhone_WriteDB", "ERR_UNDEFINED", "Undefined error.")
  Resume clsBenPhone_WriteDB_End
End Function


Public Function IBenefitClass_PrintWK() As Boolean
  On Error GoTo clsBenPhone_PrintWK_Err
  
  Call xSet("clsBenPhone_PrintWK")
  Call InitReport("Phones", gtarget, PORTRAIT, False, False)

  Call EndReport
clsBenPhone_PrintWK_End:
  Call xReturn("clsBenPhone_PrintWK")
  Exit Function
clsBenPhone_PrintWK_Err:
  Call ErrorMessage(ERR_ERROR, Err, "clsBenPhone_PrintWK", "ERR_UNDEFINED", "Undefined error.")
  Resume clsBenPhone_PrintWK_End
End Function

Private Property Get IBenefitClass_TABLE() As BENEFIT_TABLES
  IBenefitClass_TABLE = TBL_PHONE
End Property

Private Property Let IBenefitClass_InvalidFields(RHS As Long)
  m_InvalidFields = RHS
End Property

Private Property Get IBenefitClass_InvalidFields() As Long
  IBenefitClass_InvalidFields = m_InvalidFields
End Property

Public Sub Ibenefitclass_Cleardirty()
  m_dirty = False
End Sub



