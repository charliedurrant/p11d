VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CompanyDefined"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


Option Explicit
Option Base 0
Implements IBenefitClass

Public Enum CompanyDefinedItems
  cdb_item = ITEM_DESC
  cdb_value = ITEM_VALUE
  cdb_madegood = ITEM_MADEGOOD
  cdb_benefit = ITEM_BENEFIT
  cdb_uniqueid
  cdb_class
  cdb_category
  cdb_LASTITEM
End Enum

Private m_InvalidFields As Long
Private m_BenefitClass  As benClass
Private m_Parent As Object
Private m_sbookmark As String
Private m_ReadFromDB As Boolean
Private m_CompanyDefinedItems(0 To cdb_LASTITEM) As Variant
Private m_dirty As Boolean

Private Property Get IBenefitClass_BenefitClass() As benClass
  IBenefitClass_BenefitClass = BC_CDB
End Property
Private Property Let IBenefitClass_BenefitClass(NewValue As benClass)
  m_BenefitClass = NewValue
End Property
Private Property Get IBenefitClass_BenefitSubClass() As Long
 IBenefitClass_BenefitSubClass = 0
End Property

Public Function IBenefitClass_Calculate() As Variant

  On Error GoTo CompanyDefined_Error:
   
  m_CompanyDefinedItems(cdb_category) = P11d32.Rates.ClassStrToCategory(CStr(m_CompanyDefinedItems(cdb_class)))
  m_CompanyDefinedItems(cdb_benefit) = S_ERROR
  m_CompanyDefinedItems(cdb_benefit) = Max(0, m_CompanyDefinedItems(cdb_value) - m_CompanyDefinedItems(cdb_madegood))
  
  IBenefitClass_Calculate = m_CompanyDefinedItems(cdb_benefit)
  
CompanyDefined_End:
  Exit Function
CompanyDefined_Error:
  IBenefitClass_Calculate = S_ERROR
  Resume CompanyDefined_End
End Function
Private Property Let IBenefitClass_CompanyDefined(ByVal NewValue As Boolean)
  ECASE "CompanyDefined_CompanyDefined"
End Property
Private Property Get IBenefitClass_CompanyDefined() As Boolean
End Property
Private Property Let IBenefitClass_Dirty(NewValue As Boolean)
  m_dirty = NewValue
End Property
Private Property Get IBenefitClass_Dirty() As Boolean
  IBenefitClass_Dirty = m_dirty
End Property
Private Function IBenefitClass_GetItem(ByVal Item As Long) As Variant
  IBenefitClass_GetItem = m_CompanyDefinedItems(Item)
End Function
Private Property Get IBenefitClass_HasBookMark() As Boolean
  IBenefitClass_HasBookMark = Len(m_sbookmark) > 0
End Property
Private Property Get IBenefitClass_Name() As String
  IBenefitClass_Name = m_CompanyDefinedItems(cdb_item)
End Property
Private Property Set IBenefitClass_Parent(NewValue As Object)
  Set m_Parent = NewValue
End Property
Private Property Get IBenefitClass_Parent() As Object
  Set IBenefitClass_Parent = m_Parent
End Property
Private Property Get IBenefitClass_PrintHeader() As String
  
End Property
Private Function IBenefitClass_PrintWk(rep As Reporter) As Boolean
  
End Function
Private Function IBenefitClass_ReadDB() As Long
  Dim CBD As IBenefitClass
  Dim rs As Recordset
  Dim s As String
  Dim i As Long
  
  On Error GoTo CompanyDeifined_ReadDB_Err
  Call xSet("CompanyDeifined_ReadDB")
  
  Set CBD = Me
  
  If CBD.ReadFromDB Then GoTo CompanyDeifined_ReadDB_End
  
  Set rs = P11d32.CurrentEmployer.rsBenTables(CBD.TABLE)
  If Not (rs.EOF And rs.BOF) Then
    If Len(m_sbookmark) = 0 Then
      rs.MoveFirst
      CBD.RSBookMark = rs.Bookmark
      i = i + 1
      rs.MoveNext
      Do While Not rs.EOF
        i = i + 1
        Set CBD = New CompanyDefined
        CBD.RSBookMark = rs.Bookmark
        Set CBD.Parent = m_Parent
        Call m_Parent.benefits.Add(CBD)
        Set CBD = Nothing
        rs.MoveNext
      Loop
    End If
  End If
  
  If Len(m_sbookmark) > 0 And m_ReadFromDB = False Then
    rs.Bookmark = m_sbookmark
    m_CompanyDefinedItems(cdb_item) = "" & rs.Fields("Item")
    m_CompanyDefinedItems(cdb_item) = "" & rs.Fields("Item")
    m_CompanyDefinedItems(cdb_class) = "" & rs.Fields("Category")
    m_CompanyDefinedItems(cdb_uniqueid) = "" & rs.Fields("UDBCode")
    m_CompanyDefinedItems(cdb_class) = "" & rs.Fields("Class")
    
    m_CompanyDefinedItems(cdb_value) = "" & rs.Fields("Value")
    m_CompanyDefinedItems(cdb_madegood) = "" & rs.Fields("MadeGood")
    m_CompanyDefinedItems(cdb_benefit) = "" & rs.Fields("Benefit")
    m_ReadFromDB = True
  End If

CompanyDeifined_ReadDB_End:
  IBenefitClass_ReadDB = i
  Set CBD = Nothing
  Call xReturn("CompanyDeifined_ReadDB")
  Exit Function
CompanyDeifined_ReadDB_Err:
  Call ErrorMessage(ERR_ERROR, Err, "CompanyDeifined_ReadDB", "CompanyDeifined ReadDB", "Unable to load the company defined benefit details.")
  Resume CompanyDeifined_ReadDB_End
  Resume
End Function

Private Property Let IBenefitClass_ReadFromDB(ByVal NewValue As Boolean)
  m_ReadFromDB = NewValue
End Property

Private Property Get IBenefitClass_ReadFromDB() As Boolean
  IBenefitClass_ReadFromDB = m_ReadFromDB
End Property

Private Property Get IBenefitClass_Reference() As String
  
End Property

Private Property Let IBenefitClass_RSBookMark(NewValue As String)
  m_sbookmark = NewValue
End Property

Private Property Get IBenefitClass_RSBookMark() As String
  IBenefitClass_RSBookMark = m_sbookmark
End Property

Private Function IBenefitClass_SetItem(ByVal Item As Long, value As Variant) As Boolean
  m_CompanyDefinedItems(Item) = value
End Function

Private Sub IBenefitClass_Kill()
  Set m_Parent = Nothing
End Sub

Private Function IBenefitClass_DeleteDB() As Boolean
  Dim rs As Recordset

  On Error GoTo CompanyDefined_DeleteDB_Err
  Call xSet("CompanyDefined_DeleteDB")
  
  Set rs = P11d32.CurrentEmployer.rsBenTables(IBenefitClass_TABLE)
  If Len(m_sbookmark) > 0 Then
    rs.Bookmark = m_sbookmark
    rs.Delete
  End If
  IBenefitClass_DeleteDB = True
  
CompanyDefined_DeleteDB_End:
  Set rs = Nothing
  Call xReturn("CompanyDefined_DeleteDB")
  Exit Function
CompanyDefined_DeleteDB_Err:
  Call ErrorMessage(ERR_ERROR, Err, "CompanyDefined_DeleteDB", "ERR_UNDEFINED", "Undefined error.")
  Resume CompanyDefined_DeleteDB_End
End Function


Private Function IBenefitClass_WriteDB() As Boolean
  Dim rs As Recordset
  Dim s As String
  Dim ben As IBenefitClass
  
  On Error GoTo CompanyDefined_WriteDB_Err
  Call xSet("CompanyDefined_WriteDB")
  
  Set ben = Me
  
  Set rs = P11d32.CurrentEmployer.rsBenTables(ben.TABLE)
  
  If Len(ben.RSBookMark) > 0 Then
    rs.Bookmark = ben.RSBookMark
    rs.Edit
  Else
    rs.AddNew
    rs.Fields(S_EMPLOYEE_NUM_FIELD) = S_CDB_EMPLOYEE_NUMBER_PREFIX & m_CompanyDefinedItems(cdb_uniqueid)
  End If
  
  rs.Fields("Item") = m_CompanyDefinedItems(cdb_item)
  rs.Fields("Category") = m_CompanyDefinedItems(cdb_class)
  rs.Fields("UDB") = True
  rs.Fields("UDBCode") = m_CompanyDefinedItems(cdb_uniqueid)
  rs.Fields("Class") = m_CompanyDefinedItems(cdb_class)
  rs.Fields("Category") = m_CompanyDefinedItems(cdb_category)
  rs.Fields("Value") = m_CompanyDefinedItems(cdb_value)
  rs.Fields("MadeGood") = m_CompanyDefinedItems(cdb_madegood)
  rs.Fields("Benefit") = m_CompanyDefinedItems(cdb_benefit)
  rs.Update
  
  If Len(m_sbookmark) = 0 Then
    m_sbookmark = rs.LastModified
  End If
  
  m_dirty = False
  IBenefitClass_WriteDB = True
  
CompanyDefined_WriteDB_End:
  Set ben = Nothing
  Set rs = Nothing
  Call xReturn("CompanyDefined_WriteDB")
  Exit Function

CompanyDefined_WriteDB_Err:
  IBenefitClass_WriteDB = False
  Call ClearEdit(rs)
  Call ErrorMessage(ERR_ERROR, Err, "CompanyDefined_WriteDB", "ERR_UNDEFINED", "Undefined error.")
  Resume CompanyDefined_WriteDB_End
  Resume
End Function

Private Property Get IBenefitClass_TABLE() As BENEFIT_TABLES
  IBenefitClass_TABLE = TBL_CDB_FOR_CDB_EMPLOYEE
End Property

Private Property Let IBenefitClass_InvalidFields(ByVal NewValue As Long)
  m_InvalidFields = NewValue
End Property

Private Property Get IBenefitClass_InvalidFields() As Long
  IBenefitClass_InvalidFields = m_InvalidFields
End Property





