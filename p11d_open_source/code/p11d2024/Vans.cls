VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsBenVan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Option Base 0
Implements IBenefitClass

Public Enum VanItems
  van_item = ITEM_DESC
  van_GrossBenefit = ITEM_VALUE
  van_Contribution = ITEM_MADEGOOD
  van_benefit = ITEM_BENEFIT
'Van Strings
  van_employeereference
'Van Longs
  van_Unavailable
  van_RelevantDays
  van_Scalecharge
  van_DaysUnavailable
  van_Reduction
'Van Booleans
  van_newvan
  van_Multiple
  van_ShowAlternative
'Van Dates
  van_Availablefrom
  van_Availableto
'van Doubles
  van_DFraction
'Last Item
  van_LASTITEM
End Enum

Private m_InvalidFields As Long
Private m_BenefitClass  As benClass
Private m_parent As IBenefitClass
Private m_sbookmark As String
Private m_readfromdb As Boolean
Private m_VanItems(0 To van_LASTITEM) As Variant
Private m_dirty As Boolean
Private m_dateedit As Date

Private Property Get IBenefitClass_BenefitClass() As benClass
  IBenefitClass_BenefitClass = BC_VAN
End Property

Private Property Let IBenefitClass_BenefitClass(NewValue As benClass)
  m_BenefitClass = NewValue
End Property

Private Property Get IBenefitClass_BenefitSubClass() As Long
  IBenefitClass_BenefitSubClass = 0
End Property

Public Function IBenefitClass_Calculate() As Variant
  m_parent.Calculate
  IBenefitClass_Calculate = m_VanItems(van_benefit)
End Function

Private Property Let IBenefitClass_CompanyDefined(ByVal NewValue As Boolean)
  ECASE "clsBenVan_CompanyDefined"
End Property

Private Property Get IBenefitClass_CompanyDefined() As Boolean
  'ECASE "clsBenVan_CompanyDefined"
End Property

Private Property Get IBenefitClass_DateEdit() As Date
  IBenefitClass_DateEdit = m_dateedit
End Property

Private Property Let IBenefitClass_Dirty(RHS As Boolean)

End Property

Private Property Get IBenefitClass_Dirty() As Boolean
  IBenefitClass_Dirty = m_dirty
End Property

Private Function IBenefitClass_GetItem(ByVal Item As Long) As Variant
  IBenefitClass_GetItem = m_VanItems(Item)
End Function

Private Property Get IBenefitClass_HasBookMark() As Boolean
  IBenefitClass_HasBookMark = Len(m_sbookmark) > 0
End Property

Private Sub IBenefitClass_MakeDirty()
  m_dirty = True
  m_parent.Dirty = True
End Sub

Private Property Get IBenefitClass_name() As String
  IBenefitClass_name = m_VanItems(van_item)
End Property

Private Property Set IBenefitClass_Parent(NewValue As Object)
  Set m_parent = NewValue
End Property

Private Property Get IBenefitClass_Parent() As Object
  Set IBenefitClass_Parent = m_parent
End Property

Private Property Get IBenefitClass_PrintHeader() As String
  ECASE ("IBenefitClass_PrintHeader")
End Property

Private Function IBenefitClass_PrintWK(rep As Reporter) As Boolean
  Call m_parent.PrintWk(rep)
End Function

Private Function IBenefitClass_ReadDB() As Long
  Dim van As IBenefitClass
  Dim rs As Recordset
  Dim s As String
  Dim i As Long
  
  On Error GoTo clsVan_ReadDB_Err
  Call xSet("clsVan_ReadDB")
  If m_readfromdb Then GoTo clsVan_ReadDB_End
  Set rs = m_parent.Parent.Parent.rsBenTables(TBL_VANS)
  If Len(m_sbookmark) > 0 And m_readfromdb = False Then
    rs.Bookmark = m_sbookmark
    m_VanItems(van_employeereference) = "" & rs.Fields(S_EMPLOYEE_NUM_FIELD)
    m_VanItems(van_item) = "" & rs.Fields("ITEM")
    m_VanItems(van_Unavailable) = IIf(IsNull(rs.Fields("UNAVAIL")), 0, rs.Fields("UNAVAIL"))
    m_VanItems(van_Contribution) = IIf(IsNull(rs.Fields("PRIVATECONTRIB")), 0, rs.Fields("PRIVATECONTRIB"))
    m_VanItems(van_newvan) = IIf(IsNull(rs.Fields("REGAFTER")), False, rs.Fields("REGAFTER"))
    m_VanItems(van_Availablefrom) = IIf(IsNull(rs.Fields("AVAILFROM")), UNDATED, rs.Fields("AVAILFROM"))
    m_VanItems(van_Availableto) = IIf(IsNull(rs.Fields("AVAILTO")), UNDATED, rs.Fields("AVAILTO"))
    m_readfromdb = True
  End If

clsVan_ReadDB_End:
  IBenefitClass_ReadDB = i
  Set van = Nothing
  Call xReturn("clsVan_ReadDB")
  Exit Function
clsVan_ReadDB_Err:
  Call ErrorMessage(ERR_ERROR, Err, "clsVan_ReadDB", "ERR_VAN_READDB", "Unable to load the van details")
  Resume clsVan_ReadDB_End
End Function

Private Property Let IBenefitClass_ReadFromDB(ByVal NewValue As Boolean)
  m_readfromdb = NewValue
End Property

Private Property Get IBenefitClass_ReadFromDB() As Boolean
  IBenefitClass_ReadFromDB = m_readfromdb
End Property

Private Property Get IBenefitClass_Reference() As String
  ECASE "clsBenVan_Reference"
End Property

Private Property Let IBenefitClass_RSBookMark(NewValue As String)
  m_sbookmark = NewValue
End Property

Private Property Get IBenefitClass_RSBookMark() As String
  IBenefitClass_RSBookMark = m_sbookmark
End Property

Private Function IBenefitClass_SetItem(ByVal Item As Long, value As Variant) As Boolean
  m_VanItems(Item) = value
End Function

Private Sub IBenefitClass_Kill()
  Set m_parent = Nothing
End Sub

Public Function PrintWk(rep As Reporter)
Dim tmp As String
  On Error GoTo printwk_Err
  Call xSet("Printwk")
  Call rep.InitReport("Private Vans", gtarget, PORTRAIT, False, False)
  rep.Out "{Arial=8,n}{x=0}{LINE}" & vbCrLf
  rep.Out "{Arial=8,b}{x=0}Private van benefit" & vbCrLf & vbCrLf
  rep.Out "{Arial=8,n}{x=0}Reference "
  rep.Out "{Arial=8,b}" & m_VanItems(van_item) & vbCrLf & vbCrLf
  If m_VanItems(van_newvan) Then
    rep.Out "{Arial=8,n}{x=0}First registered on or after " & FormatDate(P11d32.Rates.GetItem(vanOldDate)) & vbCrLf & vbCrLf
  Else
    rep.Out "{Arial=8,n}{x=0}First registered before " & FormatDate(P11d32.Rates.GetItem(vanOldDate)) & vbCrLf & vbCrLf
  End If
  tmp = FormatDate(m_VanItems(van_Availablefrom))
  tmp = tmp & " to "
  tmp = tmp & FormatDate(m_VanItems(van_Availableto))
  rep.Out "{Arial=8}{x=0}Made available from " & tmp & vbCrLf & vbCrLf
  rep.Out "{Arial=8}{x=0}Qualifying days unavailable or shared: " & m_VanItems(van_Unavailable) & vbCrLf & vbCrLf
  rep.Out "{Arial=8,b}{x=0}Calculation of taxable benefit:" & vbCrLf & vbCrLf
  rep.Out "{Arial=8,n}{x=0}Standard annual charge"
  rep.Out "{Arial=8,rb}{x=97}" & formatworkingnumber(m_VanItems(van_GrossBenefit), "£") & vbCrLf & vbCrLf
  rep.Out "{Arial=8,n}{x=0}Less: Amount relating to days unavailable "
  rep.Out "{Arial=8} (" & (m_VanItems(van_DaysUnavailable)) & "/365)"

  rep.Out "{Arial=8,rb}{x=97}" & formatworkingnumber(m_VanItems(van_Reduction), "") & vbCrLf & vbCrLf
  rep.Out "{Arial=8,n}{x=0}Less: Payments made by the employee as a condition of private use"
  rep.Out "{Arial=8,rb}{x=97}" & formatworkingnumber(-m_VanItems(van_Contribution), "") & vbCrLf
  rep.Out "{Arial=8,nb}{x=97}{line=-7}" & vbCrLf
  rep.Out "{Arial=8,n}{x=0}Cash equivalent"
  rep.Out "{Arial=8,rb}{x=97}" & formatworkingnumber(m_VanItems(van_benefit), "£") & vbCrLf
  rep.Out "{Arial=8,nb}{x=97}{LINE=-7,d}" & vbCrLf
  rep.Out "{Arial=8,n}{x=0}{LINE}" & vbCrLf
  Call rep.EndReport
printwk_End:
  Call xReturn("printWk")
  Exit Function
printwk_Err:
  Call ErrorMessage(ERR_ERROR, Err, "PrintWK", "ERR_VAN_PRINT", "Unable to print the van worksheet")
  Resume printwk_End
End Function

Public Function VanCalculate() As Long
  Dim i As Long
  On Error GoTo VanCalculate_Err
  Call xSet("VanCalculate")
  
  m_VanItems(van_DFraction) = S_ERROR
  m_VanItems(van_DaysUnavailable) = S_ERROR
  m_VanItems(van_Scalecharge) = S_ERROR
  m_VanItems(van_Reduction) = S_ERROR
  m_VanItems(van_GrossBenefit) = S_ERROR
  m_VanItems(van_benefit) = S_ERROR
  
  m_VanItems(van_DFraction) = dGetDateFactor(i, CDate(m_VanItems(van_Availablefrom)), CDate(m_VanItems(van_Availableto)), m_VanItems(van_Unavailable), P11d32.Rates.GetItem(vanMinDays))
  m_VanItems(van_DaysUnavailable) = i
  If m_VanItems(van_newvan) Then
    m_VanItems(van_Scalecharge) = P11d32.Rates.GetItem(vanHighCharge)
  Else
    m_VanItems(van_Scalecharge) = P11d32.Rates.GetItem(vanLowCharge)
  End If
  m_VanItems(van_GrossBenefit) = CLng(m_VanItems(van_DFraction) * m_VanItems(van_Scalecharge))
  m_VanItems(van_Reduction) = CLng(m_VanItems(van_GrossBenefit) * (1 - m_VanItems(van_DFraction)))
  m_VanItems(van_benefit) = Max(0, (m_VanItems(van_GrossBenefit) - m_VanItems(van_Contribution)))
  VanCalculate = m_VanItems(van_benefit)
  
VanCalculate_End:
  Call xReturn("VanCalculate")
  Exit Function
VanCalculate_Err:
  VanCalculate = -1
  'Call ErrorMessage(ERR_ERROR, Err, "VanCalculate", "ERR_VAN_CALCULATE", "Unable to complete the van calculation on " & m_VanItems(van_item) & ".")
  Resume Next 'VanCalculate_End
End Function


Public Function IBenefitClass_DeleteDB() As Boolean
  Dim rs As Recordset
  
  On Error GoTo clsBenVan_DeleteDB_Err
  Call xSet("clsBenVan_DeleteDB")
  Set rs = m_parent.Parent.Parent.rsBenTables(TBL_VANS)
  If Len(m_sbookmark) > 0 Then
    rs.Bookmark = m_sbookmark
    rs.Delete
  End If
  IBenefitClass_DeleteDB = True
  
clsBenVan_DeleteDB_End:
  Set rs = Nothing
  Call xReturn("clsBenVan_DeleteDB")
  Exit Function
clsBenVan_DeleteDB_Err:
  Call ErrorMessage(ERR_ERROR, Err, "clsBenVan_DeleteDB", "ERR_UNDEFINED", "Undefined error.")
  Resume clsBenVan_DeleteDB_End
End Function


Public Function IBenefitClass_WriteDB() As Boolean
  Dim rs As Recordset
  Dim s As String
  
  On Error GoTo clsBenVan_WriteDB_Err
  Call xSet("clsBenVan_WriteDB")
  Set rs = m_parent.Parent.Parent.rsBenTables(TBL_VANS)
  If Len(m_sbookmark) > 0 Then
    rs.Bookmark = m_sbookmark
    rs.Edit
  Else
    rs.AddNew
    rs.Fields(S_EMPLOYEE_NUM_FIELD) = m_VanItems(van_employeereference)
  End If
  rs.Fields("ITEM") = m_VanItems(van_item)
  rs.Fields("UNAVAIL") = m_VanItems(van_Unavailable)
  rs.Fields("PRIVATECONTRIB") = m_VanItems(van_Contribution)
  rs.Fields("REGAFTER") = m_VanItems(van_newvan)
  rs.Fields("AVAILFROM") = m_VanItems(van_Availablefrom)
  rs.Fields("AVAILTO") = m_VanItems(van_Availableto)
  rs.Update
  If Len(m_sbookmark) = 0 Then
    m_sbookmark = rs.LastModified
  End If
  m_dirty = False
  IBenefitClass_WriteDB = True
  
clsBenVan_WriteDB_End:
  Set rs = Nothing
  Call xReturn("clsBenVan_WriteDB")
  Exit Function

clsBenVan_WriteDB_Err:
  IBenefitClass_WriteDB = False
  Call ClearEdit(rs)
  Call ErrorMessage(ERR_ERROR, Err, "clsBenVan_WriteDB", "Write DB", "Error writing the van benefit.")
  Resume clsBenVan_WriteDB_End
End Function

Private Property Get IBenefitClass_TABLE() As BENEFIT_TABLES
  IBenefitClass_TABLE = TBL_LASTITEM
End Property

Private Property Let IBenefitClass_InvalidFields(ByVal NewValue As Long)
  m_InvalidFields = NewValue
End Property

Private Property Get IBenefitClass_InvalidFields() As Long
  IBenefitClass_InvalidFields = m_InvalidFields
End Property

Public Sub Ibenefitclass_Cleardirty()
  m_dirty = False
End Sub

