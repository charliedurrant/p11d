VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "P"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 0
Implements IBenefitClass
Public Enum PItems
  P_item = ITEM_DESC
  P_ApportionedBen = ITEM_VALUE
  P_MadeGood = ITEM_MADEGOOD
  P_Benefit = ITEM_BENEFIT
  'Longs
    P_Value
    P_CDCKey
  'Strings
    P_EmployeeReference
    P_UDBCode
    P_CDCItem
  'Booleans
  'Dates
    P_availablefrom
    P_availableto
  'LASTITEM
    P_LASTITEM
End Enum

Private m_Parent As Object
Private m_BenefitClass As BEN_CLASS
Private m_sbookmark As String
Private m_ReadFromDB As Boolean
Private m_companydefined  As Boolean
Private m_PItems(0 To P_LASTITEM) As Variant
Private m_dirty As Boolean

Private m_InvalidFields As Long

Private Property Let IBenefitClass_BenefitClass(NewValue As BEN_CLASS)
  
  
End Property

Private Property Get IBenefitClass_BenefitClass() As BEN_CLASS
  IBenefitClass_BenefitClass = BC_POTHER_P
End Property

Private Property Get IBenefitClass_BenefitSubClass() As Long
  IBenefitClass_BenefitSubClass = P_CDCItem
End Property

Private Function IBenefitClass_Calculate() As Variant
  Dim l As Long
  On Error GoTo P_Calculate_err
   
  
  m_PItems(P_ApportionedBen) = S_ERROR
  m_PItems(P_Benefit) = S_ERROR
 
  m_PItems(P_ApportionedBen) = m_PItems(P_Value) * dGetDateFactor(l, CDate(m_PItems(P_availablefrom)), CVDate(m_PItems(P_availableto)))
  m_PItems(P_Benefit) = Max(m_PItems(P_ApportionedBen) - m_PItems(P_MadeGood), 0)
  
  'Special case for shares...
  If m_PItems(P_CDCItem) = S_SHARES_N Then
    If m_PItems(P_Value) > 0 Then
      m_PItems(P_Benefit) = S_SHARES_NRECIEVED
    End If
  End If
  
  IBenefitClass_Calculate = m_PItems(P_Benefit)
  
P_Calculate_end:
  Exit Function

P_Calculate_err:
  IBenefitClass_Calculate = S_ERROR
  'ErrorMessage ERR_ERROR, Err, "P_Calculate", "ERR_OTHER_CALCULATE", "Unable to correctly calculate the benefit on the item " & m_PItems(P_item) & "."
  Resume P_Calculate_end
End Function

Private Property Let IBenefitClass_CompanyDefined(ByVal NewValue As Boolean)
  m_companydefined = NewValue
End Property

Private Property Get IBenefitClass_CompanyDefined() As Boolean
  IBenefitClass_CompanyDefined = m_companydefined
End Property


Private Property Let IBenefitClass_Dirty(RHS As Boolean)
  m_dirty = RHS
End Property

Private Property Get IBenefitClass_Dirty() As Boolean
  IBenefitClass_Dirty = m_dirty
End Property

Private Function IBenefitClass_GetItem(ByVal Item As Long) As Variant
  IBenefitClass_GetItem = m_PItems(Item)
End Function

Private Property Get IBenefitClass_HasBookMark() As Boolean
  IBenefitClass_HasBookMark = Len(m_sbookmark) > 0
End Property

Private Sub IBenefitClass_Kill()
  Set m_Parent = Nothing
End Sub

Private Sub IBenefitClass_MakeDirty()
  m_dirty = True
End Sub

Private Property Get IBenefitClass_Name() As String
  IBenefitClass_Name = m_PItems(P_item)
End Property

Private Property Set IBenefitClass_Parent(NewValue As Object)
  Set m_Parent = NewValue
End Property

Private Property Get IBenefitClass_Parent() As Object
  Set IBenefitClass_Parent = m_Parent
End Property

Private Property Get IBenefitClass_PrintHeader() As String
   IBenefitClass_PrintHeader = "{Arial=8,n}{x=75}Amount made" & vbCrLf & "{x=60}Cost to" & _
      "{x=90}Cash{x=75}good or taxed" & vbCrLf & _
      "Description{x=60}you{x=75}under PAYE" & _
      "{x=90}equivalent" & vbCrLf & "{x=67}£   {x=82}£   {Arial=8,ri}{x=97}£   " & vbCrLf & vbCrLf
End Property

Private Function IBenefitClass_ReadDB() As Long
  Dim P As IBenefitClass
  Dim rs As Recordset
  Dim s As String
  Dim sIndex As String
  Dim i As Long
  
  On Error GoTo P_ReadDB_err
  xSet "P_ReadDB"
  If m_ReadFromDB Then GoTo P_ReadDB_end
  Set rs = m_Parent.Parent.rsBenTables(TBL_P)
  
  If Len(m_sbookmark) = 0 Then
    If m_companydefined Then
      Call rs.FindFirst("UDB = " & True)  '("LEFT$([P_NUM],4)='" & s & "'")
    Else
      s = m_Parent.PersonnelNumber
      rs.FindFirst ("P_NUM = '" & s & "'")
    End If
    If Not rs.NoMatch Then
      i = i + 1
      m_sbookmark = rs.Bookmark
      If m_companydefined Then
        Call rs.FindNext("UDB = " & True)  '("LEFT$([P_NUM],4)='" & s & "'")
      Else
        rs.FindNext ("P_NUM = '" & s & "'")
      End If
      Do While Not rs.NoMatch
        i = i + 1
        Set P = New P
        P.RSBookMark = rs.Bookmark
        P.CompanyDefined = m_companydefined
        Set P.Parent = m_Parent
        Call m_Parent.Benefits.Add(P)
        Set P = Nothing
        If m_companydefined Then
          Call rs.FindNext("UDB = " & True)  '("LEFT$([P_NUM],4)='" & s & "'")
        Else
          rs.FindNext ("P_NUM = '" & s & "'")
        End If
      Loop
    Else
      m_ReadFromDB = True
    End If
  End If
  If Len(m_sbookmark) > 0 And m_ReadFromDB = False Then
    rs.Bookmark = m_sbookmark
    m_PItems(P_EmployeeReference) = "" & rs.Fields(S_EMPLOYEE_NUM_FIELD)
    m_PItems(P_CDCKey) = "" & rs.Fields("CDCKey")
    m_PItems(P_CDCItem) = "" & rs.Fields("CDCItem")
    m_PItems(P_UDBCode) = "" & rs.Fields("UDBCode")
    m_PItems(P_item) = "" & rs.Fields("Item")
    m_PItems(P_availablefrom) = IIf(IsNull(rs.Fields("From")), P11d32.Rates.GetItem(TaxYearStart), rs.Fields("From"))
    m_PItems(P_availableto) = IIf(IsNull(rs.Fields("To")), P11d32.Rates.GetItem(TaxYearEnd), rs.Fields("To"))
    m_PItems(P_Value) = IIf(IsNull(rs.Fields("Value")), 0, rs.Fields("Value"))
    m_PItems(P_MadeGood) = IIf(IsNull(rs.Fields("Madegood")), 0, rs.Fields("Madegood"))
    m_BenefitClass = BC_POTHER_P
    
    
    m_ReadFromDB = True
    
  End If
P_ReadDB_end:
  IBenefitClass_ReadDB = i
  xReturn "P_ReadDB"
  Exit Function
P_ReadDB_err:
   ErrorMessage ERR_ERROR, Err, "P_ReadDB", "ERR_READ_P_BENEFITS", "Unable to read in the P benefits"
   Resume P_ReadDB_end
End Function

Private Property Let IBenefitClass_ReadFromDB(ByVal NewValue As Boolean)
  m_ReadFromDB = NewValue
End Property

Private Property Get IBenefitClass_ReadFromDB() As Boolean
  IBenefitClass_ReadFromDB = m_ReadFromDB
End Property


Private Property Get IBenefitClass_Reference() As String
  IBenefitClass_Reference = m_PItems(P_EmployeeReference)
End Property

Private Property Let IBenefitClass_RSBookMark(NewValue As String)
  m_sbookmark = NewValue
End Property

Private Property Get IBenefitClass_RSBookMark() As String
  IBenefitClass_RSBookMark = m_sbookmark
End Property

Private Function IBenefitClass_SetItem(ByVal Item As Long, value As Variant) As Boolean
  'If Item = P_CDCItem Then Err.Raise 1
  m_PItems(Item) = value
  IBenefitClass_SetItem = True
End Function

Public Function IBenefitClass_DeleteDB() As Boolean
  Dim rs As Recordset
  Dim s As String

  On Error GoTo P_DeleteDB_Err
  Call xSet("P_DeleteDB")
  Set rs = m_Parent.Parent.rsBenTables(TBL_P)
  If Len(m_sbookmark) > 0 Then
    rs.Bookmark = m_sbookmark
    rs.Delete
  End If
  IBenefitClass_DeleteDB = True
P_DeleteDB_End:
  Set rs = Nothing
  Call xReturn("P_DeleteDB")
  Exit Function

P_DeleteDB_Err:
  Call ErrorMessage(ERR_ERROR, Err, "P_DeleteDB", "P Delete DB", "Error deleting a 'P' benefit from the database.")
  Resume P_DeleteDB_End
End Function


Public Function IBenefitClass_WriteDB() As Boolean
  Dim rs As Recordset
  Dim s As String

  On Error GoTo P_WriteDB_Err
  Call xSet("P_WriteDB")
  Set rs = m_Parent.Parent.rsBenTables(TBL_P)
  
  If Len(m_sbookmark) > 0 Then
    rs.Bookmark = m_sbookmark
    rs.Edit
  Else
    rs.AddNew
    rs.Fields(S_EMPLOYEE_NUM_FIELD) = m_PItems(P_EmployeeReference)
  End If
  rs.Fields("CDCKey") = m_PItems(P_CDCKey)
  rs.Fields("CDCItem") = m_PItems(P_CDCItem)
  rs.Fields("UDB") = m_companydefined
  rs.Fields("UDBCode") = m_PItems(P_UDBCode)
  rs.Fields("Item") = m_PItems(P_item)
  rs.Fields("From") = IIf(m_PItems(P_availablefrom) = UNDATED, Null, m_PItems(P_availablefrom))
  rs.Fields("To") = IIf(m_PItems(P_availableto) = UNDATED, Null, m_PItems(P_availableto))
  rs.Fields("Value") = m_PItems(P_Value)
  rs.Fields("Madegood") = m_PItems(P_MadeGood)
  rs.Update
  If Len(m_sbookmark) = 0 Then
    m_sbookmark = rs.LastModified
  End If
  m_dirty = False
  IBenefitClass_WriteDB = True
P_WriteDB_End:
  Set rs = Nothing
  Call xReturn("P_WriteDB")
  Exit Function
  
P_WriteDB_Err:
  IBenefitClass_WriteDB = False
  Call ClearEdit(rs)
  Call ErrorMessage(ERR_ERROR, Err, "P_WriteDB", "P Write DB", "Error writing a 'P' benefit to the database.")
  Resume P_WriteDB_End
End Function


Public Function IBenefitClass_PrintWk(rep As Reporter) As Boolean
  On Error GoTo P_PrintWK_Err
  
  Call xSet("P_PrintWK")
  If rep.InitReport("P Expenses", gtarget, PORTRAIT, False) Then
   
    rep.Out "{Arial=8,n}{x=0}" & m_PItems(P_item)
    rep.Out "{Arial=8,rb}{x=67}" & formatworkingnumber(m_PItems(P_ApportionedBen), "")
    rep.Out "{Arial=8,rb}{x=82}" & formatworkingnumber(m_PItems(P_MadeGood), "")
    rep.Out "{Arial=8,rb}{x=97}" & formatworkingnumber(m_PItems(P_Benefit), "")
    rep.Out vbCrLf
  End If
  Call rep.EndReport
P_PrintWK_End:
  Call xReturn("P_PrintWK")
  Exit Function
P_PrintWK_Err:

  Call ErrorMessage(ERR_ERROR, Err, "P_PrintWK", "P Print WK", "Error printing a 'P' benefit.")
  Resume P_PrintWK_End
End Function

Private Property Get IBenefitClass_TABLE() As BENEFIT_TABLES
  IBenefitClass_TABLE = TBL_P
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
