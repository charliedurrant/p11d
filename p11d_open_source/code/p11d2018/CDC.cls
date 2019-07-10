VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CDC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IBenefitClass


Public Enum cdc_Items
'MP DB cdc_name = 1
  cdc_name_db = 1
  cdc_IsUsed
'MP DB - cdc_Key
  cdc_Key_db
'MP DB - cdc_LastITEM = cdc_Key
  cdc_LastITEM = cdc_Key_db
End Enum


Private m_BenItems(1 To cdc_LastITEM)
Private m_bDirty As Boolean
Private m_Parent As Employer
Private m_ReadFromDB As Boolean
Private m_NeedToCalculate As Boolean

Private Sub Class_Terminate()
 'Debug.Print"Company defined category - terminate"
End Sub

Private Function IBenefitClass_CalculateBody() As Variant

End Function

Private Property Get IBenefitClass_ImageListKey() As String
  IBenefitClass_ImageListKey = ""
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

Private Property Let IBenefitClass_BenefitClass(RHS As BEN_CLASS)

End Property

Private Property Get IBenefitClass_BenefitClass() As BEN_CLASS

End Property

Private Function IBenefitClass_Calculate() As Variant

End Function

Private Property Let IBenefitClass_CompanyDefined(ByVal RHS As Boolean)

End Property

Private Property Get IBenefitClass_CompanyDefined() As Boolean

End Property

Private Function IBenefitClass_DeleteDB() As Boolean
  On Error GoTo DeleteDB_Err
  
  Call xSet("DeleteDB")
  
'MP DB
' If Not IsEmpty(m_BenItems(cdc_Key)) Then Call m_Parent.db.Execute(sql.Queries(DELETE_CDC, m_BenItems(cdc_Key)))
  If Not IsEmpty(m_BenItems(cdc_Key_db)) Then Call m_Parent.db.Execute(sql.Queries(DELETE_CDC, m_BenItems(cdc_Key_db)))
  IBenefitClass_DeleteDB = True
DeleteDB_End:
  Call xReturn("DeleteDB")
  Exit Function
DeleteDB_Err:
'MP DB
' Call ErrorMessage(ERR_ERROR, Err, "DeleteDB", "Delete DB", "Error deleting a CDC benefit, cdc name = " & m_BenItems(cdc_name) & ".")
  Call ErrorMessage(ERR_ERROR, Err, "DeleteDB", "Delete DB", "Error deleting a CDC benefit, cdc name = " & m_BenItems(cdc_name_db) & ".")
  Resume DeleteDB_End
  Resume
End Function

Private Property Let IBenefitClass_Dirty(RHS As Boolean)
  m_bDirty = RHS
End Property

Private Property Get IBenefitClass_Dirty() As Boolean
  IBenefitClass_Dirty = m_bDirty
End Property

Private Property Get IBenefitClass_HasBookMark() As Boolean

End Property

Private Property Let IBenefitClass_InvalidFields(ByVal RHS As Long)

End Property

Private Property Get IBenefitClass_InvalidFields() As Long

End Property

Private Sub IBenefitClass_Kill()
  Set m_Parent = Nothing
End Sub

Private Property Get IBenefitClass_Name() As String
'MP DB
'  IBenefitClass_Name = m_BenItems(cdc_name)
  IBenefitClass_Name = m_BenItems(cdc_name_db)
End Property

Private Property Set IBenefitClass_Parent(RHS As Object)
  Set m_Parent = RHS
End Property

Private Property Get IBenefitClass_Parent() As Object
  Set IBenefitClass_Parent = m_Parent
End Property

Private Function IBenefitClass_PrintWk(rep As Reporter) As Boolean

End Function

Private Function IBenefitClass_PrintWkBody(rep As Reporter) As Boolean

End Function

Private Function IBenefitClass_ReadDB() As Long
  
End Function

Private Property Let IBenefitClass_ReadFromDB(ByVal RHS As Boolean)
  m_ReadFromDB = RHS
End Property

Private Property Get IBenefitClass_ReadFromDB() As Boolean
  IBenefitClass_ReadFromDB = m_ReadFromDB
End Property

Private Property Let IBenefitClass_RSBookMark(RHS As String)

End Property

Private Property Get IBenefitClass_RSBookMark() As String

End Property

Private Sub IBenefitClass_SetBenItemsInformation()

End Sub

Private Sub IBenefitClass_SetCalcDefaults()

End Sub

Private Property Get IBenefitClass_TABLE() As BENEFIT_TABLES

End Property

Private Property Let IBenefitClass_value(ByVal Item As Long, RHS As Variant)
  m_BenItems(Item) = RHS
End Property

Private Property Get IBenefitClass_value(ByVal Item As Long) As Variant
  IBenefitClass_value = m_BenItems(Item)
End Property

Private Function IBenefitClass_WriteDB() As Boolean
  Dim rs As Recordset
  Dim Employer As Employer
  Dim ben As IBenefitClass
  
  Set ben = Me
  
  If Not m_ReadFromDB Then
    Set Employer = m_Parent
    Set rs = Employer.db.OpenRecordset(sql.Queries(SELECT_CDCS))
    rs.AddNew
'MP DB
'    rs!Name = ben.value(cdc_name)
'    ben.value(cdc_Key) = rs!CDCKey
    rs!Name = ben.value(cdc_name_db)
    ben.value(cdc_Key_db) = rs!CDCKey
    rs.Update
    m_ReadFromDB = True
  End If
End Function




