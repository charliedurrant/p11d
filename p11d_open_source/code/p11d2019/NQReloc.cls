VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NonQualifyingRelocation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IBenefitClass

Public RelocationBen As IBenefitClass
Private m_CompanyDefined As Boolean
Private m_NeedToCalculate As Boolean

Private Sub Class_Terminate()
  'Debug.Print "Non qualifying relocation - terminate"
End Sub

Private Function IBenefitClass_CalculateBody() As Variant

End Function

Private Property Get IBenefitClass_ImageListKey() As String
  IBenefitClass_ImageListKey = "Accommodation"
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
  'leave as false as qualifying will deal with it
End Function

Private Function IBenefitClass_Copy(Parent As Object) As IBenefitClass

End Function

Private Function AdjustItemEnum(ByVal Item As RelocItems) As RelocItems
  On Error GoTo AdjustItemEnum_ERR
  Call xSet("AdjustItemEnum")
  
  If Item < reloc_NQ_SPECIFIC_START Then
    AdjustItemEnum = Item + (reloc_NQ_SPECIFIC_START - 1)
  ElseIf Item > reloc_NQ_SPECIFIC_END Then
    AdjustItemEnum = Item
  Else
    Call Err.Raise(ERR_ENUM__VALUE_INVALID, "AdjustItemEnum", "Enum value is not valid for NonQualifyingRelocation benefit.")
  End If
  
  
AdjustItemEnum_END:
  Call xReturn("AdjustItemEnum")
  Exit Function
AdjustItemEnum_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "AdjustItemEnum", "Adjust Item Enum", "The Item number passed is outside the enum range of RelocItems.")
  Resume AdjustItemEnum_END
  Resume
End Function

Private Sub Class_Initialize()
  Call IBenefitClass_SetBenItemsInformation
End Sub

Private Property Let IBenefitClass_BenefitClass(NewValue As BEN_CLASS)
  
End Property

Private Property Get IBenefitClass_BenefitClass() As BEN_CLASS
  IBenefitClass_BenefitClass = BC_NON_QUALIFYING_RELOCATION_N
End Property

Private Function IBenefitClass_Calculate() As Variant
  Call RelocationBen.Calculate
  IBenefitClass_Calculate = IBenefitClass_value(ITEM_BENEFIT)
End Function
Private Property Let IBenefitClass_CompanyDefined(ByVal NewValue As Boolean)
  m_CompanyDefined = NewValue
End Property
Private Property Get IBenefitClass_CompanyDefined() As Boolean
  IBenefitClass_CompanyDefined = m_CompanyDefined
End Property

Private Function IBenefitClass_DeleteDB() As Boolean
  'do no map to parent
End Function
Private Property Let IBenefitClass_Dirty(NewValue As Boolean)
  RelocationBen.Dirty = NewValue
End Property

Private Property Get IBenefitClass_Dirty() As Boolean
  IBenefitClass_Dirty = RelocationBen.Dirty
End Property

Private Function IBenefitClass_PrintWkBody(rep As Reporter) As Boolean

End Function

Private Sub IBenefitClass_SetCalcDefaults()
  Call RelocationBen.SetCalcDefaults
End Sub
Private Property Get IBenefitClass_HasBookMark() As Boolean
  IBenefitClass_HasBookMark = RelocationBen.HasBookMark
End Property
Private Property Let IBenefitClass_InvalidFields(ByVal NewValue As Long)
  RelocationBen.InvalidFields = NewValue
End Property
Private Property Get IBenefitClass_InvalidFields() As Long
  IBenefitClass_InvalidFields = RelocationBen.InvalidFields
End Property
Private Sub IBenefitClass_Kill()
  Set RelocationBen = Nothing
End Sub
Private Property Get IBenefitClass_Name() As String
  IBenefitClass_Name = RelocationBen.Name
End Property
Private Property Set IBenefitClass_Parent(NewValue As Object)
  Set RelocationBen.Parent = NewValue
End Property

Private Property Get IBenefitClass_Parent() As Object
  Set IBenefitClass_Parent = RelocationBen.Parent
End Property
Private Property Get IBenefitClass_PrintHeader() As String

End Property
Private Function IBenefitClass_PrintWk(rep As Reporter) As Boolean

End Function
Private Function IBenefitClass_ReadDB() As Long
  IBenefitClass_ReadDB = RelocationBen.ReadDB
End Function
Private Property Let IBenefitClass_ReadFromDB(ByVal NewValue As Boolean)
  'do not map to parent
End Property
Private Property Get IBenefitClass_ReadFromDB() As Boolean
  IBenefitClass_ReadFromDB = RelocationBen.ReadFromDB
End Property

Private Property Let IBenefitClass_RSBookMark(NewValue As String)
  'do not map to parent
End Property
Private Property Get IBenefitClass_RSBookMark() As String
  IBenefitClass_RSBookMark = RelocationBen.RSBookMark
End Property

Private Sub IBenefitClass_SetBenItemsInformation()
  Dim bc As BEN_CLASS
  
  On Error GoTo SetBenItemsInformation_err
  bc = BC_NON_QUALIFYING_RELOCATION_N
  m_NeedToCalculate = True
  If p11d32.DataLinkInitialised(bc) Then GoTo SetBenItemsInformation_end
  
  Call SetStandardBenItemsInformation(bc, Me)

SetBenItemsInformation_end:
  p11d32.DataLinkInitialised(bc) = True
  Exit Sub
  
SetBenItemsInformation_err:
  Call ErrorMessage(ERR_ERROR, Err, "SetBenItemsInformation", "Set Benefit Item Information", "Error setting benefit information.")
  Resume SetBenItemsInformation_end
End Sub
Private Property Get IBenefitClass_TABLE() As BENEFIT_TABLES
  IBenefitClass_TABLE = TBL_RELOC
End Property

Private Property Let IBenefitClass_value(ByVal Item As Long, RHS As Variant)
  RelocationBen.value(AdjustItemEnum(Item)) = RHS
End Property

Private Property Get IBenefitClass_value(ByVal Item As Long) As Variant
  IBenefitClass_value = RelocationBen.value(AdjustItemEnum(Item))
End Property
Private Function IBenefitClass_WriteDB() As Boolean
  IBenefitClass_WriteDB = True
End Function


