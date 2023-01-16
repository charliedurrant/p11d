VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IBenefitClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public Enum BASE_ITEMS
  ITEM_NONE = -1
  ITEM_DESC = 1
  ITEM_VALUE
  ITEM_MADEGOOD_NET
  ITEM_BENEFIT
  ITEM_MADEGOOD
  ITEM_ACTUALAMOUNTMADEGOOD
  ITEM_BENEFIT_REPORTABLE
  ITEM_UDM_BENEFIT_TITLE
  ITEM_BOX_NUMBER
  ITEM_MADEGOOD_IS_TAXDEDUCTED
  ITEM_CLASS1A_ADJUSTMENT
  ITEM_NIC_CLASS1A_ABLE
  ITEM_BENEFIT_SUBJECT_TO_CLASS1A
  ITEM_NIC_CLASS1A_BENEFIT
  ITEM_ERROR
  ITEM_VALUE_NON_OPRA
  ITEM_OPRA_AMOUNT_FOREGONE_USED_FOR_VALUE
  ITEM_OPRA_AMOUNT_FOREGONE
  [_ITEM_LAST_ITEM] = ITEM_OPRA_AMOUNT_FOREGONE
End Enum
Public Property Get ImageListKey() As String

End Property

Public Function CalculateBody() As Variant
  
End Function
Public Property Let NeedToCalculate(ByVal NewValue As Boolean)
  
End Property
Public Property Get NeedToCalculate() As Boolean
    
End Property
Public Property Get LinkBen() As Boolean

End Property
Public Property Let LinkBen(NewValue As Boolean)

End Property
Public Function CanBringForward() As Boolean
  
End Function
Public Function WriteDB() As Boolean
  
End Function
Public Function Copy(Parent As Object) As IBenefitClass
  
End Function
Public Property Get HasBookMark() As Boolean
  
End Property

Public Property Get RSBookMark() As String
  
End Property
Public Property Let RSBookMark(NewValue As String)
  
End Property
Public Sub SetBenItemsInformation()
    
End Sub
Public Sub SetCalcDefaults()
    
End Sub
Public Property Get value(ByVal Item As Long) As Variant

End Property
Public Property Let value(ByVal Item As Long, value As Variant)

End Property
Public Function ReadDB() As Long
  
End Function
Public Function DeleteDB() As Boolean
  'Delete reference in database to the benefit
End Function
Public Function Calculate() As Variant

End Function
Public Function PrintWk(rep As Reporter) As Boolean

End Function
Public Function PrintWkBody(rep As Reporter) As Boolean
  
End Function

Public Sub Kill()
  
End Sub
Public Property Get Name() As String

End Property

Public Property Set Parent(NewValue As Object)

End Property
Public Property Get Parent() As Object
  
End Property

Public Property Let CompanyDefined(ByVal NewValue As Boolean)

End Property

Public Property Get CompanyDefined() As Boolean
  
End Property

Public Property Get BenefitClass() As BEN_CLASS

End Property

Public Property Let BenefitClass(ben As BEN_CLASS)
  
End Property


Public Property Get Dirty() As Boolean
  
End Property

Public Property Let Dirty(NewValue As Boolean)

End Property

Public Property Get TABLE() As BENEFIT_TABLES

End Property

Public Property Get InvalidFields() As Long

End Property

Public Property Let InvalidFields(ByVal NewValue As Long)
  
End Property

Public Property Let ReadFromDB(ByVal NewValue As Boolean)
  
End Property

Public Property Get ReadFromDB() As Boolean
  
End Property


Private Sub Class_Initialize()
'
End Sub
