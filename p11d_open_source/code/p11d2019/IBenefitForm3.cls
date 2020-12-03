VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IBenefitForm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum BENEFIT_FORM_STATE
  FORM_ENABLED
  FORM_DISABLED
  FORM_CDB
  FORM_LINK_BEN
End Enum

Public Function AddBenefitSetDefaults(ben As IBenefitClass) As Long
    
End Function
Public Function BenefitToListView(ben As IBenefitClass, ByVal lBenefitIndex As Long) As Long

End Function

Public Function BenefitsToListView() As Long

End Function

' Enable form/Disable form/Enable CDB (disable form & enable CDB buttons)
Public Function BenefitFormState(ByVal fState As BENEFIT_FORM_STATE) As Boolean

End Function

' if BenefitIndex = -1 then Clear benefit form return to default state
' This function is resposible for calling BenefitFormState
Public Function BenefitToScreen(Optional ByVal BenefitIndex As Long = -1, Optional ByVal UpdateBenefit As Boolean = True) As Boolean
    ' Called from SelectBenefitByBenefitIndex etc..
    
End Function
    
Public Function BenefitOn() As Boolean
    
End Function
    
Public Function BenefitOff() As Boolean
    
End Function
    
Public Property Get lv() As MSComctlLib.IListView
  
End Property
    
Public Sub AddBenefit()

End Sub
Public Function ValididateBenefit(ben As IBenefitClass) As Boolean
  
End Function

' Default RemoveBenefit
Public Function RemoveBenefit(ByVal BenefitIndex As Long) As Boolean

End Function

Public Property Let benclass(ByVal NewValue As BEN_CLASS)

End Property

Public Property Get benclass() As BEN_CLASS

End Property

'default is call UpdateBenefitListViewItem from FormControl.bas
Public Function UpdateBenefitListViewItem(li As MSComctlLib.IListItem, benefit As IBenefitClass, Optional ByVal BenefitIndex As Long = 0, Optional ByVal SelectItem As Boolean = False) As Long
  'when varying benefit data we must update the list view
End Function

Public Property Get benefit() As IBenefitClass
  
End Property
Public Property Set benefit(NewValue As IBenefitClass)
  
End Property


