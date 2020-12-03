VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
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
End Enum
' Note: Default functions in FormControl

' Default BenefitsToListView
Public Function BenefitsToListView() As Long

End Function

' Enable form/Disable form/Enable CDB (disable form & enable CDB buttons)
Public Function BenefitFormState(ByVal fState As BENEFIT_FORM_STATE) As Boolean

End Function

' if BenefitIndex = -1 then Clear benefit form return to default state
' This function is resposible for calling BenefitFormState
Public Function BenefitToScreen(Optional ByVal BenefitIndex As Long = -1, Optional ByVal UpdateBenefit As Boolean = True) As IBenefitClass
    ' Called from SelectBenefit etc..
    
End Function
    
Public Property Get lv() As ListView
  
End Property
    
Public Sub AddBenefit()

End Sub

' Default RemoveBenefit
Public Function RemoveBenefit(ByVal BenefitIndex As Long) As Boolean

End Function

Public Property Let bentype(ByVal NewValue As benClass)

End Property

Public Property Get bentype() As benClass

End Property

'default is call UpdateBenefitListViewItem from FormControl.bas
Public Function UpdateBenefitListViewItem(li As ListItem, benefit As IBenefitClass, Optional ByVal BenefitIndex As Long = 0, Optional ByVal SelectItem As Boolean = False) As Boolean
  'when varying benefit data we must update the list view
End Function

Public Property Get benefit() As IBenefitClass
  
End Property
Public Property Let benefit(NewValue As IBenefitClass)
  
End Property

