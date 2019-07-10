Attribute VB_Name = "Dialog"
Option Explicit
Public Function DialogToScreen(ibfDialog As IBenefitForm2, LabelToUpdate, BCForLabel As BEN_CLASS, ibfParent As IBenefitForm2, BenefitIndex As Long, Optional bUnloadDialog As Boolean = True) As Boolean
  Dim frm As Form
  Dim bWasDirty As Boolean
  
On Error GoTo DialogToScreen_Err
  
  Call xSet("DialogToScreen")
  
  If BenefitIndex < 1 Then Call Err.Raise(ERR_INVALID_BENEFIT_INDEX, "DialogToScreen", "The benefit index is invalid, index = " & BenefitIndex)
  Set CurrentForm = Nothing
  Set frm = ibfDialog
  Set frm.Parentibf = ibfParent
  Set ibfDialog.benefit = ibfParent.benefit
  
  Call TestChangedControls(ibfParent)
  
  bWasDirty = ibfParent.benefit.Dirty
  
  ibfParent.benefit.Dirty = False
  ibfDialog.benclass = ibfParent.benclass
  ibfDialog.BenefitToScreen (BenefitIndex)
  
  'KM - added bUnloadDialog boolean for the Car Make - CO2Emissions
  'dialog.  Didn't want this to unload, just show and hide instead
  If bUnloadDialog Then 'km
    Unload frm
  End If
  Set CurrentForm = ibfParent
  
  If ibfParent.benefit.Dirty Then
    Call ibfParent.UpdateBenefitListViewItem(ibfParent.lv.SelectedItem, ibfParent.benefit)
    Call ibfParent.BenefitOn
    Call TestChangedControls(ibfParent) 'next two lines seem strange, CAD 20032004
    ibfParent.benefit.InvalidFields = InvalidFields(ibfParent)
    
    If ibfParent.benefit.Dirty Then MDIMain.SetConfirmUndo
    If Not LabelToUpdate Is Nothing Then
      If TypeOf LabelToUpdate Is ValText Then
        LabelToUpdate.Text = RoundN(ibfParent.benefit.value(BCForLabel), 0)
      Else
        LabelToUpdate.Caption = FormatWN(ibfParent.benefit.value(BCForLabel), "£")
      End If
      
    End If
  Else
    ibfParent.benefit.Dirty = bWasDirty
  End If
    
    
  DialogToScreen = True
  
DialogToScreen_End:
  Set CurrentForm = ibfParent
  Set frm = Nothing
  Call xReturn("DialogToScreen")
  Exit Function
DialogToScreen_Err:
  Call ErrorMessage(ERR_ERROR, Err, "DialogToScreen", "Fuel To Screen", "Unable to place the dialogue benefit to screen.")
  Resume DialogToScreen_End
  Resume
End Function

Public Function ChkBoxToBool(Chk As CheckBox) As Boolean
  On Error GoTo ChkBoxToBool_ERR
  
  Call xSet("ChkBoxToBool")
  
  If Chk Is Nothing Then Call Err.Raise(ERR_IS_NOTHING, "ChkBoxToBool", "The check boc is nothing.")
  
  If Chk.value = vbChecked Then ChkBoxToBool = True
  
ChkBoxToBool_END:
  Call xReturn("ChkBoxToBool")
  Exit Function
ChkBoxToBool_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "ChkBoxToBool", "Chk Box To Bool", "Error converting a checkbox value to boolean.")
  Resume ChkBoxToBool_END
  
End Function
Public Function BoolToChkBox(ByVal b As Boolean) As CheckBoxConstants
  If b Then
    BoolToChkBox = vbChecked
  Else
    BoolToChkBox = vbUnchecked
  End If
End Function
