Attribute VB_Name = "FormControl"
Option Explicit

Public Const S_CHANGEDINDICATOR As String = "C"
Private mCurrent As ListItem
Private mCurrentText As String
Public mLast As ListItem
Public Enum STDV_DEFAULTS
  STDV_STRING = 1
  STDV_UNDATED
End Enum

Public Sub SetDefaultVTDate(vtDate As ValText, Optional dMinimum As Date = UNDATED, Optional dMaximum As Date = UNDATED, Optional bNoMinimum = False, Optional bNoMaximum = False)
  
  On Error GoTo SetDefaultVTDate_ERR
  
  Call xSet("SetDefaultVTDate")
  
  If vtDate Is Nothing Then Call Err.Raise(ERR_IS_NOTHING, "SetDefaultVTDate", "The val text is nothing.")
  
  vtDate.TypeOfData = VT_DATE
  
  If dMinimum <> UNDATED Then
    vtDate.Minimum = DateStringEx(dMinimum, p11d32.Rates.value(TaxYearStart))
  Else
    If bNoMinimum Then
      vtDate.Minimum = ""
    Else
      vtDate.Minimum = DateStringEx(p11d32.Rates.value(TaxYearStart), p11d32.Rates.value(TaxYearStart))
    End If
  End If
  
  If dMaximum <> UNDATED Then
    vtDate.Maximum = DateStringEx(dMaximum, p11d32.Rates.value(TaxYearEnd))
  Else
    If bNoMaximum Then
      vtDate.Maximum = ""
    Else
      vtDate.Maximum = DateStringEx(p11d32.Rates.value(TaxYearEnd), p11d32.Rates.value(TaxYearEnd))
    End If
  End If
  
SetDefaultVTDate_END:
  Call xReturn("SetDefaultVTDate")
  Exit Sub
SetDefaultVTDate_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "SetDefaultVTDate", "Set Default VT Date", "Error setting the default from / to dates for a valtext.")
  Resume SetDefaultVTDate_END
  
End Sub
Public Function IsKeyBoardHookForm(frm As Form) As Boolean
  On Error GoTo IsKeyBoardHookForm_END
  IsKeyBoardHookForm = IsBenefitFormReal(frm) Or frm Is F_Employees
  
IsKeyBoardHookForm_END:
End Function
Public Function IsBenefitFormReal(frm As Form) As Boolean
  Dim ibf As IBenefitForm2
  
  
  On Error GoTo IsBenefitFormReal_ERR
  Call xSet("IsBenefitFormReal")
  
  If Not IsBenefitForm(frm) Then GoTo IsBenefitFormReal_END
  Set ibf = frm
  IsBenefitFormReal = ((ibf.benclass <= BC_UDM_BENEFITS_LAST_ITEM) And (ibf.benclass >= BC_FIRST_ITEM)) Or ibf.benclass = BC_CDB
  
IsBenefitFormReal_END:
  Call xReturn("IsBenefitFormReal")
  Exit Function
IsBenefitFormReal_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "IsBenefitForm", "Is Benefit Form", "Error determining if a form is real benefit form.")
  Resume IsBenefitFormReal_END
End Function
Public Function IsBenefitForm(frm As Form) As Boolean
  Dim ibf As IBenefitForm2
  
  On Error GoTo IsBenefitForm_END
  Set ibf = frm
  IsBenefitForm = True
IsBenefitForm_END:
End Function

Public Sub SetSortOrderToColumn(lv As ListView, lColIndex, SortOrder As ListSortOrderConstants)
  On Error GoTo SetSortOrderToColumn_ERR
  
  Call xSet("SetSortOrderToColumn")
  
  If lv Is Nothing Then Call Err.Raise(ERR_LV_IS_NOTHING, "SetSortOrderToColumn", "The colindex of " & lColIndex & " is outside the column headers range.")
  If lColIndex <= lv.ColumnHeaders.Count And lColIndex > 0 Then
    If lColIndex > 1 Then Call SetSortOrder(lv, lv.ColumnHeaders(lColIndex), SortOrder)
  Else
    Call Err.Raise(ERR_INVALIDCOL_INDEX, "SetSortOrderToColumn", "The colindex of " & lColIndex & " is outside the column headers range.")
  End If
  
SetSortOrderToColumn_END:
  Call xReturn("SetSortOrderToColumn")
  Exit Sub
SetSortOrderToColumn_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "SetSortOrderToColumn", "Set Sort Order To Column", "Error setting the sort order column.")
  Resume SetSortOrderToColumn_END
End Sub
Public Sub SetPanel2(Caption As String)
  If app.StartMode = vbSModeStandalone Then MDIMain.sts.Panels(S_P2).Caption = Caption
End Sub
Public Sub SetPanel1(Caption As String)
  If app.StartMode = vbSModeStandalone Then MDIMain.sts.Panels(S_P1).Caption = Caption
End Sub
Public Sub PrgStep()
  If app.StartMode = vbSModeStandalone Then Call MDIMain.sts.Step
End Sub
Public Sub PrgStepCaption(Caption As String)
  If app.StartMode = vbSModeStandalone Then Call MDIMain.sts.StepCaption(Caption)
End Sub
Public Sub PrgStopCaption()
  If app.StartMode = vbSModeStandalone Then
    Call MDIMain.sts.StopPrg
    Call SetPanel2("")
  End If
End Sub
Public Sub PrgStop()
  If app.StartMode = vbSModeStandalone Then Call MDIMain.sts.StopPrg
End Sub
Public Sub PrgStartCaption(Max As Long, Optional PanelCaption As String = "", Optional PrgCaption As String = "", Optional ByVal Indicator As Indicator = Percentage)
  If app.StartMode = vbSModeStandalone Then
    Call SetPanel2(PanelCaption)
    Call MDIMain.sts.StartPrg(Max, PrgCaption, Indicator)
  End If
End Sub
Public Sub PrgStart(Max As Long, Optional PrgCaption As String = "", Optional ByVal Indicator As Indicator = Percentage)
  If app.StartMode = vbSModeStandalone Then Call MDIMain.sts.StartPrg(Max, PrgCaption, Indicator)
End Sub

Public Sub PrgAlignment(ByVal AL As TextAlignment)
  Dim prg As Object 'TCSProgressBar
  
  On Error GoTo PrgAlignment_ERR
  
  If app.StartMode <> vbSModeStandalone Then GoTo PrgAlignment_END
  Call xSet("PrgAlignment")
  Set prg = MDIMain.sts.prg
  If AL = 0 Then
    prg.TextAlignment = Align_Centre
  Else
    prg.TextAlignment = AL
  End If
  
PrgAlignment_END:
  Call xReturn("PrgAlignment")
  Exit Sub
PrgAlignment_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "PrgAlignment", "Prg Alignment", "Error setting the alignment of the progress bar.")
  Resume PrgAlignment_END
  Resume
End Sub
Private Function GetListItemIndexbyText(lv As ListView) As Long
  Dim i As Long
  
  GetListItemIndexbyText = 1
  If (Len(mCurrentText) > 0) Then
    For i = 1 To lv.listitems.Count
      If StrComp(lv.listitems(i).Text, mCurrentText) = 0 Then
        GetListItemIndexbyText = i
      End If
    Next i
  End If
  
End Function

Public Function UpdateBenefitListViewItem(li As ListItem, benefit As IBenefitClass, Optional BenefitIndex As Long = 0, Optional ByVal SelectItem As Boolean = False) As Long
  
  On Error GoTo UpdateBenefitListViewItem_ERR
  Call xSet("UpdateBenefitListViewItem")
  
  If Not li Is Nothing And Not benefit Is Nothing Then
    If BenefitIndex > 0 Then li.Tag = BenefitIndex
    li.SmallIcon = benefit.ImageListKey
    li.SubItems(1) = FormatWN(benefit.Calculate)
    li.Text = benefit.value(ITEM_DESC)
    If SelectItem Then li.Selected = SelectItem
    UpdateBenefitListViewItem = li.Index
  End If
  
UpdateBenefitListViewItem_END:
  Call xReturn("UpdateBenefitListViewItem")
  Exit Function
UpdateBenefitListViewItem_ERR:
  UpdateBenefitListViewItem = False
  Call ErrorMessage(ERR_ERROR, Err, "UpdateBenefitListViewItem", "Updating a benefits list view text", "Unable to update the benefits list view text.")
  Resume UpdateBenefitListViewItem_END
  Resume
End Function
Public Function IRDescriptionFromCombo(cbo As ComboBox, ben As IBenefitClass) As Boolean
  
  Dim s As String
  Dim iBenITem As Long
  s = cbo.List(cbo.ListIndex)
  iBenITem = IRDescriptionBenItem(ben.BenefitClass)
  If StrComp(s, ben.value(iBenITem), vbTextCompare) <> 0 Then
    ben.value(iBenITem) = s
    IRDescriptionFromCombo = True
  End If
    
End Function

Public Sub IRDescriptionToCombo(cbo As ComboBox, ByVal ben As IBenefitClass)
  Dim i As Long
  Dim s As String
  Dim iBenITem As Long
  
  On Error GoTo err_err
  
  
  iBenITem = IRDescriptionBenItem(ben.BenefitClass)
  s = ben.value(iBenITem)
  For i = 0 To cbo.ListCount - 1
    If (StrComp(cbo.List(i), s, vbTextCompare) = 0) Then
      cbo.ListIndex = i
      GoTo err_end
    End If
  Next
  Call Err.Raise(ERR_DISPLAY, "IRDescriptionToCombo", "Failed to set the HMRC description")
  
err_end:
  Exit Sub
err_err:
  Call ErrorMessage(ERR_ERROR, Err, "IRDescriptionToCombo", "IR Description", "Error setting the IR description to the combo box.")
  Resume err_end
  Resume
End Sub


Public Function BenefitsToListView(ibf As IBenefitForm2) As Long
  Dim i As Long
  Dim ben As IBenefitClass
  Dim lst As ListItem
  
  
  On Error GoTo BenefitsToListView_err
  Call xSet("BenefitsToListView")
  
  Call ClearForm(ibf)
  Call MDIMain.SetAdd
  
  ibf.lv.Sorted = False
  
  For i = 1 To p11d32.CurrentEmployer.CurrentEmployee.benefits.Count
    Set ben = p11d32.CurrentEmployer.CurrentEmployee.benefits(i)
    BenefitsToListView = BenefitsToListView + ibf.BenefitToListView(ben, i)
  Next
  ibf.lv.Sorted = True

BenefitsToListView_end:
  Set ben = Nothing
  Set lst = Nothing
  Call xReturn("BenefitsToListView")
  Exit Function
  
BenefitsToListView_err:
  Call ErrorMessage(ERR_ERROR, Err, "BenefitsToListView", "Adding Benefits to ListView", "Error in adding benefits to listview.")
  Resume BenefitsToListView_end
End Function

Public Sub ClearForm(ibf As IBenefitForm2)
  ibf.lv.listitems.Clear
  Set ibf.lv.SmallIcons = MDIMain.imlListViewBenefits
  Call ibf.BenefitToScreen(, False)
End Sub

Private Function TestChanged(Tag As String) As Boolean
  If StrComp(Tag, S_CHANGEDINDICATOR) = 0 Then TestChanged = True
End Function

Public Function UpdateBenefitFromTags() As Boolean
  Dim ben As IBenefitClass
  Dim ibf As IBenefitForm2
  Dim bChanged As Boolean
    
  On Error GoTo UpdateBenefitFromTags_err
  Call xSet("UpdateBenefitFromTags")
  Set ibf = p11d32.GetBenefitForm
  If Not ibf Is Nothing Then
    Set ben = ibf.benefit
    If Not ben Is Nothing Or ibf.benclass = BC_NONSHAREDVANS_G Then
      bChanged = TestChangedControls(ibf)
      UpdateBenefitFromTags = bChanged
    Else
      UpdateBenefitFromTags = True
    End If
  End If
  
  Call SetLastListItemSelected(Nothing)
  
UpdateBenefitFromTags_end:
  Call xReturn("UpdateBenefitFromTags")
  Exit Function
  
UpdateBenefitFromTags_err:
  UpdateBenefitFromTags = False
  Call ErrorMessage(ERR_ERROR, Err, "UpdateBenefitFromTags", "Update benefit", "Benefit was not updated correctly on screen.")
  Resume UpdateBenefitFromTags_end
  Resume
End Function

Public Function SetLastListItemSelected(li As ListItem) As Boolean
  If mLast Is Nothing Then
    Set mLast = li
  Else
    If li Is Nothing Then
      Set mLast = Nothing
    Else
      Set mLast = mCurrent
    End If
  End If
  Set mCurrent = li
  If Not mCurrent Is Nothing Then mCurrentText = li.Text
End Function
Public Function GetNextBestListItemBenefitIndex(ibf As IBenefitForm2, ByVal CurrentBenefitIndex As Long) As Long
  Dim lv As ListView, i As Long
  
  On Error GoTo GetNextBestListItemBenefitIndex_ERR
  Call xSet("GetNextBestListItemBenefitIndex")
  
  GetNextBestListItemBenefitIndex = -1
  
  Set lv = ibf.lv
  For i = 1 To lv.listitems.Count
    If lv.listitems(i).Tag = CurrentBenefitIndex Then
      If (i + 1) < lv.listitems.Count Then
        GetNextBestListItemBenefitIndex = lv.listitems(i + 1).Tag
      ElseIf i > 1 Then
        GetNextBestListItemBenefitIndex = lv.listitems(i - 1).Tag
      End If
    End If
  Next i
  
GetNextBestListItemBenefitIndex_END:
  Call xReturn("GetNextBestListItemBenefitIndex")
  Exit Function

GetNextBestListItemBenefitIndex_ERR:
  Resume GetNextBestListItemBenefitIndex_END
End Function

Public Function GetNextBestListItem(liRet As ListItem, lv As ListView, li As ListItem, Optional SelectItem As Boolean = False) As Boolean
  Dim lCount As Long
  
  On Error GoTo GetNextBestListItem_Err
  Call xSet("GetNextBestListItem")
    
  lCount = lv.listitems.Count
  If li.Index < lCount Then
    Set liRet = lv.listitems(li.Index + 1)
  ElseIf li.Index > 1 Then
    Set liRet = lv.listitems(li.Index - 1)
  End If
    
  If Not liRet Is Nothing Then
    If SelectItem Then Set lv.SelectedItem = liRet
    GetNextBestListItem = True
  End If
  
GetNextBestListItem_End:
  Call xReturn("GetNextBestListItem")
  Exit Function
  
GetNextBestListItem_Err:
  Set liRet = Nothing
  GetNextBestListItem = False
  Call ErrorMessage(ERR_ERROR, Err, "GetNextBestListItem", "Get Next Best List Item", "Error finding the next best list item.")
  Resume GetNextBestListItem_End
  Resume
End Function
Public Function GetEmployeeIndexFromSelectedEmployee() As Long

  On Error GoTo GetEmployeeIndexFromSelectedEmployee_ERR
  
  Call xSet("GetEmployeeIndexFromSelectedEmployee")
  
  With F_Employees
    If Not (F_Employees.LB.SelectedItem Is Nothing) Then
      GetEmployeeIndexFromSelectedEmployee = F_Employees.LB.SelectedItem.Tag
    Else
      GetEmployeeIndexFromSelectedEmployee = -1
    End If
  End With
  
GetEmployeeIndexFromSelectedEmployee_END:
  Call xReturn("GetEmployeeIndexFromSelectedEmployee")
  Exit Function
GetEmployeeIndexFromSelectedEmployee_ERR:
  GetEmployeeIndexFromSelectedEmployee = -1
  Call ErrorMessage(ERR_ERROR, Err, "GetEmployeeIndexFromSelectedEmployee", "Get Employee Index From Selected Employee", "Error getting the employee index from the current selected employee")
  Resume GetEmployeeIndexFromSelectedEmployee_END
End Function

Public Sub UpdateInfoStatusBar(Optional ben As IBenefitClass = Nothing)
  On Error GoTo err_err
  
  If ben Is Nothing Then
    Call SetPanel2("")
  Else
    If ben.value(ITEM_BENEFIT) = S_ERROR Then
      Call SetPanel2(ben.value(ITEM_ERROR))
    Else
      Call SetPanel2("")
    End If
  End If
    
err_end:
  Exit Sub
err_err:
  Resume err_end
End Sub
Public Function AfterCheckChanged(c As Control, ibf As IBenefitForm2, ByVal bDirty As Boolean, Optional ctlContainer As Object = Nothing, Optional bSaveBenefitStatus As Boolean = True) As Boolean
  On Error GoTo AfterCheckChanged_ERR
  Call xSet("AfterCheckChanged")
  
  If Not ibf Is Nothing Then
    If Not ibf.benefit Is Nothing Then
      ibf.benefit.Dirty = ibf.benefit.Dirty Or bDirty
      If bDirty Then
        Call UpdateInfoStatusBar(ibf.benefit)
        Call ibf.UpdateBenefitListViewItem(ibf.lv.SelectedItem, ibf.benefit)
        c.Tag = ""
        If bSaveBenefitStatus Then Call SaveBenefitStatus(ibf.benefit)
        ibf.benefit.InvalidFields = InvalidFields(ibf, ctlContainer)
      End If
      'if we changed then check validity
      AfterCheckChanged = True
    End If
  End If
  
AfterCheckChanged_END:
 
  Call xReturn("AfterCheckChanged")
  Exit Function
  
AfterCheckChanged_ERR:
  AfterCheckChanged = False
  Call ErrorMessage(ERR_ERROR, Err, "AfterCheckChanged", "After Checkchanged", "Error changing benefit parameters after assessing if there was new user input.")
  Resume AfterCheckChanged_END
  Resume
End Function
Public Function SaveBenefitStatus(benefit As IBenefitClass) As Boolean

  On Error GoTo SaveBenefitStatus_ERR
  
  Call xSet("SaveBenefitStatus")
  
  If benefit.Dirty Then
    Call SetPanel2("")
    Call MDIMain.SetConfirmUndo
  ElseIf benefit.InvalidFields Then
    Call SetPanel2(S_NOSAVE)
    Call MDIMain.SetConfirmUndo
  End If
  
  SaveBenefitStatus = True
  
SaveBenefitStatus_END:
  Call xReturn("SaveBenefitStatus")
  Exit Function
  
SaveBenefitStatus_ERR:
  SaveBenefitStatus = False
  Call ErrorMessage(ERR_ERROR, Err, "SaveBenefitStatus", "Save Benefit Status", "Error setting the save status for a benefit")
  Resume SaveBenefitStatus_END
End Function

Public Function GetBenefitRecord(rs As Recordset, benefit As IBenefitClass)

On Error GoTo GetBenefitRecord_ERR

  Call xSet("GetBenefitRecord")
  
  If Not rs Is Nothing And Not benefit Is Nothing Then
    If benefit.HasBookMark Then
      rs.Bookmark = benefit.RSBookMark
      rs.Edit
    Else
      rs.AddNew
      rs.Fields(S_FIELD_PERSONEL_NUMBER) = GetEmployeeNumber(benefit)
    End If
    GetBenefitRecord = True
  End If
  
GetBenefitRecord_END:
  Call xReturn("GetBenefitRecord")
  Exit Function
GetBenefitRecord_ERR:
  GetBenefitRecord = False
  Call ErrorMessage(ERR_ERROR, Err, "GetBenefitRecord", "Get Benefit Record", "Error getting the record for the current benefit.")
  Resume GetBenefitRecord_END
End Function


Public Function SetChanged(Optional bSetConfirmUndo As Boolean = True) As String
  If bSetConfirmUndo Then Call MDIMain.SetConfirmUndo
  SetChanged = S_CHANGEDINDICATOR
End Function

Public Sub SetBenefitFormState(ibf As IBenefitForm2, Optional bGetInvalidFields As Boolean = True)
  On Error GoTo SetBenefitFormState_ERR
  
  Call xSet("SetBenefitFormState")
  
  Call UpdateInfoStatusBar(ibf.benefit)
      
  If Not (ibf.benefit Is Nothing) Then
    If bGetInvalidFields Then ibf.benefit.InvalidFields = InvalidFields(ibf)
    If ibf.benefit.CompanyDefined Then
      Call ibf.BenefitFormState(FORM_CDB)
    ElseIf ibf.benefit.LinkBen Then
      Call ibf.BenefitFormState(FORM_LINK_BEN)
    Else
      Call ibf.BenefitFormState(FORM_ENABLED)
    End If
  Else
    Call ibf.BenefitFormState(FORM_DISABLED)
  End If
  
SetBenefitFormState_END:
  Call xSet("SetBenefitFormState")
  Exit Sub
SetBenefitFormState_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "SetBenefitFormState", "Set Benefit Form State", "Error setting benefit form state.")
  Resume SetBenefitFormState_END
  Resume
End Sub
Public Function ScreenToDateVal(ByVal sText, ByVal STDV As STDV_DEFAULTS) As Variant
  On Error GoTo ScreenToDateVal_ERR
  
  Select Case STDV
    Case STDV_STRING
      ScreenToDateVal = TryConvertDate(sText, sText)
    Case STDV_UNDATED
      ScreenToDateVal = TryConvertDate(sText)
    Case Else
      ECASE ("Invalid Screen to date val default in ScreenToDateVal")
  End Select
  
ScreenToDateVal_END:
  Exit Function
ScreenToDateVal_ERR:
  ScreenToDateVal = sText
  Resume ScreenToDateVal_END
End Function

Public Function CheckTextInput(sText As String, benefit As IBenefitClass, ItemIndex As Long) As Boolean
  Dim bDirty As Boolean
  
  On Error GoTo CheckTextInput_ERR
  Call xSet("CheckTextInput")
  
  If p11d32.BenDataLinkDataType(benefit.BenefitClass, ItemIndex) = TYPE_DATE Then
    bDirty = StrComp(DateValReadToScreen(sText), benefit.value(ItemIndex), vbBinaryCompare)
    benefit.value(ItemIndex) = ScreenToDateVal(sText, STDV_STRING)
  Else
    bDirty = StrComp(sText, benefit.value(ItemIndex), vbBinaryCompare)
    benefit.value(ItemIndex) = sText
  End If
          
  CheckTextInput = bDirty
  
CheckTextInput_END:
  Call xReturn("CheckTextInput")
  Exit Function
CheckTextInput_ERR:
  CheckTextInput = False
  Call ErrorMessage(ERR_ERROR, Err, "CheckTextInput", "Check User Input", "Error checking the user input into the benefit array.")
  Resume CheckTextInput_END
  Resume
End Function


Public Function CheckCheckBoxInput(cbc As CheckBoxConstants, benefit As IBenefitClass, ItemIndex As Long) As Boolean
  Dim bDirty As Boolean
  
  On Error GoTo CheckCheckBoxInput_ERR
  Call xSet("CheckCheckBoxInput")
  
  bDirty = (IIf(cbc = vbChecked, True, False) <> benefit.value(ItemIndex))
  If bDirty Then
    benefit.value(ItemIndex) = IIf(cbc = vbChecked, True, False)
  End If
          
  CheckCheckBoxInput = bDirty
  
CheckCheckBoxInput_END:
  Call xReturn("CheckCheckBoxInput")
  Exit Function
CheckCheckBoxInput_ERR:
  CheckCheckBoxInput = False
  Call ErrorMessage(ERR_ERROR, Err, "CheckCheckBoxInput", "Check User Input", "Error checking the user input into the benefit array.")
  Resume CheckCheckBoxInput_END
End Function
Public Function TestChangedControls(frm As Form) As Boolean
  Dim bChanged As Boolean
  Dim ifrm As IFrmGeneral
  Dim c As Control
  Dim ibf As IBenefitForm2
  
  On Error GoTo TestChangedControls_Err
  Call xSet("TestChangedControls")

  Set ifrm = frm
  Set ibf = frm
  For Each c In frm.Controls
    If TestChanged(c.Tag) Then
      bChanged = bChanged Or ifrm.CheckChanged(c)
    End If
    c.Tag = ""
  Next c

  TestChangedControls = bChanged
  
TestChangedControls_End:
  Call xReturn("TestChangedControls")
  Exit Function
TestChangedControls_Err:
  TestChangedControls = False
  Call ErrorMessage(ERR_ERROR, Err, "TestChangedControls", "Test Changed Controls", "Error searching the tags of a control collection for signs of user input.")
  Resume TestChangedControls_End
  Resume
End Function

Public Function CheckValidity(ifg As IFrmGeneral, Optional ctlContainer As Object = Nothing, Optional bHideform As Boolean = True) As Boolean
  Dim frm As Form
  
  On Error GoTo CheckValidity_Err
  Call xSet("CheckValidity")

  If InvalidFields(ifg, ctlContainer) > 0 Then
    Call SetPanel2("There are invalid fields on this dialogue")
    If Not ifg.InvalidVT Is Nothing Then
      ifg.InvalidVT.SetFocus
      CheckValidity = False
      Beep
    End If
  Else
    Call SetPanel2("")
    CheckValidity = True
    If bHideform Then
      Set frm = ifg
      frm.Hide
    End If
  End If

CheckValidity_End:
  Call xReturn("CheckValidity")
  Exit Function

CheckValidity_Err:
  CheckValidity = False
  Call ErrorMessage(ERR_ERROR, Err, "CheckValidity", "Check Validity", "Error checking the validity of data input or setting the focus to the first invalid valtext.")
  Resume CheckValidity_End
  Resume
End Function

Public Function CheckValidityAndBenefitDirty(benefit As IBenefitClass, ifg As IFrmGeneral, Optional ctlContainer As Object = Nothing) As Boolean
  Dim frm As Form
  
  On Error GoTo CheckValidityAndBenefitDirty_Err
  Call xSet("CheckValidityAndBenefitDirty")

  If benefit.Dirty Then
    CheckValidityAndBenefitDirty = CheckValidity(ifg, ctlContainer)
  Else
    CheckValidityAndBenefitDirty = True
    Set frm = ifg
    frm.Hide
  End If
    
CheckValidityAndBenefitDirty_End:
  Call xReturn("CheckValidityAndBenefitDirty")
  Exit Function

CheckValidityAndBenefitDirty_Err:
  CheckValidityAndBenefitDirty = False
  Call ErrorMessage(ERR_ERROR, Err, "CheckValidityAndBenefitDirty", "Check Validity And Benefit Dirty", "Error checking the validity of data input and whether the benefit was dirty.")
  Resume CheckValidityAndBenefitDirty_End
End Function

Public Function AddUBGRDStandardColumn(UBGRD As Object, lIndex As Long, sngWidth As Single, sCaption As String, sNumberFormat As String) As TrueDBGrid60.column
  Dim c As TrueDBGrid60.column
  
  On Error GoTo AddUBGRDStandardColumn_Err
  Call xSet("AddUBGRDStandardColumn")

  Set c = UBGRD.Columns.Add(lIndex)
    
  With c
    .NumberFormat = sNumberFormat
    .Locked = False
    .Visible = True
    .width = sngWidth
    .Caption = sCaption
    .AllowSizing = True
    .AllowFocus = True
  End With
  
  Set AddUBGRDStandardColumn = c
  
AddUBGRDStandardColumn_End:
  Set c = Nothing
  Call xReturn("AddUBGRDStandardColumn")
  Exit Function
AddUBGRDStandardColumn_Err:
  AddUBGRDStandardColumn = Nothing
  Call ErrorMessage(ERR_ERROR, Err, "AddUBGRDStandardColumn", "Add UBGRD Standard Column", "Error adding a standard formated column to a ubgrd.")
  
  Resume AddUBGRDStandardColumn_End
  Resume
End Function


Public Function InitMilesGrid(grd As Object) As Boolean
  On Error GoTo InitMilesGrid_Err
  
  Call xSet("InitMilesGrid")
  
  Call AddUBGRDStandardColumn(grd, 0, 1244.976, "Number of miles", "")
  Call AddUBGRDStandardColumn(grd, 1, 1244.976, "Date", "")
  Call AddUBGRDStandardColumn(grd, 2, 2720, "Description", "")
  
  grd.AllowUpdate = True
  grd.ReBind
  
  InitMilesGrid = True
  
InitMilesGrid_End:
  Call xReturn("InitMilesGrid")
  Exit Function

InitMilesGrid_Err:
  Call ErrorMessage(ERR_ERROR, Err, "InitMilesGrid", "Init Miles Grid", "Error initialising the miles grid.")
  Resume InitMilesGrid_End
End Function

Public Function MilesDelete(lblTotalMiles As Label, benefit As IBenefitClass, lTotalMilesEnumKey As Long, ObjectList As ObjectList, ObjectListIndex As Long) As Boolean
  Dim Miles As MileDetail
  
On Error GoTo MilesDelete_Err
  
  Call xSet("MilesDelete")

  Set Miles = ObjectList(ObjectListIndex)
  benefit.value(lTotalMilesEnumKey) = benefit.value(lTotalMilesEnumKey) - Miles.MileAmount
  Call ObjectList.Remove(ObjectListIndex)
  
  lblTotalMiles = benefit.value(lTotalMilesEnumKey)
  
  benefit.Dirty = benefit.Dirty Or True
  
MilesDelete_End:
  Set Miles = Nothing
  Call xReturn("MilesDelete")
  Exit Function
MilesDelete_Err:
  Call ErrorMessage(ERR_ERROR, Err, "MilesDelete", "Miles Delete", "Error deleting a mileage item from the miles object list.")
  Resume MilesDelete_End
End Function

Public Function MilesRead(RowBuf As TrueDBGrid60.RowBuffer, RowBufRowIndex As Long, ObjectList As ObjectList, ObjectListIndex As Long) As Boolean
  Dim l As Long
  Dim Miles As MileDetail

  On Error GoTo MilesRead_Err

  Call xSet("MilesRead")

  Set Miles = ObjectList(ObjectListIndex)

  For l = 0 To (RowBuf.ColumnCount - 1)
    Select Case l
      Case 0
        RowBuf.value(RowBufRowIndex, l) = Miles.MileAmount
      Case 1
        RowBuf.value(RowBufRowIndex, l) = IIf(Miles.MileDate = UNDATED, "", DateValReadToScreen(Miles.MileDate))
      Case 2
        RowBuf.value(RowBufRowIndex, l) = Miles.MileItem
      Case Else
        ECASE ("Invalid column ubgrd read data.")
    End Select
  Next

  MilesRead = True

MilesRead_End:
  Set Miles = Nothing
  Call xReturn("MilesRead")
  Exit Function
MilesRead_Err:
  Call ErrorMessage(ERR_ERROR, Err, "MilesRead", "Miles Read", "Error reading the miles object list to the mileage ubgrd.")
  Resume MilesRead_End
End Function

Public Function MilesWrite(lblTotalMiles As Label, lTotalMilesEnumKey As Long, RowBuf As TrueDBGrid60.RowBuffer, RowBufRowIndex As Long, ObjectList As ObjectList, ObjectListIndex As Long, benefit As IBenefitClass) As Boolean
  Dim Miles As MileDetail
  
  On Error GoTo MilesWrite_Err
  
  Call xSet("MilesWrite")

  If ObjectListIndex = -1 Then
    Set Miles = New MileDetail
    ObjectListIndex = ObjectList.Add(Miles)
  Else
    Set Miles = ObjectList(ObjectListIndex)
    benefit.value(lTotalMilesEnumKey) = benefit.value(lTotalMilesEnumKey) - Miles.MileAmount
  End If
       
  With Miles
    If Not IsNull(RowBuf.value(RowBufRowIndex, 0)) Then .MileAmount = RowBuf.value(RowBufRowIndex, 0)
    If Not IsNull(RowBuf.value(RowBufRowIndex, 1)) Then .MileDate = TryConvertDate(RowBuf.value(RowBufRowIndex, 1))
    If Not IsNull(RowBuf.value(RowBufRowIndex, 2)) Then .MileItem = RowBuf.value(RowBufRowIndex, 2)
  End With

  benefit.value(lTotalMilesEnumKey) = benefit.value(lTotalMilesEnumKey) + Miles.MileAmount
  
  lblTotalMiles = benefit.value(lTotalMilesEnumKey)
  
  benefit.Dirty = benefit.Dirty Or True
  
  MilesWrite = True

MilesWrite_End:
  Set Miles = Nothing
  Call xReturn("MilesWrite")
  Exit Function

MilesWrite_Err:
  Call ErrorMessage(ERR_ERROR, Err, "MilesWrite", "Miles Write", "Error writing a mileage item to the mileage object list.")
  Resume MilesWrite_End
  Resume
End Function

Public Function MilesValidate(FirstColIndexInError As Long, ValidateMessage As String, RowBuf As TrueDBGrid60.RowBuffer, RowBufRowIndex As Long, ByVal ObectListIndex As Long) As Boolean
  Dim l As Long
  
  On Error GoTo MilesValidate_Err
  
  Call xSet("MilesValidate")
  
  With RowBuf
    For l = 0 To RowBuf.ColumnCount - 1
      Select Case l
        Case 0
          'no of miles
          If GridIsNotNumericOrLong(ValidateMessage, RowBuf.value(RowBufRowIndex, l), ObectListIndex) Then
            FirstColIndexInError = l
            GoTo MilesValidate_End
          End If
        Case 1
          'date
          If Not IsNull(RowBuf.value(RowBufRowIndex, l)) Then
            If Len(RowBuf.value(RowBufRowIndex, l)) > 0 Then
              If GridIsNotDate(ValidateMessage, RowBuf.value(RowBufRowIndex, l), ObectListIndex, False) Then
                FirstColIndexInError = l
                GoTo MilesValidate_End
              End If
            End If
          End If
        Case 2
          'description
          If GrisIsTooLong(ValidateMessage, RowBuf, RowBufRowIndex, l) Then
            FirstColIndexInError = l
            GoTo MilesValidate_End
          End If
      End Select
    Next
  End With

MilesValidate_End:
  Call xReturn("MilesValidate")
  Exit Function

MilesValidate_Err:
  Call ErrorMessage(ERR_ERROR, Err, "MilesValidate", "Miles Validate", "Error validating a mileage item.")
  Resume MilesValidate_End
End Function

Public Function BenefitToListView(ben As IBenefitClass, ibf As IBenefitForm2, BenefitIndex As Long) As Long
  Dim lst As ListItem
  
  
  
  On Error GoTo BenefitToListView_Err
  Call xSet("BenefitToListView")
  
  If Not ben Is Nothing Then
    If ibf.ValididateBenefit(ben) Then
      Set lst = ibf.lv.listitems.Add(, , ben.Name)
      DoEvents
      BenefitToListView = ibf.UpdateBenefitListViewItem(lst, ben, BenefitIndex)
      
    End If
  End If

BenefitToListView_End:
  Set lst = Nothing
  Call xReturn("BenefitToListView")
  Exit Function

BenefitToListView_Err:
  Call ErrorMessage(ERR_ERROR, Err, "BenefitToListView", "Benefit To List View", "Error placing a benefit to a list view.")
  Resume BenefitToListView_End
  Resume
End Function

Public Function AddAddDelete(tbar As ToolBar) As Boolean
  Dim b As Button
  
  On Error GoTo AddAddDelete_Err
  Call xSet("AddAddDelete")

  Set tbar.ImageList = MDIMain.ImgToolbar
  
  Set b = tbar.Buttons.Add(1, , , , 14)
  b.ToolTipText = "Add"
  Set b = tbar.Buttons.Add(2, , , , 13)
  b.ToolTipText = "Delete"

  AddAddDelete = True
  
AddAddDelete_End:
  Set b = Nothing
  Call xReturn("AddAddDelete")
  Exit Function

AddAddDelete_Err:
  Call ErrorMessage(ERR_ERROR, Err, "AddAddDelete", "Add Add Delete", "Error adding the add/delete buttons to a toolbar.")
  Resume AddAddDelete_End
  Resume
End Function

Public Function AddDeleteClick(ByVal lButtonIndex As Long, ibf As IBenefitForm2) As Boolean

  On Error GoTo AddDeleteClick_Err
  Call xSet("AddDeleteClick")

  With ibf
    Select Case lButtonIndex
      Case 1 'tick
          Call .AddBenefit
      Case 2 'cross
        If Not .lv.SelectedItem Is Nothing Then Call .RemoveBenefit(.lv.SelectedItem.Tag)
    End Select
  End With

AddDeleteClick_End:
  Call xReturn("AddDeleteClick")
  Exit Function

AddDeleteClick_Err:
  Call ErrorMessage(ERR_ERROR, Err, "AddDeleteClick", "Add Delete Click", "Error interpreting a click for add/delete buttons.")
  Resume AddDeleteClick_End
End Function

Public Function SelectBenefitByListItem(ibf As IBenefitForm2, li As ListItem) As Boolean
  Dim Selected As Boolean
  
  On Error GoTo SelectBenefitByListItem_Err
  Call xSet("SelectBenefitByListItem")

  If Not li Is Nothing Then
    Set ibf.lv.SelectedItem = li
    Call ibf.BenefitToScreen(li.Tag)
    
    Call BenefitFormSelectDefaultControl(ibf)
    Call SetLastListItemSelected(li)
    DoEvents 'added to enble ensure visible to work, take it out and see
    Call ibf.lv.SelectedItem.EnsureVisible
    Selected = True
  ElseIf ibf.lv.listitems.Count > 0 Then
    Set ibf.lv.SelectedItem = ibf.lv.listitems(1)
    Call ibf.BenefitToScreen(ibf.lv.SelectedItem.Tag)
    
    Call BenefitFormSelectDefaultControl(ibf)
    Call SetLastListItemSelected(ibf.lv.SelectedItem)
    DoEvents 'added to enble ensure visible to work, take it out and see
    Call ibf.lv.SelectedItem.EnsureVisible
    Selected = True
  Else
    Call ibf.BenefitToScreen(-1)
  End If
  
SelectBenefitByListItem_End:
  Call xReturn("SelectBenefitByListItem")
  Exit Function

SelectBenefitByListItem_Err:
  Call ErrorMessage(ERR_ERROR, Err, "SelectBenefitByListItem", "Select Benefit By ListItem", "Error seleting a benefit via the listitem.")
  Resume SelectBenefitByListItem_End
  Resume
End Function

Public Function SelectBenefitByBenefitIndex(ibf As IBenefitForm2, Optional ByVal lBenefitIndex As Long = -1) As Boolean
  Dim Selected As Boolean
  Dim lListCount As Long
  Dim lListItemIndex As Long
  Dim li As ListItem
  
  On Error GoTo SelectBenefitByBenefitIndex_ERR

  Call xSet("SelectBenefitByBenefitIndex")
  lListCount = ibf.lv.listitems.Count
  If lListCount > 0 Then
    If (lBenefitIndex < 1) Then
      lListItemIndex = GetListItemIndexbyText(ibf.lv) ' returns 1 if no matches found
      If lListItemIndex <> -1 Then
        Set li = ibf.lv.listitems(lListItemIndex)
        lBenefitIndex = li.Tag
      Else
        lBenefitIndex = -1
      End If
    Else
      For Each li In ibf.lv.listitems
        If li.Tag = lBenefitIndex Then
          Exit For
        End If
      Next
    End If
    
    Set ibf.lv.SelectedItem = li
    Call ibf.BenefitToScreen(lBenefitIndex)
    ibf.lv.SelectedItem.EnsureVisible
    Call BenefitFormSelectDefaultControl(ibf)
    Selected = True
    Call SetLastListItemSelected(ibf.lv.SelectedItem)
  Else
    Call ibf.BenefitToScreen(-1)
  End If
  If Not Selected Then Call ibf.BenefitFormState(FORM_DISABLED)
  
SelectBenefitByBenefitIndex_END:
  SelectBenefitByBenefitIndex = Selected
  Call xSet("SelectBenefitByBenefitIndex")
  Exit Function
  
SelectBenefitByBenefitIndex_ERR:
  SelectBenefitByBenefitIndex = False
  Call ErrorMessage(ERR_ERROR, Err, "SelectBenefitByBenefitIndex", "Select Benefit", "Unable to select benefit " & CStr(lBenefitIndex))
  Resume SelectBenefitByBenefitIndex_END
  Resume
End Function
Public Function LastListItemDifferent() As Boolean
  On Error GoTo LastListItemDifferent_Err
  Call xSet("LastListItemDifferent")

  If Not mLast Is mCurrent Then LastListItemDifferent = True

LastListItemDifferent_End:
  Call xReturn("LastListItemDifferent")
  Exit Function

LastListItemDifferent_Err:
  Call ErrorMessage(ERR_ERROR, Err, "LastListItemDifferent", "Last List Item Different", "Error determining whether the last list item selected is different from the current listitem.")
  Resume LastListItemDifferent_End
End Function
Private Sub SetControlBooleanProperty(ByVal CP As CONTROL_PROPERTY, ByVal bProperyValue As Boolean, v As Variant)
  Dim c As Control
  On Error GoTo SetControlBooleanProperty_ERR
    
  Set c = v
  
  If c Is Nothing Then Call Err.Raise(ERR_IS_NOTHING, "SetControlBooleanProperty", "The control is nothing.")
    
  Select Case CP
    Case CP_ENABLED
      c.Enabled = bProperyValue
    Case CP_VISIBLE
      c.Visible = bProperyValue
    Case Else
      Call ECASE("Invalid property value.")
  End Select
  
SetControlBooleanProperty_END:
  Exit Sub
SetControlBooleanProperty_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "SetControlBooleanProperty", "Set Control Boolean Property", "Error setting a controls property.")
  Resume SetControlBooleanProperty_END
End Sub
Private Function IsVaraintArrayWithElements(v As Variant)
  Dim v1 As Variant
  
  On Error GoTo IsVaraintArrayWithElements_ERR
  
  If Not IsArray(v) Then GoTo IsVaraintArrayWithElements_END
  If (UBound(v) < 0) Then
    IsVaraintArrayWithElements = False
  Else
    v1 = v(LBound(v))
    IsVaraintArrayWithElements = True
  
  End If
  
  
IsVaraintArrayWithElements_END:
  Exit Function
IsVaraintArrayWithElements_ERR:
  Resume IsVaraintArrayWithElements_END
End Function
Private Sub SetControlsBooleanProperty(ByVal CP As CONTROL_PROPERTY, ByVal bPropertyValue As Boolean, vControls() As Variant)
  Dim i As Long
  
  'need this rubbish as tab controls are not parented to a frame thus need to be disabled specifically
  'can't use form as cdbs need to be allowed to copy but input disabled
  
  On Error GoTo SetControlsBooleanProperty_ERR
    
  If Not IsVaraintArrayWithElements(vControls) Then GoTo SetControlsBooleanProperty_END
  
  For i = LBound(vControls) To UBound(vControls)
    Call SetControlBooleanProperty(CP, bPropertyValue, vControls(i))
  Next
    
SetControlsBooleanProperty_END:
  Exit Sub
SetControlsBooleanProperty_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "SetControlsBooleanProperty", "Set Controls Boolean Property", "Error setting controls property.")
  Resume SetControlsBooleanProperty_END
End Sub
Public Function BenefitFormStateEx(ByVal fState As BENEFIT_FORM_STATE, benefit As IBenefitClass, ParamArray Controls()) As Boolean
  Dim vControls()
  On Error GoTo BenefitFormStateEx_err
  
  Call xSet("BenefitFormStateEx")
  
  vControls = Controls
  Call SetControlsBooleanProperty(CP_VISIBLE, True, vControls)
  If (fState = FORM_ENABLED) Or (fState = FORM_CDB) Then
    
    If fState = FORM_ENABLED Then
      Call SetControlsBooleanProperty(CP_ENABLED, True, vControls)
    Else
      Call SetControlsBooleanProperty(CP_ENABLED, False, vControls)
    End If
    Call MDIMain.SetDelete
  ElseIf fState = FORM_DISABLED Then
    Set benefit = Nothing
    Call SetControlsBooleanProperty(CP_ENABLED, False, vControls)
    Call MDIMain.ClearDelete
    Call MDIMain.ClearConfirmUndo
  End If
  
  BenefitFormStateEx = True
    
BenefitFormStateEx_end:
  Call xReturn("BenefitFormStateEx")
  Exit Function
  
BenefitFormStateEx_err:
  BenefitFormStateEx = False
  Call ErrorMessage(ERR_ERROR, Err, "BenefitFormStateEx", "Benefit Form State Ex", "Error setting the benefit form state.")
  Resume BenefitFormStateEx_end
  Resume

End Function
Public Function SelectLastListItem(lv As ListView) As Boolean

  On Error GoTo SelectLastListItem_Err
  Call xSet("SelectLastListItem")
  
  If Not mLast Is Nothing Then
    Set lv.SelectedItem = mLast
  End If

SelectLastListItem_End:
  Call xReturn("SelectLastListItem")
  Exit Function

SelectLastListItem_Err:
  Call ErrorMessage(ERR_ERROR, Err, "SelectLastListItem", "Select Last List Item", "Error selecting the last listitem selected.")
  Resume SelectLastListItem_End
End Function

Public Sub SetSortOrderEmployees(LB As ListView, ByVal ColumnHeader As MSComctlLib.ColumnHeader, Optional OverrideSortOrder As ListSortOrderConstants = -1)
  
  On Error GoTo SetSortOrderEmployees_ERR
  
  
  Call SetSortOrder(LB, ColumnHeader, OverrideSortOrder)
  p11d32.EmployeeSortOrder = LB.SortOrder
  p11d32.EmployeeSortOrderColumn = ColumnHeader.Index - 1
  
SetSortOrderEmployees_END:
  Exit Sub
SetSortOrderEmployees_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "SetSortOrderEmployees", "Set Sort Order Employees", "Error setting the sort order for the employees.")
  Resume SetSortOrderEmployees_END
  Resume
End Sub
