Attribute VB_Name = "ScreenCTRL"
Option Explicit

Public Enum SELECT_MODE
  SELECT_ALPHABETICALLY_START = 1
  SELECT_ALPHABETICALLY_A_C = 1 'keep this as 1
  SELECT_ALPHABETICALLY_D_F
  SELECT_ALPHABETICALLY_G_I
  SELECT_ALPHABETICALLY_J_L
  SELECT_ALPHABETICALLY_M_O
  SELECT_ALPHABETICALLY_P_R
  SELECT_ALPHABETICALLY_S_U
  SELECT_ALPHABETICALLY_V_X
  SELECT_ALPHABETICALLY_Y_Z
  SELECT_ALPHABETICALLY_END = SELECT_ALPHABETICALLY_Y_Z
  SELECT_ALL
  SELECT_NONE
  SELECT_REVERSE
  SELECT_GROUP_1
  SELECT_GROUP_2
  SELECT_GROUP_3
  SELECT_CURRENT_EMPLOYED
  SELECT_LEFT
  SELECT_NO_EMAIL
  SELECT_EMAIL
  SELECT_BY_REPORT
  
End Enum

Public Enum LIST_EMPLOYER_FUNCTIONS
  LEF_SIZE_COLUMNS
  LEF_LISTVIEW_COLUMN
End Enum

Public Enum TVM_GET_SETIMAGELIST_wParam
  TVSIL_NORMAL = 0
  TVSIL_STATE = 2
End Enum
Public Const TV_FIRST = &H1100
Public Const TVM_SETIMAGELIST = (TV_FIRST + 9)

Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
                            (ByVal hwnd As Long, _
                            ByVal wMsg As Long, _
                            wParam As Any, _
                            lParam As Any) As Long


Public Function TreeView_SetImageList(hwnd As Long, himl As Long, iImage As Long) As Long
  TreeView_SetImageList = SendMessage(hwnd, TVM_SETIMAGELIST, ByVal iImage, ByVal himl)
End Function
Public Function GetRGB(ByVal RGBval As Long, ByRef r As Integer, ByRef G As Integer, ByRef b As Integer) As Boolean
  r = RGBval \ 256 ^ (0) And 255
  G = RGBval \ 256 ^ (1) And 255
  b = RGBval \ 256 ^ (2) And 255
  GetRGB = True
End Function
Public Function HTMLColor(l As Long) As String
  Dim r As Integer
  Dim G As Integer
  Dim b As Integer
  
  Call GetRGB(l, r, G, b)
  HTMLColor = HTMLColorEx(r, G, b)
  
End Function
Public Function HTMLColorEx(ByRef r As Integer, ByRef G As Integer, ByRef b As Integer) As String
  HTMLColorEx = "#" & Hex(r) & Hex(G) & Hex(b)
End Function

Public Sub BenefitInErrorRow(ByVal ben As IBenefitClass, ByVal li As ListItem)
  Dim c As ColorConstants
  Dim i As Long
  Dim lisub As ListSubItem
  
  If (StrComp(ben.value(ITEM_BENEFIT), S_ERROR) = 0) Then
    c = vbRed
  Else
    c = vbBlack
  End If
  li.ForeColor = c
  
  For i = 1 To li.ListSubItems.Count
    Set lisub = li.ListSubItems(i)
    lisub.ForeColor = c
  Next
End Sub
Public Sub LabelToolTip(ByVal lbl As Label, ByVal sCaption As String)
  lbl.Caption = sCaption
  lbl.ToolTipText = sCaption
  
End Sub
Public Sub AddComboItemAndItemData(cbo As ComboBox, s As String, iItemData As Long)
    Call cbo.AddItem(s)
    cbo.ItemData(cbo.ListCount - 1) = iItemData
End Sub
Public Sub ComboBoxItemDataToScreen(cbo As ComboBox, iItemData As Long)
  Dim i As Long
  For i = 0 To cbo.ListCount - 1
    If (cbo.ItemData(i) = iItemData) Then
      cbo.ListIndex = i
      Exit For
    End If
    
  Next
    
End Sub

Public Function ListViewFastKey(lv As ListView, ByVal LVI As LV_EE_ITEMS, ByVal KeyCode As Integer, ByVal KeyAscii As Integer, sToSearch As String) As Long

  Dim li As ListItem
  Dim i As Long, lStart As Long, lRet As Long, lEnd As Long
  Dim s As String


  On Error GoTo ListViewFastKey_ERR

  If KeyAscii = vbKeySpace Then GoTo ListViewFastKey_END
  If KeyAscii = vbKeyEscape Then GoTo ListViewFastKey_END

  ListViewFastKey = KeyCode

  If lv.listitems.Count = 0 Then GoTo ListViewFastKey_END

  s = lv.ColumnHeaders(LVI + 1).Text
  ListViewFastKey = 0

  If (KeyCode <> vbKeyF3) Then
    F_Input.ValText.TypeOfData = VT_STRING
    If F_Input.Start("Search " & s, "Enter string to search for...", Chr$(KeyAscii), False) = False Then GoTo ListViewFastKey_END
    sToSearch = F_Input.ValText.Text
  End If

  lStart = 1
  lEnd = lv.listitems.Count
  If Not lv.SelectedItem Is Nothing Then
    If lv.SelectedItem.Index <> lEnd Then
      lStart = lv.SelectedItem.Index + 1
    End If
  End If

  lRet = LVFoundMatch(lStart, lEnd, sToSearch, LVI, lv)
  If lRet > 0 Then ListViewFastKey = lRet: GoTo ListViewFastKey_END
  If lStart > 1 Then
    lEnd = lStart - 1
    lStart = 1
    lRet = LVFoundMatch(lStart, lEnd, sToSearch, LVI, lv)
    If lRet > 0 Then ListViewFastKey = lRet: GoTo ListViewFastKey_END
  End If


ListViewFastKey_END:
  Set F_Input = Nothing
  Exit Function
ListViewFastKey_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "ListViewFastKey", "List View Fast Key", "Error performing list view fast key.")
  Resume ListViewFastKey_END
  Resume
End Function
Private Function LVFoundMatch(lStart As Long, lEnd As Long, sToMatch As String, LVI As LV_EE_ITEMS, lv As ListView) As Long
  Dim i As Long
  Dim li As ListItem
  Dim listitems As listitems
  Dim s1 As String
  
  On Error GoTo LVFoundMatch_ERR
  
  Set listitems = lv.listitems
  For i = lStart To lEnd
    Set li = listitems(i)
    
    If LVI = LV_EE_NAME Then
      s1 = listitems(i)
    Else
      s1 = li.SubItems(LVI)
    End If
    
    If InStr(1, s1, sToMatch, vbTextCompare) = 1 Then
      Set lv.SelectedItem = li
      LVFoundMatch = i
      Call li.EnsureVisible
      GoTo LVFoundMatch_END
    End If
  Next
  
LVFoundMatch_END:
  Exit Function
LVFoundMatch_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "LVFoundMatch", "LV Found Match", "Error findng a match by string in a listitem subitem.")
  Resume LVFoundMatch_END
  Resume
End Function

Public Sub SetNameOrder()
  On Error GoTo SetNameOrder_ERR
  
  Call xSet("SetNameOrder")
  
'  F_NameOrder.Show
  Call p11d32.Help.ShowForm(F_NameOrder, vbModal)
  
SetNameOrder_END:
  Set F_NameOrder = Nothing
  Exit Sub
SetNameOrder_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "SetNameOrder", "Set Name Order", "Error setting the name order.")
End Sub
Public Sub SetLVEnabled(lv As ListView, ByVal bEnabledValue As Boolean)
  On Error GoTo SetLVEnabled_ERR
  
  If lv Is Nothing Then Call Err.Raise(ERR_IS_NOTHING, "SetLVEnabled", "The list view is nothing.")
  If lv.Enabled <> bEnabledValue Then lv.Enabled = bEnabledValue
  
SetLVEnabled_END:
  Exit Sub
SetLVEnabled_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "SetLVEnabled", "Set LV Enabled", "Error setting the list views enabled status.")
  Resume SetLVEnabled_END
End Sub

Public Function RTSelText(rt As RichTextBox, lSelStart As Long, lSelLen As Long, Optional lcolor As Long = vbBlue) As Boolean
  rt.SelStart = lSelStart
  rt.SelLength = lSelLen
  'rt.SelProtected = True 'km - removed Protection
  rt.SelColor = lcolor
  rt.SelStart = rt.SelStart + lSelLen
  
  rt.SelLength = 0
  
  RTSelText = True
End Function
Public Function OnlyFromForm(frm As Form, Optional ByVal sDescriptionOfForm As String = "employer screen")
  If Not (CurrentForm Is frm) Then Call Err.Raise(ERR_CURENTFORM_NOT_EMPLOYERS, "OnlyFromForm", "Can only be run from the " & sDescriptionOfForm & ".")
End Function

Public Function ToolBarButton(Index As Long, param As Long) As Boolean
  Static bInFunc As Boolean
  Dim benfrm2 As IBenefitForm2
  Dim ben As IBenefitClass
  Dim ee As Employee
  Dim rep As Reporter
  
  On Error GoTo ToolBarButton_Err
  Call xSet("ToolBarButton")
  
  If bInFunc Then GoTo ToolBarButton_End
  bInFunc = True
  Select Case Index
    Case TBR_OPEN_EMPLOYER
      Call p11d32.LoadEmployer(p11d32.Employers(param))
    Case TBR_EDIT_EMPLOYER
      Call p11d32.EditEmployer(param)
    Case TBR_SEPERATOR1
    Case TBR_REFRESH_EMPLOYERS
      Call p11d32.LoadEmployers
    Case TBR_CONFIRM
      Call UpdateBenefitFromTags
      Call p11d32.CurrentEmployer.SaveCurrentEmployee
    Case TBR_UNDO
      Call p11d32.CurrentEmployer.ReloadCurrentEmployee
    Case TBR_SEPERATOR2
    Case TBR_ADD_BENEFIT
      If Not CurrentForm Is Nothing Then
        If IsBenefitFormReal(CurrentForm) Then MDIMain.SetConfirmUndo
        Set benfrm2 = CurrentForm
        CurrentForm.Enabled = True
        benfrm2.AddBenefit
      End If
    Case TBR_REMOVE_BENEFIT
      If IsBenefitForm(CurrentForm) Then
        Set benfrm2 = CurrentForm
        If benfrm2.benclass = BC_EMPLOYEE Then
          If MultiDialog("Delete", "Are you sure you want to delete " & p11d32.CurrentEmployer.CurrentEmployee.FullName & " from the employer file?", "&Yes", "&No") = 1 Then
            benfrm2.RemoveBenefit (param)
          End If
        Else
          If MultiDialog("Delete", "Are you sure you want to delete the selected item(s)", "&Yes", "&No") = 1 Then benfrm2.RemoveBenefit (param)
        End If
      End If
    Case TBR_SEPERATOR3
    Case TBR_PRINT
      Call p11d32.ReportPrint.InitPrintDialog
    Case TBR_PREVIEW
      Call p11d32.ReportPrint.ToolBarPreview(CurrentForm)
    Case TBR_EMPLOYERSCREEN
      Call p11d32.EmployerScreen
    Case TBR_SEPERATOR4
    Case TBR_SHAREDVANS
      Call p11d32.CurrentEmployer.EditSharedVans
    Case TBR_EMPLOYEESCREEN
      If p11d32.CurrentEmployer Is Nothing Then
          If Not F_Employers.LB.SelectedItem Is Nothing Then
            Call p11d32.LoadEmployer(p11d32.Employers(F_Employers.LB.SelectedItem.Tag))
          End If
      Else
        Call p11d32.CurrentEmployer.EmployeeScreen
      End If
  End Select
  
  
  bInFunc = False
  
ToolBarButton_End:
  Call xReturn("ToolBarButton")
  Exit Function
ToolBarButton_Err:
  bInFunc = False
  Call ErrorMessage(ERR_ERROR, Err, "ToolBarButton", "Toolbar Action", "Unable to complete toolbar action.")
  Resume ToolBarButton_End
  Resume
End Function
Public Function BenScreenSwitch(benclass As BEN_CLASS, Optional NewEE As Employee = Nothing) As Boolean
  Dim frm As Form
  Dim benfrm2 As IBenefitForm2
  Dim Caption As String
  
  On Error GoTo BenScreenSwitch_Err
  
  Call xSet("BenScreenSwitch")
  If NewEE Is Nothing Then Set NewEE = p11d32.CurrentEmployer.employees(GetEmployeeIndexFromSelectedEmployee)
    
  If Not (NewEE Is Nothing) Then
    Call NewEE.LoadBenefits(TBL_ALLBENEFITS, True)
      Select Case benclass
        Case BC_ALL
          Set benfrm2 = F_AllBenefits
          MDIMain.chkMoveToNextEmployeeWithBenefit.Visible = False
        Case BC_COMPANY_CARS_F, BC_FUEL_F
          Set benfrm2 = F_CompanyCar
          MDIMain.chkMoveToNextEmployeeWithBenefit.Visible = True
        Case BC_PHONE_HOME_N
          Set benfrm2 = F_Phone
          MDIMain.chkMoveToNextEmployeeWithBenefit.Visible = True
        Case BC_PAYMENTS_ON_BEFALF_B, BC_VOUCHERS_AND_CREDITCARDS_C, BC_CLASS_1A_M, BC_PRIVATE_MEDICAL_I, BC_GENERAL_EXPENSES_BUSINESS_N, BC_NON_CLASS_1A_M, BC_INCOME_TAX_PAID_NOT_DEDUCTED_M, BC_OOTHER_N, BC_ENTERTAINMENT_N, BC_TRAVEL_AND_SUBSISTENCE_N, BC_CHAUFFEUR_OTHERO_N, BC_TAX_NOTIONAL_PAYMENTS_B ' , BC_SHARES_M
          Set benfrm2 = F_Other
          MDIMain.chkMoveToNextEmployeeWithBenefit.Visible = True
        Case BC_NONSHAREDVANS_G
          Set benfrm2 = F_NonSharedVans
          MDIMain.chkMoveToNextEmployeeWithBenefit.Visible = True
        Case BC_LOAN_OTHER_H
          Set benfrm2 = F_Loan
          MDIMain.chkMoveToNextEmployeeWithBenefit.Visible = True
        Case BC_SERVICES_PROVIDED_K
          Set benfrm2 = F_ServicesProvided
          Caption = "Services provided to the employee"
          MDIMain.chkMoveToNextEmployeeWithBenefit.Visible = True
        Case BC_ASSETSATDISPOSAL_L
          Set benfrm2 = F_AssetsAtDisposal
          Caption = "Assets placed at the employees disposal"
          MDIMain.chkMoveToNextEmployeeWithBenefit.Visible = True
        Case BC_ASSETSTRANSFERRED_A
          Set benfrm2 = F_AssetsTransferred
          MDIMain.chkMoveToNextEmployeeWithBenefit.Visible = True
        Case BC_LIVING_ACCOMMODATION_D
          Set benfrm2 = F_Accommodation
          MDIMain.chkMoveToNextEmployeeWithBenefit.Visible = True
        Case BC_EMPLOYEE_CAR_E
          Set benfrm2 = F_EmployeeCar
          MDIMain.chkMoveToNextEmployeeWithBenefit.Visible = True
        Case BC_QUALIFYING_RELOCATION_J
          Set benfrm2 = F_Relocation
          MDIMain.chkMoveToNextEmployeeWithBenefit.Visible = True
        Case BC_NON_QUALIFYING_RELOCATION_N
          benclass = BC_QUALIFYING_RELOCATION_J
          Set benfrm2 = F_Relocation
          MDIMain.chkMoveToNextEmployeeWithBenefit.Visible = True
        Case BC_CDB
          Set benfrm2 = F_CompanyDefined
          MDIMain.chkMoveToNextEmployeeWithBenefit.Visible = True
        Case Else
          ECASE "Unknown Benefit Screen benclass = " & benclass
          GoTo BenScreenSwitch_End
      End Select
      
      Call MDIMain.CutCopyPasteVisible(benclass <> BC_ALL)
      
      benfrm2.benclass = benclass
      
      Call UpdateBenefitFromTags
      
      If p11d32.CurrentEmployer.LoadEmployeeEx(NewEE) Then
        Set frm = CurrentForm
        Set CurrentForm = benfrm2
        CurrentForm.Caption = p11d32.Rates.BenClassTo(benfrm2.benclass, BCT_FORM_CAPTION)
        Call MDIMain.NavigateBarUpdate(NewEE)
        Call ShowMaximized(CurrentForm, frm, D_BENEFIT)
        Call BenScreenSwitchEnd(benfrm2)
        DoEvents
     End If
     
  End If
  
BenScreenSwitch_End:
  Set frm = Nothing
  Set benfrm2 = Nothing
  Set NewEE = Nothing
  Call xReturn("BenScreenSwitch")
  Exit Function
BenScreenSwitch_Err:
  Call ErrorMessage(ERR_ERROR, Err, "BenScreenSwitch", "Benefit Screen Switch", "Unable to switch benefit screens")
  Resume BenScreenSwitch_End
  Resume
End Function
Public Function BenefitFormSelectedIndex() As Long
  Dim l As Long
  Dim ibf As IBenefitForm2
  
  
  l = -1
  If (IsBenefitForm(CurrentForm)) Then
    Set ibf = CurrentForm
    If Not ibf.lv.SelectedItem Is Nothing Then
      l = CLng(ibf.lv.SelectedItem.Tag)
    End If
  End If

  BenefitFormSelectedIndex = l
End Function
Public Function GotoScreen() As Boolean
  Dim sNewKey As String
  Dim sCurrentKey As String
  
  On Error GoTo GotoScreen_Err
  Call xSet("GotoScreenScreen")
  If MoveGotoCheck Then GoTo GotoScreen_End
  
  Call CopyListView(F_Goto.LB, F_Employees.LB, True, True)
  
  Call SetSortOrderEmployees(F_Goto.LB, F_Goto.LB.ColumnHeaders(p11d32.EmployeeSortOrderColumn + 1), p11d32.EmployeeSortOrder)
  sCurrentKey = F_Employees.LB.SelectedItem.Key
  F_Goto.LB.SelectedItem.EnsureVisible
'  F_Goto.Show vbModal
  Call p11d32.Help.ShowForm(F_Goto, vbModal)
  If Not F_Goto.NewEmployeeListItem Is Nothing Then
    sNewKey = F_Goto.NewEmployeeListItem.Key
    If StrComp(sNewKey, sCurrentKey) Then
      Call RemotelySelectAnotherEmployee(True, , sNewKey)
    End If
  End If
  Call SetSortOrderEmployees(F_Employees.LB, F_Employees.LB.ColumnHeaders(p11d32.EmployeeSortOrderColumn + 1), p11d32.EmployeeSortOrder)
  
  
GotoScreen_End:
  If Not F_Goto Is Nothing Then Call F_Goto.Hide
  Call xReturn("GotoScreen")
  Exit Function

GotoScreen_Err:
  Call ErrorMessage(ERR_ERROR, Err, "GotoScreen", "Goto Screen", "Error using the goto screen.")
  Resume GotoScreen_End
  Resume
End Function
Public Sub UpdateMDICaption(ey As Employer)
  Dim ben As IBenefitClass
  Dim sCaption As String
  
  On Error GoTo UpdateMDICaption_ERR
  
  Set ben = ey
  
  If Not ben Is Nothing Then sCaption = " " & ben.value(employer_Name_db)
  
  MDIMain.Caption = p11d32.Rates.value(TaxFormYear) & sCaption

UpdateMDICaption_END:
  Exit Sub
UpdateMDICaption_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "UpdateMDICaption", "Update MDI Caption", "Error updating the MDI forms caption.")
  Resume UpdateMDICaption_END
End Sub
Public Function BenScreenSwitchEnd(ibf As IBenefitForm2) As Boolean
  
  
  On Error GoTo BenScreenSwitchEnd_Err
  Call xSet("BenScreenSwitchEnd")
  
  ibf.BenefitsToListView
  Call SelectBenefitByListItem(ibf, Nothing)
  
  Call UpdateMDICaption(p11d32.CurrentEmployer)
  
BenScreenSwitchEnd_End:
  Call xReturn("BenScreenSwitchEnd")
  Exit Function

BenScreenSwitchEnd_Err:
  Call ErrorMessage(ERR_ERROR, Err, "BenScreenSwitchEnd", "Ben Screen Switch End", "Error finishing the benefit screen switch.")
  Resume BenScreenSwitchEnd_End
End Function


Public Function MoveEmployee(bBackWards As Boolean) As Boolean
  Dim ibf As IBenefitForm2
  Dim iEmployeeObjectListIndex As Long
  Dim lNextEmployeeListItemIndex As Long
  Dim lCurrentListItemIndex As Long
  Dim ben As IBenefitClass
  Dim i As Long
  Dim iStep As Long
  Dim ee As Employee
  Dim bc As BEN_CLASS
  Dim bt As BENEFIT_TABLES
  Dim bBenefitsShortLoad As Boolean
  On Error GoTo MoveEmployee_Err
  Call xSet("MoveEmployee")
  Call SetCursor
  
  
  If MoveGotoCheck Then GoTo MoveEmployee_End
  
  iStep = 1
  If (bBackWards) Then iStep = -1
  
Set ibf = F_Employees
  If ibf.lv.listitems.Count > 0 Then
     
    If Not ibf.lv.SelectedItem Is Nothing Then
      lCurrentListItemIndex = ibf.lv.SelectedItem.Index
      If bBackWards Then
        If lCurrentListItemIndex > 1 Then
          lNextEmployeeListItemIndex = lCurrentListItemIndex + iStep
        End If
        'get the employees collection
      Else
        If lCurrentListItemIndex < ibf.lv.listitems.Count Then
          lNextEmployeeListItemIndex = lCurrentListItemIndex + iStep
        End If
        'forwards
      End If
      
      If (lNextEmployeeListItemIndex <> lCurrentListItemIndex) Then
        Set ibf = CurrentForm
        bc = ibf.benclass
        'ibf.benefit
        
        If p11d32.MoveToNextEmployeeWithBenefit And bc <> BC_EMPLOYEE Then
          Set ibf = F_Employees
          If bc <> BC_ALL Then
            bt = p11d32.Rates.BenClassTo(bc, BCT_BENEFIT_TABLE)
            If bc = BC_NONSHAREDVANS_G Then bc = BC_nonSHAREDVAN_G
          Else
            bt = TBL_ALLBENEFITS
          End If

          Do While (lNextEmployeeListItemIndex > 0 And lNextEmployeeListItemIndex < ibf.lv.listitems.Count + 1)
            Set ee = p11d32.CurrentEmployer.employees.Item(ibf.lv.listitems(lNextEmployeeListItemIndex).Tag)
            If Not ee.BenefitsLoaded Then
              'only load from the relevant table
              Call ee.LoadBenefits(bt, False, False)
            End If
            If (ee.HasBenefit(bc, True)) Then
              If (bt <> TBL_ALLBENEFITS) Then Call ee.LoadBenefits(TBL_ALLBENEFITS, True, False)
              Call RemotelySelectAnotherEmployee(True, lNextEmployeeListItemIndex)
              Exit Do
            End If
            lNextEmployeeListItemIndex = lNextEmployeeListItemIndex + iStep
          Loop
        Else
          Call RemotelySelectAnotherEmployee(True, lNextEmployeeListItemIndex)
        End If
      End If
      
    End If
  End If
  
MoveEmployee_End:
  Call ClearCursor
  Call xReturn("MoveEmployee")
  Exit Function

MoveEmployee_Err:
  Call ErrorMessage(ERR_ERROR, Err, "MoveEmployee", "Move Employee", "Error moving backward or forward through the employee list.")
  Resume MoveEmployee_End
  Resume
End Function

Public Function RemotelySelectAnotherEmployee(bLoadBenefits As Boolean, Optional lNewEmployeeListItemIndex As Long = 0, Optional sNewEmployeeListItemKey As String = "") As Boolean
  Dim ibf As IBenefitForm2
  Dim lNewEmployeeIndex As Long, lListIndex As Long
  Dim ee As Employee
  
  On Error GoTo RemotelySelectAnotherEmployee_Err
  Call xSet("RemotelySelectAnotherEmployee")

  If Len(sNewEmployeeListItemKey) = 0 And (lNewEmployeeListItemIndex < 1) Then GoTo RemotelySelectAnotherEmployee_End
  
  Set ibf = F_Employees
  
  'if change here then change below
  If Len(sNewEmployeeListItemKey) > 0 Then
    lNewEmployeeIndex = ibf.lv.listitems(sNewEmployeeListItemKey).Tag
    lListIndex = ibf.lv.listitems(sNewEmployeeListItemKey).Index
  Else
    lNewEmployeeIndex = ibf.lv.listitems(lNewEmployeeListItemIndex).Tag
    lListIndex = lNewEmployeeListItemIndex
  End If
    
  Set ee = p11d32.CurrentEmployer.employees(lNewEmployeeIndex)
  If p11d32.CurrentEmployer.LoadEmployeeEx(ee) Then
     If Not ee Is Nothing Then Call ee.LoadBenefits(TBL_ALLBENEFITS, True)
     Set ibf.lv.SelectedItem = ibf.lv.listitems(lListIndex)
     Call ibf.lv.SelectedItem.EnsureVisible
     
     Call ibf.BenefitToScreen(lNewEmployeeIndex)
     
     Set ibf = p11d32.GetBenefitForm
     If Not ibf Is Nothing Then Call BenScreenSwitchEnd(ibf)
     
  End If

RemotelySelectAnotherEmployee_End:
  Set ibf = Nothing
  Call xReturn("RemotelySelectAnotherEmployee")
  Exit Function

RemotelySelectAnotherEmployee_Err:
  Call ErrorMessage(ERR_ERROR, Err, "RemotelySelectAnotherEmployee", "Remotely Select Another Employee", "Error remotely selecting another employee via the employees list item index.")
  Resume RemotelySelectAnotherEmployee_End
  Resume
End Function


Public Function MoveGotoCheck() As Boolean

  On Error GoTo MoveGotoCheck_Err
  Call xSet("MoveGotoCheck")
  
  If p11d32.CurrentEmployer Is Nothing Or (CurrentForm Is F_Employees Or CurrentForm Is F_CompanyDefined) Then MoveGotoCheck = True
  
MoveGotoCheck_End:
  Call xReturn("MoveGotoCheck")
  Exit Function

MoveGotoCheck_Err:
  Call ErrorMessage(ERR_ERROR, Err, "MoveGotoCheck", "Move Goto Check", "Error determining if it is possible to remotely select another employee.")
  Resume MoveGotoCheck_End
End Function

Public Sub FixLevelShowFunction(ByVal LEF As LIST_EMPLOYER_FUNCTIONS, ByVal bNewFixLevelShow As Boolean)
  Dim ibf As IBenefitForm2
  
  Set ibf = F_Employers
  
  Select Case LEF
    Case LEF_LISTVIEW_COLUMN
      If bNewFixLevelShow <> p11d32.FixLevelsShow Then
        If bNewFixLevelShow = False Then
          ibf.lv.ColumnHeaders.Remove (5)
        Else
          ibf.lv.ColumnHeaders.Add (5)
          ibf.lv.ColumnHeaders(5).Text = "Fix level"
        End If
      End If
    Case LEF_SIZE_COLUMNS
      If bNewFixLevelShow Then
        Call ColumnWidths(ibf.lv, 30, 20, 20, 10, 10)
      Else
        Call ColumnWidths(ibf.lv, 40, 20, 20, 20)
      End If
  End Select
  
  p11d32.FixLevelsShow = bNewFixLevelShow
End Sub
Public Sub BenefitFormSelectDefaultControl(frm As Form)
  Dim c As Control, cToSelect As Control
  Dim lTabIndex As Long
  Dim b As Boolean
  
  On Error GoTo BenefitFormSelectDefaultControl_ERR
  
  Call xSet("BenefitFormSelectDefaultControl")
  
  If IsBenefitForm(frm) Then
    For Each c In frm.Controls
      If IsVisible(c) And HasTabStop(c) Then
        If Not TypeOf c Is ListView And (Not TypeOf c Is CommandButton) Then
          If Not b Then
            lTabIndex = c.TabIndex
            Set cToSelect = c
            b = True
          Else
            If c.TabIndex < lTabIndex Then
              lTabIndex = c.TabIndex
              Set cToSelect = c
            End If
          End If
        End If
      End If
    Next
  End If
  
  If Not cToSelect Is Nothing Then Call ControlSetFocus(cToSelect)
  
BenefitFormSelectDefaultControl_END:
  Call xReturn("BenefitFormSelectDefaultControl")
  Exit Sub
BenefitFormSelectDefaultControl_ERR:
  If Not frm Is Nothing Then
    Call ErrorMessage(ERR_ERROR, Err, "BenefitFormSelectDefaultControl", "Benefit Form Select Default Control", "Error selecting the benefit forms default control. Form = " & frm.Name)
  Else
    Call ErrorMessage(ERR_ERROR, Err, "BenefitFormSelectDefaultControl", "Benefit Form Select Default Control", "Error selecting the benefit forms default control.")
  End If
  Resume BenefitFormSelectDefaultControl_END
  Resume
End Sub
Private Sub ControlSetFocus(c As Control)
  On Error GoTo ControlSetFocus_ERR
  
  'may not be able to setfocus when getting the default control as maybe CDB and disabled
  c.SetFocus
  
ControlSetFocus_END:
  Exit Sub
ControlSetFocus_ERR:
  Resume ControlSetFocus_END
End Sub
Public Sub EnableFrame(ByVal frm As Form, ByVal fra As frame, ByVal bEnable As Boolean)
  Dim c As Control
  
On Error GoTo err_err
  For Each c In frm.Controls
    If (c.Container Is fra) Then
      If TypeOf c Is Label Then
        Set c = c
      End If
      c.Enabled = bEnable
      If (TypeOf c Is ValText) Then
         If (bEnable) Then
          c.ForeColor = vbBlack
         Else
          c.ForeColor = vbGrayText
         End If
      End If
    End If
NEXT_ITEM:
  Next
  fra.Enabled = bEnable
  
err_end:
  Exit Sub
err_err:
  Resume NEXT_ITEM:
End Sub
Private Function IsVisible(c As Control)
  On Error GoTo IsVisible_ERR
  
  If (Not TypeOf c Is HOOK) Then
    IsVisible = c.Visible
    IsVisible = True
  End If
  
  
IsVisible_END:
  Exit Function
IsVisible_ERR:
  Resume IsVisible_END
End Function
Private Function HasTabStop(c As Control)
  On Error GoTo HasTabStop_ERR
  
  HasTabStop = False
  If (Not (TypeOf c Is Label)) And (Not (TypeOf c Is frame)) And (Not (TypeOf c Is HOOK)) Then
    
    HasTabStop = c.TabStop
    HasTabStop = True
  End If
  
  
  
HasTabStop_END:
  Exit Function
HasTabStop_ERR:
  Resume HasTabStop_END
End Function
