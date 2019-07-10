VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{E297AE83-F913-4A8C-873C-EDEAC00CB9AC}#2.1#0"; "atc3ubgrd.ocx"
Begin VB.Form F_FPCS 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Company-Defined rates"
   ClientHeight    =   4050
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6525
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4050
   ScaleWidth      =   6525
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkApproved 
      Caption         =   "Approved scheme"
      Height          =   285
      Left            =   2340
      TabIndex        =   5
      Top             =   3510
      Width           =   2580
   End
   Begin atc3ubgrd.UBGRD UBGRD 
      Height          =   2940
      Left            =   1710
      TabIndex        =   4
      Top             =   495
      Width           =   4740
      _ExtentX        =   8361
      _ExtentY        =   5186
   End
   Begin MSComctlLib.Toolbar tbar 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   6525
      _ExtentX        =   11509
      _ExtentY        =   741
      ButtonWidth     =   635
      ButtonHeight    =   582
      Appearance      =   1
      _Version        =   393216
   End
   Begin VB.TextBox txtFPCS 
      Height          =   375
      Left            =   45
      TabIndex        =   1
      Text            =   "txtFPCS"
      Top             =   3465
      Width           =   2175
   End
   Begin MSComctlLib.ListView LB 
      Height          =   2940
      Left            =   45
      TabIndex        =   0
      ToolTipText     =   "Car schemes"
      Top             =   495
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   5186
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Schemes"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   375
      Left            =   5220
      TabIndex        =   3
      Top             =   3645
      Width           =   1245
   End
End
Attribute VB_Name = "F_FPCS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Implements IBenefitForm2
Implements IFrmGeneral

Private m_FPCS As FPCS
Private m_InvalidVT As Control

Private Sub B_Cancel_Click()
 Call F_FPCS.Hide
End Sub

Private Sub B_OK_Click()
  
End Sub
Private Function ValidateFPCSName(ByVal sFPCSName As String)
  Dim CS As FPCS
  Dim l As Long
  
On Error GoTo ValidateFPCSName_Err

  Call xSet("ValidateFPCSName")
  
  For l = 1 To p11d32.CurrentEmployer.FPCSchemes.Count
    Set CS = p11d32.CurrentEmployer.FPCSchemes(l)
    If Not CS Is Nothing Then
      If StrComp(CS.Name, sFPCSName) = 0 Then
        ValidateFPCSName = False
        GoTo ValidateFPCSName_End
      End If
    End If
  Next
  
  ValidateFPCSName = True
  
ValidateFPCSName_End:
  Set CS = Nothing
  Call xReturn("ValidateFPCSName")
  Exit Function

ValidateFPCSName_Err:
  Call ErrorMessage(ERR_ERROR, Err, "ValidateFPCSName", "Validate FPCS Name", "Error validating a scheme name.")
  Resume ValidateFPCSName_End
  Resume
End Function
Private Function AddNewFPCS(sNewFPCS As String) As Boolean
  Dim l As Long
  Dim CS As FPCS
  Dim li As ListItem
  Dim ibf As IBenefitForm2
  
  On Error GoTo AddNewFPCS_Err
  Call xSet("AddNewFPCS")
  
  If ValidateFPCSName(sNewFPCS) = False Then GoTo AddNewFPCS_End
  
  Set CS = New FPCS
  CS.Dirty = True
  CS.Name = sNewFPCS
  l = p11d32.CurrentEmployer.FPCSchemes.Add(CS)
  Set li = lb.listitems.Add(, , sNewFPCS)
  li.Tag = l
  Set ibf = Me
  Set ibf.lv.SelectedItem = li
  ibf.BenefitToScreen (li.Tag)
  AddNewFPCS = True
  
AddNewFPCS_End:
  Set li = Nothing
  Set CS = Nothing
  Call xReturn("AddNewFPCS")
  Exit Function

AddNewFPCS_Err:
  Call ErrorMessage(ERR_ERROR, Err, "AddNewFPCS", "Add New FPCS", "Error adding a new car scheme.")
  Resume AddNewFPCS_End
End Function

Private Sub chkApproved_Click()
  Call IFrmGeneral_CheckChanged(chkApproved)
End Sub

Private Sub cmdClose_Click()
  On Error GoTo cmdClose_ERR
  
  Call xSet("cmdClose")
  
  p11d32.CurrentEmployer.WriteFPCS
  
cmdClose_END:
  Unload Me
  Call xReturn("cmdClose")
  Exit Sub
cmdClose_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "cmdClose", "cmd Close", "Error writing or finishing the car schemes.")
  Resume cmdClose_END
End Sub


Private Sub Form_Load()
  Dim grd As Object
  
  On Error GoTo F_FPCS_Load_ERR

  Call xSet("F_FPCS_Load")
  
  Call AddAddDelete(tbar)
  
  Set grd = UBGRD.Grid
  
  Call AddUBGRDStandardColumn(grd, 0, 1244.976, "Band Name", "")
  Call AddUBGRDStandardColumn(grd, 1, 1019.906, "Miles above", "")
  Call AddUBGRDStandardColumn(grd, 2, 945.0709, "CC above", "")
  Call AddUBGRDStandardColumn(grd, 3, 945.0709, "Rate (£)", "")
  
  grd.AllowUpdate = True
  
F_FPCS_Load_END:
  Set grd = Nothing
  Call xReturn("F_FPCS_Load")
  Exit Sub
F_FPCS_Load_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "F_FPCS_Load", "F_FPCS Load", "Unable to load the car scheme form.")
  Resume F_FPCS_Load_END
  Resume
End Sub


Private Function IBenefitForm2_AddBenefitSetDefaults(ben As IBenefitClass) As Long

End Function

Private Function IBenefitForm2_BenefitOff() As Boolean
  'not used
End Function

Private Function IBenefitForm2_BenefitOn() As Boolean
  'not used
End Function

Private Function IBenefitForm2_BenefitToListView(ben As IBenefitClass, ByVal lBenefitIndex As Long) As Long
End Function

'Private Property Get IBenefitForm2_ControlDefault() As Control
'  Set IBenefitForm2_ControlDefault = lb
'End Property

Private Function IBenefitForm2_UpdateBenefitListViewItem(li As MSComctlLib.IListItem, benefit As IBenefitClass, Optional ByVal BenefitIndex As Long = 0&, Optional ByVal SelectItem As Boolean = False) As Long
  ECASE "UpdateBenefitListViewItem"
End Function
Private Function IBenefitForm2_ValididateBenefit(ben As IBenefitClass) As Boolean
End Function

Private Function IFrmGeneral_CheckChanged(c As Control) As Boolean
  
  Dim ibf As IBenefitForm2
  
  On Error GoTo CheckChanged_Err
  Call xSet("CheckChanged")
  
  With c
    If m_FPCS Is Nothing Then
      GoTo CheckChanged_End
    End If
    If lb.SelectedItem Is Nothing Then
      GoTo CheckChanged_End
    End If
    'we are asking if the value has changed and if it is valid thus save
    
    Select Case .Name
      Case "txtFPCS"
        If StrComp(lb.SelectedItem.Text, .Text) Then
          If ValidateFPCSName(txtFPCS) = False Then
            Call ErrorMessage(ERR_INFO, Err, "txtFPCS_KeyDown", "txtFPCS KeyDown", "The car scheme name you chosen is already in use. Press escape to cancel change.")
            txtFPCS.SetFocus
            txtFPCS.SelLength = Len(txtFPCS)
          Else
            m_FPCS.Name = txtFPCS
            lb.SelectedItem.Text = txtFPCS
          End If
        End If
      Case "chkApproved"
        m_FPCS.Approved = ChkBoxToBool(chkApproved)
      Case Else
         ECASE "Unknown"
     End Select
    
    'must be required in all check changed
    Set ibf = Me
    IFrmGeneral_CheckChanged = True 'ibf.UpdateBenefitListViewItem(ibf.lv.SelectedItem, ibf.benefit)

  End With
  
CheckChanged_End:
  Set ibf = Nothing
  Call xReturn("CheckChanged")
  Exit Function
  
CheckChanged_Err:
  IFrmGeneral_CheckChanged = False
  Call ErrorMessage(ERR_ERROR, Err, "CheckChanged", "ERR_CHECKCHANGED", "This function has failed for the form " & Me.Name & ".")
  Resume CheckChanged_End

End Function

Private Property Get IFrmGeneral_InvalidVT() As Control
  IFrmGeneral_InvalidVT = m_InvalidVT
End Property

Private Property Set IFrmGeneral_InvalidVT(NewValue As Control)
  Set m_InvalidVT = NewValue
  
End Property

Private Sub LB_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  Call TestChangedControls(Me)
End Sub

Private Sub txtFPCS_KeyDown(KeyCode As Integer, Shift As Integer)
  txtFPCS.Tag = SetChanged(False)
  If KeyCode = vbKeyEscape Then
    txtFPCS = m_FPCS.Name
    txtFPCS.SelLength = Len(txtFPCS)
  End If
End Sub

Private Sub IBenefitForm2_AddBenefit()
  Dim sNewScheme As String
  Dim sPrompt As String
  
  
  sPrompt = "Please enter a name for a company car scheme."
RETRY:
  sNewScheme = InputBox(sPrompt, "New Scheme", "")
  
  If Len(sNewScheme) Then
    'try and add then scheme
    If Not AddNewFPCS(sNewScheme) Then
      sPrompt = "The name you chose is already present, please try again."
      GoTo RETRY
    Else
    End If
  End If
  
  
  
End Sub

Private Property Set IBenefitForm2_benefit(RHS As IBenefitClass)

End Property

Private Property Get IBenefitForm2_benefit() As IBenefitClass
  ECASE "IBenefitForm2_benefit"
End Property

Private Function IBenefitForm2_BenefitFormState(ByVal fState As BENEFIT_FORM_STATE) As Boolean
  'do nothing here as is first selected then disble cross / done in benefit to screen
  
End Function

Private Function IBenefitForm2_BenefitsToListView() As Long
  Dim CS As FPCS
  Dim l As Long
  Dim li As ListItem
  
  
  lb.listitems.Clear
  
  For l = 1 To p11d32.CurrentEmployer.FPCSchemes.Count
    Set CS = p11d32.CurrentEmployer.FPCSchemes(l)
    If Not CS Is Nothing Then
      Set li = lb.listitems.Add(, , CS.Name)
      li.Tag = p11d32.CurrentEmployer.FPCSchemes.ItemIndex(CS)
    End If
  Next
  
  Set CS = Nothing
  Set li = Nothing
End Function

Private Function IBenefitForm2_BenefitToScreen(Optional ByVal BenefitIndex As Long = -1&, Optional ByVal UpdateBenefit As Boolean = True) As Boolean
  Dim b As Boolean
  
  On Error GoTo BenefitToScreen_Err
  
  Call xSet("BenefitToScreen")
  
  If BenefitIndex <> -1 Then
    Set m_FPCS = p11d32.CurrentEmployer.FPCSchemes(BenefitIndex)
    
    If m_FPCS Is Nothing Then Call Err.Raise(ERR_IS_NOTHING, "BenefitToScreen", "The FPCS is nothing.")
    If m_FPCS.Bands Is Nothing Then Call Err.Raise(ERR_IS_NOTHING, "BenefitToScreen", "The bands objectlist is nothing.")
    
    Set UBGRD.ObjectList = m_FPCS.Bands
    b = BenefitIndex <> 1
    
    txtFPCS.Visible = b
    UBGRD.Grid.Enabled = b
    tbar.Buttons(2).Enabled = b
    chkApproved.Visible = b
    
    'Call m_FPCS.DebugScheme(m_FPCS) 'leave
    Call UBGRD.Grid.ReBind 'paint me / fills the grid
    txtFPCS = m_FPCS.Name
    chkApproved = BoolToChkBox(m_FPCS.Approved)
  End If
  
BenefitToScreen_End:
  Call xSet("BenefitToScreen")
  Exit Function
BenefitToScreen_Err:
  Call ErrorMessage(ERR_ERROR, Err, "BenefitToScreen", "Benefit To Screen", "Error placing a FPCS to the screen.")
  Resume BenefitToScreen_End
End Function

Private Property Let IBenefitForm2_benclass(ByVal RHS As BEN_CLASS)

End Property

Private Property Get IBenefitForm2_benclass() As BEN_CLASS

End Property

Private Property Get IBenefitForm2_lv() As MSComctlLib.IListView
  Set IBenefitForm2_lv = lb
End Property

Private Function IBenefitForm2_RemoveBenefit(ByVal BenefitIndex As Long) As Boolean
  Dim ibf As IBenefitForm2
  Dim NextBenefitIndex As Long
  Dim rs As Recordset
  Dim FPCS As FPCS
  Dim benEECar As IBenefitClass, benEE As IBenefitClass
  Dim ee As Employee
  Dim i As Long, j As Long
  
On Error GoTo RemoveBenefit_ERR
  Call xSet("RemoveBenefit")
  
  If MsgBox("Removing the scheme may have an impact on current employee cars." & vbCrLf & vbCrLf & "Removal from all the cars may take some time, do you wish to continue?", vbYesNo, "RemoveBenefit") = vbNo Then GoTo RemoveBenefit_END
  
  Call SetCursor(vbHourglass)
  
  Set FPCS = p11d32.CurrentEmployer.FPCSchemes(BenefitIndex)
  
  Set rs = p11d32.CurrentEmployer.db.OpenRecordset(sql.Queries(SELECT_FPCS_EECARS, FPCS.Name), dbOpenDynaset)
  'reset eecars using scheme to IR
  Do While Not rs.EOF
    For i = 1 To p11d32.CurrentEmployer.employees.Count
      Set ee = p11d32.CurrentEmployer.employees(i)
      Set benEE = p11d32.CurrentEmployer.employees(i)
      If StrComp(benEE.value(ee_PersonnelNumber_db), rs!P_Num) <> 0 Then GoTo NEXT_EMPLOYEE
      If ee.BenefitsLoaded Then
        For j = 1 To ee.benefits.Count
          Set benEECar = ee.benefits(j)
          If Not benEECar Is Nothing Then
            If benEECar.BenefitClass = BC_EMPLOYEE_CAR_E Then
              If benEECar.value(eecar_FPCS_db) = BenefitIndex Then
                benEECar.value(eecar_FPCS_db) = L_IRFPCS
              End If
            End If
          End If
        Next
        Exit For
      End If
NEXT_EMPLOYEE:
    Next
    rs.Edit
    rs!FPCS = S_IRFPCS
    rs.Update

    rs.MoveNext
  Loop
  
  Set ibf = Me
  NextBenefitIndex = GetNextBestListItemBenefitIndex(ibf, BenefitIndex)
  p11d32.CurrentEmployer.FPCSchemes.Remove (BenefitIndex)
  
  ibf.lv.listitems.Remove (ibf.lv.SelectedItem.Index)
  
  Call SelectBenefitByBenefitIndex(ibf, NextBenefitIndex)
  
  If CurrentForm Is F_EmployeeCar Then
    Set ibf = CurrentForm
    Call BenScreenSwitchEnd(CurrentForm)
  End If
  
  IBenefitForm2_RemoveBenefit = True
    
    
RemoveBenefit_END:
  Call ClearCursor
  Call xReturn("RemoveBenefit")
  Exit Function
  
RemoveBenefit_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "RemoveBenefit", "Remove Benefit", "Error removing benefit.")
  Resume RemoveBenefit_END
  Resume
End Function
Private Sub LB_ItemClick(ByVal Item As MSComctlLib.ListItem)
  If Not Item Is Nothing Then
    Call IBenefitForm2_BenefitToScreen(Item.Tag)
  End If
End Sub
Private Sub tbar_ButtonClick(ByVal Button As MSComctlLib.Button)
  Call AddDeleteClick(Button.Index, Me)
End Sub

Private Sub txtFPCS_Validate(Cancel As Boolean)
  Call IFrmGeneral_CheckChanged(txtFPCS)
End Sub

Private Sub ubgrd_DeleteData(ObjectList As ObjectList, ObjectListIndex As Long)
  Call UBGRD.ObjectList.Remove(ObjectListIndex)
End Sub


Private Sub UBGRD_ReadData(RowBuf As TrueDBGrid60.RowBuffer, ByVal RowBufRowIndex As Long, ObjectList As ObjectList, ByVal ObjectListIndex As Long)
  Dim i As Long
  Dim Band As FPCSBand

  Set Band = ObjectList(ObjectListIndex)

  For i = 0 To (RowBuf.ColumnCount - 1)
    Select Case i
      Case 0
        RowBuf.value(RowBufRowIndex, i) = Band.Name
      Case 1
        RowBuf.value(RowBufRowIndex, i) = Band.GreaterThanMiles
      Case 2
        RowBuf.value(RowBufRowIndex, i) = Band.GreaterThanCC
      Case 3
        RowBuf.value(RowBufRowIndex, i) = Band.Rate
      Case Else
        ECASE ("Invalid column if get user data")
    End Select
  Next i
    
  Set Band = Nothing

End Sub


Private Sub UBGRD_ValidateTCS(FirstColIndexInError As Long, ValidateMessage As String, ByVal RowBuf As TrueDBGrid60.RowBuffer, ByVal RowBufRowIndex As Long, ByVal ObjectListIndex As Long)

  Dim i As Long

  With RowBuf
    For i = 0 To RowBuf.ColumnCount - 1
      Select Case i
        Case 0
          If GridIsZeroLength(ValidateMessage, RowBuf.value(RowBufRowIndex, i), ObjectListIndex) Then
            FirstColIndexInError = i
            Exit Sub
          End If
        Case 1, 2, 3
          If GridIsNotNumericOrLong(ValidateMessage, RowBuf.value(RowBufRowIndex, i), ObjectListIndex) Then
            FirstColIndexInError = i
            Exit Sub
          End If
      End Select
    Next
  End With

  Call m_FPCS.CheckForDuplicates(FirstColIndexInError, ValidateMessage, RowBuf, RowBufRowIndex, ObjectListIndex)


End Sub


Private Sub ubgrd_WriteData(ByVal RowBuf As TrueDBGrid60.RowBuffer, ByVal RowBufRowIndex As Long, ObjectList As ObjectList, ObjectListIndex As Long)
  Dim Band As FPCSBand
  Dim lNewCC As Long, lNewMiles As Long
  
  If ObjectListIndex = -1 Then
    Set Band = New FPCSBand
    ObjectListIndex = ObjectList.Add(Band)
  Else
    Set Band = ObjectList(ObjectListIndex)
  End If
  
  With Band
    If Not IsNull(RowBuf.value(RowBufRowIndex, 0)) Then .Name = RowBuf.value(RowBufRowIndex, 0)
    If Not IsNull(RowBuf.value(RowBufRowIndex, 1)) Then .GreaterThanMiles = IIf(RowBuf.value(RowBufRowIndex, 1) <= 0, 0, RowBuf.value(RowBufRowIndex, 1))
    If Not IsNull(RowBuf.value(RowBufRowIndex, 2)) Then .GreaterThanCC = IIf(RowBuf.value(RowBufRowIndex, 2) <= 0, 0, RowBuf.value(RowBufRowIndex, 2))
    If Not IsNull(RowBuf.value(RowBufRowIndex, 3)) Then .Rate = RowBuf.value(RowBufRowIndex, 3)
  End With
  
  
  
End Sub
