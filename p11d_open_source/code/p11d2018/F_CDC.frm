VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{8D988532-0F0C-460C-B00E-7B5637E97680}#1.0#0"; "ATC2VTEXT.OCX"
Begin VB.Form F_CDC 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Company Defined Categories"
   ClientHeight    =   4440
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3555
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4440
   ScaleWidth      =   3555
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   2340
      TabIndex        =   3
      Top             =   4005
      Width           =   1095
   End
   Begin atc2valtext.ValText txt 
      Height          =   330
      Left            =   90
      TabIndex        =   2
      Top             =   3555
      Width           =   3390
      _ExtentX        =   5980
      _ExtentY        =   582
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxLength       =   50
      Text            =   ""
      TypeOfData      =   4
      AllowEmpty      =   0   'False
      AutoSelect      =   0
   End
   Begin MSComctlLib.ListView lv 
      Height          =   3030
      Left            =   45
      TabIndex        =   1
      Top             =   495
      Width           =   3480
      _ExtentX        =   6138
      _ExtentY        =   5345
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Item"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Used"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbar 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3555
      _ExtentX        =   6271
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      _Version        =   393216
   End
End
Attribute VB_Name = "F_CDC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Implements IBenefitForm2
Implements IFrmGeneral

Private m_InvalidVT As Control
Private benefit As IBenefitClass
Private Sub cmdOK_Click()
  If lv.listitems.Count > 0 And txt.FieldInvalid Then
    Call MsgBox("Invalid data, all descriptions must be different.", vbInformation, "Check data")
  Else
    Call TestChangedControls(Me)
    Call p11d32.CurrentEmployer.WriteCDCs
    Set benefit = Nothing
    Me.Hide
  End If
End Sub
Private Sub Form_Load()
  Call AddAddDelete(tbar)
End Sub
Private Sub IBenefitForm2_AddBenefit()
  Dim ben As IBenefitClass
  Dim ibf As IBenefitForm2
  
  Set ibf = Me
  
  Set ben = New CDC
  
  
  Call ibf.AddBenefitSetDefaults(ben)
  
  Call ibf.BenefitToListView(ben, p11d32.CurrentEmployer.CDCs.Add(ben))
  tbar.Buttons(2).Enabled = True
  ben.Dirty = True
  Call SelectBenefitByListItem(ibf, ibf.lv.listitems(ibf.lv.listitems.Count))
  
End Sub

Private Function IBenefitForm2_AddBenefitSetDefaults(ben As IBenefitClass) As Long
  Set ben.Parent = p11d32.CurrentEmployer
  ben.Value(cdc_name_db) = "Please enter a description..."
  ben.Value(cdc_IsUsed) = False
End Function

Private Property Let IBenefitForm2_benclass(ByVal RHS As BEN_CLASS)

End Property

Private Property Get IBenefitForm2_benclass() As BEN_CLASS

End Property

Private Property Set IBenefitForm2_benefit(RHS As IBenefitClass)
  Set benefit = RHS
End Property

Private Property Get IBenefitForm2_benefit() As IBenefitClass
  Set IBenefitForm2_benefit = benefit
End Property

Private Function IBenefitForm2_BenefitFormState(ByVal fState As BENEFIT_FORM_STATE) As Boolean
  On Error GoTo BenefitFormState_err

  Call xSet("IBenefitForm2_EnableBenefitForm")
  
  If (fState = FORM_ENABLED) Or (fState = FORM_CDB) Then
    If fState = FORM_ENABLED Then
      txt.Enabled = True
    Else
      ECASE ("Car CBD?") 'CAD
    End If
    lv.Enabled = True
    
  ElseIf fState = FORM_DISABLED Then
    Set benefit = Nothing
    txt.Enabled = False
    lv.Enabled = False 'new
  End If
  
  IBenefitForm2_BenefitFormState = True
    
BenefitFormState_end:
  Call xReturn("BenefitFormState")
  Exit Function
  
BenefitFormState_err:
  IBenefitForm2_BenefitFormState = False
  Call ErrorMessage(ERR_ERROR, Err, "BenefitFormState", "Benefit Form State", "Error setting the benefit form state.")
  Resume BenefitFormState_end
End Function

Private Function IBenefitForm2_BenefitOff() As Boolean
  txt.Text = ""
End Function

Private Function IBenefitForm2_BenefitOn() As Boolean
  txt.Text = benefit.Value(cdc_name_db)
End Function

Private Function IBenefitForm2_BenefitsToListView() As Long
  Dim ben As IBenefitClass
  Dim i As Long
  Dim ibf As IBenefitForm2
  
  Set ibf = Me
  
  Call ClearForm(ibf)
  For i = 1 To p11d32.CurrentEmployer.CDCs.Count
     Set ben = p11d32.CurrentEmployer.CDCs(i)
     IBenefitForm2_BenefitsToListView = IBenefitForm2_BenefitsToListView + ibf.BenefitToListView(ben, i)
  Next
  tbar.Buttons(2).Enabled = IBenefitForm2_BenefitsToListView <> 0
  
End Function

Private Function IBenefitForm2_BenefitToListView(ben As IBenefitClass, ByVal lBenefitIndex As Long) As Long
  Dim ibf As IBenefitForm2
  Dim lst As ListItem
  
  Set ibf = Me
  
  If Not ben Is Nothing Then
    If ibf.ValididateBenefit(ben) Then
      Set lst = ibf.lv.listitems.Add(, , ben.Name)
      IBenefitForm2_BenefitToListView = ibf.UpdateBenefitListViewItem(lst, ben, lBenefitIndex)
    End If
  End If
End Function

Private Function IBenefitForm2_BenefitToScreen(Optional ByVal BenefitIndex As Long = -1&, Optional ByVal UpdateBenefit As Boolean = True) As Boolean
  Dim ben As IBenefitClass
  
  Dim ibf As IBenefitForm2
  
  On Error GoTo BenefitToScreen_Err
  
  Call xSet("BenefitToScreen")
  
  Set ibf = Me
  
  If BenefitIndex <> -1 Then
    Set ben = p11d32.CurrentEmployer.CDCs(BenefitIndex)
    If Not ibf.ValididateBenefit(ben) Then Call Err.Raise(ERR_INVALIDBENCLASS, "BenefitToScreenHelper", "Benefit To Screen Helper", "Invalid benefit type.")
    Set ibf.benefit = ben
    Call ibf.BenefitOn
  Else
    Set ibf.benefit = Nothing
    Call ibf.BenefitOff
  End If
  
  Call SetBenefitFormState(ibf)
  IBenefitForm2_BenefitToScreen = True
  
BenefitToScreen_End:
  Call xReturn("BenefitToScreen")
  Exit Function
BenefitToScreen_Err:
  Call ErrorMessage(ERR_ERROR, Err, "BenefitToScreen", "Benefit To Screen", "Error placing a CDC to the screen.")
  Resume BenefitToScreen_End
End Function

'Private Property Get IBenefitForm2_ControlDefault() As Control
'  Set IBenefitForm2_ControlDefault = lv
'End Property

Private Property Get IBenefitForm2_lv() As MSComctlLib.IListView
  Set IBenefitForm2_lv = lv
End Property

Private Function IBenefitForm2_RemoveBenefit(ByVal BenefitIndex As Long) As Boolean
  Dim NextBenefitIndex As Long
  Dim ibf As IBenefitForm2
  Dim ben As IBenefitClass, ben2 As IBenefitClass, benEmployee As IBenefitClass
  Dim rs As Recordset
  Dim emp As Employee
  Dim i As Long, j As Long
  
  On Error GoTo RemoveBenefit_END
  
  Call xSet("RemoveBenefit")
  
  
  Set ibf = Me
  
  NextBenefitIndex = GetNextBestListItemBenefitIndex(ibf, BenefitIndex)
  Set ben = p11d32.CurrentEmployer.CDCs(BenefitIndex)
  If ben.Value(cdc_IsUsed) Then
    If MsgBox("The CDC selected for removal is used in current O Other benefits..." & vbCrLf & vbCrLf & "If you continue the benefits will be reset to no category (the process may take some time).", vbOKCancel, "RemoveBenefit") = vbCancel Then GoTo RemoveBenefit_END
    'find the employees load the benefits and reset CDCCategory to ""
    Call SetCursor(vbArrowHourglass)
    Set rs = p11d32.CurrentEmployer.db.OpenRecordset(sql.Queries(SELECT_CDC_EMPLOYEES, ben.Value(cdc_name_db)))
    If rs Is Nothing Then Call Err.Raise(ERR_RS_IS_NOTHING, "RemoveBenefit", "The rs is nothing when trying to select the employees with as CDC of " & ben.Value(cdc_name_db) & " to remove the CDC.")
    
    Do While Not rs.EOF
      For i = 1 To p11d32.CurrentEmployer.employees.Count
        Set benEmployee = p11d32.CurrentEmployer.employees(i)
        If Not benEmployee Is Nothing Then
            If StrComp(benEmployee.Value(ee_PersonnelNumber_db), rs!P_Num, vbTextCompare) = 0 Then
              Set emp = benEmployee
              Call emp.LoadBenefits(TBL_ALLBENEFITS, False)
              For j = 1 To emp.benefits.Count
                Set ben2 = emp.benefits(j)
                If Not ben2 Is Nothing Then
                  If ben2.BenefitClass = BC_OOTHER_N Then
                    ben2.Value(oth_CompanyDefinedCategoryKey_db) = 0
                    ben2.WriteDB
                  End If
                End If
              Next
              If Not emp Is p11d32.CurrentEmployer.CurrentEmployee Then Call emp.KillBenefitsEx
            End If
        End If
      Next
      rs.MoveNext
    Loop
    'refresh the current form
    
    
  End If
  
  Set ben = p11d32.CurrentEmployer.CDCs(BenefitIndex)
  Call ben.DeleteDB
  Call p11d32.CurrentEmployer.CDCs.Remove(BenefitIndex)
  If CurrentForm Is F_Other Then Call BenScreenSwitchEnd(CurrentForm)
  Call ibf.BenefitsToListView
  'select an item
  Call SelectBenefitByBenefitIndex(ibf, NextBenefitIndex)
  IBenefitForm2_RemoveBenefit = True

RemoveBenefit_END:
  Call ClearCursor
  Call xReturn("RemoveBenefit")
  Exit Function
RemoveBenefit_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "RemoveBenefit", "Remove Benefit", "Error removing the benefit for a CDC.")
  Resume RemoveBenefit_END
  Resume
End Function

Private Function IBenefitForm2_UpdateBenefitListViewItem(li As MSComctlLib.IListItem, benefit As IBenefitClass, Optional ByVal BenefitIndex As Long = 0&, Optional ByVal SelectItem As Boolean = False) As Long
  On Error GoTo UpdateBenefitListViewItemCDC_ERR
  
  Call xSet("UpdateBenefitListViewItemCDC")
  
  If Not li Is Nothing And Not benefit Is Nothing Then
    If BenefitIndex > 0 Then li.Tag = BenefitIndex
    li.Text = benefit.Name
    li.SubItems(1) = benefit.Value(cdc_IsUsed)
    If SelectItem Then li.Selected = SelectItem
    IBenefitForm2_UpdateBenefitListViewItem = li.Index
  End If

UpdateBenefitListViewItemCDC_END:
  Call xReturn("UpdateBenefitListViewItemCDC")
  Exit Function
UpdateBenefitListViewItemCDC_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "UpdateBenefitListViewItemCDC", "Update Benefit List View Item CDC", "Error updating the benefit list view item for a CDC.")
  Resume UpdateBenefitListViewItemCDC_END
  
End Function

Private Function IBenefitForm2_ValididateBenefit(ben As IBenefitClass) As Boolean
  IBenefitForm2_ValididateBenefit = Not ben Is Nothing
End Function

Private Function IFrmGeneral_CheckChanged(c As Control) As Boolean
  Dim bDirty As Boolean
  
  On Error GoTo CheckChanged_Err
  
  Call xSet("CheckChanged")
  
  
  If benefit Is Nothing Then GoTo CheckChanged_End
  
  Select Case c.Name
    Case "txt"
      bDirty = StrComp(c.Text, benefit.Value(cdc_name_db)) <> 0
      If bDirty Then
        benefit.Value(cdc_name_db) = c.Text
      End If
    Case Else
      ECASE ("Invalid control in check changed")
  End Select
  
  Call AfterCheckChanged(c, Me, bDirty, , False)
  
CheckChanged_End:
  
  Call xReturn("CheckChanged")
  Exit Function
CheckChanged_Err:
  Call ErrorMessage(ERR_ERROR, Err, "CheckChanged", "ERR_CHECKCHANGED", "This function has failed for the form " & Me.Name & ".")
  Resume CheckChanged_End
End Function

Private Property Get IFrmGeneral_InvalidVT() As Control
  IFrmGeneral_InvalidVT = m_InvalidVT
End Property

Private Property Set IFrmGeneral_InvalidVT(RHS As Control)
  Set m_InvalidVT = RHS
End Property

Private Sub lv_ItemClick(ByVal item As MSComctlLib.ListItem)
  If Not (lv.SelectedItem Is Nothing) Then
    Call IBenefitForm2_BenefitToScreen(item.Tag)
  End If
End Sub

Private Sub tbar_ButtonClick(ByVal Button As MSComctlLib.Button)
  Call AddDeleteClick(Button.Index, Me)
End Sub

Private Sub txt_KeyDown(KeyCode As Integer, Shift As Integer)
  txt.Tag = SetChanged(False)
End Sub

Private Sub txt_UserValidate(Valid As Boolean, Message As String, sTextEntered As String)
  Dim i As Long
  
  Valid = True
  
  If Len(sTextEntered) = 0 Or Len(sTextEntered) > txt.MaxLength Then
    Valid = False
    Exit Sub
  End If
  
  For i = 1 To lv.listitems.Count
    If Not lv.SelectedItem Is lv.listitems(i) Then
      If StrComp(lv.listitems(i).Text, sTextEntered) = 0 Then
        Valid = False
        Message = "No 2 company defined categories can have the same description."
        Exit For
      End If
    End If
  Next
End Sub

Private Sub txt_Validate(Cancel As Boolean)
  Call IFrmGeneral_CheckChanged(txt)
End Sub
