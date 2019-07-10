VERSION 5.00
Object = "{8D988532-0F0C-460C-B00E-7B5637E97680}#1.0#0"; "atc2vtext.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form F_CompanyDefined 
   Caption         =   " "
   ClientHeight    =   5760
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8325
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5760
   ScaleWidth      =   8325
   Begin MSComctlLib.ListView lb 
      Height          =   3885
      Left            =   0
      TabIndex        =   6
      Tag             =   "free,font"
      Top             =   120
      Width           =   8580
      _ExtentX        =   15134
      _ExtentY        =   6853
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Item"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Unique ID"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Benefit"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Frame fmeInput 
      ForeColor       =   &H00FF0000&
      Height          =   1830
      Left            =   45
      TabIndex        =   7
      Tag             =   "FREE,FONT"
      Top             =   3915
      Width           =   8220
      Begin VB.CommandButton cmdApply 
         Caption         =   "&Apply Benefit"
         Height          =   315
         Left            =   6435
         TabIndex        =   5
         Tag             =   "FREE,FONT"
         Top             =   1395
         Width           =   1500
      End
      Begin atc2valtext.ValText txt 
         Height          =   330
         Index           =   0
         Left            =   1170
         TabIndex        =   0
         Tag             =   "FREE,FONT"
         Top             =   230
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   582
         BackColor       =   255
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   ""
         TypeOfData      =   3
         AllowEmpty      =   0   'False
         AutoSelect      =   0
      End
      Begin VB.ComboBox cbx 
         ForeColor       =   &H00000080&
         Height          =   315
         ItemData        =   "F_CBDNEW.frx":0000
         Left            =   1170
         List            =   "F_CBDNEW.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Tag             =   "FREE,FONT"
         Top             =   1010
         Width           =   2925
      End
      Begin atc2valtext.ValText txt 
         Height          =   330
         Index           =   1
         Left            =   1170
         TabIndex        =   1
         Tag             =   "FREE,FONT"
         Top             =   620
         Width           =   2895
         _ExtentX        =   5106
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
         Text            =   ""
         TypeOfData      =   4
         AllowEmpty      =   0   'False
         AutoSelect      =   0
      End
      Begin atc2valtext.ValText txt 
         Height          =   330
         Index           =   2
         Left            =   6435
         TabIndex        =   3
         Tag             =   "FREE,FONT"
         Top             =   225
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   582
         BackColor       =   255
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   ""
         AllowEmpty      =   0   'False
         TXTAlign        =   2
         AutoSelect      =   0
      End
      Begin atc2valtext.ValText txt 
         Height          =   330
         Index           =   3
         Left            =   6435
         TabIndex        =   4
         Tag             =   "FREE,FONT"
         Top             =   615
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   582
         BackColor       =   255
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   ""
         AllowEmpty      =   0   'False
         TXTAlign        =   2
         AutoSelect      =   0
      End
      Begin P11D2018.ValCombo cboIRDesc2 
         Height          =   315
         Left            =   1170
         TabIndex        =   12
         Tag             =   "free,font"
         Top             =   1400
         Width           =   2925
         _ExtentX        =   5159
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin atc2valtext.ValText txt 
         Height          =   330
         Index           =   4
         Left            =   6435
         TabIndex        =   14
         Tag             =   "FREE,FONT"
         Top             =   990
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   582
         BackColor       =   255
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   ""
         AllowEmpty      =   0   'False
         TXTAlign        =   2
         AutoSelect      =   0
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "OpRA amount foregone"
         ForeColor       =   &H8000000D&
         Height          =   195
         Index           =   6
         Left            =   4500
         TabIndex        =   16
         Tag             =   "FREE,FONT"
         Top             =   1080
         Width           =   1680
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Made good"
         ForeColor       =   &H8000000D&
         Height          =   195
         Index           =   3
         Left            =   4500
         TabIndex        =   15
         Tag             =   "FREE,FONT"
         Top             =   675
         Width           =   810
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "IR Description"
         ForeColor       =   &H8000000D&
         Height          =   195
         Index           =   5
         Left            =   135
         TabIndex        =   13
         Tag             =   "FREE,FONT"
         Top             =   1440
         Width           =   1005
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Unique id"
         ForeColor       =   &H8000000D&
         Height          =   240
         Index           =   4
         Left            =   135
         TabIndex        =   11
         Tag             =   "FREE,FONT"
         Top             =   660
         Width           =   675
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Value"
         ForeColor       =   &H8000000D&
         Height          =   195
         Index           =   2
         Left            =   4500
         TabIndex        =   10
         Tag             =   "FREE,FONT"
         Top             =   270
         Width           =   405
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Description"
         ForeColor       =   &H8000000D&
         Height          =   195
         Index           =   1
         Left            =   135
         TabIndex        =   9
         Tag             =   "FREE,FONT"
         Top             =   270
         Width           =   795
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Category"
         ForeColor       =   &H8000000D&
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   8
         Tag             =   "FREE,FONT"
         Top             =   1050
         Width           =   630
      End
   End
End
Attribute VB_Name = "F_CompanyDefined"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Implements IBenefitForm2
Implements IFrmGeneral
Public benefit As IBenefitClass

Private mclsResize As New clsFormResize
Private m_BenClass As BEN_CLASS
Private m_InvalidVT As Control

Private Const L_DES_HEIGHT  As Long = 6090
Private Const L_DES_WIDTH  As Long = 8445

Private Sub cboIRDesc2_Click()
  Call IFrmGeneral_CheckChanged(cboIRDesc2)
End Sub



Private Sub cbx_Click()
  Call IFrmGeneral_CheckChanged(cbx)
End Sub

Private Sub cmdApply_Click()
  If F_Employees.lb.listitems.Count > 0 Then
    If p11d32.CurrentEmployer.MoveMenuUpdateEmployee Then
      Set F_ApplyCompanyDefined.benefit = benefit
'      F_ApplyCompanyDefined.Show vbModal
      Call p11d32.Help.ShowForm(F_ApplyCompanyDefined, vbModal)
      Unload F_ApplyCompanyDefined
    End If
  Else
    Call ErrorMessage(ERR_ERROR, Err, "cmdApply", "cmd Apply", "No employees to apply company defined benefits to.")
  End If
End Sub

Private Sub FillComboBox()
  Dim sl As StringList
  Dim l As Integer
  Set sl = New StringList
  cboIRDesc2.Clear
  Select Case benefit.BenefitClass
    Case BC_CLASS_1A_M
      Call sl.Add(S_IR_DESC_M_C1A_SUBS_AND_FEES)
      Call sl.Add(S_IR_DESC_M_C1A_ED_ASS)
      Call sl.Add(S_IR_DESC_M_C1A_NON_QUAL_RELOC)
      Call sl.Add(S_IR_DESC_M_C1A_STOP_LOSS_CHARGES)
    Case BC_NON_CLASS_1A_M
      Call sl.Add(S_IR_DESC_M_NC1A_SUBS_AND_FEES)
      Call sl.Add(S_IR_DESC_M_NC1A_NURSERY)
      Call sl.Add(S_IR_DESC_M_NC1A_ED_ASS)
      Call sl.Add(S_IR_DESC_M_NC1A_LOANS_WRIT_WAIV)
    Case BC_OOTHER_N
      Call sl.Add(S_IR_DESC_N_PERSONAL_INC_EXP)
      Call sl.Add(S_IR_DESC_N_WORK_HOME)
    Case Else
  End Select
  
  If (sl.Count > 0) Then
    Call cboIRDesc2.AddItem(S_IR_DESC_PLEASE_SELECT)
    For l = 1 To sl.Count
      Call cboIRDesc2.AddItem(sl.Item(l))
    Next
    Call cboIRDesc2.AddItem(S_IR_DESC_OTHER)
  End If
  cboIRDesc2.InvalidValue = S_IR_DESC_PLEASE_SELECT
End Sub





Private Sub Form_Load()
  Dim i As Long
  Dim ben As IBenefitClass
  
  On Error GoTo Form_Load_ERR
  
  Call xSet("Form_Load")
  
  Set ben = New other
  
  With cbx
    For i = BC_FIRST_ITEM To BC_LAST_ITEM
      ben.BenefitClass = i
      If IBenefitForm2_ValididateBenefit(ben) Then
        .AddItem (p11d32.Rates.BenClassTo(i, BCT_DBCLASS))
        .ItemData(.ListCount - 1) = i
      End If
    Next
  End With
  Call SetupOpraInput(lbl(6), txt(4))
  Call mclsResize.InitResize(Me, L_DES_HEIGHT, L_DES_WIDTH, , , , MDIMain)
  
Form_Load_END:
  Call xSet("Form_Load")
  Exit Sub
Form_Load_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "Form_Load", "Form Load", "Error loading the form F_Companydefined.")
  Resume Form_Load_END
End Sub

Private Sub Form_Resize()
  Call mclsResize.Resize
  Call ColumnWidths(lb, 50, 25, 25)
End Sub

Private Sub IBenefitForm2_AddBenefit()
  Dim ben As IBenefitClass
  
On Error GoTo AddBenefit_Err

  Call xSet("AddBenefit")
  
  Set ben = New other
  Call AddBenefitHelper(Me, ben)
  Call txt(0).SetFocus
  
AddBenefit_End:
  Set ben = Nothing
  Call xReturn("AddBenefit")
  Exit Sub
AddBenefit_Err:
  Call ErrorMessage(ERR_ERROR, Err, "AddBenefit", "ERR_ADDBENEFIT", "Error in the AddBenefit function, called from the form " & Me.Name & ".")
  Resume AddBenefit_End
  Resume
End Sub

Private Function IBenefitForm2_AddBenefitSetDefaults(ben As IBenefitClass) As Long
  
  With ben
    .value(oth_item_db) = "Please enter a description..."
    .value(oth_class_db) = p11d32.Rates.BenClassTo(BC_PRIVATE_MEDICAL_I, BCT_DBCLASS)
    .value(oth_GrossAmountPaidByEmployer_db) = 0&
    .value(oth_madegood_db) = 0&
    .value(oth_HMITCode_UDBCode_db) = "Please enter a unique code..."
    .value(oth_category_db) = p11d32.Rates.BenClassTo(BC_PRIVATE_MEDICAL_I, BCT_DBCLASS)
    .value(oth_NIC_Class1A_Able) = p11d32.Rates.BenClassTo(BC_PRIVATE_MEDICAL_I, BCT_CLASS1A_ABLE)
    .value(oth_availablefrom_db) = p11d32.Rates.value(TaxYearStart)
    .value(oth_availableto_db) = p11d32.Rates.value(TaxYearEnd)
    .value(oth_CompanyDefinedCategoryKey_db) = 0
    .value(oth_OPRA_Ammount_Foregone_db) = 0
    
    ben.CompanyDefined = True
    .BenefitClass = BC_PRIVATE_MEDICAL_I
  End With
  
End Function

Private Property Set IBenefitForm2_benefit(NewValue As IBenefitClass)
  Set benefit = NewValue
End Property

Private Property Get IBenefitForm2_benefit() As IBenefitClass
  Set IBenefitForm2_benefit = benefit
End Property

Private Function IBenefitForm2_BenefitFormState(ByVal fState As BENEFIT_FORM_STATE) As Boolean
  If fState = FORM_CDB Then fState = FORM_ENABLED
  IBenefitForm2_BenefitFormState = BenefitFormStateEx(fState, benefit, fmeInput)
  
End Function

Private Function IBenefitForm2_BenefitOff() As Boolean
    txt(0).Text = ""
    txt(1).Text = ""
    txt(2).Text = ""
    txt(3).Text = ""
    txt(4).Text = ""
    cbx.Enabled = False
    cboIRDesc2.Enabled = False
End Function

Private Function IBenefitForm2_BenefitOn() As Boolean
  Dim i As Long
  Dim l As Integer
  
  With benefit
    txt(0).Text = .value(oth_item_db)
    txt(1).Text = .value(oth_HMITCode_UDBCode_db)
    txt(2).Text = .value(oth_GrossAmountPaidByEmployer_db)
    txt(3).Text = .value(oth_madegood_db)
    txt(4).Text = .value(oth_OPRA_Ammount_Foregone_db)
    cbx.Enabled = True
    For i = 0 To cbx.ListCount - 1
      If cbx.ItemData(i) = benefit.BenefitClass Then
        cbx.ListIndex = i
        Exit For
      End If
    Next
    If benefit.BenefitClass = BC_NON_CLASS_1A_M Or _
        benefit.BenefitClass = BC_OOTHER_N Or _
        benefit.BenefitClass = BC_CLASS_1A_M Then
      Call FillComboBox
      cboIRDesc2.Visible = True
      cboIRDesc2.Enabled = True
      lbl(5).Visible = True
      Call IRDescriptionToCombo(cboIRDesc2.ComboBox, benefit)
    Else
      cboIRDesc2.Visible = False
      lbl(5).Visible = False
    End If
        
  End With
End Function

Private Function IBenefitForm2_BenefitsToListView() As Long
  IBenefitForm2_BenefitsToListView = BenefitsToListView(Me)
End Function

Private Function IBenefitForm2_BenefitToListView(ben As IBenefitClass, ByVal lBenefitIndex As Long) As Long
  IBenefitForm2_BenefitToListView = BenefitToListView(ben, Me, lBenefitIndex)
End Function

Private Function IBenefitForm2_BenefitToScreen(Optional ByVal BenefitIndex As Long = -1&, Optional ByVal UpdateBenefit As Boolean = True) As Boolean
  Dim ben As IBenefitClass
  
  On Error GoTo BenefitToScreen_Err
  
  Call xSet("BenefitToScreen")
  
  If BenefitIndex > 0 Then
    Set ben = p11d32.CurrentEmployer.CDBEmployee.benefits(BenefitIndex)
    m_BenClass = ben.BenefitClass
  End If
  IBenefitForm2_BenefitToScreen = BenefitToScreenHelper(Me, BenefitIndex, UpdateBenefit)
  
BenefitToScreen_End:
  Call xReturn("BenefitToScreen")
  Exit Function
BenefitToScreen_Err:
  Set ben = Nothing
  Call ErrorMessage(ERR_ERROR, Err, "BenefitToScreen", "Benefit To Screen", "Error placing a CDB benefit to the screen.")
  Resume BenefitToScreen_End
End Function

Private Property Let IBenefitForm2_benclass(ByVal NewValue As BEN_CLASS)
  m_BenClass = NewValue
End Property

Private Property Get IBenefitForm2_benclass() As BEN_CLASS
  IBenefitForm2_benclass = m_BenClass
End Property

Private Property Get IBenefitForm2_lv() As MSComctlLib.IListView
  Set IBenefitForm2_lv = lb
End Property

Private Function IBenefitForm2_RemoveBenefit(ByVal BenefitIndex As Long) As Boolean
  Set F_ApplyCompanyDefined.benefit = benefit
  Call F_ApplyCompanyDefined.ViewAsignments(True)
  Call F_ApplyCompanyDefined.ApplyCompanyDefinedBenefits
  Unload F_ApplyCompanyDefined
  Set F_ApplyCompanyDefined = Nothing
  IBenefitForm2_RemoveBenefit = p11d32.CurrentEmployer.CurrentEmployee.RemoveBenefit(Me, benefit, BenefitIndex)
  'also remove from any employees with benefit
End Function
Private Function IBenefitForm2_UpdateBenefitListViewItem(li As MSComctlLib.IListItem, benefit As IBenefitClass, Optional ByVal BenefitIndex As Long = 0&, Optional ByVal SelectItem As Boolean = False) As Long
  Dim v As Variant
  
  On Error GoTo UpdateBenefitListViewItem_ERR
  Call xSet("UpdateBenefitListViewItem")
  
  If Not li Is Nothing And Not benefit Is Nothing Then
    If BenefitIndex > 0 Then li.Tag = BenefitIndex
    
    li.SmallIcon = benefit.ImageListKey
    li.Text = benefit.Name
    v = benefit.Calculate
    li.SubItems(1) = benefit.value(oth_HMITCode_UDBCode_db)
    
    If VarType(v) = vbString Then
      li.SubItems(2) = v
    Else
      li.SubItems(2) = FormatWN(v, "£")
    End If
    
    If SelectItem Then li.Selected = SelectItem
    IBenefitForm2_UpdateBenefitListViewItem = li.Index
  End If
  
UpdateBenefitListViewItem_END:
  Call xReturn("UpdateBenefitListViewItem")
  Exit Function
  
UpdateBenefitListViewItem_ERR:
  IBenefitForm2_UpdateBenefitListViewItem = False
  'If Err <> 35605 Then  's control has been deleted
    Call ErrorMessage(ERR_ERROR, Err, "UpdateBenefitListViewItem", "Updating a benefits list view text", "Unable to update the benefits list view text.")
  'End If
  Resume UpdateBenefitListViewItem_END
  Resume

End Function

Private Function IBenefitForm2_ValididateBenefit(ben As IBenefitClass) As Boolean
  IBenefitForm2_ValididateBenefit = IsBenOtherClass(ben.BenefitClass)
End Function

Private Function IFrmGeneral_CheckChanged(c As Control) As Boolean
  Dim bDirty As Boolean
  Dim i As Long
  Dim v As Variant
  Dim iBenITem As Long
  On Error GoTo CheckChanged_Err
  Call xSet("CheckChanged")
  Dim sUDBCode As String
  With c
    If p11d32.CurrentEmployeeIsNothing Then
      GoTo CheckChanged_End
    End If
    If benefit Is Nothing Then
      GoTo CheckChanged_End
    End If
    
    'we are asking if the value has changed and if it is valid thus save
    
    Select Case .Name
      Case "txt"
        Select Case .Index
          Case 0
            bDirty = CheckTextInput(.Text, benefit, oth_item_db)
          Case 1
            sUDBCode = benefit.value(oth_HMITCode_UDBCode_db)
            bDirty = CheckTextInput(.Text, benefit, oth_HMITCode_UDBCode_db)
            'Need to update the t_benstd table, do we need to do this here as
            'we need to know prior value before updating the table?!
            p11d32.CurrentEmployer.db.Execute (sql.Queries(UPDATE_CDB_LINKS, "CDB_" & benefit.value(oth_HMITCode_UDBCode_db), "CDB_" & sUDBCode))
          Case 2
            bDirty = CheckTextInput(.Text, benefit, oth_GrossAmountPaidByEmployer_db)
          Case 3
            bDirty = CheckTextInput(.Text, benefit, oth_madegood_db)
          Case 4
            bDirty = CheckTextInput(.Text, benefit, oth_OPRA_Ammount_Foregone_db)
          Case Else
            ECASE "Unknown control"
        End Select
      Case "cbx"
        bDirty = CheckTextInput(.Text, benefit, oth_class_db)
        benefit.BenefitClass = .ItemData(.ListIndex)
        If bDirty Then
          benefit.value(oth_category_db) = p11d32.Rates.BenClassTo(benefit.BenefitClass, BCT_DBCATEGORY)
          benefit.value(oth_NIC_Class1A_Able) = p11d32.Rates.BenClassTo(benefit.BenefitClass, BCT_CLASS1A_ABLE)
          If benefit.BenefitClass = BC_NON_CLASS_1A_M Or _
            benefit.BenefitClass = BC_OOTHER_N Or _
            benefit.BenefitClass = BC_CLASS_1A_M Then
            Call FillComboBox
            cboIRDesc2.Visible = True
            cboIRDesc2.Enabled = True
            lbl(5).Visible = True
            iBenITem = IRDescriptionBenItem(benefit.BenefitClass)
            cboIRDesc2.ComboBox.Text = S_IR_DESC_PLEASE_SELECT
          Else
            cboIRDesc2.Visible = False
            cboIRDesc2.Enabled = False
            lbl(5).Visible = False
          End If
        End If
      Case "cboIRDesc2"
        bDirty = IRDescriptionFromCombo(cboIRDesc2.ComboBox, benefit)
      Case Else
        ECASE "Unknown control"
    End Select
    
    IFrmGeneral_CheckChanged = AfterCheckChanged(c, Me, bDirty)
  
  End With
  
CheckChanged_End:
  Call xReturn("CheckChanged")
  Exit Function
  
CheckChanged_Err:
  IFrmGeneral_CheckChanged = False
  Call ErrorMessage(ERR_ERROR, Err, "CheckChanged", "ERR_CHECKCHANGED", "This function has failed for the form " & Me.Name & ".")
  Resume CheckChanged_End
  Resume
End Function

Private Property Get IFrmGeneral_InvalidVT() As Control
  Set IFrmGeneral_InvalidVT = m_InvalidVT
End Property

Private Property Set IFrmGeneral_InvalidVT(NewValue As Control)
  Set m_InvalidVT = NewValue
End Property

Private Sub lb_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
  Call SetSortOrder(lb, ColumnHeader)
End Sub

Private Sub LB_ItemClick(ByVal Item As MSComctlLib.ListItem)
  Call SetLastListItemSelected(Item)
  If Not (lb.SelectedItem Is Nothing) Then
    Call IBenefitForm2_BenefitToScreen(Item.Tag)
  End If
End Sub

Private Sub lb_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then Call SendKeys(vbTab)
End Sub






Private Sub txt_FieldInvalid(Index As Integer, Valid As Boolean, Message As String)
  Call SetPanel2(Message)
End Sub

Private Sub txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  txt(Index).Tag = SetChanged
End Sub

Private Sub txt_UserValidate(Index As Integer, Valid As Boolean, Message As String, sTextEntered As String)
  Dim i As Long
  Dim ibf As IBenefitForm2
  Dim ben As IBenefitClass
  
On Error GoTo UserValidate_ERR
  
  
Call xSet("UserValidate")
  
  If Index <> 1 Then GoTo UserValidate_END
  
  Valid = True
  
  If Len(sTextEntered) = 0 Then
    Valid = False
    Message = "Unique id can not be zero length."
    GoTo UserValidate_END:
  End If
  
  Set ibf = Me
  
  With p11d32.CurrentEmployer.CurrentEmployee
    For i = 1 To ibf.lv.listitems.Count
      If Not ibf.lv.listitems(i) Is ibf.lv.SelectedItem Then
        Set ben = .benefits(ibf.lv.listitems(i).Tag)
        If StrComp(ben.value(oth_HMITCode_UDBCode_db), sTextEntered, vbTextCompare) = 0 Then
          Valid = False
          Message = "Unique id is the same as another company defined benefits."
          GoTo UserValidate_END
        End If
      End If
    Next
  End With
  
UserValidate_END:
  Call xReturn("UserValidate")
  Exit Sub
UserValidate_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "UserValidate", "User Validate", "Error in user validate of unique id.")
  Resume UserValidate_END
End Sub

Private Sub txt_Validate(Index As Integer, Cancel As Boolean)
  Call IFrmGeneral_CheckChanged(txt(Index))
End Sub

