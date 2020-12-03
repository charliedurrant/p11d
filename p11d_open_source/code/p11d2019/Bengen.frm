VERSION 5.00
Object = "{8D988532-0F0C-460C-B00E-7B5637E97680}#1.0#0"; "atc2vtext.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form F_Other 
   Caption         =   " "
   ClientHeight    =   5685
   ClientLeft      =   345
   ClientTop       =   2130
   ClientWidth     =   8325
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   ScaleHeight     =   5685
   ScaleWidth      =   8325
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ListView lb 
      Height          =   5505
      Left            =   45
      TabIndex        =   0
      Tag             =   "free,font"
      Top             =   90
      Width           =   4035
      _ExtentX        =   7117
      _ExtentY        =   9710
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
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Item"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Benefit"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Frame P_NoBenefits 
      ForeColor       =   &H00FF0000&
      Height          =   5580
      Left            =   4140
      TabIndex        =   11
      Top             =   0
      Width           =   4125
      Begin VB.Frame fmeApportion 
         Caption         =   "Note: Only annualised values require apportionment."
         Height          =   1065
         Left            =   120
         TabIndex        =   17
         Tag             =   "free,font"
         Top             =   4440
         Width           =   3915
         Begin atc2valtext.ValText TB_Data 
            Height          =   285
            Index           =   4
            Left            =   2550
            TabIndex        =   10
            Tag             =   "FREE,FONT"
            Top             =   660
            Width           =   1305
            _ExtentX        =   0
            _ExtentY        =   0
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
            TypeOfData      =   2
            Maximum         =   "5/4/1999"
            Minimum         =   "6/4/1998"
            AllowEmpty      =   0   'False
         End
         Begin atc2valtext.ValText TB_Data 
            Height          =   285
            Index           =   3
            Left            =   2550
            TabIndex        =   9
            Tag             =   "FREE,FONT"
            Top             =   315
            Width           =   1305
            _ExtentX        =   0
            _ExtentY        =   0
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
            TypeOfData      =   2
            Maximum         =   "5/4/1999"
            Minimum         =   "6/4/1998"
            AllowEmpty      =   0   'False
         End
         Begin VB.Label Label7 
            Appearance      =   0  'Flat
            BackStyle       =   0  'Transparent
            Caption         =   "Available to"
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   2
            Left            =   60
            TabIndex        =   19
            Tag             =   "FREE,FONT"
            Top             =   660
            Width           =   1950
         End
         Begin VB.Label Label7 
            Appearance      =   0  'Flat
            BackStyle       =   0  'Transparent
            Caption         =   "Available from"
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   1
            Left            =   60
            TabIndex        =   18
            Tag             =   "FREE,FONT"
            Top             =   330
            Width           =   1950
         End
      End
      Begin VB.Frame fmeInput 
         BorderStyle     =   0  'None
         Height          =   3345
         Left            =   135
         TabIndex        =   21
         Top             =   270
         Width           =   3855
         Begin P11D2019.ValCombo cboIRDesc2 
            Height          =   315
            Left            =   1425
            TabIndex        =   1
            Tag             =   "free,font"
            Top             =   555
            Width           =   2430
            _ExtentX        =   4286
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
         Begin VB.CheckBox ChkBx 
            Alignment       =   1  'Right Justify
            Caption         =   "Is the amount above an amount subjected to PAYE?"
            Height          =   320
            Left            =   0
            TabIndex        =   5
            Tag             =   "free,font"
            Top             =   2115
            Width           =   3855
         End
         Begin VB.ComboBox cbo 
            Height          =   315
            Left            =   2520
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Tag             =   "FREE,FONT"
            Top             =   2880
            Width           =   1335
         End
         Begin atc2valtext.ValText TB_Data 
            Height          =   285
            Index           =   1
            Left            =   2520
            TabIndex        =   3
            Tag             =   "FREE,FONT"
            Top             =   1320
            Width           =   1305
            _ExtentX        =   0
            _ExtentY        =   0
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
            MaxLength       =   10
            Text            =   ""
            AllowEmpty      =   0   'False
            TXTAlign        =   2
         End
         Begin atc2valtext.ValText TB_Data 
            Height          =   285
            Index           =   2
            Left            =   2520
            TabIndex        =   4
            Tag             =   "FREE,FONT"
            Top             =   1725
            Width           =   1305
            _ExtentX        =   0
            _ExtentY        =   0
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
            MaxLength       =   10
            Text            =   ""
            AllowEmpty      =   0   'False
            TXTAlign        =   2
         End
         Begin atc2valtext.ValText TB_Data 
            Height          =   285
            Index           =   0
            Left            =   1440
            TabIndex        =   2
            Tag             =   "FREE,FONT"
            Top             =   960
            Width           =   2385
            _ExtentX        =   0
            _ExtentY        =   0
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
            MaxLength       =   50
            Text            =   ""
            TypeOfData      =   3
            AllowEmpty      =   0   'False
         End
         Begin atc2valtext.ValText TB_Data 
            Height          =   285
            Index           =   5
            Left            =   2520
            TabIndex        =   6
            Tag             =   "FREE,FONT"
            Top             =   2520
            Width           =   1305
            _ExtentX        =   0
            _ExtentY        =   0
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
            MaxLength       =   10
            Text            =   ""
            AllowEmpty      =   0   'False
            TXTAlign        =   2
         End
         Begin VB.Label Lab 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "OpRA amount foregone"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   2
            Left            =   0
            TabIndex        =   26
            Tag             =   "free,font"
            Top             =   2520
            Width           =   1680
         End
         Begin VB.Label Lab 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Category"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   0
            Left            =   0
            TabIndex        =   25
            Tag             =   "free,font"
            Top             =   600
            Width           =   630
         End
         Begin VB.Label Label6 
            Appearance      =   0  'Flat
            BackStyle       =   0  'Transparent
            Caption         =   "Company defined category"
            ForeColor       =   &H80000008&
            Height          =   495
            Index           =   1
            Left            =   0
            TabIndex        =   24
            Tag             =   "FREE,FONT"
            Top             =   2880
            Width           =   2295
         End
         Begin VB.Label lblFullDescription 
            BorderStyle     =   1  'Fixed Single
            Height          =   465
            Left            =   1440
            TabIndex        =   23
            Tag             =   "FREE,FONT"
            Top             =   0
            Width           =   2400
         End
         Begin VB.Label Label3 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Description"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   1
            Left            =   0
            TabIndex        =   22
            Tag             =   "FREE,FONT"
            Top             =   930
            Width           =   1095
         End
         Begin VB.Label Label7 
            Appearance      =   0  'Flat
            BackStyle       =   0  'Transparent
            Caption         =   "Gross annual amount paid by the employer "
            ForeColor       =   &H80000008&
            Height          =   510
            Index           =   0
            Left            =   0
            TabIndex        =   13
            Tag             =   "FREE,FONT"
            Top             =   1275
            Width           =   2415
         End
         Begin VB.Label Label3 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "P11D Class"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   0
            Left            =   0
            TabIndex        =   12
            Tag             =   "FREE,FONT"
            Top             =   0
            Width           =   825
         End
         Begin VB.Label Label6 
            Appearance      =   0  'Flat
            BackStyle       =   0  'Transparent
            Caption         =   "Actual amount made good, or amount subjected to PAYE"
            ForeColor       =   &H80000008&
            Height          =   495
            Index           =   0
            Left            =   0
            TabIndex        =   14
            Tag             =   "FREE,FONT"
            Top             =   1695
            Width           =   2535
         End
      End
      Begin VB.Frame fmeCDB 
         BorderStyle     =   0  'None
         Height          =   975
         Left            =   120
         TabIndex        =   20
         Top             =   3480
         Width           =   3975
         Begin VB.CommandButton cmdCopyCDB 
            Caption         =   "Copy"
            Height          =   375
            Left            =   2550
            TabIndex        =   8
            Tag             =   "FREE,FONT"
            Top             =   540
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.Label PushPullText 
            Appearance      =   0  'Flat
            BackStyle       =   0  'Transparent
            Caption         =   "Copy the benefit to the individual"
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   0
            TabIndex        =   16
            Tag             =   "FREE,FONT"
            Top             =   540
            Visible         =   0   'False
            Width           =   2535
         End
         Begin VB.Label lblCDB 
            Caption         =   "Company defined benefit"
            ForeColor       =   &H80000008&
            Height          =   270
            Left            =   45
            TabIndex        =   15
            Tag             =   "FREE,FONT"
            Top             =   240
            Visible         =   0   'False
            Width           =   2655
         End
      End
   End
End
Attribute VB_Name = "F_Other"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Implements IBenefitForm2
Implements IFrmGeneral

Public benefit As IBenefitClass

Private m_BenClass As BEN_CLASS
Private mclsResize As New clsFormResize
Private m_ValidIRDesc As Boolean 'RC 2/04 TTP#224

Private Const L_DES_HEIGHT  As Long = 6090
Private Const L_DES_WIDTH  As Long = 8445
Private m_InvalidVT As Control

Private Sub CB_Category_KeyPress(KeyAscii As Integer)
  'Check for return key - tab to next field
  If KeyAscii = 13 Then Call SendKeys(vbTab)
End Sub
Private Sub cboIRDesc2_Click()
  Call IFrmGeneral_CheckChanged(cboIRDesc2)
End Sub
Private Sub cbo_KeyDown(KeyCode As Integer, Shift As Integer)
  cbo.Tag = SetChanged
End Sub

Private Sub cbo_Validate(Cancel As Boolean)
  Call IFrmGeneral_CheckChanged(cbo)
End Sub

Private Sub CopyCDBToEmployee()
  Dim ibf As IBenefitForm2
  Dim rs As Recordset
  Dim ee As Employee
  Dim other As other
  On Error GoTo CopyCDBToEmployee_ERR
  
  Call xSet("CopyCDBToEmployee")
  
  Set ibf = Me
  If benefit Is Nothing Then Call Err.Raise(ERR_IS_NOTHING, "CopyCDBToEmployee", "The benefit for the form is nothing.")
  If benefit.LinkBen Then Call Err.Raise(ERR_IS_LINK_BEN, "CopyCDBToEmployee", "Benefit is LinkBen.")
  If Not benefit.CompanyDefined Then Call Err.Raise(ERR_IS_NOT_CDB, "CopyCDBToEmployee", "Benefit is not company defined.")
  If ibf.lv.SelectedItem Is Nothing Then Call Err.Raise(ERR_IS_NOTHING, "CopyCDBToEmployee", "The selected list item is nothing.")
  'need to execute query to remove from cdb list
  Set rs = p11d32.CurrentEmployer.rsBenTables(TBL_CDB_LINKS)
  
  Set ee = p11d32.CurrentEmployer.CurrentEmployee
  
  Set other = benefit
  Set other = other.CDBMasterBenefitLink
  Call RemoveCDBAssignment(rs, ee, other.PersonnelNumber)
  'need to copy benefit
  Set benefit = benefit.Copy(ee) 'nothing means it will not be added to ee.benefits collection
  'assign back to benefits collection
  Set benefit.Parent = ee
  Call ee.benefits.Remove(ibf.lv.SelectedItem.Tag)
  ibf.lv.SelectedItem.Tag = ee.benefits.ItemIndex(benefit)
  'need to change parent to me and set to dirty
  Call SetAvaialbleRange(benefit, ee, oth_availablefrom_db, oth_availableto_db)
  benefit.RSBookMark = ""
  benefit.CompanyDefined = False
  benefit.Dirty = True
  Call MDIMain.SetConfirmUndo
  'need to update the screen
  Call ibf.BenefitToScreen(ibf.lv.SelectedItem.Tag, True)
  Call ibf.UpdateBenefitListViewItem(ibf.lv.SelectedItem, benefit)
  
CopyCDBToEmployee_END:
  Call xReturn("CopyCDBToEmployee")
  Exit Sub
CopyCDBToEmployee_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "CopyCDBToEmployee", "Copy CDB To Employee", "Error copying cdb to employee.")
  Resume CopyCDBToEmployee_END
End Sub

Private Sub ChkBx_Click()
  Call IFrmGeneral_CheckChanged(ChkBx)
End Sub

Private Sub cmdCopyCDB_Click()
  Call CopyCDBToEmployee
End Sub

Private Sub Form_Resize()
  mclsResize.Resize
  Call ColumnWidths(lb, 75, 25)
End Sub
Private Sub Form_Load()
  If Not (mclsResize.InitResize(Me, L_DES_HEIGHT, L_DES_WIDTH, DESIGN, , , MDIMain)) Then
    Err.Raise ERR_Application
  End If
  Call SetDefaultVTDate(TB_Data(3))
  Call SetDefaultVTDate(TB_Data(4))
  Call SetupOpraInput(Lab(2), TB_Data(5))
  
End Sub
Private Function IBenefitForm2_AddBenefitSetDefaults(ben As IBenefitClass) As Long
  With ben
    Call StandardReadData(ben)
    
    'possible add to F12 menu?
    Select Case ben.BenefitClass
      Case BC_PRIVATE_MEDICAL_I, BC_NON_CLASS_1A_M, BC_CLASS_1A_M  'RC - 20030109
        Call SetAvaialbleRange(ben, ben.Parent, oth_availablefrom_db, oth_availableto_db)
      Case Else
        .value(oth_availablefrom_db) = p11d32.Rates.value(TaxYearStart)
        .value(oth_availableto_db) = p11d32.Rates.value(TaxYearEnd)
    End Select
    .value(oth_item_db) = "Please enter description..."
    .value(oth_GrossAmountPaidByEmployer_db) = 0
    .value(oth_madegood_db) = 0
    .value(oth_OPRA_Ammount_Foregone_db) = 0
    .value(oth_class_db) = p11d32.Rates.BenClassTo(m_BenClass, BCT_DBCLASS)
    .value(oth_CompanyDefinedCategoryKey_db) = 0
    .value(oth_HMITCode_UDBCode_db) = p11d32.Rates.BenClassTo(m_BenClass, BCT_HMIT_SECTION_STRING)
    If HasIRDescription(ben.BenefitClass) And Not ben.BenefitClass = BC_CHAUFFEUR_OTHERO_N Then
      .value(IRDescriptionBenItem(ben.BenefitClass)) = S_IR_DESC_PLEASE_SELECT
    End If
  End With
End Function
Private Property Set IBenefitForm2_benefit(NewValue As IBenefitClass)
  Set benefit = NewValue
End Property
Private Property Get IBenefitForm2_benefit() As IBenefitClass
  Set IBenefitForm2_benefit = benefit
End Property
Private Function IBenefitForm2_BenefitOff() As Boolean
  TB_Data(0).Text = ""
  TB_Data(1).Text = ""
  TB_Data(2).Text = ""
  TB_Data(3).Text = ""
  TB_Data(4).Text = ""
  TB_Data(5).Text = ""
  
End Function
Private Sub OPRADisplay()
  Dim b As Boolean
  
  b = OPRAVisible()
  Lab(2).Visible = b
  TB_Data(5).Visible = b
  TB_Data(5).Enabled = b
  
  
End Sub
Private Function OPRAVisible() As Boolean
  OPRAVisible = BenCTRL.IsOpRABenefitClassDataBase(m_BenClass)
End Function
Private Function IBenefitForm2_BenefitOn() As Boolean
  Dim i As Long
  Dim other As other
  
  Set other = benefit
  
  If benefit.BenefitClass = BC_OOTHER_N And (other.Accommodation Is Nothing) Then
    If cbo.ListCount > 0 Then
      For i = 0 To cbo.ListCount - 1
        If benefit.value(oth_CompanyDefinedCategoryKey_db) = cbo.ItemData(i) Then
          cbo.ListIndex = i
        End If
      Next
    Else
      ECASE ("List count for company defined category where benclass is O Other should be atleast 1.")
    End If
  End If
  
  Call SetClassCaption(benefit)
  TB_Data(0).Text = benefit.value(oth_item_db)
  TB_Data(1).Text = benefit.value(oth_GrossAmountPaidByEmployer_db)
  TB_Data(2).Text = benefit.value(oth_madegood_db)
  TB_Data(3).Text = DateValReadToScreen(benefit.value(oth_availablefrom_db))
  TB_Data(4).Text = DateValReadToScreen(benefit.value(oth_availableto_db))
  If (OPRAVisible()) Then
    TB_Data(5).Text = benefit.value(oth_OPRA_Ammount_Foregone_db)
  Else
    TB_Data(5).Text = "0"
  End If
  
  If IRDescriptionSelectorIsAvailable Then
    Call IRDescriptionToCombo(cboIRDesc2.ComboBox, benefit)
  End If
  ChkBx = BoolToChkBox(benefit.value(ITEM_MADEGOOD_IS_TAXDEDUCTED))
  
End Function
Private Property Get IRDescriptionSelectorIsAvailable() As Boolean
  IRDescriptionSelectorIsAvailable = cboIRDesc2.ComboBox.ListCount > 0
End Property

Private Function IBenefitForm2_BenefitToListView(ben As IBenefitClass, ByVal lBenefitIndex As Long) As Long
  IBenefitForm2_BenefitToListView = BenefitToListView(ben, Me, lBenefitIndex)
End Function

'Private Property Get IBenefitForm2_ControlDefault() As Control
'  Set IBenefitForm2_ControlDefault = TB_Data(0)
'End Property

Private Function IBenefitForm2_UpdateBenefitListViewItem(li As MSComctlLib.IListItem, benefit As IBenefitClass, Optional ByVal BenefitIndex As Long = 0&, Optional ByVal SelectItem As Boolean = False) As Long
  Dim i As Long
  
  IBenefitForm2_UpdateBenefitListViewItem = UpdateBenefitListViewItem(li, benefit, BenefitIndex, SelectItem)
End Function
Private Function IBenefitForm2_ValididateBenefit(ben As IBenefitClass) As Boolean
  If m_BenClass = ben.BenefitClass Then IBenefitForm2_ValididateBenefit = True
End Function
Public Function IFrmGeneral_CheckChanged(c As Control) As Boolean
  Dim bDirty As Boolean
  
  On Error GoTo CheckChanged_Err
  Call xSet("CheckChanged")
  
  With c
    If p11d32.CurrentEmployeeIsNothing Then
      GoTo CheckChanged_End
    End If
    If benefit Is Nothing Then
      GoTo CheckChanged_End
    End If
    
    'we are asking if the value has changed and if it is valid thus save
    Select Case .Name
      Case "cbo"
        If cbo.ItemData(cbo.ListIndex) <> benefit.value(oth_CompanyDefinedCategoryKey_db) Then
          bDirty = True
          benefit.value(oth_CompanyDefinedCategoryKey_db) = cbo.ItemData(cbo.ListIndex)
        End If
      Case "TB_Data"
        Select Case .Index
          Case 0
            bDirty = CheckTextInput(.Text, benefit, oth_item_db)
          Case 1
            bDirty = CheckTextInput(.Text, benefit, oth_GrossAmountPaidByEmployer_db)
          Case 2
            bDirty = CheckTextInput(.Text, benefit, oth_madegood_db)
          Case 3
            bDirty = CheckTextInput(.Text, benefit, oth_availablefrom_db)
          Case 4
            bDirty = CheckTextInput(.Text, benefit, oth_availableto_db)
          Case 5
            bDirty = CheckTextInput(.Text, benefit, oth_OPRA_Ammount_Foregone_db)
          Case Else
            ECASE "Unknown control"
        End Select
      Case "ChkBx"
        bDirty = CheckCheckBoxInput(.value, benefit, ITEM_MADEGOOD_IS_TAXDEDUCTED)
      Case "cboIRDesc2"
        ' May need to do some checking here -   Cf CheckCO2figure
        bDirty = IRDescriptionFromCombo(cboIRDesc2.ComboBox, benefit)
      Case Else
        ECASE "Unknown control"
    End Select
    'must be required in all check changed
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
Public Property Get IFrmGeneral_InvalidVT() As Control
  Set IFrmGeneral_InvalidVT = m_InvalidVT
End Property
Public Property Set IFrmGeneral_InvalidVT(NewValue As Control)
  Set m_InvalidVT = NewValue
End Property

Private Sub LB_ItemClick(ByVal Item As MSComctlLib.ListItem)
  Call SetLastListItemSelected(Item)
  If Not (lb.SelectedItem Is Nothing) Then
    IBenefitForm2_BenefitToScreen (Item.Tag)
  End If
End Sub
Private Sub lb_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
  Call SetSortOrder(lb, ColumnHeader)

End Sub

Private Sub LB_KeyDown(KeyCode As Integer, Shift As Integer)
  Call LVKeyDown(KeyCode, Shift)
End Sub

Private Sub lb_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then Call SendKeys(vbTab)
End Sub

Private Sub TB_Data_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  TB_Data(Index).Tag = SetChanged
End Sub
Private Sub SetClassCaption(ben As IBenefitClass)
  'CAD error handler
  Dim sBasicCaption As String
  Dim other As other
  Dim b As Boolean
  Dim sl As StringList
  Dim i As Long
  
  sBasicCaption = p11d32.Rates.BenClassTo(m_BenClass, BCT_FORM_CAPTION) & " - " & p11d32.Rates.BenClassTo(m_BenClass, BCT_HMIT_SECTION_STRING)
  
  If Not ben Is Nothing Then
    Set other = ben
    If Not other.Accommodation Is Nothing Then
      lblFullDescription.Caption = p11d32.Rates.BenClassTo(m_BenClass, BCT_FORM_CAPTION) & " (Accommodation expenses)"
    ElseIf Not other.Loan Is Nothing Then
      lblFullDescription.Caption = p11d32.Rates.BenClassTo(m_BenClass, BCT_FORM_CAPTION) & " (Loan amount waived)"
    Else
      lblFullDescription.Caption = sBasicCaption
    End If
  Else
    lblFullDescription.Caption = sBasicCaption
  End If
  
  cboIRDesc2.Clear
  Set sl = New StringList
  
  Select Case m_BenClass
    Case BC_PAYMENTS_ON_BEFALF_B
      Call sl.Add(S_IR_DESC_B_DOMESTIC_BILLS)
      Call sl.Add(S_IR_DESC_B_ACCOUNTANCY_FEES)
      Call sl.Add(S_IR_DESC_B_PRIVATE_ED)
      Call sl.Add(S_IR_DESC_B_PRIVATE_CAR_EX)
      Call sl.Add(S_IR_DESC_B_SEASON_TICKET)
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
    For i = 1 To sl.Count
      Call cboIRDesc2.AddItem(sl.Item(i))
    Next
    Call cboIRDesc2.AddItem(S_IR_DESC_OTHER)
  End If
  cboIRDesc2.InvalidValue = S_IR_DESC_PLEASE_SELECT
  
  b = IRDescriptionSelectorIsAvailable
  
  cboIRDesc2.Visible = b
  cboIRDesc2.Enabled = b
  Lab(0).Visible = b
  
End Sub

Private Property Let IBenefitForm2_benclass(ByVal NewValue As BEN_CLASS)
  m_BenClass = NewValue
  Call SetClassCaption(Nothing)
  Call OPRADisplay
  
End Property
Private Property Get IBenefitForm2_benclass() As BEN_CLASS
  IBenefitForm2_benclass = m_BenClass
End Property
Private Sub IBenefitForm2_AddBenefit()
  Dim ben As IBenefitClass
    
  On Error GoTo AddBenefit_Err
  Call xSet("AddBenefit")
  Set ben = New other
  Call AddBenefitHelper(Me, ben)
  
AddBenefit_End:
  Set ben = Nothing
  Call xReturn("AddBenefit")
  Exit Sub
AddBenefit_Err:
  Call ErrorMessage(ERR_ERROR, Err, "AddBenefit", "ERR_ADDBENEFIT", "Error in AddBenefit function, called from the form " & Me.Name & ".")
  Resume AddBenefit_End
  Resume
End Sub
Private Function IBenefitForm2_BenefitFormState(ByVal fState As BENEFIT_FORM_STATE) As Boolean
  On Error GoTo BenefitFormState_err
  Call xSet("IBenefitForm2_EnableBenefitForm")
  
  Select Case fState
    Case FORM_ENABLED
      If m_BenClass = BC_OOTHER_N Then
        cbo.Visible = True
        cbo.Enabled = True
        Label6(1).Visible = True
      End If
      fmeCDB.Visible = False
      fmeInput.Enabled = True
      fmeApportion.Visible = True
      fmeApportion.Enabled = True
      Call MDIMain.SetDelete
      fmeApportion.Enabled = True
      cmdCopyCDB.Visible = False
    Case FORM_CDB
      Label6(1).Visible = False
      cbo.Visible = False
      cbo.Enabled = False
      fmeCDB.Visible = True
      fmeCDB.Enabled = True
      fmeInput.Enabled = False
      fmeApportion.Visible = False
      Call MDIMain.ClearDelete
      cmdCopyCDB.Visible = True
      fmeApportion.Enabled = True
    Case FORM_LINK_BEN
      Label6(1).Visible = False
      cbo.Visible = False
      cbo.Enabled = False
      fmeCDB.Visible = False
      fmeInput.Enabled = False
      fmeApportion.Visible = False
      Call MDIMain.ClearDelete
      cmdCopyCDB.Visible = False
      fmeApportion.Enabled = True
    Case FORM_DISABLED
      Set benefit = Nothing
      cbo.Enabled = False
      If m_BenClass = BC_OOTHER_N Then
        cbo.Visible = True
        Label6(1).Visible = True
      Else
        Label6(1).Visible = False
        cbo.Visible = False
      End If
      fmeInput.Enabled = False
      fmeApportion.Enabled = False
      cmdCopyCDB.Visible = False
      Call MDIMain.ClearDelete
      Call MDIMain.ClearConfirmUndo
    Case Else
      Call ECASE("Invalid Form state in Other benefit form state.")
  End Select
  
  IBenefitForm2_BenefitFormState = True
    
BenefitFormState_end:
  Call xReturn("BenefitFormState")
  Exit Function
  
BenefitFormState_err:
  IBenefitForm2_BenefitFormState = False
  Call ErrorMessage(ERR_ERROR, Err, "BenefitFormState", "BenefitFormState", "Error setting the benefit form state to the screen.")
  Resume BenefitFormState_end
  Resume
End Function
Private Function IBenefitForm2_BenefitsToListView() As Long
  Dim i As Long
  Dim benCDC As IBenefitClass
  
  On Error GoTo BenefitsToListView_err
  
  Call xSet("BenefitsToListView")
  
  If m_BenClass = BC_OOTHER_N Then
    Label6(1).Visible = True
    cbo.Visible = True
  Else
    Label6(1).Visible = False
    cbo.Visible = False
  End If
  
  IBenefitForm2_BenefitsToListView = BenefitsToListView(Me)
  
  If m_BenClass = BC_OOTHER_N Then
    cbo.Clear
    Call cbo.AddItem(S_NO_COMPANYDEFINED_CATEGORY)
    cbo.ItemData(cbo.ListCount - 1) = 0
    For i = 1 To p11d32.CurrentEmployer.CDCs.Count
      Set benCDC = p11d32.CurrentEmployer.CDCs(i)
      If Not benCDC Is Nothing Then
'MP DB
'        Call cbo.AddItem(benCDC.value(cdc_name))
'        cbo.ItemData(cbo.ListCount - 1) = benCDC.value(cdc_Key)
        Call cbo.AddItem(benCDC.value(cdc_name_db))
        cbo.ItemData(cbo.ListCount - 1) = benCDC.value(cdc_Key_db)
      End If
    Next
  End If
  
BenefitsToListView_end:
  Call xReturn("BenefitsToListView")
  Exit Function
BenefitsToListView_err:
  Call ErrorMessage(ERR_ERROR, Err, "BenefitsToListView", "Benefits To List View", "Error placing other type benefits to the listview.")
  Resume BenefitsToListView_end
  Resume
End Function
Private Function IBenefitForm2_BenefitToScreen(Optional ByVal BenefitIndex As Long = -1&, Optional ByVal UpdateBenefit As Boolean = True) As Boolean
  IBenefitForm2_BenefitToScreen = BenefitToScreenHelper(Me, BenefitIndex, UpdateBenefit)
End Function
Private Property Get IBenefitForm2_lv() As MSComctlLib.IListView
  Set IBenefitForm2_lv = Me.lb
End Property
Private Function IBenefitForm2_RemoveBenefit(ByVal BenefitIndex As Long) As Boolean
  IBenefitForm2_RemoveBenefit = p11d32.CurrentEmployer.CurrentEmployee.RemoveBenefit(Me, benefit, BenefitIndex)
 
End Function
Private Sub TB_Data_Validate(Index As Integer, Cancel As Boolean)
  Call IFrmGeneral_CheckChanged(TB_Data(Index))
End Sub


