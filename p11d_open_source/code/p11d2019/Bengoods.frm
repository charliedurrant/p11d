VERSION 5.00
Object = "{8D988532-0F0C-460C-B00E-7B5637E97680}#1.0#0"; "atc2vtext.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form F_AssetsAtDisposal 
   Caption         =   " "
   ClientHeight    =   6060
   ClientLeft      =   870
   ClientTop       =   2250
   ClientWidth     =   8415
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   ScaleHeight     =   6060
   ScaleWidth      =   8415
   WindowState     =   2  'Maximized
   Begin VB.Frame fmeInput 
      ForeColor       =   &H00FF0000&
      Height          =   5565
      Left            =   3960
      TabIndex        =   11
      Tag             =   "free,font"
      Top             =   90
      Width           =   4305
      Begin VB.CheckBox Chk 
         Alignment       =   1  'Right Justify
         Caption         =   "Is the amount above an amount subjected to PAYE?"
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   4
         Tag             =   "free,font"
         Top             =   2400
         Width           =   4095
      End
      Begin VB.Frame fmeAssets 
         BorderStyle     =   0  'None
         Height          =   2700
         Left            =   120
         TabIndex        =   20
         Top             =   2715
         Width           =   4110
         Begin atc2valtext.ValText TxtBx 
            Height          =   315
            Index           =   6
            Left            =   2745
            TabIndex        =   7
            Tag             =   "free,font"
            Top             =   810
            Width           =   1305
            _ExtentX        =   1323
            _ExtentY        =   476
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
         End
         Begin atc2valtext.ValText TxtBx 
            Height          =   315
            Index           =   5
            Left            =   2745
            TabIndex        =   6
            Tag             =   "free,font"
            Top             =   410
            Width           =   1305
            _ExtentX        =   1323
            _ExtentY        =   476
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
         End
         Begin atc2valtext.ValText TxtBx 
            Height          =   315
            Index           =   4
            Left            =   2745
            TabIndex        =   5
            Tag             =   "free,font"
            Top             =   10
            Width           =   1305
            _ExtentX        =   1323
            _ExtentY        =   476
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "0"
            TXTAlign        =   2
         End
         Begin atc2valtext.ValText TxtBx 
            Height          =   315
            Index           =   7
            Left            =   2745
            TabIndex        =   8
            Tag             =   "free,font"
            Top             =   1210
            Width           =   1305
            _ExtentX        =   1323
            _ExtentY        =   476
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
            AllowEmpty      =   0   'False
         End
         Begin atc2valtext.ValText TxtBx 
            Height          =   315
            Index           =   8
            Left            =   2745
            TabIndex        =   9
            Tag             =   "free,font"
            Top             =   1845
            Width           =   1305
            _ExtentX        =   1323
            _ExtentY        =   476
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
            AllowEmpty      =   0   'False
         End
         Begin VB.Label Lab 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "OpRa amount foregone"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   9
            Left            =   0
            TabIndex        =   23
            Tag             =   "free,font"
            Top             =   1890
            Width           =   1665
         End
         Begin VB.Label Lab 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Available from"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   6
            Left            =   0
            TabIndex        =   17
            Tag             =   "free,font"
            Top             =   405
            Width           =   990
         End
         Begin VB.Label Lab 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Annual rent paid by employer"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   5
            Left            =   0
            TabIndex        =   16
            Tag             =   "free,font"
            Top             =   15
            Width           =   2040
         End
         Begin VB.Label Lab 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Available to"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   7
            Left            =   0
            TabIndex        =   18
            Tag             =   "free,font"
            Top             =   810
            Width           =   825
         End
         Begin VB.Label Lab 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Date first available as a benefit"
            ForeColor       =   &H00800000&
            Height          =   240
            Index           =   8
            Left            =   0
            TabIndex        =   19
            Tag             =   "free,font"
            Top             =   1200
            Width           =   2370
         End
      End
      Begin atc2valtext.ValText TxtBx 
         Height          =   315
         Index           =   3
         Left            =   2880
         TabIndex        =   3
         Tag             =   "free,font"
         Top             =   1920
         Width           =   1305
         _ExtentX        =   1323
         _ExtentY        =   476
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "0"
         AllowEmpty      =   0   'False
         TXTAlign        =   2
      End
      Begin atc2valtext.ValText TxtBx 
         Height          =   285
         Index           =   0
         Left            =   1200
         TabIndex        =   0
         Tag             =   "free,font"
         Top             =   720
         Width           =   2970
         _ExtentX        =   1323
         _ExtentY        =   476
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
      End
      Begin atc2valtext.ValText TxtBx 
         Height          =   315
         Index           =   2
         Left            =   2880
         TabIndex        =   2
         Tag             =   "free,font"
         Top             =   1515
         Width           =   1305
         _ExtentX        =   1323
         _ExtentY        =   476
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "0"
         AllowEmpty      =   0   'False
         TXTAlign        =   2
      End
      Begin atc2valtext.ValText TxtBx 
         Height          =   315
         Index           =   1
         Left            =   2880
         TabIndex        =   1
         Tag             =   "free,font"
         Top             =   1125
         Width           =   1305
         _ExtentX        =   1323
         _ExtentY        =   476
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "0"
         AllowEmpty      =   0   'False
         TXTAlign        =   2
      End
      Begin P11D2019.ValCombo cboIRDesc2 
         Height          =   315
         Left            =   1200
         TabIndex        =   22
         Tag             =   "free,font"
         Top             =   240
         Width           =   2895
         _ExtentX        =   3201
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
      Begin VB.Label Lab 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Category"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   21
         Tag             =   "free,font"
         Top             =   345
         Width           =   630
      End
      Begin VB.Label Lab 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Equivalent annual marginal costs"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   14
         Tag             =   "free,font"
         Top             =   1515
         Width           =   2325
      End
      Begin VB.Label Lab 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   12
         Tag             =   "free,font"
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Lab 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Actual amount made good, or amount subjected to PAYE"
         ForeColor       =   &H00800000&
         Height          =   390
         Index           =   4
         Left            =   120
         TabIndex        =   15
         Tag             =   "free,font"
         Top             =   1875
         Width           =   2535
         WordWrap        =   -1  'True
      End
      Begin VB.Label Lab 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Market value when provided"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   13
         Tag             =   "free,font"
         Top             =   1125
         Width           =   2010
      End
   End
   Begin MSComctlLib.ListView lb 
      Height          =   5505
      Left            =   45
      TabIndex        =   10
      Tag             =   "free,font"
      Top             =   180
      Width           =   3855
      _ExtentX        =   6800
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
         Text            =   "Benefit Value"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "F_AssetsAtDisposal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Implements IBenefitForm2
Implements IFrmGeneral

Public benefit As IBenefitClass

Private mclsResize As New clsFormResize
Private m_InvalidVT As Control
Private m_ValidIRDesc As Boolean 'EK 2/04 TTP#224

Private Const L_DES_HEIGHT As Long = 6090
Private Const L_DES_WIDTH As Long = 8445
Private Sub cboIRDesc2_Click()
  Call IFrmGeneral_CheckChanged(cboIRDesc2)
End Sub
Private Sub chk_Click(Index As Integer)
  Chk(Index).Tag = SetChanged
  Call IFrmGeneral_CheckChanged(Chk(Index))
End Sub

Private Sub Form_Load()
  If Not (mclsResize.InitResize(Me, L_DES_HEIGHT, L_DES_WIDTH, DESIGN, , , MDIMain)) Then
    Err.Raise ERR_Application
  End If
  Call SetDefaultVTDate(TxtBx(5))
  Call SetDefaultVTDate(TxtBx(6))
  
  
  cboIRDesc2.AddItem (S_IR_DESC_PLEASE_SELECT)
  cboIRDesc2.AddItem (S_IR_DESC_L_HOLIDAYACCOM)
  cboIRDesc2.AddItem (S_IR_DESC_L_TIMESHAREACCOM)
  cboIRDesc2.AddItem (S_IR_DESC_L_AIRCRAFT)
  cboIRDesc2.AddItem (S_IR_DESC_L_BOAT)
  cboIRDesc2.AddItem (S_IR_DESC_L_CORPORATEHOSP)
  cboIRDesc2.AddItem (S_IR_DESC_OTHER)
  cboIRDesc2.InvalidValue = S_IR_DESC_PLEASE_SELECT
  Call SetupOpraInput(Lab(9), TxtBx(8))
End Sub
Private Sub Form_Resize()
  mclsResize.Resize
  Call ColumnWidths(lb, 75, 25)
End Sub
Private Sub IBenefitForm2_AddBenefit()
  Dim ben As IBenefitClass
  
  On Error GoTo AddBenefit_Err
  Call xSet("AddBenefit")
  
  Set ben = New AssetsAtDisposal
  Call AddBenefitHelper(Me, ben)
  
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
    Call StandardReadData(ben)
    
    Call SetAvaialbleRange(ben, ben.Parent, AssetsAtDisposal_availablefrom_db, AssetsAtDisposal_availableto_db)
    .value(AssetsAtDisposal_item_db) = "Please enter description..."
    .value(AssetsAtDisposal_MarketValue_db) = 0
    .value(AssetsAtDisposal_Marginal_db) = 0
    .value(AssetsAtDisposal_Rent_db) = 0
    .value(AssetsAtDisposal_MadeGood_db) = 0
    .value(AssetsAtDisposal_DateAvail_db) = p11d32.Rates.value(goodDEFAULTDATE)
    .value(AssetsAtDisposal_OPRA_Ammount_Foregone_db) = 0
    '.value(AssetsAtDisposal_ComputerRelated_db) = False
    
    .value(AssetsAtDisposal_IRDesc_db) = S_IR_DESC_PLEASE_SELECT
  End With
End Function

Private Property Set IBenefitForm2_benefit(NewValue As IBenefitClass)
  Set benefit = NewValue
End Property

Private Property Get IBenefitForm2_benefit() As IBenefitClass
  Set IBenefitForm2_benefit = benefit
End Property

Private Function IBenefitForm2_BenefitFormState(ByVal fState As BENEFIT_FORM_STATE) As Boolean
  IBenefitForm2_BenefitFormState = BenefitFormStateEx(fState, benefit, fmeInput)
  
End Function

Private Function IBenefitForm2_BenefitOff() As Boolean
    TxtBx(0).Text = ""
    TxtBx(1).Text = ""
    TxtBx(2).Text = ""
    TxtBx(3).Text = ""
    TxtBx(4).Text = ""
    TxtBx(5).Text = ""
    TxtBx(6).Text = ""
    TxtBx(7).Text = ""
    TxtBx(8).Text = ""
End Function

Private Function IBenefitForm2_BenefitOn() As Boolean
  With benefit
    TxtBx(0).Text = .value(AssetsAtDisposal_item_db)
    TxtBx(1).Text = .value(AssetsAtDisposal_MarketValue_db)
    TxtBx(2).Text = .value(AssetsAtDisposal_Marginal_db)
    TxtBx(3).Text = .value(AssetsAtDisposal_MadeGood_db)
    TxtBx(4).Text = .value(AssetsAtDisposal_Rent_db)
    TxtBx(5).Text = DateValReadToScreen(.value(AssetsAtDisposal_availablefrom_db))
    TxtBx(6).Text = DateValReadToScreen(.value(AssetsAtDisposal_availableto_db))
    TxtBx(7).Text = DateValReadToScreen(.value(AssetsAtDisposal_DateAvail_db))
    TxtBx(8).Text = .value(AssetsAtDisposal_OPRA_Ammount_Foregone_db)
    Chk(1).value = BoolToChkBox(.value(ITEM_MADEGOOD_IS_TAXDEDUCTED))
    
    Call IRDescriptionToCombo(cboIRDesc2.ComboBox, benefit)
    
    
  End With
End Function

Private Function IBenefitForm2_BenefitsToListView() As Long
  IBenefitForm2_BenefitsToListView = BenefitsToListView(Me)
End Function

Private Function IBenefitForm2_BenefitToListView(ben As IBenefitClass, ByVal lBenefitIndex As Long) As Long
  IBenefitForm2_BenefitToListView = BenefitToListView(ben, Me, lBenefitIndex)
End Function

Private Function IBenefitForm2_BenefitToScreen(Optional ByVal BenefitIndex As Long = -1&, Optional ByVal UpdateBenefit As Boolean = True) As Boolean
  IBenefitForm2_BenefitToScreen = BenefitToScreenHelper(Me, BenefitIndex, UpdateBenefit)
End Function

Private Property Get IBenefitForm2_benclass() As BEN_CLASS
  IBenefitForm2_benclass = BC_ASSETSATDISPOSAL_L
End Property

Private Property Let IBenefitForm2_benclass(ByVal NewValue As BEN_CLASS)
End Property

'Private Property Get IBenefitForm2_ControlDefault() As Control
'  Set IBenefitForm2_ControlDefault = TxtBx(0)
'End Property

Private Property Get IBenefitForm2_lv() As MSComctlLib.IListView
  Set IBenefitForm2_lv = lb
End Property

Private Function IBenefitForm2_RemoveBenefit(ByVal BenefitIndex As Long) As Boolean
  IBenefitForm2_RemoveBenefit = p11d32.CurrentEmployer.CurrentEmployee.RemoveBenefit(Me, benefit, BenefitIndex)
End Function

Private Function IBenefitForm2_UpdateBenefitListViewItem(li As MSComctlLib.IListItem, benefit As IBenefitClass, Optional ByVal BenefitIndex As Long = 0&, Optional ByVal SelectItem As Boolean = False) As Long
  IBenefitForm2_UpdateBenefitListViewItem = UpdateBenefitListViewItem(li, benefit, BenefitIndex, SelectItem)
End Function

Private Function IBenefitForm2_ValididateBenefit(ben As IBenefitClass) As Boolean
  If ben.BenefitClass = BC_ASSETSATDISPOSAL_L Then IBenefitForm2_ValididateBenefit = True
End Function

Private Function IFrmGeneral_CheckChanged(c As Control) As Boolean
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
    
    Select Case UCASE$(.Name)
      Case "TXTBX"
        Select Case .Index
          Case 0
            bDirty = CheckTextInput(.Text, benefit, AssetsAtDisposal_item_db)
          Case 1
            bDirty = CheckTextInput(.Text, benefit, AssetsAtDisposal_MarketValue_db)
          Case 2
            bDirty = CheckTextInput(.Text, benefit, AssetsAtDisposal_Marginal_db)
          Case 3
            bDirty = CheckTextInput(.Text, benefit, AssetsAtDisposal_MadeGood_db)
          Case 4
            bDirty = CheckTextInput(.Text, benefit, AssetsAtDisposal_Rent_db)
          Case 5
            bDirty = CheckTextInput(.Text, benefit, AssetsAtDisposal_availablefrom_db)
          Case 6
            bDirty = CheckTextInput(.Text, benefit, AssetsAtDisposal_availableto_db)
          Case 7
            bDirty = CheckTextInput(.Text, benefit, AssetsAtDisposal_DateAvail_db)
          Case 8
            bDirty = CheckTextInput(.Text, benefit, AssetsAtDisposal_OPRA_Ammount_Foregone_db)
          Case Else
            ECASE "Unknown control index"
        End Select
      Case "CHK"
        Select Case .Index
          'Case 0
           ' bDirty = CheckCheckBoxInput(.value, benefit, AssetsAtDisposal_ComputerRelated_db)
          Case 1
            bDirty = CheckCheckBoxInput(.value, benefit, ITEM_MADEGOOD_IS_TAXDEDUCTED)
          Case Else
            ECASE "Unknown control index"
        End Select
      Case "CBOIRDESC2"
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
  
End Function

Private Property Get IFrmGeneral_InvalidVT() As Control
  Set IFrmGeneral_InvalidVT = m_InvalidVT
End Property

Private Property Set IFrmGeneral_InvalidVT(NewValue As Control)
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
  'Check for return key - tab to next field
  If KeyAscii = 13 Then Call SendKeys(vbTab)
End Sub

Private Property Get IBenefitForm_benclass() As BEN_CLASS
  IBenefitForm_benclass = BC_ASSETSATDISPOSAL_L
End Property

Private Sub TxtBx_FieldInvalid(Index As Integer, Valid As Boolean, Message As String)
  Call SetPanel2(Message)
End Sub

Private Sub TxtBx_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  TxtBx(Index).Tag = SetChanged
End Sub


Private Sub TxtBx_Validate(Index As Integer, Cancel As Boolean)
  Call IFrmGeneral_CheckChanged(TxtBx(Index))
End Sub


