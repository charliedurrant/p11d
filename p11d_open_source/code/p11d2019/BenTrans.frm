VERSION 5.00
Object = "{8D988532-0F0C-460C-B00E-7B5637E97680}#1.0#0"; "atc2vtext.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form F_AssetsTransferred 
   Caption         =   " "
   ClientHeight    =   5910
   ClientLeft      =   60
   ClientTop       =   240
   ClientWidth     =   8490
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5910
   ScaleWidth      =   8490
   WindowState     =   2  'Maximized
   Begin VB.Frame fmeInput 
      ForeColor       =   &H00FF0000&
      Height          =   5835
      Left            =   3915
      TabIndex        =   9
      Tag             =   "free,font"
      Top             =   0
      Width           =   4530
      Begin P11D2019.ValCombo cboIRDesc2 
         Height          =   315
         Left            =   1200
         TabIndex        =   16
         Tag             =   "free,font"
         Top             =   360
         Width           =   3015
         _ExtentX        =   5318
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
      Begin VB.CheckBox Op_Data 
         Alignment       =   1  'Right Justify
         Caption         =   "Is the amount above an amount subjected to PAYE?"
         DataSource      =   "DB"
         ForeColor       =   &H00800000&
         Height          =   320
         Index           =   1
         Left            =   120
         TabIndex        =   4
         Tag             =   "free,font"
         Top             =   2800
         Width           =   4050
      End
      Begin VB.CheckBox Op_Data 
         Alignment       =   1  'Right Justify
         Caption         =   "Is this a car or has this not been a benefit previously?"
         DataField       =   "IsCar"
         DataSource      =   "DB"
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Tag             =   "free,font"
         Top             =   1320
         Width           =   4050
      End
      Begin atc2valtext.ValText TxtBx 
         Height          =   315
         Index           =   2
         Left            =   2880
         TabIndex        =   3
         Tag             =   "free,font"
         Top             =   2350
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
         Top             =   840
         Width           =   2970
         _ExtentX        =   5239
         _ExtentY        =   503
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
      Begin atc2valtext.ValText TxtBx 
         Height          =   315
         Index           =   1
         Left            =   2880
         TabIndex        =   2
         Tag             =   "free,font"
         Top             =   1900
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
         Index           =   3
         Left            =   2880
         TabIndex        =   5
         Tag             =   "free,font"
         Top             =   3250
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
         Index           =   4
         Left            =   2880
         TabIndex        =   6
         Tag             =   "free,font"
         Top             =   3700
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
         Index           =   5
         Left            =   2880
         TabIndex        =   7
         Tag             =   "free,font"
         Top             =   4200
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
      Begin VB.Label Lab 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "OpRA amount forgone"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   6
         Left            =   135
         TabIndex        =   17
         Tag             =   "free,font"
         Top             =   4230
         Width           =   2355
         WordWrap        =   -1  'True
      End
      Begin VB.Label Lab 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Category"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   5
         Left            =   120
         TabIndex        =   15
         Tag             =   "free,font"
         Top             =   360
         Width           =   630
      End
      Begin VB.Label Lab 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Benefit in respect of private use prior to transfer"
         ForeColor       =   &H00800000&
         Height          =   390
         Index           =   4
         Left            =   135
         TabIndex        =   14
         Tag             =   "free,font"
         Top             =   3650
         Width           =   2355
         WordWrap        =   -1  'True
      End
      Begin VB.Label Lab 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Market value when first provided"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   13
         Tag             =   "free,font"
         Top             =   3250
         Width           =   2310
      End
      Begin VB.Label Lab 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Market value when transferred"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   12
         Tag             =   "free,font"
         Top             =   1900
         Width           =   2160
      End
      Begin VB.Label Lab 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Actual amount made good, or amount subjected to PAYE"
         ForeColor       =   &H00800000&
         Height          =   390
         Index           =   2
         Left            =   120
         TabIndex        =   11
         Tag             =   "free,font"
         Top             =   2300
         Width           =   2535
         WordWrap        =   -1  'True
      End
      Begin VB.Label Lab 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   10
         Tag             =   "free,font"
         Top             =   840
         Width           =   975
      End
   End
   Begin MSComctlLib.ListView lb 
      Height          =   5775
      Left            =   0
      TabIndex        =   8
      Tag             =   "free,font"
      Top             =   120
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   10186
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
Attribute VB_Name = "F_AssetsTransferred"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Implements IFrmGeneral
Implements IBenefitForm2

Public benefit As IBenefitClass
Private mclsResize As New clsFormResize
Private m_ValidIRDesc As Boolean 'EK 2/04 TTP#224

Private Const L_DES_HEIGHT  As Long = 6090
Private Const L_DES_WIDTH  As Long = 8445
Private m_InvalidVT As Control

Private Sub cboIRDesc2_Click()
  Call IFrmGeneral_CheckChanged(cboIRDesc2)
End Sub
Private Sub Form_Load()
  Dim i As Integer
  If Not (mclsResize.InitResize(Me, L_DES_HEIGHT, L_DES_WIDTH, DESIGN, , , MDIMain)) Then
    Err.Raise ERR_Application
  End If
  
  'EK 2/04 TTP#224
  cboIRDesc2.AddItem (S_IR_DESC_PLEASE_SELECT)
  cboIRDesc2.AddItem (S_IR_DESC_A_CARS)
  cboIRDesc2.AddItem (S_IR_DESC_A_PROPERTY)
  cboIRDesc2.AddItem (S_IR_DESC_A_PRECIOUSMETALS)
  cboIRDesc2.AddItem (S_IR_DESC_OTHER)
  
  cboIRDesc2.InvalidValue = S_IR_DESC_PLEASE_SELECT
  'cboIRDesc.AddItem (S_IR_DESC_A_MULTIPLE)
  Call SetupOpraInput(Lab(6), TxtBx(5))
End Sub

Private Sub Form_Resize()
  Call mclsResize.Resize
  Call ColumnWidths(lb, 75, 25)
End Sub
  

Private Sub IBenefitForm2_AddBenefit()
  Dim ben As IBenefitClass
  
  On Error GoTo AddBenefit_Err
  Call xSet("AddBenefit")
  
  Set ben = New AssetsTransferred
  'Put in defaults for benefit
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

Private Function IBenefitForm2_AddBenefitSetDefaults(ben As IBenefitClass) As Long
  With ben
    Call StandardReadData(ben)
    .value(trans_EmployeeReference_db) = p11d32.CurrentEmployer.CurrentEmployee.PersonnelNumber
    
    .value(trans_MarketValueatTrans_db) = 0
    .value(trans_Item_db) = "Please enter description..."
    .value(trans_MarketValueorig_db) = 0
    .value(trans_BenefitAlready_db) = 0
    .value(trans_IsCAr_db) = 0
    .value(Trans_MadeGood_db) = 0
    .value(trans_IRDesc_db) = S_IR_DESC_PLEASE_SELECT
    .value(ITEM_OPRA_AMOUNT_FOREGONE) = 0
    
    
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
  
End Function

Private Function IBenefitForm2_BenefitOn() As Boolean
  With benefit
    TxtBx(0).Text = .value(trans_Item_db)
    TxtBx(1).Text = .value(trans_MarketValueatTrans_db)
    TxtBx(2).Text = .value(Trans_MadeGood_db)
    TxtBx(3).Text = .value(trans_MarketValueorig_db)
    TxtBx(4).Text = .value(trans_BenefitAlready_db)
    TxtBx(5).Text = .value(ITEM_OPRA_AMOUNT_FOREGONE)
    Op_Data(0) = BoolToChkBox(.value(trans_IsCAr_db))
    Op_Data(1) = BoolToChkBox(.value(ITEM_MADEGOOD_IS_TAXDEDUCTED))
    Call IsCarDisplay
    
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
  IBenefitForm2_benclass = BC_ASSETSTRANSFERRED_A
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
  If ben.BenefitClass = BC_ASSETSTRANSFERRED_A Then IBenefitForm2_ValididateBenefit = True
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
    
    Select Case .Name
      Case "TxtBx"
        Select Case .Index
          Case 0
            bDirty = CheckTextInput(.Text, benefit, trans_Item_db)
          Case 1
            bDirty = CheckTextInput(.Text, benefit, trans_MarketValueatTrans_db)
          Case 2
            bDirty = CheckTextInput(.Text, benefit, Trans_MadeGood_db)
          Case 3
            bDirty = CheckTextInput(.Text, benefit, trans_MarketValueorig_db)
          Case 4
            bDirty = CheckTextInput(.Text, benefit, trans_BenefitAlready_db)
          Case 5
            bDirty = CheckTextInput(.Text, benefit, ITEM_OPRA_AMOUNT_FOREGONE)
          Case Else
            ECASE "Unknown control"
        End Select
      Case "Op_Data"
        Select Case .Index
          Case 0
            bDirty = CheckCheckBoxInput(.value, benefit, trans_IsCAr_db)
            Call IsCarDisplay
          Case 1
            bDirty = CheckCheckBoxInput(.value, benefit, ITEM_MADEGOOD_IS_TAXDEDUCTED)
          Case Else
            ECASE "Unknown control"
        End Select
      Case "cboIRDesc2"
          'EK 2/04 TTP#224
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
End Function

Private Property Get IFrmGeneral_InvalidVT() As Control
  IFrmGeneral_InvalidVT = m_InvalidVT
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
    IBenefitForm2_BenefitToScreen (Item.Tag)
  End If
End Sub

Private Sub LB_KeyDown(KeyCode As Integer, Shift As Integer)
   Call LVKeyDown(KeyCode, Shift)
End Sub

Private Sub lb_KeyPress(KeyAscii As Integer)
  'Check for return key - tab to next field
  If KeyAscii = 13 Then Call SendKeys(vbTab)
End Sub

Private Sub lboIRDesc_Click()

End Sub

Private Sub Op_Data_Click(Index As Integer)
  Op_Data(Index).Tag = SetChanged
  Call IFrmGeneral_CheckChanged(Op_Data(Index))
End Sub

Private Sub TxtBx_FieldInvalid(Index As Integer, Valid As Boolean, Message As String)
  Call SetPanel2(Message)
End Sub

Private Sub TxtBx_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  TxtBx(Index).Tag = SetChanged
End Sub

Private Function IsCarDisplay() As Boolean

  On Error GoTo IsCarDisplay_Err
  Call xSet("IsCarDisplay")
  
  If benefit.value(trans_IsCAr_db) <> 0 Then
    TxtBx(3).Visible = False
    TxtBx(4).Visible = False
    Lab(3).Visible = False
    Lab(4).Visible = False
  Else
    TxtBx(3).Visible = True
    TxtBx(4).Visible = True
    Lab(3).Visible = True
    Lab(4).Visible = True
  End If

IsCarDisplay_End:
  Call xReturn("IsCarDisplay")
  Exit Function

IsCarDisplay_Err:
  Call ErrorMessage(ERR_ERROR, Err, "IsCarDisplay", "Is Car Display", "Error changing the display as the 'asset transferred' is a car.")
  Resume IsCarDisplay_End
  Resume
End Function

Private Sub TxtBx_Validate(Index As Integer, Cancel As Boolean)
  Call IFrmGeneral_CheckChanged(TxtBx(Index))
End Sub


