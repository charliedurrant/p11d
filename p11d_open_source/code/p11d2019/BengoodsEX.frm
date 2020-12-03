VERSION 5.00
Object = "{8D988532-0F0C-460C-B00E-7B5637E97680}#1.0#0"; "atc2vtext.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form F_ServicesProvided 
   Caption         =   " "
   ClientHeight    =   6060
   ClientLeft      =   2340
   ClientTop       =   4875
   ClientWidth     =   8415
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   6060
   ScaleWidth      =   8415
   WindowState     =   2  'Maximized
   Begin VB.Frame fmeInput 
      ForeColor       =   &H00FF0000&
      Height          =   5565
      Left            =   3960
      TabIndex        =   7
      Tag             =   "free,font"
      Top             =   90
      Width           =   4305
      Begin VB.CheckBox ChkBx 
         Alignment       =   1  'Right Justify
         Caption         =   "Is the amount above an amount subjected to PAYE?"
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Tag             =   "free,font"
         Top             =   2200
         Width           =   3975
      End
      Begin atc2valtext.ValText TxtBx 
         Height          =   315
         Index           =   3
         Left            =   2880
         TabIndex        =   3
         Tag             =   "free,font"
         Top             =   1800
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
         Top             =   600
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
         MaxLength       =   50
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
         Top             =   1400
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
         Top             =   1000
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
         TabIndex        =   5
         Tag             =   "free,font"
         Top             =   2800
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
         Caption         =   "OpRA amount foregone"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   12
         Tag             =   "free,font"
         Top             =   2805
         Width           =   1680
      End
      Begin VB.Label Lab 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Equivalent annual marginal costs"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   3
         Left            =   135
         TabIndex        =   10
         Tag             =   "free,font"
         Top             =   1395
         Width           =   2325
      End
      Begin VB.Label Lab 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   1
         Left            =   135
         TabIndex        =   8
         Tag             =   "free,font"
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Lab 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Actual amount made good, or amount subjected to PAYE"
         ForeColor       =   &H00800000&
         Height          =   390
         Index           =   4
         Left            =   135
         TabIndex        =   11
         Tag             =   "free,font"
         Top             =   1750
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
         Left            =   135
         TabIndex        =   9
         Tag             =   "free,font"
         Top             =   1000
         Width           =   2010
      End
   End
   Begin MSComctlLib.ListView lb 
      Height          =   5505
      Left            =   45
      TabIndex        =   6
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
Attribute VB_Name = "F_ServicesProvided"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Implements IBenefitForm2
Implements IFrmGeneral

Public benefit As IBenefitClass
Public GoodsDefault As String
Private mclsResize As New clsFormResize
Private m_InvalidVT As Control

Private Const L_DES_HEIGHT  As Long = 6090
Private Const L_DES_WIDTH  As Long = 8445

Private Sub fmeAssets_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Sub ChkBx_Click()
  Call IFrmGeneral_CheckChanged(chkbx)
End Sub

Private Sub Form_Load()
  If Not (mclsResize.InitResize(Me, L_DES_HEIGHT, L_DES_WIDTH, DESIGN, , , MDIMain)) Then
    Err.Raise ERR_Application
  End If
  Call SetupOpraInput(Lab(0), TxtBx(4))
  
End Sub
Private Sub Form_Resize()
  mclsResize.Resize
  Call ColumnWidths(LB, 50, 25)
End Sub
Private Sub IBenefitForm2_AddBenefit()
  Dim ben As IBenefitClass
  
  On Error GoTo AddBenefit_Err
  Call xSet("AddBenefit")
  
  Set ben = New ServicesProvided
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
    
    .value(ServicesProvided_MarketValue_db) = 0
    .value(ServicesProvided_item_db) = "Please enter description..."
    .value(ServicesProvided_Marginal_db) = 0
    .value(ServicesProvided_MadeGood_db) = 0
    .value(ServicesProvided_OPRA_Ammount_Foregone_db) = 0
    
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
End Function

Private Function IBenefitForm2_BenefitOn() As Boolean
  With benefit
    TxtBx(0).Text = .value(ServicesProvided_item_db)
    TxtBx(1).Text = .value(ServicesProvided_MarketValue_db)
    TxtBx(2).Text = .value(ServicesProvided_Marginal_db)
    TxtBx(3).Text = .value(ServicesProvided_MadeGood_db)
    TxtBx(4).Text = .value(ServicesProvided_OPRA_Ammount_Foregone_db)
    
    chkbx.value = BoolToChkBox(.value(ITEM_MADEGOOD_IS_TAXDEDUCTED))
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
  IBenefitForm2_benclass = BC_SERVICES_PROVIDED_K
End Property

Private Property Let IBenefitForm2_benclass(ByVal NewValue As BEN_CLASS)
End Property

'Private Property Get IBenefitForm2_ControlDefault() As Control
'  Set IBenefitForm2_ControlDefault = TxtBx(0)
'End Property

Private Property Get IBenefitForm2_lv() As MSComctlLib.IListView
  Set IBenefitForm2_lv = LB
End Property

Private Function IBenefitForm2_RemoveBenefit(ByVal BenefitIndex As Long) As Boolean
  IBenefitForm2_RemoveBenefit = p11d32.CurrentEmployer.CurrentEmployee.RemoveBenefit(Me, benefit, BenefitIndex)
End Function

Private Function IBenefitForm2_UpdateBenefitListViewItem(li As MSComctlLib.IListItem, benefit As IBenefitClass, Optional ByVal BenefitIndex As Long = 0&, Optional ByVal SelectItem As Boolean = False) As Long
  IBenefitForm2_UpdateBenefitListViewItem = UpdateBenefitListViewItem(li, benefit, BenefitIndex, SelectItem)
End Function

Private Function IBenefitForm2_ValididateBenefit(ben As IBenefitClass) As Boolean
  If ben.BenefitClass = BC_SERVICES_PROVIDED_K Then IBenefitForm2_ValididateBenefit = True
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
    
    Select Case .Name
      Case "TxtBx"
        Select Case .Index
          Case 0
            bDirty = CheckTextInput(.Text, benefit, ServicesProvided_item_db)
          Case 1
            bDirty = CheckTextInput(.Text, benefit, ServicesProvided_MarketValue_db)
          Case 2
            bDirty = CheckTextInput(.Text, benefit, ServicesProvided_Marginal_db)
          Case 3
            bDirty = CheckTextInput(.Text, benefit, ServicesProvided_MadeGood_db)
          Case 4
            bDirty = CheckTextInput(.Text, benefit, ServicesProvided_OPRA_Ammount_Foregone_db)
          Case Else
            ECASE "Unknown control index"
        End Select
      Case "ChkBx"
        bDirty = CheckCheckBoxInput(.value, benefit, ITEM_MADEGOOD_IS_TAXDEDUCTED)
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

Private Property Get IFrmGeneral_InvalidVT() As Control
  Set IFrmGeneral_InvalidVT = m_InvalidVT
End Property

Private Property Set IFrmGeneral_InvalidVT(NewValue As Control)
  Set m_InvalidVT = NewValue
End Property

Private Sub LB_ItemClick(ByVal Item As MSComctlLib.ListItem)
  Call SetLastListItemSelected(Item)
  If Not (LB.SelectedItem Is Nothing) Then
    IBenefitForm2_BenefitToScreen (Item.Tag)
  End If
End Sub

Private Sub lb_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
  Call SetSortOrder(LB, ColumnHeader)
End Sub

Private Sub LB_KeyDown(KeyCode As Integer, Shift As Integer)
  Call LVKeyDown(KeyCode, Shift)
End Sub

Private Sub lb_KeyPress(KeyAscii As Integer)
  'Check for return key - tab to next field
  If KeyAscii = 13 Then Call SendKeys(vbTab)
End Sub


Private Sub TxtBx_FieldInvalid(Index As Integer, Valid As Boolean, Message As String)
  Call SetPanel2(Message)
End Sub

Private Sub TxtBx_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  TxtBx(Index).Tag = SetChanged
End Sub


Private Sub TxtBx_Validate(Index As Integer, Cancel As Boolean)
  Call IFrmGeneral_CheckChanged(TxtBx(Index))
End Sub
