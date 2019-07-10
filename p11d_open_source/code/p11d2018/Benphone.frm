VERSION 5.00
Object = "{8D988532-0F0C-460C-B00E-7B5637E97680}#1.0#0"; "atc2vtext.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form F_Phone 
   Caption         =   " "
   ClientHeight    =   5325
   ClientLeft      =   1335
   ClientTop       =   1725
   ClientWidth     =   8250
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   ScaleHeight     =   5325
   ScaleWidth      =   8250
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ListView LB 
      Height          =   2265
      Left            =   45
      TabIndex        =   31
      Tag             =   "free,font"
      Top             =   45
      Width           =   8130
      _ExtentX        =   14340
      _ExtentY        =   3995
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
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Phone Reference"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   1
         Text            =   "Benefit"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Frame fmeInput 
      ForeColor       =   &H00FF0000&
      Height          =   2955
      Left            =   45
      TabIndex        =   7
      Top             =   2295
      Width           =   8130
      Begin VB.Frame P_Nobenefits 
         BorderStyle     =   0  'None
         Height          =   2190
         Index           =   0
         Left            =   90
         TabIndex        =   25
         Tag             =   "free,font"
         Top             =   585
         Visible         =   0   'False
         Width           =   7905
         Begin VB.CheckBox ChkBx 
            Alignment       =   1  'Right Justify
            Caption         =   "Is the amount an amount subjected to PAYE?"
            DataSource      =   "DB"
            ForeColor       =   &H00800000&
            Height          =   375
            Index           =   4
            Left            =   60
            TabIndex        =   5
            Tag             =   "free,font"
            Top             =   1350
            Width           =   3720
         End
         Begin atc2valtext.ValText TxtBx 
            Height          =   315
            Index           =   8
            Left            =   2160
            TabIndex        =   2
            Tag             =   "free,font"
            Top             =   915
            Width           =   1680
            _ExtentX        =   2963
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
            Text            =   "0"
            AllowEmpty      =   0   'False
            TXTAlign        =   2
         End
         Begin atc2valtext.ValText TxtBx 
            Height          =   315
            Index           =   6
            Left            =   2160
            TabIndex        =   1
            Tag             =   "free,font"
            Top             =   480
            Width           =   1680
            _ExtentX        =   2963
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
            Text            =   "0"
            AllowEmpty      =   0   'False
            TXTAlign        =   2
         End
         Begin atc2valtext.ValText TxtBx 
            Height          =   315
            Index           =   10
            Left            =   6210
            TabIndex        =   3
            Tag             =   "free,font"
            Top             =   495
            Width           =   1680
            _ExtentX        =   2963
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
            Text            =   "0"
            AllowEmpty      =   0   'False
            TXTAlign        =   2
         End
         Begin atc2valtext.ValText TxtBx 
            Height          =   315
            Index           =   11
            Left            =   6210
            TabIndex        =   4
            Tag             =   "free,font"
            Top             =   945
            Width           =   1680
            _ExtentX        =   2963
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
            Text            =   "0"
            AllowEmpty      =   0   'False
            TXTAlign        =   2
         End
         Begin atc2valtext.ValText TxtBx 
            Height          =   315
            Index           =   2
            Left            =   2160
            TabIndex        =   6
            Tag             =   "free,font"
            Top             =   1755
            Width           =   1680
            _ExtentX        =   2963
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
            Left            =   45
            TabIndex        =   33
            Tag             =   "free,font"
            Top             =   1800
            Width           =   1680
         End
         Begin VB.Label Lab 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Employee contribution (calls)/collected via PAYE"
            ForeColor       =   &H00800000&
            Height          =   390
            Index           =   17
            Left            =   4005
            TabIndex        =   30
            Tag             =   "free,font"
            Top             =   915
            Width           =   2085
            WordWrap        =   -1  'True
         End
         Begin VB.Label Lab 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Calls paid by employer"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   15
            Left            =   4005
            TabIndex        =   29
            Tag             =   "free,font"
            Top             =   495
            Width           =   2190
         End
         Begin VB.Label Lab 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "O - Home telephones"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   6
            Left            =   45
            TabIndex        =   28
            Tag             =   "free,font"
            Top             =   225
            Width           =   1500
         End
         Begin VB.Label Lab 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Employee contribution (rent)/collected via PAYE"
            ForeColor       =   &H00800000&
            Height          =   390
            Index           =   12
            Left            =   60
            TabIndex        =   27
            Tag             =   "free,font"
            Top             =   915
            Width           =   2175
            WordWrap        =   -1  'True
         End
         Begin VB.Label Lab 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Rent paid by employer"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   9
            Left            =   60
            TabIndex        =   26
            Tag             =   "free,font"
            Top             =   480
            Width           =   2190
         End
      End
      Begin atc2valtext.ValText TxtBx 
         Height          =   315
         Index           =   1
         Left            =   1560
         TabIndex        =   0
         Tag             =   "free,font"
         Top             =   255
         Width           =   3435
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
      Begin VB.Frame P_Nobenefits 
         BorderStyle     =   0  'None
         Height          =   2100
         Index           =   1
         Left            =   135
         TabIndex        =   9
         Tag             =   "free,font"
         Top             =   675
         Width           =   7455
         Begin VB.CheckBox ChkBx 
            Alignment       =   1  'Right Justify
            Caption         =   "Was the relevant fraction of the capital cost of the phone made good?"
            DataField       =   "CapMadeGood"
            DataSource      =   "DB"
            ForeColor       =   &H00800000&
            Height          =   630
            Index           =   3
            Left            =   3375
            TabIndex        =   20
            Tag             =   "free,font"
            Top             =   1260
            Width           =   3285
         End
         Begin VB.CheckBox ChkBx 
            Alignment       =   1  'Right Justify
            Caption         =   "Required to make good?"
            DataField       =   "MakeGood"
            DataSource      =   "DB"
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   1
            Left            =   3360
            TabIndex        =   18
            Tag             =   "free,font"
            Top             =   660
            Width           =   3285
         End
         Begin VB.CheckBox ChkBx 
            Alignment       =   1  'Right Justify
            Caption         =   "Was cost made good?"
            DataField       =   "madeGood"
            DataSource      =   "DB"
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   2
            Left            =   3360
            TabIndex        =   19
            Tag             =   "free,font"
            Top             =   1020
            Width           =   3285
         End
         Begin VB.CheckBox ChkBx 
            Alignment       =   1  'Right Justify
            Caption         =   "Available for Private use?"
            DataField       =   "PvtUse"
            DataSource      =   "DB"
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   0
            Left            =   3360
            TabIndex        =   17
            Tag             =   "free,font"
            Top             =   330
            Width           =   3285
         End
         Begin atc2valtext.ValText TxtBx 
            Height          =   315
            Index           =   9
            Left            =   1530
            TabIndex        =   16
            Tag             =   "free,font"
            Top             =   1095
            Width           =   1680
            _ExtentX        =   0
            _ExtentY        =   0
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
            Maximum         =   "365"
            AllowEmpty      =   0   'False
            TXTAlign        =   2
            Validate        =   0   'False
         End
         Begin atc2valtext.ValText TxtBx 
            Height          =   315
            Index           =   7
            Left            =   1530
            TabIndex        =   14
            Tag             =   "free,font"
            Top             =   735
            Width           =   1680
            _ExtentX        =   0
            _ExtentY        =   0
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
            Validate        =   0   'False
         End
         Begin atc2valtext.ValText TxtBx 
            Height          =   315
            Index           =   5
            Left            =   1530
            TabIndex        =   12
            Tag             =   "free,font"
            Top             =   360
            Width           =   1680
            _ExtentX        =   0
            _ExtentY        =   0
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
            Validate        =   0   'False
         End
         Begin atc2valtext.ValText TxtBx 
            Height          =   315
            Index           =   0
            Left            =   0
            TabIndex        =   32
            Tag             =   "free,font"
            Top             =   0
            Width           =   1680
            _ExtentX        =   2963
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
            Text            =   "0"
            AllowEmpty      =   0   'False
            TXTAlign        =   2
         End
         Begin VB.Label Lab 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "I - Mobile Phones"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   5
            Left            =   0
            TabIndex        =   10
            Tag             =   "free,font"
            Top             =   75
            Width           =   1500
         End
         Begin VB.Label Lab 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Unavailable"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   13
            Left            =   0
            TabIndex        =   15
            Tag             =   "free,font"
            Top             =   1080
            Width           =   1200
         End
         Begin VB.Label Lab 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Available To"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   10
            Left            =   0
            TabIndex        =   13
            Tag             =   "free,font"
            Top             =   735
            Width           =   1155
         End
         Begin VB.Label Lab 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Available From"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   7
            Left            =   0
            TabIndex        =   11
            Tag             =   "free,font"
            Top             =   375
            Width           =   1275
         End
      End
      Begin VB.Label Lab 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   8
         Tag             =   "free,font"
         Top             =   255
         Width           =   1335
      End
      Begin VB.Label Lab 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   8
         Left            =   4080
         MousePointer    =   1  'Arrow
         TabIndex        =   23
         Tag             =   "[Class=MOBILE], [Type=LABEL],[Field=PVTUSE]"
         Top             =   1080
         Width           =   45
      End
      Begin VB.Label Lab 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   11
         Left            =   4080
         TabIndex        =   21
         Tag             =   "[Class=MOBILE], [Type=LABEL],[Field=MAKEGOOD]"
         Top             =   1440
         Width           =   45
      End
      Begin VB.Label Lab 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   14
         Left            =   4080
         TabIndex        =   22
         Tag             =   "[Class=MOBILE], [Type=LABEL],[Field=MADEGOOD]"
         Top             =   1800
         Width           =   45
      End
      Begin VB.Label Lab 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   16
         Left            =   4080
         TabIndex        =   24
         Tag             =   "[Class=MOBILE], [Type=LABEL],[Field=MADEGOOD]"
         Top             =   2040
         Width           =   1935
         WordWrap        =   -1  'True
      End
   End
End
Attribute VB_Name = "F_Phone"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Implements IBenefitForm2
Implements IFrmGeneral

Public benefit As IBenefitClass
Public PhoneDefault As String
Private m_BenClass As BEN_CLASS
Private mclsResize As New clsFormResize

Private m_InvalidVT As Control
Private Const L_DES_HEIGHT  As Long = 6090
Private Const L_DES_WIDTH  As Long = 8445

'MP DB removed as CB_Category control does not exist
'Private Sub CB_Category_KeyPress(KeyAscii As Integer)
'  'Check for return key - tab to next field
'  If KeyAscii = 13 Then Call SendKeys(vbTab)
'End Sub

Private Sub ChkBx_Click(Index As Integer)
  Call IFrmGeneral_CheckChanged(chkbx(Index))
End Sub

Private Sub Form_Resize()
  mclsResize.Resize
  Call ColumnWidths(LB, 75, 25)
End Sub
Private Sub Form_Load()
  If Not (mclsResize.InitResize(Me, L_DES_HEIGHT, L_DES_WIDTH, DESIGN, , , MDIMain)) Then
    Err.Raise ERR_Application
  End If
  Call SetupOpraInput(Lab(0), TxtBx(2))
'MP DB commented as these controls were part of P_Nobenefits(1) frame that was never displayed
'  Call SetDefaultVTDate(TxtBx(5))
'  Call SetDefaultVTDate(TxtBx(7))
'  TxtBx(9).Maximum = p11d32.Rates.value(DaysInYearLeap)
End Sub
Private Sub IBenefitForm2_AddBenefit()
  Dim ben As IBenefitClass
  
On Error GoTo AddBenefit_Err

  Call xSet("AddBenefit")
  
  Set ben = New phone
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
'MP DB (never displayed) - were part of P_Nobenefits(1) frame
'MP DB  Call SetAvaialbleRange(ben, ben.Parent, pho_availablefrom_db, pho_availableto_db)
    .value(Pho_CallsValue_db) = 0
    .value(Pho_CallsMadeGood_db) = 0
    .value(Pho_RentValue_db) = 0
    .value(Pho_RentMadeGood_db) = 0
    .value(pho_OPRA_Ammount_Foregone_db) = 0
'MP DB (never displayed)    .value(pho_unavailable_db) = 0
    .value(pho_item_db) = "Please enter description..."
'MP DB (never displayed)    .value(Pho_PrivateUse_db) = True
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

Private Function IBenefitForm2_BenefitsToListView() As Long
  IBenefitForm2_BenefitsToListView = BenefitsToListView(Me)
End Function

Private Function IBenefitForm2_BenefitToListView(ben As IBenefitClass, ByVal lBenefitIndex As Long) As Long
  IBenefitForm2_BenefitToListView = BenefitToListView(ben, Me, lBenefitIndex)
End Function

Private Function IBenefitForm2_BenefitToScreen(Optional ByVal BenefitIndex As Long = -1&, Optional ByVal UpdateBenefit As Boolean = True) As Boolean
  IBenefitForm2_BenefitToScreen = BenefitToScreenHelper(Me, BenefitIndex, UpdateBenefit)
End Function

Private Function IBenefitForm2_BenefitOn() As Boolean
  m_BenClass = benefit.BenefitClass
  Call HomePhoneDisplay
  With benefit
    TxtBx(1).Text = .value(pho_item_db)
'MP DB (never displayed)        TxtBx(5).Text = DateValReadToScreen(.value(pho_availablefrom_db))
'MP DB (never displayed)        TxtBx(7).Text = DateValReadToScreen(.value(pho_availableto_db))
'MP DB (never displayed)    TxtBx(9).Text = .value(pho_unavailable_db)
    TxtBx(6).Text = .value(Pho_RentValue_db)
    TxtBx(8).Text = .value(Pho_RentMadeGood_db)
    TxtBx(10).Text = .value(Pho_CallsValue_db)
    TxtBx(11).Text = .value(Pho_CallsMadeGood_db)
    TxtBx(2).Text = .value(pho_OPRA_Ammount_Foregone_db)
'MP DB (never displayed)    ChkBx(0) = IIf(.value(Pho_PrivateUse_db), vbChecked, vbUnchecked)
'MP DB (never displayed)    ChkBx(1) = IIf(.value(Pho_reqmakegood_db), vbChecked, vbUnchecked)
'MP DB (never displayed)    ChkBx(2) = IIf(.value(Pho_ActMadeGood_db), vbChecked, vbUnchecked)
'MP DB (never displayed)    ChkBx(3) = IIf(.value(Pho_CapMadeGood_db), vbChecked, vbUnchecked)
    chkbx(4) = BoolToChkBox(.value(ITEM_MADEGOOD_IS_TAXDEDUCTED))

  End With
End Function

Private Function IBenefitForm2_BenefitOff() As Boolean
  TxtBx(1).Text = ""
'MP DB (never displayed)      TxtBx(5).Text = ""
'MP DB (never displayed)      TxtBx(7).Text = ""
'MP DB (never displayed)      TxtBx(9).Text = ""
  TxtBx(6).Text = ""
  TxtBx(8).Text = ""
  TxtBx(10).Text = ""
  TxtBx(11).Text = ""
  TxtBx(2).Text = ""
  chkbx(0) = vbUnchecked
  chkbx(1) = vbUnchecked
  chkbx(2) = vbUnchecked
  chkbx(3) = vbUnchecked
End Function


Private Function HomePhoneDisplay() As Boolean

On Error GoTo HomePhoneDisplay_ERR
  
  Call xSet("HomePhoneDisplay")
    P_NoBenefits(0).Visible = True
    P_NoBenefits(0).Enabled = True
'MP DB - deleted frame & controls within
'    P_Nobenefits(1).Visible = False
'    P_Nobenefits(1).Enabled = False
    HomePhoneDisplay = True
  
HomePhoneDisplay_END:
  Call xSet("HomePhoneDisplay")
  Exit Function
HomePhoneDisplay_ERR:
  HomePhoneDisplay = False
  Call ErrorMessage(ERR_ERROR, Err, "HomePhoneDisplay", "Home phone display", "Error setting the benefit form for home telephones")
  Resume HomePhoneDisplay_END:
End Function
Private Property Let IBenefitForm2_benclass(ByVal NewValue As BEN_CLASS)
  m_BenClass = NewValue
  Call HomePhoneDisplay
End Property

Private Property Get IBenefitForm2_benclass() As BEN_CLASS
  IBenefitForm2_benclass = m_BenClass
End Property

'Private Property Get IBenefitForm2_ControlDefault() As Control
'  Set IBenefitForm2_ControlDefault = TxtBx(0)
'End Property

Private Property Get IBenefitForm2_lv() As MSComctlLib.IListView
  Set IBenefitForm2_lv = LB
End Property

Private Function IBenefitForm2_RemoveBenefit(ByVal BenefitIndex As Long) As Boolean
  IBenefitForm2_RemoveBenefit = p11d32.CurrentEmployer.CurrentEmployee.RemoveBenefit(Me, p11d32.CurrentEmployer.CurrentEmployee.benefits(BenefitIndex), BenefitIndex)
End Function

Private Function IBenefitForm2_UpdateBenefitListViewItem(li As MSComctlLib.IListItem, benefit As IBenefitClass, Optional ByVal BenefitIndex As Long = 0&, Optional ByVal SelectItem As Boolean = False) As Long
  IBenefitForm2_UpdateBenefitListViewItem = UpdateBenefitListViewItem(li, benefit, BenefitIndex, SelectItem)
End Function
Private Function IBenefitForm2_ValididateBenefit(ben As IBenefitClass) As Boolean
  If (m_BenClass = ben.BenefitClass) Then IBenefitForm2_ValididateBenefit = True
End Function

Private Function IFrmGeneral_CheckChanged(c As Control) As Boolean
  Dim lst As ListItem
  Dim i As Long
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
           Case 1
              bDirty = CheckTextInput(.Text, benefit, pho_item_db)
'MP DB (never displayed)
'           Case 5
'              bDirty = CheckTextInput(.Text, benefit, pho_availablefrom_db)
           Case 6
              bDirty = CheckTextInput(.Text, benefit, Pho_RentValue_db)
'MP DB (never displayed)
'           Case 7
'              bDirty = CheckTextInput(.Text, benefit, pho_availableto_db)
           Case 8
              bDirty = CheckTextInput(.Text, benefit, Pho_RentMadeGood_db)
'MP DB (never displayed)
'           Case 9
'              bDirty = CheckTextInput(.Text, benefit, pho_unavailable_db)
           Case 10
              bDirty = CheckTextInput(.Text, benefit, Pho_CallsValue_db)
           Case 11
              bDirty = CheckTextInput(.Text, benefit, Pho_CallsMadeGood_db)
           Case 2
              bDirty = CheckTextInput(.Text, benefit, pho_OPRA_Ammount_Foregone_db)
           Case Else
           ECASE "Unknown control"
         End Select
       Case "ChkBx"
         Select Case .Index
'MP DB (never displayed)
'          Case 0
'            bDirty = CheckCheckBoxInput(.value, benefit, Pho_PrivateUse_db)
'MP DB (never displayed)
'          Case 1
'            bDirty = CheckCheckBoxInput(.value, benefit, Pho_reqmakegood_db)
'MP DB (never displayed)
'          Case 2
'            bDirty = CheckCheckBoxInput(.value, benefit, Pho_ActMadeGood_db)
'MP DB (never displayed)
'          Case 3
'            bDirty = CheckCheckBoxInput(.value, benefit, Pho_CapMadeGood_db)
          Case 4
            bDirty = CheckCheckBoxInput(.value, benefit, ITEM_MADEGOOD_IS_TAXDEDUCTED)
          Case Else
            ECASE "Unknown index"
         End Select
       Case Else
         ECASE "Unknown"
     End Select
    
    'must be required in all check changed
    IFrmGeneral_CheckChanged = AfterCheckChanged(c, Me, bDirty)

  End With
  
CheckChanged_End:
  Set lst = Nothing
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
  If Not (LB.SelectedItem Is Nothing) Then
    Call IBenefitForm2_BenefitToScreen(Item.Tag)
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
Private Sub TxtBx_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  TxtBx(Index).Tag = SetChanged
End Sub
Private Sub TxtBx_Validate(Index As Integer, Cancel As Boolean)
  Call IFrmGeneral_CheckChanged(TxtBx(Index))
End Sub


