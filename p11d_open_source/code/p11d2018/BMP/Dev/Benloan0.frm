VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.2#0"; "comctl32.ocx"
Object = "{4582CA9E-1A45-11D2-8D2F-00C04FA9DD6F}#1.0#0"; "ATC2VTEXT.OCX"
Begin VB.Form F_BenLoan 
   Caption         =   "Loans"
   ClientHeight    =   5715
   ClientLeft      =   1650
   ClientTop       =   2115
   ClientWidth     =   8355
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5715
   ScaleWidth      =   8355
   WindowState     =   2  'Maximized
   Begin ComctlLib.ListView LB 
      Height          =   2535
      Left            =   90
      TabIndex        =   0
      Tag             =   "free,font"
      Top             =   120
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   4471
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      _Version        =   327682
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
      NumItems        =   4
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Loan Reference"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Alignment       =   1
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Amount Waived"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Alignment       =   1
         SubItemIndex    =   2
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Beneficial interest"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Alignment       =   1
         SubItemIndex    =   3
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "P/Y value"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Frame fmeLoans 
      Height          =   2835
      Left            =   90
      TabIndex        =   1
      Top             =   2760
      Width           =   8145
      Begin VB.CheckBox ChkBx 
         Alignment       =   1  'Right Justify
         Caption         =   "Is this within MIRAS?"
         DataField       =   "MIRAS"
         DataSource      =   "DB"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   4
         Left            =   120
         TabIndex        =   9
         Tag             =   "free,font"
         Top             =   1800
         Width           =   4125
      End
      Begin atc2valtext.ValText TxtBx 
         Height          =   285
         Index           =   4
         Left            =   3120
         TabIndex        =   12
         Tag             =   "free,font"
         Top             =   2280
         Width           =   1095
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
         Minimum         =   "0"
         TXTAlign        =   2
      End
      Begin atc2valtext.ValText TxtBx 
         Height          =   285
         Index           =   3
         Left            =   1560
         TabIndex        =   5
         Tag             =   "free,font"
         Top             =   720
         Width           =   2715
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
         TypeOfData      =   3
      End
      Begin VB.ComboBox CboBx 
         DataField       =   "LType"
         DataSource      =   "DB"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         ItemData        =   "Benloan0.frx":0000
         Left            =   5535
         List            =   "Benloan0.frx":0002
         TabIndex        =   14
         Tag             =   "free,font"
         Top             =   360
         Width           =   2235
      End
      Begin VB.CommandButton B_Balance 
         Caption         =   "&Loan Balance Sheet -->"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   5595
         TabIndex        =   22
         Tag             =   "free,font"
         Top             =   2280
         Width           =   2265
      End
      Begin VB.CheckBox ChkBx 
         Alignment       =   1  'Right Justify
         Caption         =   "Disclose regardless of de minimus?"
         DataField       =   "DiscloseRegardless"
         DataSource      =   "DB"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   2
         Left            =   4680
         TabIndex        =   20
         Tag             =   "free,font"
         Top             =   1800
         Visible         =   0   'False
         Width           =   3165
      End
      Begin VB.ComboBox CboBx 
         BackColor       =   &H00FFFFFF&
         DataField       =   "Curr"
         DataSource      =   "DB"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   2760
         TabIndex        =   7
         Tag             =   "free,font"
         Top             =   1080
         Width           =   1485
      End
      Begin atc2valtext.ValText TxtBx 
         Height          =   315
         Index           =   2
         Left            =   6840
         TabIndex        =   18
         Tag             =   "free,font"
         Top             =   1110
         Width           =   945
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
         Minimum         =   "0"
         AllowEmpty      =   0   'False
         TXTAlign        =   2
      End
      Begin atc2valtext.ValText TxtBx 
         Height          =   315
         Index           =   1
         Left            =   6840
         TabIndex        =   16
         Tag             =   "free,font"
         Top             =   720
         Width           =   945
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
         Minimum         =   "0"
         AllowEmpty      =   0   'False
         TXTAlign        =   2
      End
      Begin VB.CheckBox ChkBx 
         Alignment       =   1  'Right Justify
         Caption         =   "Use daily calculation method only?"
         DataField       =   "DailyOnly"
         DataSource      =   "DB"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   1
         Left            =   4680
         TabIndex        =   19
         Tag             =   "free,font"
         Top             =   1560
         Width           =   3165
      End
      Begin VB.CheckBox ChkBx 
         Alignment       =   1  'Right Justify
         Caption         =   "Was the loan part of a relocation package?"
         DataField       =   "Relocation"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   0
         Left            =   120
         TabIndex        =   8
         Tag             =   "free,font"
         Top             =   1440
         Width           =   4125
      End
      Begin VB.CheckBox ChkBx 
         Alignment       =   1  'Right Justify
         Caption         =   "Use normal calculation method only?"
         DataField       =   "NormalOnly"
         DataSource      =   "DB"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   3
         Left            =   4680
         TabIndex        =   21
         Tag             =   "free,font"
         Top             =   2040
         Visible         =   0   'False
         Width           =   3165
      End
      Begin atc2valtext.ValText TxtBx 
         Height          =   315
         Index           =   0
         Left            =   1560
         TabIndex        =   3
         Tag             =   "free,font"
         Top             =   330
         Width           =   2715
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
         TypeOfData      =   3
      End
      Begin VB.Label Lab 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "How many borrowers?"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   10
         Left            =   480
         TabIndex        =   11
         Tag             =   "free,font"
         Top             =   2400
         Width           =   1575
      End
      Begin VB.Label Lab 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "If this is a joint loan:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   10
         Tag             =   "free,font"
         Top             =   2160
         Width           =   1380
      End
      Begin VB.Label Lab 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name of Lender"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   4
         Tag             =   "free,font"
         Top             =   720
         Width           =   1140
      End
      Begin VB.Label Lab 
         Caption         =   "Loan Type"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   2
         Left            =   4680
         TabIndex        =   13
         Tag             =   "free,font"
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Lab 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   6
         Left            =   4680
         TabIndex        =   24
         Tag             =   "[FIELD=DISCLOSE]"
         Top             =   1320
         Visible         =   0   'False
         Width           =   45
      End
      Begin VB.Label Lab 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Interest paid by employee during the year"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   390
         Index           =   9
         Left            =   4680
         TabIndex        =   17
         Tag             =   "free,font"
         Top             =   1080
         Width           =   1935
         WordWrap        =   -1  'True
      End
      Begin VB.Label Lab 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Amount waived / written off"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   8
         Left            =   4680
         TabIndex        =   15
         Tag             =   "free,font"
         Top             =   720
         Width           =   2055
      End
      Begin VB.Label Lab 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Currency of Loan"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   5
         Left            =   120
         TabIndex        =   6
         Tag             =   "free,font"
         Top             =   1080
         Width           =   2535
      End
      Begin VB.Label Lab 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   4
         Left            =   4680
         TabIndex        =   25
         Tag             =   "[FIELD=DAILYONLY]"
         Top             =   1080
         Width           =   45
      End
      Begin VB.Label Lab 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   7
         Left            =   4680
         TabIndex        =   23
         Tag             =   "[FIELD=NORMALONLY]"
         Top             =   1560
         Visible         =   0   'False
         Width           =   45
      End
      Begin VB.Label Lab 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Loan Description"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   2
         Tag             =   "free,font"
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.Label Lab 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "There are no H(D) benefits for this employee.  To create one, click the '+' button on the toolbar above."
      ForeColor       =   &H00800000&
      Height          =   525
      Index           =   11
      Left            =   1800
      TabIndex        =   26
      Top             =   3000
      Width           =   4665
   End
End
Attribute VB_Name = "F_BenLoan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Implements IBenefitForm2
Implements IFrmGeneral

Public benefit As IBenefitClass
Public m_loans As clsLoansCollection

Private m_InvalidVT As atc2valtext.ValText
Private m_bentype As benClass
Private mclsResize As New clsFormResize
Private Const L_DES_HEIGHT = 6090
Private Const L_DES_WIDTH = 8445

Private Sub B_Balance_Click()
  Dim lLoansBenefit As Long
  Dim lLoanBenefit As Long
  Dim lBenefitIndex As Long
  
  lLoanBenefit = LB.SelectedItem.Tag
  lLoansBenefit = CurrentEmployee.benefits.ItemIndex(m_loans)
  
  Call DialogToScreen(F_Balance, Nothing, 0, Me, TwoLongsToHiAndLow(lLoansBenefit, lLoanBenefit))
End Sub

Private Sub TxtBx_Data_FieldInvalid(Index As Integer, Valid As Boolean, Message As String)
  Call MDIMain.sts.SetStatus(0, Message)
End Sub

Private Sub CboBx_Click(Index As Integer)
  Call IFrmGeneral_CheckChanged(CboBx(Index), True)
End Sub

Private Sub CboBx_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  CboBx(Index).Tag = SetChanged
End Sub

Private Sub CboBx_KeyPress(Index As Integer, KeyAscii As Integer)
  'Check for return key - tab to next field
  If KeyAscii = 13 Then Call SendKeys(vbTab)
  
End Sub

Private Sub ChkBx_Click(Index As Integer)
  Call IFrmGeneral_CheckChanged(ChkBx(Index), True)
End Sub

Private Sub ChkBx_KeyPress(Index As Integer, KeyAscii As Integer)
  'Check for return key - tab to next field
  If KeyAscii = 13 Then Call SendKeys(vbTab)
End Sub

Private Sub IBenefitForm2_AddBenefit()
  Dim ben As IBenefitClass
  Dim benLoan As clsBenLoan
  Dim load As clsBenLoan
  
  Dim lst As ListItem, i As Long
  Dim ibf As IBenefitForm2
  
  On Error GoTo AddBenefit_Err
  Call xSet("AddBenefit")
  
  If m_loans Is Nothing Then
    Set m_loans = New clsLoansCollection
    Set ben = m_loans
    Set ben.Parent = CurrentEmployee

    Call CurrentEmployee.benefits.Add(m_loans)
  End If
  
  Set benLoan = New clsBenLoan
  Set ben = benLoan
  Set ben.Parent = m_loans
  
  ben.BenefitClass = m_bentype
  i = m_loans.Add(ben)
  
  Call ben.SetItem(ln_EmployeeReference, CurrentEmployee.PersonelNo)
  Call ben.SetItem(ln_item, "Please enter description")
  Call ben.SetItem(ln_LoanCurrency, S_LOANSTERLING)
  Call ben.SetItem(ln_amountwaived, 0&)
  Call ben.SetItem(ln_Relocation, False)
  Call ben.SetItem(ln_Miras, False)
  Call ben.SetItem(ln_InterestPaid, 0&)

  If m_bentype = BC_HOMELOAN Then
    'Home Loan specifics
    Call ben.SetItem(ln_LoanType, S_LOANBSHORT)  ' Home loan
    Call ben.SetItem(ln_Lender, "Lender's name")
    Call ben.SetItem(ln_ACNumber, "Account number")
    Call ben.SetItem(ln_nborrowers, 0&)
    Call ben.SetItem(ln_default, 30000&)
    benLoan.HomeLoan = True
  ElseIf m_bentype = BC_OTHERLOAN Then
    'Beneficial Loan specifics
    Call ben.SetItem(ln_LoanType, S_LOANFSHORT) 'Unknown
    Call ben.SetItem(ln_nborrowers, 0&)
    benLoan.HomeLoan = False
  End If
  
  ben.ReadFromDB = True
  
  Set lst = LB.ListItems.Add(, , ben.name)
  Set ibf = Me
  Call ibf.UpdateBenefitListViewItem(lst, ben, i, True)
  ben.Dirty = True
  Call ibf.BenefitToScreen(i)
  Call MDIMain.SetDelete
  LB.Enabled = True
  
AddBenefit_End:
  Set ibf = Nothing
  Set ben = Nothing
  Call xReturn("AddBenefit")
  Exit Sub
AddBenefit_Err:
  Call ErrorMessage(ERR_ERROR, Err, "AddBenefit", "ERR_ADDBENEFIT", "Error in AddBenefit function, called from the form " & Me.name & ".")
  Resume AddBenefit_End
  Resume
End Sub

Private Property Let IBenefitForm2_benefit(NewValue As IBenefitClass)
  Set benefit = NewValue
End Property

Private Property Get IBenefitForm2_benefit() As IBenefitClass
  Set IBenefitForm2_benefit = benefit
End Property

Private Function IBenefitForm2_BenefitFormState(ByVal fState As BENEFIT_FORM_STATE) As Boolean
   On Error GoTo BenefitFormState_err
  Call xSet("IBenefitForm2_EnableBenefitForm")
  
  If (fState = FORM_ENABLED) Or (fState = FORM_CDB) Then
    LB.Enabled = True
    fmeLoans.Enabled = True
    Call MDIMain.SetDelete
  ElseIf fState = FORM_DISABLED Then
    Set benefit = Nothing
    LB.Enabled = False
    fmeLoans.Enabled = False
    Call MDIMain.ClearDelete
    Call MDIMain.ClearConfirmUndo
  End If
  IBenefitForm2_BenefitFormState = True
    
BenefitFormState_end:
  Call xReturn("BenefitFormState")
  Exit Function
  
BenefitFormState_err:
  IBenefitForm2_BenefitFormState = False
  Call ErrorMessage(ERR_ERROR, Err, "BenefitFormState", "ERR_UNDEFINED", "Undefined error.")
  Resume BenefitFormState_end

End Function

Private Function IBenefitForm2_BenefitsToListView() As Long
  'not the default IBenefitForm2_BenefitsToListView = BenefitsToListView(Me, m_bentype)
  Dim v As Variant
  Dim i As Long
  Dim lst As ListItem
  Dim ibf As IBenefitForm2
  Dim ben As IBenefitClass
  
On Error GoTo BenefitsToListView_err
  
  Call xSet("BenefitsToListView")
  
  Set ibf = Me
  
  Call ClearForm(ibf)
  Call MDIMain.SetAdd

  i = GetLoansBenefitIndex
  If i > 0 Then
    Set m_loans = CurrentEmployee.benefits(i)
    For i = 1 To m_loans.count
    Set ben = m_loans.Item(i)
    If Not ben Is Nothing Then
      If ben.BenefitClass = m_bentype Then
        v = ben.Calculate
        Set lst = LB.ListItems.Add(, , ben.name)
        lst.Tag = i
        lst.SubItems(1) = formatworkingnumber(ben.GetItem(ln_amountwaived), "£")
        lst.SubItems(2) = formatworkingnumber(v, "£")
        IBenefitForm2_BenefitsToListView = IBenefitForm2_BenefitsToListView + 1
      End If
    End If
    Next i
  End If
  
BenefitsToListView_end:
  Set ben = Nothing
  Set lst = Nothing
  Call xReturn("BenefitsToListView")
  Exit Function
BenefitsToListView_err:
  Call ErrorMessage(ERR_ERROR, Err, "BenefitsToListView", "Adding Benefits to ListView", "Error in adding benefits to listview.")
  Resume BenefitsToListView_end
 End Function

Private Function IBenefitForm2_BenefitToScreen(Optional ByVal BenefitIndex As Long = -1&, Optional ByVal UpdateBenefit As Boolean = True) As IBenefitClass
  Dim ben As IBenefitClass
  Dim ibf As IBenefitForm2
  Dim s As String
On Error GoTo BenefitToScreen_Err

  Call xSet("BenefitToScreen")
  
  Set ibf = Me
  If UpdateBenefit Then Call UpdateBenefitFromTags
  
  If m_bentype = BC_HOMELOAN Then
      TxtBx(4).Visible = True
      ChkBx(4).Visible = True
      Lab(3).Visible = True
      Lab(10).Visible = True
      Lab(0).Visible = True
      TxtBx(3).Visible = True
      CboBx(0).Enabled = False
    ElseIf m_bentype = BC_OTHERLOAN Then
      TxtBx(4).Visible = False
      ChkBx(4).Visible = False
      Lab(3).Visible = False
      Lab(10).Visible = False
      Lab(0).Visible = False
      TxtBx(3).Visible = False
      CboBx(0).Enabled = True
    Else
      ECASE ("Unknwon benefit type within loans.")
    End If

  
  If BenefitIndex <> -1 Then
    Set ben = m_loans.Item(BenefitIndex)
    If ben.BenefitClass <> m_bentype Then Call Err.Raise(ERR_INVALIDBENTYPE, "BenefitToScreen", "Benefit type invalid")
    Set IBenefitForm2_BenefitToScreen = ben
    Set benefit = ben
    TxtBx(0).Text = ben.GetItem(ln_item)
    TxtBx(1).Text = ben.GetItem(ln_amountwaived)
    TxtBx(2).Text = ben.GetItem(ln_InterestPaid)
    TxtBx(3).Text = ben.GetItem(ln_Lender)
    TxtBx(4).Text = ben.GetItem(ln_nborrowers)
    ChkBx(0).value = IIf(ben.GetItem(ln_Relocation), vbChecked, vbUnchecked)
    ChkBx(1).value = IIf(ben.GetItem(ln_DiscloseDaily), vbChecked, vbUnchecked)
    ChkBx(2).value = vbUnchecked
    ChkBx(3).value = vbUnchecked
    ChkBx(4).value = IIf(ben.GetItem(ln_Miras), vbChecked, vbUnchecked)
    
    'if left$(ben.GetItem(ln_LoanType),1
    If LoanTypeShortToLong(s, ben.GetItem(ln_LoanType)) Then
      CboBx(0).Text = s
    Else
      If m_bentype = BC_HOMELOAN Then
        CboBx(0).Text = S_LOANBLONG
      Else
        CboBx(0).Text = S_LOANFLONG
      End If
      
    End If
    CboBx(1).Text = ben.GetItem(ln_LoanCurrency)
  Else
    TxtBx(0).Text = ""
    TxtBx(1).Text = ""
    TxtBx(2).Text = ""
    TxtBx(3).Text = ""
    TxtBx(4).Text = ""
    ChkBx(0).value = vbUnchecked
    ChkBx(1).value = vbUnchecked
    ChkBx(2).value = vbUnchecked
    ChkBx(3).value = vbUnchecked
    ChkBx(4).value = vbUnchecked
    CboBx(0).Text = ""
    CboBx(1).Text = ""
  End If
  
  Call SetBenefitFormState(ibf, ben)
  
BenefitToScreen_End:
  Set ibf = Nothing
  Set ben = Nothing
  Call xReturn("BenefitToScreen")
  Exit Function

BenefitToScreen_Err:
  Call ErrorMessage(ERR_ERROR, Err, "BenefitToScreen", "ERR_UNDEFINED", "Unable to place then chosen benefit to the screen. Benefit index = " & BenefitIndex & ".")
  Resume BenefitToScreen_End
  Resume
End Function

Private Property Let IBenefitForm2_bentype(ByVal NewValue As benClass)
  m_bentype = NewValue
End Property

Private Property Get IBenefitForm2_bentype() As benClass
  IBenefitForm2_bentype = m_bentype
End Property

Private Property Get IBenefitForm2_lv() As ComctlLib.IListView
  Set IBenefitForm2_lv = LB
End Property

Private Function IBenefitForm2_RemoveBenefit(ByVal BenefitIndex As Long) As Boolean
  IBenefitForm2_RemoveBenefit = RemoveBenefit(Me, benefit, BenefitIndex)
End Function

Private Function IBenefitForm2_UpdateBenefitListViewItem(li As ComctlLib.IListItem, benefit As IBenefitClass, Optional ByVal BenefitIndex As Long = 0&, Optional ByVal SelectItem As Boolean = False) As Boolean
  Dim v As Variant
  'NOT DEFAULT IBenefitForm2_UpdateBenefitListViewItem = UpdateBenefitListViewItem(li, benefit, BenefitIndex, SelectItem)
  On Error GoTo UpdateBenefitListViewItem_ERR
  
Call xSet("UpdateBenefitListViewItem")
  
  If Not li Is Nothing And Not benefit Is Nothing Then
    If BenefitIndex > 0 Then li.Tag = BenefitIndex
    li.Text = benefit.name
    li.SubItems(2) = formatworkingnumber(benefit.Calculate, "£")
    v = benefit.GetItem(ln_amountwaived)
    If IsNumeric(v) Then
      li.SubItems(1) = formatworkingnumber(v, "£")
    Else
      li.SubItems(1) = S_ERROR
    End If
    
    If SelectItem Then li.Selected = SelectItem
    IBenefitForm2_UpdateBenefitListViewItem = True
  End If
  
UpdateBenefitListViewItem_END:
  Call xReturn("UpdateBenefitListViewItem")
  Exit Function
  
UpdateBenefitListViewItem_ERR:
  IBenefitForm2_UpdateBenefitListViewItem = False
  'If Err <> 35605 Then li's control has been deleted
  Call ErrorMessage(ERR_ERROR, Err, "UpdateBenefitListViewItem", "Updating a benefits list view text", "Unable to update then benefits list view text.")
  Resume UpdateBenefitListViewItem_END

End Function

Public Function IFrmGeneral_CheckChanged(c As Control, ByVal UpdateCurrentListItem As Boolean) As Boolean
  Dim lst As ListItem
  Dim i As Long
  Dim bDirty As Boolean
  Dim s As String
  
  On Error GoTo CheckChanged_Err
  Call xSet("CheckChanged")
  
  With c
    If CurrentEmployee Is Nothing Then
      GoTo CheckChanged_End
    End If
    If benefit Is Nothing Then
      GoTo CheckChanged_End
    End If
    'we are asking if the value has changed and if it is valid thus save
    Select Case .name
      Case "TxtBx"
        Select Case .Index
          Case 0
            bDirty = CheckTextInput(.Text, benefit, ln_item)
          Case 1
            bDirty = CheckTextInput(.Text, benefit, ln_amountwaived)
          Case 2
            bDirty = CheckTextInput(.Text, benefit, ln_InterestPaid)
          Case 3
            bDirty = CheckTextInput(.Text, benefit, ln_Lender)
          Case 4
            bDirty = CheckTextInput(.Text, benefit, ln_nborrowers)
          Case Else
            ECASE "Unknown control"
        End Select
      Case "ChkBx"
        Select Case .Index
          Case 0
            bDirty = CheckCheckBoxInput(.value, benefit, ln_Relocation)
          Case 1
            bDirty = CheckCheckBoxInput(.value, benefit, ln_DiscloseDaily)
          Case 2
          '  i = (IIf(.Value = vbChecked, True, False) <> benefit.GetItem())
          '  If i <> 0 Then Call benefit.SetItem(car_Second, IIf(.Value = vbChecked, True, False))
          Case 3
          '  i = (IIf(.Value = vbChecked, True, False) <> benefit.GetItem(car_Replaced))
          '  If i <> 0 Then Call benefit.SetItem(car_Replaced, IIf(.Value = vbChecked, True, False))
          Case 4
            bDirty = CheckCheckBoxInput(.value, benefit, ln_Miras)
          Case Else
            ECASE "Unknown control"
        End Select
      Case "CboBx"
        Select Case .Index
          Case 0
            If LoanTypeLongToShort(s, .Text) Then
              bDirty = StrComp(s, benefit.GetItem(ln_LoanType))
              If bDirty Then Call benefit.SetItem(ln_LoanType, s)
            Else
              If m_bentype = BC_HOMELOAN Then
                .Text = S_LOANBLONG
                Call benefit.SetItem(ln_LoanType, S_LOANBSHORT)
              Else
                .Text = S_LOANFLONG
                Call benefit.SetItem(ln_LoanType, S_LOANFSHORT)
              End If
              bDirty = True
            End If
          Case 1
            bDirty = CheckTextInput(.Text, benefit, ln_LoanCurrency)
          Case Else
            ECASE "Unknown control"
        End Select
      Case Else
        ECASE "Unknown control"
    End Select
    
    'must be required in all check changed
    IFrmGeneral_CheckChanged = AfterCheckChanged(c, Me, bDirty, UpdateCurrentListItem)
    
  End With
  
CheckChanged_End:
  Set lst = Nothing
  Call xReturn("CheckChanged")
  Exit Function
  
CheckChanged_Err:
  IFrmGeneral_CheckChanged = False
  Call ErrorMessage(ERR_ERROR, Err, "CheckChanged", "ERR_CHECKCHANGED", "This function has failed for the form " & Me.name & ".")
  Resume CheckChanged_End
  Resume
  
End Function


Private Property Get IFrmGeneral_InvalidVT() As atc2valtext.ValText
  Set IFrmGeneral_InvalidVT = m_InvalidVT
End Property

Private Property Set IFrmGeneral_InvalidVT(NewValue As atc2valtext.ValText)
  Set m_InvalidVT = NewValue
End Property

Private Sub LB_Click()
  If Not (LB.SelectedItem Is Nothing) Then
    LB.SelectedItem.Selected = True
  End If
End Sub

Private Sub LB_ItemClick(ByVal Item As ComctlLib.ListItem)
  Call SetLastListItemSelected(Item)
  If Not (LB.SelectedItem Is Nothing) Then
    IBenefitForm2_BenefitToScreen (Item.Tag)
  End If
End Sub

Private Sub LB_KeyPress(KeyAscii As Integer)
  'Check for return key - tab to next field
  If KeyAscii = 13 Then Call SendKeys(vbTab)
End Sub


Private Sub TxtBx_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  TxtBx(Index).Tag = SetChanged
End Sub

Private Sub TxtBx_LostFocus(Index As Integer)
  Call IFrmGeneral_CheckChanged(TxtBx(Index), True)
End Sub

Private Sub CboBx_Lostfocus(Index As Integer)
  Call IFrmGeneral_CheckChanged(CboBx(Index), True)
End Sub

Private Sub ChkBx_LostFocus(Index As Integer)
  Call IFrmGeneral_CheckChanged(ChkBx(Index), True)
End Sub

Private Sub Form_Resize()
  mclsResize.Resize
  Call ColumnWidths(LB, 50, 15, 20, 10)
End Sub





Private Sub Form_Load()
  If Not (mclsResize.InitResize(Me, L_DES_HEIGHT, L_DES_WIDTH, DESIGN, , , MDIMain)) Then
    Err.Raise ERR_Application
  End If
  
  With CboBx(0)
    .AddItem (S_LOANALONG)
    .AddItem (S_LOANCLONG)
    .AddItem (S_LOANDLONG)
    .AddItem (S_LOANELONG)
    .AddItem (S_LOANFLONG)
  End With
  
  With CboBx(1)
    .AddItem (S_LOANSTERLING)
    .AddItem (S_LOANFRANC)
    .AddItem (S_LOANYEN)
  End With
  
End Sub
Private Function LoanTypeShortToLong(sLongLoanType As String, ByVal sShortLoanType As String) As Boolean
  If Len(sShortLoanType) Then
    Select Case sShortLoanType
      Case S_LOANASHORT
        sLongLoanType = S_LOANALONG
      Case S_LOANBSHORT
        sLongLoanType = S_LOANBLONG
      Case S_LOANCSHORT
        sLongLoanType = S_LOANCLONG
      Case S_LOANDSHORT
        sLongLoanType = S_LOANDLONG
      Case S_LOANESHORT
        sLongLoanType = S_LOANELONG
      Case S_LOANFSHORT
        sLongLoanType = S_LOANFLONG
      Case Else
        Exit Function
    End Select
    LoanTypeShortToLong = True
  End If
End Function
Private Function LoanTypeLongToShort(sShortLoanType, ByVal sLongLoanType As String) As Boolean
  If Len(sLongLoanType) Then
    sShortLoanType = Left$(sLongLoanType, 1)
    LoanTypeLongToShort = True
  End If
End Function
Private Function GetLoansBenefitIndex() As Long
  Dim i As Long
  Dim ben As IBenefitClass
  
  On Error GoTo GetLoansBenefitIndex_Err
  Call xSet("GetLoansBenefitIndex")
  
  For i = 1 To CurrentEmployee.benefits.count
    Set ben = CurrentEmployee.benefits(i)
    If Not (ben Is Nothing) Then
      If ben.BenefitClass = BC_LOANCOL Then
        GetLoansBenefitIndex = i
        Exit For
      End If
    End If
  Next i
  
GetLoansBenefitIndex_End:
  Call xReturn("GetLoansBenefitIndex")
  Exit Function

GetLoansBenefitIndex_Err:
  Call ErrorMessage(ERR_ERROR, Err, "GetLoansBenefitIndex", "ERR_UNDEFINED", "Undefined error.")
  Resume GetLoansBenefitIndex_End
End Function

Private Sub LB_ColumnClick(ByVal ColumnHeader As ComctlLib.ColumnHeader)
  LB.SortKey = ColumnHeader.Index - 1
  LB.SelectedItem.EnsureVisible
End Sub


