VERSION 5.00
Object = "{8D988532-0F0C-460C-B00E-7B5637E97680}#1.0#0"; "atc2vtext.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form F_Loan 
   Caption         =   " "
   ClientHeight    =   5715
   ClientLeft      =   1650
   ClientTop       =   2115
   ClientWidth     =   8325
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
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5715
   ScaleWidth      =   8325
   WindowState     =   2  'Maximized
   Begin VB.CheckBox ChkBx 
      Alignment       =   1  'Right Justify
      Caption         =   "Is the amount above an amount subjected to PAYE?"
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
      Height          =   405
      Index           =   3
      Left            =   4680
      TabIndex        =   9
      Tag             =   "free,font"
      Top             =   3650
      Width           =   3120
   End
   Begin MSComctlLib.ListView LB 
      Height          =   2535
      Left            =   45
      TabIndex        =   13
      Tag             =   "free,font"
      Top             =   45
      Width           =   8265
      _ExtentX        =   14579
      _ExtentY        =   4471
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
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Loan Reference"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   1
         Text            =   "Amount Waived"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "Beneficial Interest"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Frame fmeInput 
      Height          =   3105
      Left            =   45
      TabIndex        =   14
      Top             =   2580
      Width           =   8235
      Begin VB.CheckBox ChkBx 
         Alignment       =   1  'Right Justify
         Caption         =   "Did loan commence on first day of tax year"
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
         Height          =   375
         Index           =   2
         Left            =   4680
         TabIndex        =   12
         Tag             =   "free,font"
         Top             =   2400
         Width           =   3120
      End
      Begin atc2valtext.ValText TxtBx 
         Height          =   330
         Index           =   4
         Left            =   3150
         TabIndex        =   5
         Tag             =   "free,font"
         Top             =   2250
         Width           =   1095
         _ExtentX        =   1931
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
         MouseIcon       =   "Benloan.frx":0000
         Text            =   "1"
         Minimum         =   "1"
         AllowEmpty      =   0   'False
         TXTAlign        =   2
      End
      Begin atc2valtext.ValText TxtBx 
         Height          =   285
         Index           =   3
         Left            =   1575
         TabIndex        =   1
         Tag             =   "free,font"
         Top             =   630
         Width           =   2670
         _ExtentX        =   4710
         _ExtentY        =   503
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
         MouseIcon       =   "Benloan.frx":001C
         Text            =   ""
         TypeOfData      =   3
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
         Height          =   390
         Left            =   5550
         TabIndex        =   11
         Tag             =   "free,font"
         Top             =   1920
         Width           =   2265
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
         Left            =   2745
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Tag             =   "free,font"
         Top             =   1395
         Width           =   1485
      End
      Begin atc2valtext.ValText TxtBx 
         Height          =   315
         Index           =   2
         Left            =   6840
         TabIndex        =   8
         Tag             =   "free,font"
         Top             =   630
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
         MouseIcon       =   "Benloan.frx":0038
         Text            =   "0"
         Minimum         =   "0"
         AllowEmpty      =   0   'False
         TXTAlign        =   2
      End
      Begin atc2valtext.ValText TxtBx 
         Height          =   315
         Index           =   1
         Left            =   6840
         TabIndex        =   7
         Tag             =   "free,font"
         Top             =   240
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
         MouseIcon       =   "Benloan.frx":0054
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
         Left            =   4640
         TabIndex        =   10
         Tag             =   "free,font"
         Top             =   1560
         Width           =   3120
      End
      Begin VB.CheckBox ChkBx 
         Alignment       =   1  'Right Justify
         Caption         =   "Is the loan a 'Taxable cheap loan'?"
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
         Left            =   90
         TabIndex        =   4
         Tag             =   "free,font"
         Top             =   1755
         Width           =   4170
      End
      Begin atc2valtext.ValText TxtBx 
         Height          =   315
         Index           =   0
         Left            =   1575
         TabIndex        =   0
         Tag             =   "free,font"
         Top             =   240
         Width           =   2670
         _ExtentX        =   4710
         _ExtentY        =   556
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
         MouseIcon       =   "Benloan.frx":0070
         Text            =   ""
         TypeOfData      =   3
         AllowEmpty      =   0   'False
      End
      Begin atc2valtext.ValText TxtBx 
         Height          =   285
         Index           =   5
         Left            =   1575
         TabIndex        =   2
         Tag             =   "free,font"
         Top             =   1020
         Width           =   2670
         _ExtentX        =   4710
         _ExtentY        =   503
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
         MouseIcon       =   "Benloan.frx":008C
         Text            =   ""
         TypeOfData      =   3
      End
      Begin atc2valtext.ValText TxtBx 
         Height          =   330
         Index           =   6
         Left            =   3150
         TabIndex        =   6
         Tag             =   "free,font"
         Top             =   2655
         Width           =   1095
         _ExtentX        =   1931
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
         MouseIcon       =   "Benloan.frx":00A8
         Text            =   "1"
         Minimum         =   "1"
         AllowEmpty      =   0   'False
         TXTAlign        =   2
      End
      Begin VB.Label Lab 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "OpRA amount foregone"
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
         Index           =   2
         Left            =   135
         TabIndex        =   27
         Tag             =   "free,font"
         Top             =   2745
         Width           =   1680
      End
      Begin VB.Label Lab 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "A/C number"
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
         Index           =   12
         Left            =   120
         TabIndex        =   26
         Tag             =   "free,font"
         Top             =   1050
         Width           =   855
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
         Left            =   360
         TabIndex        =   19
         Tag             =   "free,font"
         Top             =   2295
         Width           =   1845
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
         TabIndex        =   18
         Tag             =   "free,font"
         Top             =   2070
         Width           =   1380
      End
      Begin VB.Label Lab 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name of lender"
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
         TabIndex        =   16
         Tag             =   "free,font"
         Top             =   660
         Width           =   1080
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
         TabIndex        =   23
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
         TabIndex        =   21
         Tag             =   "free,font"
         Top             =   580
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
         TabIndex        =   20
         Tag             =   "free,font"
         Top             =   270
         Width           =   2055
      End
      Begin VB.Label Lab 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Currency of loan"
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
         TabIndex        =   17
         Tag             =   "free,font"
         Top             =   1440
         Width           =   1155
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
         TabIndex        =   24
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
         TabIndex        =   22
         Tag             =   "[FIELD=NORMALONLY]"
         Top             =   1560
         Visible         =   0   'False
         Width           =   45
      End
      Begin VB.Label Lab 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Loan description"
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
         TabIndex        =   15
         Tag             =   "free,font"
         Top             =   270
         Width           =   1170
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
      TabIndex        =   25
      Top             =   3000
      Width           =   4665
   End
End
Attribute VB_Name = "F_Loan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Implements IBenefitForm2
Implements IFrmGeneral

Public benefit As IBenefitClass
Public m_loans As loans

Private m_InvalidVT As Control
Private m_BenClass As BEN_CLASS

Private mclsResize As New clsFormResize
Private Const L_DES_HEIGHT As Long = 6090
Private Const L_DES_WIDTH As Long = 8445

Private Sub B_Balance_Click()
  Dim lLoansBenefit As Long
  Dim lLoanBenefit As Long
  Dim lBenefitIndex As Long
  
  lLoanBenefit = LB.SelectedItem.Tag
  lLoansBenefit = p11d32.CurrentEmployer.CurrentEmployee.benefits.ItemIndex(m_loans)
  
  Call DialogToScreen(F_BalanceSheet, Nothing, 0, Me, TwoLongsToHiAndLow(lLoansBenefit, lLoanBenefit))
End Sub

Private Sub TxtBx_Data_FieldInvalid(Index As Integer, Valid As Boolean, Message As String)
  Call SetPanel2(Message)
End Sub

Private Sub CboBx_Click(Index As Integer)
  Call IFrmGeneral_CheckChanged(CboBx(Index))
End Sub

Private Sub CboBx_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  CboBx(Index).Tag = SetChanged
End Sub

Private Sub CboBx_KeyPress(Index As Integer, KeyAscii As Integer)
  'Check for return key - tab to next field
  If KeyAscii = 13 Then Call SendKeys(vbTab)
  
End Sub

Private Sub CboBx_Validate(Index As Integer, Cancel As Boolean)
  Call IFrmGeneral_CheckChanged(CboBx(Index))
End Sub

Private Sub ChkBx_Click(Index As Integer)
  Call IFrmGeneral_CheckChanged(chkbx(Index))
End Sub

Private Sub ChkBx_KeyPress(Index As Integer, KeyAscii As Integer)
  'Check for return key - tab to next field
  If KeyAscii = 13 Then Call SendKeys(vbTab)
End Sub

Private Sub IBenefitForm2_AddBenefit()
  Dim ben As IBenefitClass
  Dim benLoans As IBenefitClass
  Dim benLoan As Loan
  Dim lst As ListItem, i As Long
  Dim ibf As IBenefitForm2
  
  On Error GoTo AddBenefit_Err
  Call xSet("AddBenefit")
  
  Set m_loans = LoansAddToCurrentEmployee(p11d32.CurrentEmployer.CurrentEmployee)
  If m_loans Is Nothing Then GoTo AddBenefit_End
  
  Set benLoan = New Loan
  Set ben = benLoan
  Set ben.Parent = m_loans
  ben.BenefitClass = m_BenClass
  
  i = m_loans.Add(ben)
  Set benLoans = m_loans
  benLoans.Dirty = True
  Set ibf = Me
  Call ibf.AddBenefitSetDefaults(ben)
  
  If m_BenClass = BC_LOAN_OTHER_H Then
    ben.value(ln_LoanType_db) = S_LOANFSHORT 'Unknown
  End If
  
  ben.ReadFromDB = True
  Set lst = LB.listitems.Add(, , ben.Name)
  Call ibf.UpdateBenefitListViewItem(lst, ben, i, True)
  ben.Dirty = True
  Call SelectBenefitByListItem(ibf, lst)
  
  Call benLoan.AddNOther
  
  Call StandardReadData(benLoan.NOther)
  
AddBenefit_End:
  Set ibf = Nothing
  Set ben = Nothing
  Set benLoan = Nothing
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
    
    .value(ln_item_db) = "Please enter loan description"
    .value(ln_LoanCurrency_db) = S_LOANSTERLING
    .value(ln_amountwaived_db) = 0&
    
    .value(ln_InterestPaid_db) = 0&
    .value(ln_UseDailyCalculationOnly_db) = False
    
    .value(ln_Lender) = "Lender's name"
    .value(ln_ACNumber) = "Account number"
    .value(ln_nborrowers_db) = 1&
    .value(ln_DidLoanCommenceOnFirstDayOfTaxYear_db) = False
    .value(ln_CheapTaxableLoan_db) = True
    .value(ln_OPRA_Ammount_Foregone_db) = 0
    .value(ln_OPRA_Ammount_Foregone_Used_For_Value) = False
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
  
  chkbx(0) = vbUnchecked
  chkbx(1) = vbUnchecked
  
  
End Function

Private Function IBenefitForm2_BenefitOn() As Boolean
  Dim s As String
  
  TxtBx(0).Text = benefit.value(ln_item_db)
  TxtBx(1).Text = benefit.value(ln_amountwaived_db)
  TxtBx(2).Text = benefit.value(ln_InterestPaid_db)
  TxtBx(3).Text = benefit.value(ln_Lender)
  TxtBx(4).Text = benefit.value(ln_nborrowers_db)
  TxtBx(5).Text = benefit.value(ln_ACNumber)
  TxtBx(6).Text = benefit.value(ln_OPRA_Ammount_Foregone_db)
  
  chkbx(0) = BoolToChkBox(benefit.value(ln_CheapTaxableLoan_db))
  chkbx(1) = BoolToChkBox(benefit.value(ln_UseDailyCalculationOnly_db))
  chkbx(2) = BoolToChkBox(benefit.value(ln_DidLoanCommenceOnFirstDayOfTaxYear_db))
  chkbx(3) = BoolToChkBox(benefit.value(ITEM_MADEGOOD_IS_TAXDEDUCTED))
  CboBx(1).Text = benefit.value(ln_LoanCurrency_db)
  
End Function

Private Function IBenefitForm2_BenefitsToListView() As Long
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

  i = p11d32.CurrentEmployer.CurrentEmployee.GetLoansBenefitIndex
  
  If i > 0 Then
    Set m_loans = p11d32.CurrentEmployer.CurrentEmployee.benefits(i)
    For i = 1 To m_loans.Count
    Set ben = m_loans.Item(i)
    If Not ben Is Nothing Then
      If ben.BenefitClass = ibf.benclass Then
        IBenefitForm2_BenefitsToListView = IBenefitForm2_BenefitsToListView + ibf.BenefitToListView(ben, i)
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

Private Function IBenefitForm2_BenefitToListView(ben As IBenefitClass, ByVal lBenefitIndex As Long) As Long
  IBenefitForm2_BenefitToListView = BenefitToListView(ben, Me, lBenefitIndex)
End Function

Private Function IBenefitForm2_BenefitToScreen(Optional ByVal BenefitIndex As Long = -1&, Optional ByVal UpdateBenefit As Boolean = True) As Boolean
  Dim ibf As IBenefitForm2
  Dim ben As IBenefitClass
  Dim loans As loans
On Error GoTo BenefitToScreen_Err

  Call xSet("BenefitToScreen")
  
  Set ibf = Me
  
  TxtBx(4).Visible = True
  TxtBx(5).Visible = True
      
  Lab(12).Visible = True

  Lab(3).Visible = True
  Lab(10).Visible = True
  Lab(0).Visible = True
  TxtBx(3).Visible = True

  If ibf.benclass = BC_LOAN_OTHER_H Then
    TxtBx(4).Visible = False
    TxtBx(5).Visible = False
    TxtBx(3).Visible = False
    
    Lab(12).Visible = False
    Lab(3).Visible = False
    Lab(10).Visible = False
    Lab(0).Visible = False

  Else
    ECASE ("Unknown benefit type within loans.")
  End If

  If UpdateBenefit Then Call UpdateBenefitFromTags
  
  If BenefitIndex <> -1 Then
    Set ben = m_loans.Item(LowWord(BenefitIndex))
    If ben.BenefitClass <> ibf.benclass Then Call Err.Raise(ERR_INVALIDBENCLASS, "BenefitToScreen", "Benefit type invalid")
    Set ibf.benefit = ben
    Call ibf.BenefitOn
  Else
    Set ibf.benefit = Nothing
    Call ibf.BenefitOff
  End If
  Call SetBenefitFormState(ibf)
  
  IBenefitForm2_BenefitToScreen = True
  
BenefitToScreen_End:
  Set ibf = Nothing
  Call xReturn("BenefitToScreen")
  Exit Function
BenefitToScreen_Err:
  Call ErrorMessage(ERR_ERROR, Err, "BenefitToScreen", "Benefit To Screen", "Unable to place the chosen benefit to the screen. Benefit index = " & BenefitIndex & ".")
  Resume BenefitToScreen_End
  Resume
End Function

Private Property Let IBenefitForm2_benclass(ByVal NewValue As BEN_CLASS)
  m_BenClass = NewValue
End Property

Private Property Get IBenefitForm2_benclass() As BEN_CLASS
  IBenefitForm2_benclass = m_BenClass
End Property

Private Property Get IBenefitForm2_lv() As MSComctlLib.IListView
  Set IBenefitForm2_lv = LB
End Property

Private Function IBenefitForm2_RemoveBenefit(ByVal BenefitIndex As Long) As Boolean
  Dim Loan As Loan
  Dim benLoans As IBenefitClass
  Dim loans As loans
  
  Dim ibf As IBenefitForm2
  Dim i As Long
  On Error GoTo RemoveBenefit_ERR
  Call xSet("RemoveBenefit")
  Set Loan = benefit
  
  Set benLoans = benefit.Parent
  IBenefitForm2_RemoveBenefit = p11d32.CurrentEmployer.CurrentEmployee.RemoveBenefitWithLinks(Me, benefit, BenefitIndex, Loan.NOther)
  Call benLoans.Calculate
  Set loans = benLoans
  Call loans.RedrawFormLoanItems
RemoveBenefit_END:
  Call xReturn("RemoveBenefit")
  Exit Function
RemoveBenefit_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "RemoveBenefit", "Remove Benefit", "Error removing a loan benefit.")
  Resume RemoveBenefit_END
End Function

Private Function IBenefitForm2_UpdateBenefitListViewItem(li As MSComctlLib.IListItem, benefit As IBenefitClass, Optional ByVal BenefitIndex As Long = 0&, Optional ByVal SelectItem As Boolean = False) As Long
  
  On Error GoTo UpdateBenefitListViewItem_ERR
  Call xSet("UpdateBenefitListViewItem")
  
  If Not li Is Nothing And Not benefit Is Nothing Then
    li.SmallIcon = benefit.ImageListKey
  End If
    
  IBenefitForm2_UpdateBenefitListViewItem = UpdateBenefitListViewItemEx(li, benefit, BenefitIndex, SelectItem, False)
  
UpdateBenefitListViewItem_END:
  Call xReturn("UpdateBenefitListViewItem")
  Exit Function
  
UpdateBenefitListViewItem_ERR:
  IBenefitForm2_UpdateBenefitListViewItem = False
  'If Err <> 35605 Then li's control has been deleted
  Call ErrorMessage(ERR_ERROR, Err, "UpdateBenefitListViewItem", "Updating a benefits list view text", "Unable to update then benefits list view text.")
  Resume UpdateBenefitListViewItem_END
  Resume
End Function

Private Function IBenefitForm2_ValididateBenefit(ben As IBenefitClass) As Boolean
  If ben.BenefitClass = m_BenClass Then IBenefitForm2_ValididateBenefit = True
End Function

Public Function IFrmGeneral_CheckChanged(c As Control) As Boolean
  Dim lst As ListItem
  Dim i As Long
  Dim bDirty As Boolean
  Dim s As String
  
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
            bDirty = CheckTextInput(.Text, benefit, ln_item_db)
          Case 1
            bDirty = CheckTextInput(.Text, benefit, ln_amountwaived_db)
          Case 2
            bDirty = CheckTextInput(.Text, benefit, ln_InterestPaid_db)
          Case 3
            bDirty = CheckTextInput(.Text, benefit, ln_Lender)
          Case 4
            bDirty = CheckTextInput(.Text, benefit, ln_nborrowers_db)
          Case 5
            bDirty = CheckTextInput(.Text, benefit, ln_ACNumber)
          Case 6
            bDirty = CheckTextInput(.Text, benefit, ln_OPRA_Ammount_Foregone_db)
          Case Else
            ECASE "Unknown control"
        End Select
      Case "ChkBx"
        Select Case .Index
          Case 0
            bDirty = CheckCheckBoxInput(.value, benefit, ln_CheapTaxableLoan_db)
          Case 1
            bDirty = CheckCheckBoxInput(.value, benefit, ln_UseDailyCalculationOnly_db)
          Case 2
            bDirty = CheckCheckBoxInput(.value, benefit, ln_DidLoanCommenceOnFirstDayOfTaxYear_db)
          Case 3
            bDirty = CheckCheckBoxInput(.value, benefit, ITEM_MADEGOOD_IS_TAXDEDUCTED)
          Case Else
            ECASE "Unknown control"
        End Select
      Case "CboBx"
        Select Case .Index
          Case 0
          Case 1
            bDirty = CheckTextInput(.Text, benefit, ln_LoanCurrency_db)
            If bDirty Then benefit.value(ln_LoanCurrencyIndex) = CboBx(1).ListIndex
          Case Else
            ECASE "Unknown control"
        End Select
      Case Else
        ECASE "Unknown control"
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

Private Sub Form_Resize()
  mclsResize.Resize
  Call ColumnWidths(LB, 50, 25, 25)
End Sub

Private Sub Form_Load()
  If Not (mclsResize.InitResize(Me, L_DES_HEIGHT, L_DES_WIDTH, DESIGN, , , MDIMain)) Then
    Err.Raise ERR_Application
  End If
  
  
  With CboBx(1)
    .AddItem (S_LOANSTERLING)
    .AddItem (S_LOANFRANC)
    .AddItem (S_LOANYEN)
  End With
  Call SetupOpraInput(Lab(2), TxtBx(6))
  
  
  
End Sub
Private Function LoanTypeShortToLong(sLongLoanType As String, ByVal sShortLoanType As String) As Boolean
  If Len(sShortLoanType) Then
    Select Case UCASE$(sShortLoanType)
      Case UCASE$(S_LOANASHORT)
        sLongLoanType = S_LOANALONG
      Case UCASE$(S_LOANBSHORT)
        sLongLoanType = S_LOANBLONG
      Case UCASE$(S_LOANCSHORT)
        sLongLoanType = S_LOANCLONG
      Case UCASE$(S_LOANDSHORT)
        sLongLoanType = S_LOANDLONG
      Case UCASE$(S_LOANESHORT)
        sLongLoanType = S_LOANELONG
      Case UCASE$(S_LOANFSHORT)
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

Private Sub lb_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
  Call SetSortOrder(LB, ColumnHeader)
End Sub

Public Function UpdateBenefitListViewItemEx(li As MSComctlLib.IListItem, benefit As IBenefitClass, Optional ByVal BenefitIndex As Long = 0&, Optional ByVal SelectItem As Boolean = False, Optional NoCalculate As Boolean = False) As Long
  Dim v As Variant
  
  On Error GoTo UpdateBenefitListViewItemEx_Err
  Call xSet("UpdateBenefitListViewItemEx")

  If Not li Is Nothing And Not benefit Is Nothing Then
    If BenefitIndex > 0 Then li.Tag = BenefitIndex
    li.Text = benefit.Name
    If NoCalculate Then
      li.SubItems(2) = FormatWN(benefit.value(ln_Benefit))
    Else
      li.SubItems(2) = FormatWN(benefit.Calculate)
    End If
    v = benefit.value(ln_amountwaived_db)
    If IsNumeric(v) Then
      li.SubItems(1) = FormatWN(v)
    Else
      li.SubItems(1) = S_ERROR
    End If
    If SelectItem Then Set IBenefitForm2_lv.SelectedItem = li
    UpdateBenefitListViewItemEx = li.Index
  End If


UpdateBenefitListViewItemEx_End:
  Call xReturn("UpdateBenefitListViewItemEx")
  Exit Function
UpdateBenefitListViewItemEx_Err:
  Call ErrorMessage(ERR_ERROR, Err, "UpdateBenefitListViewItemEx", "Update Benefit List View ItemEx", "Error updating the benefit list view item.")
  Resume UpdateBenefitListViewItemEx_End
End Function

Private Sub TxtBx_Validate(Index As Integer, Cancel As Boolean)
  Call IFrmGeneral_CheckChanged(TxtBx(Index))
End Sub
