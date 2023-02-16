VERSION 5.00
Object = "{8D988532-0F0C-460C-B00E-7B5637E97680}#1.0#0"; "atc2vtext.ocx"
Begin VB.Form F_P11Db 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "P11D(b)"
   ClientHeight    =   9750
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6975
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9750
   ScaleWidth      =   6975
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      Caption         =   "Adjustments to Class 1A NICs"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   4305
      Left            =   120
      TabIndex        =   27
      Top             =   5025
      Width           =   6705
      Begin VB.Frame fraDeduections 
         Caption         =   "Add any amounts not included on which Class 1A NICs are due"
         ForeColor       =   &H00800000&
         Height          =   1860
         Left            =   165
         TabIndex        =   31
         Top             =   2325
         Width           =   6375
         Begin P11D2022.P11DbAdjustmentEditor gridDeductions 
            Height          =   1050
            Left            =   165
            TabIndex        =   35
            Top             =   660
            Width           =   5970
            _ExtentX        =   10530
            _ExtentY        =   1852
         End
         Begin atc2valtext.ValText txtDeductClass1ADescription 
            Height          =   285
            Left            =   1440
            TabIndex        =   32
            Top             =   285
            Width           =   4710
            _ExtentX        =   8308
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
            MaxLength       =   100
            MouseIcon       =   "F_P11Db.frx":0000
            Text            =   ""
            TypeOfData      =   3
         End
         Begin VB.Label Label3 
            Caption         =   "Brief description:"
            ForeColor       =   &H00800000&
            Height          =   345
            Left            =   90
            TabIndex        =   33
            Top             =   300
            Width           =   5040
         End
      End
      Begin VB.Frame fraAdditions 
         Caption         =   "Add any amounts not included on which Class 1A NICs are due"
         ForeColor       =   &H00800000&
         Height          =   1950
         Left            =   180
         TabIndex        =   28
         Top             =   330
         Width           =   6375
         Begin P11D2022.P11DbAdjustmentEditor gridAdditions 
            Height          =   1110
            Left            =   165
            TabIndex        =   34
            Top             =   660
            Width           =   5970
            _ExtentX        =   10530
            _ExtentY        =   1958
         End
         Begin atc2valtext.ValText txtAddClass1ADescription 
            Height          =   285
            Left            =   1440
            TabIndex        =   30
            Top             =   270
            Width           =   4710
            _ExtentX        =   8308
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
            MaxLength       =   100
            MouseIcon       =   "F_P11Db.frx":001C
            Text            =   ""
            TypeOfData      =   3
         End
         Begin VB.Label Label15 
            Caption         =   "Brief description:"
            ForeColor       =   &H00800000&
            Height          =   345
            Left            =   90
            TabIndex        =   29
            Top             =   300
            Width           =   5040
         End
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Employer Declaration"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   1800
      Left            =   120
      TabIndex        =   23
      Top             =   3150
      Width           =   6705
      Begin VB.CheckBox chkEerDeclSect2Chk1 
         Alignment       =   1  'Right Justify
         Caption         =   "No expenses payments or benefits of the type to be returned on form P11D have been or will be provided for the year end"
         ForeColor       =   &H00800000&
         Height          =   405
         Left            =   120
         TabIndex        =   26
         Top             =   270
         Width           =   6435
      End
      Begin VB.CheckBox chkEerDeclSect2Chk2 
         Alignment       =   1  'Right Justify
         Caption         =   "I confirm that all details of expenses payments and benefits that have to be returned on froms P11D for the year end are enclosed"
         ForeColor       =   &H00800000&
         Height          =   405
         Left            =   105
         TabIndex        =   25
         Top             =   825
         Width           =   6435
      End
      Begin VB.CheckBox chkEerDeclSect2Chk3 
         Alignment       =   1  'Right Justify
         Caption         =   "Forms P11D for the year end have been sent"
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   120
         TabIndex        =   24
         Top             =   1305
         Width           =   6435
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Inland Revenue Office details:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   3090
      Left            =   105
      TabIndex        =   4
      Top             =   60
      Width           =   6705
      Begin atc2valtext.ValText txtTaxOfficeNumber 
         Height          =   285
         Left            =   4200
         TabIndex        =   5
         Top             =   240
         Width           =   2325
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
         MaxLength       =   50
         MouseIcon       =   "F_P11Db.frx":0038
         Text            =   ""
         TypeOfData      =   3
      End
      Begin atc2valtext.ValText txtIRAddressLine1 
         Height          =   285
         Left            =   4200
         TabIndex        =   6
         Top             =   540
         Width           =   2325
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
         MaxLength       =   50
         MouseIcon       =   "F_P11Db.frx":0054
         Text            =   ""
         TypeOfData      =   3
      End
      Begin atc2valtext.ValText txtIRAddressLine2 
         Height          =   285
         Left            =   4200
         TabIndex        =   7
         Top             =   840
         Width           =   2325
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
         MaxLength       =   50
         MouseIcon       =   "F_P11Db.frx":0070
         Text            =   ""
         TypeOfData      =   3
      End
      Begin atc2valtext.ValText txtIRAddressLine3 
         Height          =   285
         Left            =   4200
         TabIndex        =   8
         Top             =   1140
         Width           =   2325
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
         MaxLength       =   50
         MouseIcon       =   "F_P11Db.frx":008C
         Text            =   ""
         TypeOfData      =   3
      End
      Begin atc2valtext.ValText txtIRAddressLine4 
         Height          =   285
         Left            =   4200
         TabIndex        =   9
         Top             =   1440
         Width           =   2325
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
         MaxLength       =   50
         MouseIcon       =   "F_P11Db.frx":00A8
         Text            =   ""
         TypeOfData      =   3
      End
      Begin atc2valtext.ValText txtIRPostCode 
         Height          =   285
         Left            =   4200
         TabIndex        =   10
         Top             =   1740
         Width           =   2325
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
         MaxLength       =   50
         MouseIcon       =   "F_P11Db.frx":00C4
         Text            =   ""
         TypeOfData      =   3
      End
      Begin atc2valtext.ValText txtIRTelephoneNumber 
         Height          =   285
         Left            =   4200
         TabIndex        =   11
         Top             =   2040
         Width           =   2325
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
         MaxLength       =   50
         MouseIcon       =   "F_P11Db.frx":00E0
         Text            =   ""
         TypeOfData      =   3
      End
      Begin atc2valtext.ValText txtIRYourReference 
         Height          =   285
         Left            =   4200
         TabIndex        =   12
         Top             =   2340
         Width           =   2325
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
         MaxLength       =   50
         MouseIcon       =   "F_P11Db.frx":00FC
         Text            =   ""
         TypeOfData      =   3
      End
      Begin atc2valtext.ValText txtIRAccountsOfficeReference 
         Height          =   285
         Left            =   4200
         TabIndex        =   13
         Top             =   2640
         Width           =   2325
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
         MaxLength       =   50
         MouseIcon       =   "F_P11Db.frx":0118
         Text            =   ""
         TypeOfData      =   3
      End
      Begin VB.Label Label12 
         Caption         =   "Tax office number"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   105
         TabIndex        =   22
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label Label8 
         Caption         =   "Address line 2"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   780
         Width           =   1695
      End
      Begin VB.Label Label7 
         Caption         =   "Address line 1"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   510
         Width           =   1935
      End
      Begin VB.Label Label11 
         Caption         =   "Postcode"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   1770
         Width           =   1815
      End
      Begin VB.Label Label10 
         Caption         =   "Address line 4"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   1455
         Width           =   1935
      End
      Begin VB.Label Label9 
         Caption         =   "Address line 3"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   1140
         Width           =   2055
      End
      Begin VB.Label Label6 
         Caption         =   "Tax office reference"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   2355
         Width           =   2175
      End
      Begin VB.Label Label13 
         Caption         =   "Telephone number"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   2040
         Width           =   1935
      End
      Begin VB.Label Label14 
         Caption         =   "Accounts office reference"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   2655
         Width           =   2055
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Default         =   -1  'True
      Height          =   280
      Left            =   5955
      TabIndex        =   1
      Top             =   9390
      Width           =   855
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Height          =   280
      Left            =   4860
      TabIndex        =   0
      Top             =   9390
      Width           =   930
   End
   Begin VB.Label Label2 
      Caption         =   "Amounts not included on which Class 1A is due"
      ForeColor       =   &H00800000&
      Height          =   495
      Left            =   360
      TabIndex        =   3
      Tag             =   "free,font"
      Top             =   2040
      Width           =   5760
   End
   Begin VB.Label Label1 
      Caption         =   "Additions"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   2880
      Width           =   1215
   End
End
Attribute VB_Name = "F_P11Db"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_Cancelled As Boolean

Private Sub ubgrdP11DbAdditions_ReadData(RowBuf As TrueDBGrid60.RowBuffer, ByVal RowBufRowIndex As Long, ObjectList As ATC2CORE.ObjectList, ByVal ObjectListIndex As Long)

End Sub

Public Property Set Additions(ByVal value As ObjectList)
  Set gridAdditions.Adjustments = value
End Property
Public Property Get Additions() As ObjectList
  Set Additions = gridAdditions.Adjustments
End Property
Public Property Set Deductions(ByVal value As ObjectList)
  Set gridDeductions.Adjustments = value
End Property
Public Property Get Deductions() As ObjectList
  Set Deductions = gridDeductions.Adjustments
End Property
Public Property Get AdditionsDescription() As String
  AdditionsDescription = txtAddClass1ADescription
End Property
Public Property Let AdditionsDescription(ByVal value As String)
  txtAddClass1ADescription.Text = value
End Property
Public Property Get DeductionsDescription() As String
  DeductionsDescription = txtDeductClass1ADescription.Text
End Property
Public Property Let DeductionsDescription(ByVal value As String)
  txtDeductClass1ADescription.Text = value
End Property
Public Property Get IRTaxOfficeNumber() As String
  IRTaxOfficeNumber = txtTaxOfficeNumber.Text
End Property
Public Property Let IRTaxOfficeNumber(ByVal value As String)
  txtTaxOfficeNumber.Text = value
End Property
Public Property Get IRAddressLine1() As String
  IRAddressLine1 = txtIRAddressLine1.Text
End Property
Public Property Let IRAddressLine1(ByVal value As String)
  txtIRAddressLine1.Text = value
End Property
Public Property Get IRAddressLine2() As String
  IRAddressLine2 = txtIRAddressLine2.Text
End Property
Public Property Let IRAddressLine2(ByVal value As String)
  txtIRAddressLine2.Text = value
End Property
Public Property Get IRAddressLine3() As String
  IRAddressLine3 = txtIRAddressLine3.Text
End Property
Public Property Let IRAddressLine3(ByVal value As String)
  txtIRAddressLine3.Text = value
End Property
Public Property Get IRAddressLine4() As String
  IRAddressLine4 = txtIRAddressLine4.Text
End Property
Public Property Let IRAddressLine4(ByVal value As String)
  txtIRAddressLine4.Text = value
End Property
Public Property Get IRPostCode() As String
  IRPostCode = txtIRPostCode.Text
End Property
Public Property Let IRPostCode(ByVal value As String)
  txtIRPostCode.Text = value
End Property
Public Property Get IRTelephoneNumber() As String
  IRTelephoneNumber = txtIRTelephoneNumber.Text
End Property
Public Property Let IRTelephoneNumber(ByVal value As String)
  txtIRTelephoneNumber.Text = value
End Property
Public Property Get IRYourReference() As String
  IRYourReference = txtIRYourReference.Text
End Property
Public Property Let IRYourReference(ByVal value As String)
  txtIRYourReference.Text = value
End Property

Public Property Get IRAccountsOfficeReference() As String
  IRAccountsOfficeReference = txtIRAccountsOfficeReference.Text
End Property
Public Property Let IRAccountsOfficeReference(ByVal value As String)
  txtIRAccountsOfficeReference.Text = value
End Property
Public Property Let AddClass1ADescription(ByVal value As String)
  txtAddClass1ADescription.Text = value
End Property
Public Property Get AddClass1ADescription() As String
  AddClass1ADescription = txtAddClass1ADescription.Text
End Property
Public Property Let DeductClass1ADescription(ByVal value As String)
  txtDeductClass1ADescription.Text = value
End Property
Public Property Get DeductClass1ADescription() As String
  DeductClass1ADescription = txtDeductClass1ADescription.Text
End Property
Public Property Let EerDeclSect2Chk1(ByVal value As Boolean)
  chkEerDeclSect2Chk1.value = BoolToChkBox(value)
End Property
Public Property Get EerDeclSect2Chk1() As Boolean
  EerDeclSect2Chk1 = ChkBoxToBool(chkEerDeclSect2Chk1)
End Property
Public Property Let EerDeclSect2Chk2(ByVal value As Boolean)
  chkEerDeclSect2Chk2.value = BoolToChkBox(value)
End Property
Public Property Get EerDeclSect2Chk2() As Boolean
  EerDeclSect2Chk2 = ChkBoxToBool(chkEerDeclSect2Chk2)
End Property
Public Property Let EerDeclSect2Chk3(ByVal value As Boolean)
  chkEerDeclSect2Chk3.value = BoolToChkBox(value)
End Property
Public Property Get EerDeclSect2Chk3() As Boolean
  EerDeclSect2Chk3 = ChkBoxToBool(chkEerDeclSect2Chk3)
End Property
Public Function ShowDialog(ByVal formParent As Form) As VbMsgBoxResult
  m_Cancelled = False
  Call p11d32.Help.ShowForm(Me, FormShowConstants.vbModal)
  ShowDialog = IIf(m_Cancelled, vbCancel, vbOK)
End Function

Private Sub cmdCancel_Click()
  m_Cancelled = True
  Me.Hide
End Sub

Private Sub cmdOk_Click()
  Me.Hide
End Sub

Private Sub Form_Load()
  fraAdditions.caption = S_P11DB_ADDITIONS_BOX_B
  fraDeduections.caption = S_P11DB_DEDUCTIONS_BOX_C
  
  
End Sub
