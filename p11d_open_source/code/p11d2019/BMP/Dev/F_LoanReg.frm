VERSION 5.00
Object = "{8D988532-0F0C-460C-B00E-7B5637E97680}#1.0#0"; "ATC2VTEXT.OCX"
Begin VB.Form F_Loan_RegularPayments 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Regular Payments"
   ClientHeight    =   3720
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4080
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3720
   ScaleWidth      =   4080
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkForceEndOfMonth 
      Caption         =   "Force end of month"
      Height          =   420
      Left            =   90
      TabIndex        =   4
      Top             =   3240
      Width           =   1185
   End
   Begin atc2valtext.ValText txtAmountPerInstallment 
      Height          =   330
      Left            =   2700
      TabIndex        =   1
      Top             =   540
      Width           =   1320
      _ExtentX        =   2328
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
      TypeOfData      =   1
      AllowEmpty      =   0   'False
      TXTAlign        =   2
      AutoSelect      =   0
   End
   Begin VB.Frame fraRepaymentMethod 
      Caption         =   "Repayment method"
      Height          =   1005
      Left            =   45
      TabIndex        =   9
      Top             =   1080
      Width           =   3975
      Begin VB.PictureBox pctFrame 
         BorderStyle     =   0  'None
         Height          =   690
         Left            =   45
         ScaleHeight     =   690
         ScaleWidth      =   3795
         TabIndex        =   12
         Top             =   180
         Width           =   3795
         Begin VB.OptionButton optRepaymentMethod 
            Caption         =   "Monthly method"
            Height          =   375
            Index           =   0
            Left            =   45
            TabIndex        =   14
            Top             =   0
            Value           =   -1  'True
            Width           =   2220
         End
         Begin VB.OptionButton optRepaymentMethod 
            Caption         =   "Interval method"
            Height          =   375
            Index           =   1
            Left            =   45
            TabIndex        =   13
            Top             =   315
            Width           =   2220
         End
      End
   End
   Begin atc2valtext.ValText txtDateFirstInstallment 
      Height          =   330
      Left            =   2700
      TabIndex        =   0
      Top             =   135
      Width           =   1320
      _ExtentX        =   2328
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
      TypeOfData      =   2
      AllowEmpty      =   0   'False
      AutoSelect      =   0
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2925
      TabIndex        =   6
      Top             =   3285
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   1710
      TabIndex        =   5
      Top             =   3285
      Width           =   1140
   End
   Begin atc2valtext.ValText txtDayOfMonthOrInterval 
      Height          =   330
      Left            =   2700
      TabIndex        =   2
      Top             =   2295
      Width           =   1320
      _ExtentX        =   2328
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
      Minimum         =   "1"
      AllowEmpty      =   0   'False
      TXTAlign        =   2
      AutoSelect      =   0
   End
   Begin atc2valtext.ValText txtDateOfLastInstallment 
      Height          =   330
      Left            =   2700
      TabIndex        =   3
      Top             =   2700
      Width           =   1320
      _ExtentX        =   2328
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
   Begin VB.Label lblNoOfInstallmentsOrDateOfLast 
      Caption         =   "Date of last instalment"
      Height          =   285
      Left            =   90
      TabIndex        =   11
      Top             =   2745
      Width           =   2445
   End
   Begin VB.Label lblDayOfMonthOrInterval 
      Caption         =   "lblDayOfMonthOrInterval"
      Height          =   330
      Left            =   90
      TabIndex        =   10
      Top             =   2295
      Width           =   2445
   End
   Begin VB.Label lblAmountPaid 
      Caption         =   "Amount paid/(repaid)"
      Height          =   285
      Left            =   135
      TabIndex        =   8
      Top             =   585
      Width           =   2040
   End
   Begin VB.Label lblDAteFirstInstallment 
      Caption         =   "Date of first instalment"
      Height          =   240
      Left            =   135
      TabIndex        =   7
      Top             =   225
      Width           =   1770
   End
End
Attribute VB_Name = "F_Loan_RegularPayments"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Implements IFrmGeneral
Private m_InvalidVT As Control
Private m_ok As Boolean

Private Function validatedata() As Boolean
  Dim dFrom As Date, dTo As Date
  
  On Error GoTo ValidateData_ERR
  
  Call xSet("ValidateData")
  
  If Not CheckValidity(Me, , False) Then GoTo ValidateData_END
  dFrom = ScreenToDateVal(txtDateFirstInstallment.Text, STDV_STRING)
  dTo = ScreenToDateVal(txtDateOfLastInstallment.Text, STDV_STRING)
  
  If dFrom > dTo Then Call Err.Raise(ERR_DATES, "ValidateData", "The date of first instalment is greater than the date of last instalment.")
  
  validatedata = True
  
ValidateData_END:
  Call xReturn("ValidateData")
  Exit Function
ValidateData_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "ValidateData", "Validate Data", "Error validating data for regular payments")
  Resume ValidateData_END
End Function

Private Sub chkForceEndOfMonth_Click()
  p11d32.RegularPaymentsForceEndOfMonth = ChkBoxToBool(chkForceEndOfMonth)
End Sub

Private Sub cmdCancel_Click()
  Me.Hide
End Sub

Private Sub cmdOK_Click()
  If validatedata Then
    m_ok = True
    Me.Hide
  End If
End Sub

Private Sub Form_Load()
  Call SetDefaultVTDate(txtDateFirstInstallment)
  Call SetDefaultVTDate(txtDateOfLastInstallment)
  Call SettingsToScreen
End Sub
Private Sub SettingsToScreen()
  On Error GoTo SettingsToScreen_Err
  
  Call xSet("SettingsToScreen")
  
  Call optRepaymentMethod_Click(-1)
  
  Me.txtDateFirstInstallment.Text = DateValReadToScreen(p11d32.Rates.value(TaxYearStart))
  Me.txtDateOfLastInstallment.Text = DateValReadToScreen(p11d32.Rates.value(TaxYearEnd))
  Me.txtAmountPerInstallment.Text = 0
  Me.txtDayOfMonthOrInterval.Text = 1
  chkForceEndOfMonth.value = BoolToChkBox(p11d32.RegularPaymentsForceEndOfMonth)
  
SettingsToScreen_End:
  Exit Sub
SettingsToScreen_Err:
  Call ErrorMessage(ERR_ERROR, Err, "SettingsToScreen", "Settings To Screen", "Error placing the screen settings to the screen for regular loan payments.")
  Resume SettingsToScreen_End
End Sub
Private Function IsLastDayOfMonth(dDate As Date) As Boolean
  
  On Error GoTo IsLastDayOfMonth_ERR
  If Month(dDate) < 12 Then 'JN
    IsLastDayOfMonth = (Month(DateAdd("d", 1, dDate)) = Month(dDate) + 1)
  Else
    IsLastDayOfMonth = (Day(dDate) = 31)
  End If 'JN
  
IsLastDayOfMonth_END:
  Exit Function
IsLastDayOfMonth_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "IsLastDayOfMonth", "Is Last Day Of Month", "Error determining if " & dDate & " is last day of month.")
End Function
Private Function ForceEndOfMonth(d As Date) As Date
  On Error GoTo ForceEndOfMonth_ERR
  
  If (Not p11d32.RegularPaymentsForceEndOfMonth) Or (d = UNDATED) Then GoTo ForceEndOfMonth_END
  
  Do Until IsLastDayOfMonth(d)
    d = DateAdd("d", 1, d)
  Loop
  
ForceEndOfMonth_END:
  ForceEndOfMonth = d
  Exit Function
ForceEndOfMonth_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "ForceEndOfMonth", "Force End Of Month", "Error forcing payment to end of month.")
  Resume ForceEndOfMonth_END
End Function
Private Function MonthDate(ByVal dDateFrom As Date, lInterval As Long) As Date
  Dim sDay As String
  Dim sMonth As String
  Dim sYear As String
  Dim l As Long
  Dim d As Date
  
  On Error GoTo MonthDate_ERR
  
  dDateFrom = DateAdd("m", lInterval, dDateFrom)
  sMonth = Month(dDateFrom)
  sYear = Year(dDateFrom)
  sDay = Day(dDateFrom)
  
  d = TryConvertDate(sDay & "/" & sMonth & "/" & sYear)
  d = ForceEndOfMonth(d)
  
  Do Until (d <> UNDATED)
    l = l + 1
    sDay = Day(dDateFrom) - l
    d = TryConvertDate(sDay & "/" & sMonth & "/" & sYear)
    d = ForceEndOfMonth(d)
  Loop
  
  MonthDate = d
  
MonthDate_END:
  Exit Function
MonthDate_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "MonthDate", "Month Date", "Error getting the month date for a regular payment.")
  Resume MonthDate_END
End Function
Private Function DateExists(ByVal Interval As String, ByVal Number As Long, ByVal d As Date) As Boolean
  On Error GoTo DateExists_ERR
  
  d = DateAdd(Interval, Number, d)
  
DateExists_END:
  Exit Function
DateExists_ERR:
  Resume DateExists_END
End Function
Public Function RegularPayments(Loan As Loan) As Boolean
  Dim BI As BalanceItem
  Dim dFrom As Date, dTo As Date
  Dim dTemp As Date
  Dim lInterval As Long
  Dim dblPayment As Double
  
  On Error GoTo RegularPayments_ERR
  
  Call xSet("RegularPayments")
  
  If Loan Is Nothing Then Call Err.Raise(ERR_IS_NOTHING, "RegularPayments", "The loan is nothing.")
'  Me.Show vbModal
  Call p11d32.Help.ShowForm(Me, vbModal)
  
  If Not m_ok Then GoTo RegularPayments_END
  dFrom = ScreenToDateVal(txtDateFirstInstallment.Text, STDV_STRING)
  dTo = ScreenToDateVal(txtDateOfLastInstallment.Text, STDV_STRING)
  dTemp = dFrom
  dblPayment = txtAmountPerInstallment.Text
  
  Select Case p11d32.RegularPaymentsMethod
    Case RPM_INTERVAL
      lInterval = txtDayOfMonthOrInterval.Text
      Do While dTemp <= dTo
        Set BI = New BalanceItem
        BI.Payment = dblPayment
        BI.DateFrom = dTemp
        BI.RegularPayment = True
        dTemp = DateAdd("d", lInterval, dTemp)
        Call Loan.BalanceSheet.Add(BI)
      Loop
    Case RPM_MONTHLY
      Do While dTemp <= dTo
        Set BI = New BalanceItem
        BI.Payment = dblPayment
        BI.DateFrom = dTemp
        BI.RegularPayment = True
        Call Loan.BalanceSheet.Add(BI)
        lInterval = lInterval + 1
        dTemp = MonthDate(dFrom, lInterval)
      Loop
    Case Else
      Call ECASE("Invalid regaular payments method, index = " & p11d32.RegularPaymentsMethod & ".")
  End Select
  
  RegularPayments = True
  
  
RegularPayments_END:
  Call xReturn("RegularPayments")
  Exit Function
RegularPayments_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "RegularPayments", "Regular Payments", "Error setting regular payments for a loan.")
  Resume RegularPayments_END
  Resume
End Function
Private Function IsMaxOfMonth() As Date
  
End Function
Private Function IFrmGeneral_CheckChanged(c As Control) As Boolean
  
End Function

Private Property Get IFrmGeneral_InvalidVT() As Control
  Set IFrmGeneral_InvalidVT = m_InvalidVT
End Property

Private Property Set IFrmGeneral_InvalidVT(RHS As Control)
  Set m_InvalidVT = RHS
End Property
Private Sub optRepaymentMethod_Click(Index As Integer)
  Dim i As Integer
  
  Select Case Index
    Case RPM_MONTHLY
      lblDayOfMonthOrInterval.Visible = False
      txtDayOfMonthOrInterval.Visible = False
      txtDayOfMonthOrInterval.TypeOfData = VT_STRING
      txtDayOfMonthOrInterval.AllowEmpty = True
      p11d32.RegularPaymentsMethod = Index
    Case RPM_INTERVAL
      lblDayOfMonthOrInterval = "Interval between payments (days)"
      txtDayOfMonthOrInterval.TypeOfData = VT_LONG
      lblDayOfMonthOrInterval.Visible = True
      txtDayOfMonthOrInterval.Visible = True
      txtDayOfMonthOrInterval.Minimum = 1
      p11d32.RegularPaymentsMethod = Index
    Case -1
      For i = 0 To 1
        If i = p11d32.RegularPaymentsMethod Then
          Call optRepaymentMethod_Click(i)
          optRepaymentMethod(i).value = True
          Exit For
        End If
      Next
  End Select
End Sub

