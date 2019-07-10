VERSION 5.00
Object = "{E297AE83-F913-4A8C-873C-EDEAC00CB9AC}#2.1#0"; "atc3ubgrd.ocx"
Begin VB.Form F_BalanceSheet 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Balance Sheet"
   ClientHeight    =   4230
   ClientLeft      =   2310
   ClientTop       =   1605
   ClientWidth     =   8445
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
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4230
   ScaleWidth      =   8445
   StartUpPosition =   2  'CenterScreen
   Begin atc3ubgrd.UBGRD UBGRD 
      Height          =   3015
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   5318
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear All"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   45
      TabIndex        =   7
      Top             =   3780
      Width           =   1590
   End
   Begin VB.CommandButton B_Regular 
      Caption         =   "&Regular Payments..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   5130
      TabIndex        =   3
      Top             =   3780
      Width           =   1590
   End
   Begin VB.CommandButton B_Ok 
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   6795
      TabIndex        =   4
      Top             =   3780
      Width           =   1590
   End
   Begin VB.Label lblNoOfRows 
      Alignment       =   1  'Right Justify
      Caption         =   "lblNoOfRows"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   6660
      TabIndex        =   6
      Top             =   3060
      Width           =   1680
   End
   Begin VB.Label Lab 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Interest - 'Daily' method : "
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
      Left            =   3960
      TabIndex        =   0
      Top             =   3420
      Width           =   1770
   End
   Begin VB.Label lblAverageInterest 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "lblAverageInterest"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2385
      TabIndex        =   1
      Top             =   3420
      Width           =   1080
   End
   Begin VB.Label Lab 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Interest - 'Averaging' method : "
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
      Left            =   120
      TabIndex        =   5
      Top             =   3420
      Width           =   2385
   End
   Begin VB.Label lblDailyInterest 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "lblDailyInterest"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   7110
      TabIndex        =   2
      Top             =   3375
      Width           =   1185
   End
End
Attribute VB_Name = "F_BalanceSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public benefit As IBenefitClass
Implements IBenefitForm2
Implements IFrmGeneral
Private m_InvalidVT As Control
Public Parentibf As IBenefitForm2
Private m_BenClass As BEN_CLASS
Public m_dirty As Boolean

Private Sub B_OK_Click()
  Me.Hide
End Sub

Private Sub B_Regular_Click()
  If F_Loan_RegularPayments.RegularPayments(benefit) Then Call ChangedData(True)
  Set F_Loan_RegularPayments = Nothing
End Sub
Private Sub Clear()
  Dim Loan As Loan
  On Error GoTo Clear_ERR
  
  Set Loan = benefit
  Call Loan.BalanceSheet.RemoveAll
  Call ChangedData(True)
  
Clear_END:
  Exit Sub
Clear_ERR:
  Call ErrorMessage(ERR_INFO, Err, "Clear", "Clear", "Error clearing the balances.")
  Resume Clear_END
End Sub

Private Sub cmdClear_Click()
  If MsgBox("Are you sure you want to clear all your entries?", vbYesNo, "P11D") = vbYes Then
    Call Clear 'JN code
  End If
End Sub

Private Sub Form_Load()
  Dim c As TrueDBGrid60.Column
  Dim grd As Object
  
 On Error GoTo F_BalanceSheet_Load_ERR
  
On Error GoTo F_BalanceSheet_Load_ERR

  Call xSet("F_BalanceSheet_Load")

  Set grd = UBGRD.Grid
  
  
  Call AddUBGRDStandardColumn(grd, 0, 1244.976, "Date from", "")
  Call AddUBGRDStandardColumn(grd, 1, 2200, "Amount received/(repaid) (£)", "General Number") '"##,##0;(##,##0)")
  
  Set c = AddUBGRDStandardColumn(grd, 2, 1500, "Balance (£)", "General Number")
  Set c.Style = grd.Styles.Item("Heading")
  Set c = AddUBGRDStandardColumn(grd, 3, 1800, "Number of days", "")
  Set c.Style = grd.Styles.Item("Heading")
  Set c = AddUBGRDStandardColumn(grd, 4, 1500, "Daily interest", "Currency")
  Set c.Style = grd.Styles.Item("Heading")
  grd.AllowUpdate = True
  
  
F_BalanceSheet_Load_END:
  Set c = Nothing
  Set grd = Nothing
  Call xReturn("F_BalanceSheet_Load")
  Exit Sub
F_BalanceSheet_Load_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "F_BalanceSheet_Load", "F_BalanceSheet Load", "Unable to load the balance sheet form.")
  Resume F_BalanceSheet_Load_END
  Resume
End Sub

Private Sub IBenefitForm2_AddBenefit()

End Sub

Private Function IBenefitForm2_AddBenefitSetDefaults(ben As IBenefitClass) As Long

End Function

Private Property Set IBenefitForm2_benefit(NewValue As IBenefitClass)
  Set benefit = NewValue
End Property

Private Property Get IBenefitForm2_benefit() As IBenefitClass
  Set IBenefitForm2_benefit = benefit
End Property

Private Function IBenefitForm2_BenefitFormState(ByVal fState As BENEFIT_FORM_STATE) As Boolean

End Function

Private Function IBenefitForm2_BenefitOff() As Boolean
  'note used
End Function

Private Function IBenefitForm2_BenefitOn() As Boolean
  'note used
End Function

Private Function IBenefitForm2_BenefitsToListView() As Long
  'note used
End Function

Private Function IBenefitForm2_BenefitToListView(ben As IBenefitClass, ByVal lBenefitIndex As Long) As Long

End Function

Private Function IBenefitForm2_BenefitToScreen(Optional ByVal BenefitIndex As Long = -1&, Optional ByVal UpdateBenefit As Boolean = True) As Boolean
  Dim ben As IBenefitClass
  Dim loans As loans
  Dim Loan As Loan
  
'On Error GoTo BenefitToScreen_Err

  Call xSet("BenefitToScreen")

  If BenefitIndex <> -1 Then
    Me.Enabled = True
    'benfit index is the current loan within the loans collection
    'so the BenefitIndex has the benefits collection index and the loans collectionindex
    Set loans = p11d32.CurrentEmployer.CurrentEmployee.benefits(HiWord(BenefitIndex))
    Set ben = loans.Item(LowWord(BenefitIndex))
    If ben.BenefitClass <> m_BenClass Then Call Err.Raise(ERR_INVALIDBENCLASS, "BenefitToScreen", "Benefit type invalid.")
    Set benefit = ben
    Set Loan = ben
    Set UBGRD.ObjectList = Loan.BalanceSheet
    Call UpdateInterestLabels
'    Me.Show vbModal
    Call p11d32.Help.ShowForm(Me, vbModal)
    If m_dirty Then Loan.FinishBalances (True)
  Else
    Me.Enabled = False
  End If

  
  

BenefitToScreen_End:
  m_dirty = False
  Set ben = Nothing
  Set loans = Nothing
  Call xReturn("BenefitToScreen")
  Exit Function
BenefitToScreen_Err:
  
  Call ErrorMessage(ERR_ERROR, Err, "BenefitToScreen", "Benefit To Screen", "Unlable to place the balances to the screen.")
  Resume BenefitToScreen_End
  Resume
End Function

Private Property Let IBenefitForm2_benclass(ByVal NewValue As BEN_CLASS)
  m_BenClass = NewValue
End Property

Private Property Get IBenefitForm2_benclass() As BEN_CLASS
  IBenefitForm2_benclass = m_BenClass
End Property

'Private Property Get IBenefitForm2_ControlDefault() As Control''
'
'End Property

Private Property Get IBenefitForm2_lv() As MSComctlLib.IListView

End Property

Private Function IBenefitForm2_RemoveBenefit(ByVal BenefitIndex As Long) As Boolean

End Function

Private Function IBenefitForm2_UpdateBenefitListViewItem(li As MSComctlLib.IListItem, benefit As IBenefitClass, Optional ByVal BenefitIndex As Long = 0&, Optional ByVal SelectItem As Boolean = False) As Long

End Function

Private Function IBenefitForm2_ValididateBenefit(ben As IBenefitClass) As Boolean

End Function

Private Function IFrmGeneral_CheckChanged(c As Control) As Boolean

End Function

Private Property Get IFrmGeneral_InvalidVT() As Control
  Set IFrmGeneral_InvalidVT = m_InvalidVT
End Property

Private Property Set IFrmGeneral_InvalidVT(NewValue As Control)
  Set m_InvalidVT = NewValue
End Property

Private Sub ubgrd_DeleteData(ObjectList As ObjectList, ObjectListIndex As Long)
  Dim Loan As Loan
  
On Error GoTo F_BalanceSheet_ubgrd_DeleteData_ERR
  
  Call xSet("F_BalanceSheet_ubgrd_DeleteData")
  
  Call UBGRD.ObjectList.Remove(ObjectListIndex)
  
  Call ChangedData(False)
  
F_BalanceSheet_ubgrd_DeleteData_END:
  Call xReturn("F_BalanceSheet_ubgrd_DeleteData")
  Exit Sub
F_BalanceSheet_ubgrd_DeleteData_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "F_BalanceSheet_ubgrd_DeleteData", "F_BalanceSheet_ubgrd_DeleteData", "Error deleting a balance item or refreshing the grid.")
  Resume F_BalanceSheet_ubgrd_DeleteData_END
End Sub

Private Sub UBGRD_ReadData(RowBuf As TrueDBGrid60.RowBuffer, ByVal RowBufRowIndex As Long, ObjectList As ObjectList, ByVal ObjectListIndex As Long)
  Dim l As Long
  Dim BI As BalanceItem

On Error GoTo Balaces_ubgrd_ReadData_ERR

  Call xSet("Balaces_ubgrd_ReadData")

  Set BI = ObjectList(ObjectListIndex)

  For l = 0 To (RowBuf.ColumnCount - 1)
    Select Case l
      Case 0
        RowBuf.value(RowBufRowIndex, l) = DateValReadToScreen(BI.DateFrom)
      Case 1
        RowBuf.value(RowBufRowIndex, l) = BI.Payment
      Case 2
        RowBuf.value(RowBufRowIndex, l) = BI.Balance
      Case 3
        RowBuf.value(RowBufRowIndex, l) = BI.days
      Case 4
        RowBuf.value(RowBufRowIndex, l) = BI.Interest(benefit.value(ln_LoanCurrencyIndex))
      Case Else
        ECASE ("Invalid column if get user data")
    End Select
  Next

Balaces_ubgrd_ReadData_END:
  Set BI = Nothing
  Call xReturn("ubgrd_ReadData")
  Exit Sub
Balaces_ubgrd_ReadData_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "Balaces_ubgrd_ReadData", "Balaces_ubgrd_ReadData", "Error reading a BalanceItem")
  Resume Balaces_ubgrd_ReadData_END
End Sub


Private Sub UBGRD_ValidateTCS(FirstColIndexInError As Long, ValidateMessage As String, ByVal RowBuf As TrueDBGrid60.RowBuffer, ByVal RowBufRowIndex As Long, ByVal ObjectListIndex As Long)


  Dim l As Long

On Error GoTo ubgrd_ValidateTCS_Err

  Call xSet("ubgrd_ValidateTCS")

  With RowBuf
    For l = 0 To 1
      Select Case l
        Case 0
          If GridIsNotDate(ValidateMessage, RowBuf.value(RowBufRowIndex, l), ObjectListIndex, True) Then
            FirstColIndexInError = l
            GoTo ubgrd_ValidateTCS_End
          End If
        Case 1
          If GridIsNotNumericOrLong(ValidateMessage, RowBuf.value(RowBufRowIndex, l), ObjectListIndex) Then
            FirstColIndexInError = l
            GoTo ubgrd_ValidateTCS_End
          End If
      End Select
    Next
  End With


ubgrd_ValidateTCS_End:
  Call xReturn("ubgrd_ValidateTCS")
  Exit Sub

ubgrd_ValidateTCS_Err:
  Call ErrorMessage(ERR_ERROR, Err, "ubgrd_ValidateTCS", "ubgrd Validate TCS", "Error validating a balance sheet entry.")
  Resume ubgrd_ValidateTCS_End
End Sub



Private Sub ubgrd_WriteData(ByVal RowBuf As TrueDBGrid60.RowBuffer, ByVal RowBufRowIndex As Long, ObjectList As ObjectList, ObjectListIndex As Long)
  Dim BI As BalanceItem
  Dim Loan As Loan
  
On Error GoTo Balaces_ubgrd_WriteData_ERR
  Dim BINext As BalanceItem
  
  Call xSet("Balaces_ubgrd_WriteData")
  
  If ObjectListIndex = -1 Then
    Set BI = New BalanceItem
    ObjectListIndex = ObjectList.Add(BI)
  Else
    Set BI = ObjectList(ObjectListIndex)
  End If
  
  With BI
    If Not IsNull(RowBuf.value(RowBufRowIndex, 0)) Then .DateFrom = TryConvertDate((RowBuf.value(RowBufRowIndex, 0)))
    If Not IsNull(RowBuf.value(RowBufRowIndex, 1)) Then .Payment = RoundN(RowBuf.value(RowBufRowIndex, 1))
    benefit.Dirty = True
    
  End With
  
  Call ChangedData(False)
  
Balaces_ubgrd_WriteData_END:
  Set BI = Nothing
  Call xReturn("Balaces_ubgrd_WriteData")
  Exit Sub
Balaces_ubgrd_WriteData_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "Balaces_ubgrd_WriteData", "Balaces_ubgrd_WriteData", "Error writing the ubgrd to the BalanceSheet collection.")
  Resume Balaces_ubgrd_WriteData_END
  Resume
End Sub
Private Sub ChangedData(bRebind As Boolean)
  Dim Loan As Loan
  
  On Error GoTo ChangedData_ERR
  
  If benefit Is Nothing Then Call Err.Raise(ERR_IS_NOTHING, "ChagedData", "The benefit is nothing.")
  Set Loan = benefit
  Call Loan.FinishBalances(False)
  If bRebind Then
    UBGRD.Grid.ReBind
  Else
    UBGRD.Grid.Refresh
  End If
  Call UpdateInterestLabels
  benefit.value(ln_BalancesAmmended) = True
  benefit.Dirty = True
  m_dirty = True
  
ChangedData_END:
  Exit Sub
  
ChangedData_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "ChangedData", "Changed Data", "Error setting benefit info after balances written.")
  Resume ChangedData_END
End Sub
Private Function UpdateInterestLabels() As Boolean
  Dim Loan As Loan
  
  On Error GoTo UpdateInterestLabels_Err
  Call xSet("UpdateInterestLabels")

  Set Loan = benefit
  lblNoOfRows = Loan.BalanceSheet.Count & " row(s)"
  lblDailyInterest.Caption = FormatWN(benefit.value(ln_InterestDaily))
  lblAverageInterest.Caption = FormatWN(benefit.value(ln_InterestNormal))
  UpdateInterestLabels = True
  
UpdateInterestLabels_End:
  Call xReturn("UpdateInterestLabels")
  Exit Function

UpdateInterestLabels_Err:
  Call ErrorMessage(ERR_ERROR, Err, "UpdateInterestLabels", "Update Interest Labels", "Error updating the balance sheet interest labels.")
  Resume UpdateInterestLabels_End
End Function


