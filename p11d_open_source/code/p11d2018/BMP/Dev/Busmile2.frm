VERSION 5.00
Object = "{E297AE83-F913-4A8C-873C-EDEAC00CB9AC}#2.1#0"; "atc3ubgrd.ocx"
Begin VB.Form F_CompanyCarMiles 
   Caption         =   "Business Mileage"
   ClientHeight    =   3855
   ClientLeft      =   1170
   ClientTop       =   2685
   ClientWidth     =   5625
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
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3855
   ScaleWidth      =   5625
   StartUpPosition =   1  'CenterOwner
   Begin atc3ubgrd.UBGRD UBGRD 
      Height          =   3135
      Left            =   0
      TabIndex        =   3
      Top             =   120
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   5530
   End
   Begin VB.CommandButton B_Ok 
      Appearance      =   0  'Flat
      Caption         =   "&Ok"
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
      Left            =   4050
      TabIndex        =   2
      Top             =   3375
      Width           =   1455
   End
   Begin VB.Label lblTotalMiles 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "lblTotalMiles"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   990
      TabIndex        =   1
      Top             =   3375
      Width           =   1695
   End
   Begin VB.Label L_Miles 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Total miles"
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
      Height          =   195
      Left            =   90
      TabIndex        =   0
      Top             =   3420
      Width           =   750
   End
End
Attribute VB_Name = "F_CompanyCarMiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Implements IBenefitForm2

Public Parentibf As IBenefitForm2

Public benefit As IBenefitClass

Private m_bDirty As Boolean
Private m_Miles As ObjectList
Private Sub B_OK_Click()
  Me.Hide
End Sub
Private Sub efgMiles_FieldInvalid(Valid As Boolean, Message As String)
  Call SetPanel2(Message)
End Sub
Private Sub Form_Load()
  Call InitMilesGrid(UBGRD.Grid)
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
  'not used
End Function

Private Function IBenefitForm2_BenefitOff() As Boolean
  'not used
End Function

Private Function IBenefitForm2_BenefitOn() As Boolean
  'NOT USED
End Function

Private Function IBenefitForm2_BenefitsToListView() As Long
  'not used
End Function

Private Function IBenefitForm2_BenefitToListView(ben As IBenefitClass, ByVal lBenefitIndex As Long) As Long

End Function

Private Function IBenefitForm2_BenefitToScreen(Optional ByVal BenefitIndex As Long = -1&, Optional ByVal UpdateBenefit As Boolean = True) As Boolean
  Dim ben As IBenefitClass
  Dim benCar As CompanyCar
  Dim ibf As IBenefitForm2
  
On Error GoTo BenefitToScreen_Err:
  
  Call xSet("BenefitToScreen")
  
  If BenefitIndex <> -1 Then
    Set ibf = Me
    
    Set ben = p11d32.CurrentEmployer.CurrentEmployee.benefits(BenefitIndex)
    If ben.BenefitClass <> ibf.benclass Then Call Err.Raise(ERR_INVALIDBENCLASS, "BenefitToScreen", "Benefit type invalid.")
    Set benCar = ben
    
    Set UBGRD.ObjectList = benCar.mileage
'AM To be removed    lblTotalMiles = benefit.value(car_SumMiles)
'    Me.Show vbModal
  Call p11d32.Help.ShowForm(Me, vbModal)
  End If
  
BenefitToScreen_End:
  Set ibf = Nothing
  Set ben = Nothing
  Set benCar = Nothing
  Call xReturn("BenefitToScreen")
  Exit Function
BenefitToScreen_Err:
  
  Call ErrorMessage(ERR_ERROR, Err, "BenefitToScreen", "Benefit To Screen", "Unable to place the mileage benefit to the screen.")
  Resume BenefitToScreen_End
  Resume
End Function

Private Property Let IBenefitForm2_benclass(ByVal NewValue As BEN_CLASS)
End Property

Private Property Get IBenefitForm2_benclass() As BEN_CLASS
  IBenefitForm2_benclass = BC_COMPANY_CARS_F
End Property

'Private Property Get IBenefitForm2_ControlDefault() As Control
'  Set IBenefitForm2_ControlDefault = UBGRD
'End Property

Private Property Get IBenefitForm2_lv() As MSComctlLib.IListView

End Property

Private Function IBenefitForm2_RemoveBenefit(ByVal BenefitIndex As Long) As Boolean

End Function

Private Function IBenefitForm2_UpdateBenefitListViewItem(li As MSComctlLib.IListItem, benefit As IBenefitClass, Optional ByVal BenefitIndex As Long = 0&, Optional ByVal SelectItem As Boolean = False) As Long

End Function

Private Function IBenefitForm2_ValididateBenefit(ben As IBenefitClass) As Boolean
    
End Function
Private Sub ubgrd_DeleteData(ObjectList As ObjectList, ObjectListIndex As Long)
'AM To be removed  Call MilesDelete(lblTotalMiles, benefit, car_SumMiles, ObjectList, ObjectListIndex)
End Sub

Private Sub UBGRD_ReadData(RowBuf As TrueDBGrid60.RowBuffer, ByVal RowBufRowIndex As Long, ObjectList As ObjectList, ByVal ObjectListIndex As Long)
  Call MilesRead(RowBuf, RowBufRowIndex, ObjectList, ObjectListIndex)
End Sub

Private Sub UBGRD_ValidateTCS(FirstColIndexInError As Long, ValidateMessage As String, ByVal RowBuf As TrueDBGrid60.RowBuffer, ByVal RowBufRowIndex As Long, ByVal ObjectListIndex As Long)

  Call MilesValidate(FirstColIndexInError, ValidateMessage, RowBuf, RowBufRowIndex, ObjectListIndex)
End Sub

Private Sub ubgrd_WriteData(ByVal RowBuf As TrueDBGrid60.RowBuffer, ByVal RowBufRowIndex As Long, ObjectList As ObjectList, ObjectListIndex As Long)
'AM To be removed  Call MilesWrite(lblTotalMiles, car_SumMiles, RowBuf, RowBufRowIndex, ObjectList, ObjectListIndex, benefit)
End Sub
