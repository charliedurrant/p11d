VERSION 5.00
Object = "{8D988532-0F0C-460C-B00E-7B5637E97680}#1.0#0"; "atc2vtext.ocx"
Begin VB.Form F_EmployeeCarExtras 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Employee Car Extras"
   ClientHeight    =   2385
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4635
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2385
   ScaleWidth      =   4635
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox ChkBx 
      Alignment       =   1  'Right Justify
      Caption         =   "Is the amount above an amount subjected to PAYE?"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   135
      TabIndex        =   1
      Top             =   450
      Width           =   4410
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   3360
      TabIndex        =   5
      Top             =   1920
      Width           =   1185
   End
   Begin atc2valtext.ValText TB_Data 
      Height          =   285
      Index           =   0
      Left            =   3480
      TabIndex        =   0
      Tag             =   "FREE,FONT"
      Top             =   90
      Width           =   1065
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
      MaxLength       =   10
      Text            =   "0"
      Minimum         =   "0"
      AllowEmpty      =   0   'False
      TXTAlign        =   2
   End
   Begin atc2valtext.ValText TB_Data 
      Height          =   285
      Index           =   1
      Left            =   3480
      TabIndex        =   2
      Tag             =   "FREE,FONT"
      Top             =   800
      Width           =   1065
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
      MaxLength       =   10
      Text            =   "0"
      Minimum         =   "0"
      AllowEmpty      =   0   'False
      TXTAlign        =   2
   End
   Begin atc2valtext.ValText TB_Data 
      Height          =   285
      Index           =   2
      Left            =   3480
      TabIndex        =   3
      Tag             =   "FREE,FONT"
      Top             =   1150
      Width           =   1065
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
      MaxLength       =   10
      Text            =   "0"
      Minimum         =   "0"
      AllowEmpty      =   0   'False
      TXTAlign        =   2
   End
   Begin atc2valtext.ValText TB_Data 
      Height          =   285
      Index           =   3
      Left            =   3480
      TabIndex        =   4
      Tag             =   "FREE,FONT"
      Top             =   1500
      Width           =   1065
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
      MaxLength       =   10
      Text            =   "0"
      Minimum         =   "0"
      AllowEmpty      =   0   'False
      TXTAlign        =   2
   End
   Begin VB.Label L_Data 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Any hire cost amount made good"
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   9
      Left            =   135
      TabIndex        =   9
      Tag             =   "FREE,FONT"
      Top             =   1500
      Width           =   2325
   End
   Begin VB.Label L_Data 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Hire cost"
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   8
      Left            =   135
      TabIndex        =   8
      Tag             =   "FREE,FONT"
      Top             =   1155
      Width           =   630
   End
   Begin VB.Label L_Data 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Lump sum"
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   7
      Left            =   135
      TabIndex        =   7
      Tag             =   "FREE,FONT"
      Top             =   795
      Width           =   720
   End
   Begin VB.Label L_Data 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Amount made good, or amount subjected to PAYE"
      ForeColor       =   &H00800000&
      Height          =   390
      Index           =   6
      Left            =   135
      TabIndex        =   6
      Tag             =   "FREE,FONT"
      Top             =   50
      Width           =   2950
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "F_EmployeeCarExtras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Implements IBenefitForm2
Implements IFrmGeneral

Private mInvalidVT As Control
Public Parentibf As IBenefitForm2

Public benefit As IBenefitClass


Private Sub ChkBx_Click()
  Call IFrmGeneral_CheckChanged(ChkBx)
End Sub

Private Sub cmdOK_Click()
  Dim eecar As EmployeeCar
  
  If CheckValidityAndBenefitDirty(benefit, Me) Then
    With benefit
      Set eecar = benefit
      .value(eecar_TotalExtras) = eecar.GetTotalOfExtrasPostMadeGood
    End With
  End If
  
  Set eecar = Nothing
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
  'not used
End Function

Private Function IBenefitForm2_BenefitOn() As Boolean
  With benefit
    TB_Data(0).Text = .value(eecar_CarMadeGood_db)
    TB_Data(1).Text = .value(eecar_LumpSum_db)
    TB_Data(2).Text = .value(eecar_HireCost_db)
    TB_Data(3).Text = .value(eecar_HireCostMadeGood_db)
    ChkBx = BoolToChkBox(benefit.value(ITEM_MADEGOOD_IS_TAXDEDUCTED))
  End With
End Function

Private Function IBenefitForm2_BenefitsToListView() As Long
  'not used
End Function

Private Function IBenefitForm2_BenefitToListView(ben As IBenefitClass, ByVal lBenefitIndex As Long) As Long

End Function

Private Function IBenefitForm2_BenefitToScreen(Optional ByVal BenefitIndex As Long = -1&, Optional ByVal UpdateBenefit As Boolean = True) As Boolean
  Dim ibf As IBenefitForm2
On Error GoTo BenefitToScreen_Err
    
  Call xSet("BenefitToScreen")
  
  If BenefitIndex <> -1 Then
    Set ibf = Me
    Set benefit = p11d32.CurrentEmployer.CurrentEmployee.benefits(BenefitIndex)
    If Not benefit Is Nothing Then
      If benefit.BenefitClass <> ibf.benclass Then Call Err.Raise(ERR_INVALIDBENCLASS, "BenefitToScreen", "Benefit type invalid.")
      Call ibf.BenefitOn
'      Me.Show vbModal
      Call p11d32.Help.ShowForm(Me, vbModal)
      IBenefitForm2_BenefitToScreen = True
    End If
  End If
  
BenefitToScreen_End:
  Set ibf = Nothing
  Call xReturn("BenefitToScreen")
  Exit Function
BenefitToScreen_Err:
  
  Call ErrorMessage(ERR_ERROR, Err, "BenefitToScreen", "Benefit To Screen", "Unlable to place the fuel benefit onto the screen.")
  Resume BenefitToScreen_End
  Resume


End Function
Private Property Let IBenefitForm2_benclass(ByVal NewValue As BEN_CLASS)
End Property
Private Property Get IBenefitForm2_benclass() As BEN_CLASS
  IBenefitForm2_benclass = BC_EMPLOYEE_CAR_E
End Property

'Private Property Get IBenefitForm2_ControlDefault() As Control
'  IBenefitForm2_ControlDefault = TB_Data(0)
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
  Dim bDirty As Boolean
  
  On Error GoTo CheckChanged_Err
  Call xSet("CheckChanged")
  
  
  If p11d32.CurrentEmployeeIsNothing Then GoTo CheckChanged_End
  If benefit Is Nothing Then GoTo CheckChanged_End
  
  With c
    Select Case .Name
      Case "TB_Data"
        Select Case .Index
          Case 0
            bDirty = CheckTextInput(.Text, benefit, eecar_CarMadeGood_db)
          Case 1
            bDirty = CheckTextInput(.Text, benefit, eecar_LumpSum_db)
          Case 2
            bDirty = CheckTextInput(.Text, benefit, eecar_HireCost_db)
          Case 3
            bDirty = CheckTextInput(.Text, benefit, eecar_HireCostMadeGood_db)
          Case Else
            ECASE "Unknown control"
        End Select
      Case "ChkBx"
        bDirty = CheckCheckBoxInput(.value, benefit, ITEM_MADEGOOD_IS_TAXDEDUCTED)
      Case Else
        ECASE "Unknown control"
    End Select
  End With
  
  IFrmGeneral_CheckChanged = AfterCheckChanged(c, Me.Parentibf, bDirty)
   
CheckChanged_End:
  Call xReturn("CheckChanged")
  Exit Function
CheckChanged_Err:
  IFrmGeneral_CheckChanged = False
  Call ErrorMessage(ERR_ERROR, Err, "CheckChanged", "ERR_CHECKCHANGED", "This function has failed for the form " & Me.Name & ".")
  Resume CheckChanged_End

End Function

Private Property Get IFrmGeneral_InvalidVT() As Control
  Set IFrmGeneral_InvalidVT = mInvalidVT
End Property

Private Property Set IFrmGeneral_InvalidVT(NewValue As Control)
  Set mInvalidVT = NewValue
End Property

Private Sub TB_Data_FieldInvalid(Index As Integer, Valid As Boolean, Message As String)
  Call SetPanel2(Message)
End Sub

Private Sub TB_Data_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   TB_Data(Index).Tag = SetChanged
End Sub

Private Sub TB_Data_Validate(Index As Integer, Cancel As Boolean)
  Call IFrmGeneral_CheckChanged(TB_Data(Index))
End Sub
