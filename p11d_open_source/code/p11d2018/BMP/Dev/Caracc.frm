VERSION 5.00
Object = "{8D988532-0F0C-460C-B00E-7B5637E97680}#1.1#0"; "atc2vtext.OCX"
Begin VB.Form F_CompanyCarAcc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Accessories"
   ClientHeight    =   2625
   ClientLeft      =   885
   ClientTop       =   2385
   ClientWidth     =   6915
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
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2625
   ScaleWidth      =   6915
   StartUpPosition =   1  'CenterOwner
   Begin atc2valtext.ValText TB_Acc 
      Height          =   315
      Index           =   2
      Left            =   5700
      TabIndex        =   2
      Top             =   1080
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
      MouseIcon       =   "Caracc.frx":0000
      Text            =   "0"
      AllowEmpty      =   0   'False
      TXTAlign        =   2
   End
   Begin atc2valtext.ValText TB_Acc 
      Height          =   315
      Index           =   1
      Left            =   5700
      TabIndex        =   1
      Top             =   615
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
      MouseIcon       =   "Caracc.frx":001C
      Text            =   "0"
      AllowEmpty      =   0   'False
      TXTAlign        =   2
   End
   Begin atc2valtext.ValText TB_Acc 
      Height          =   315
      Index           =   0
      Left            =   5700
      TabIndex        =   0
      Top             =   150
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
      MouseIcon       =   "Caracc.frx":0038
      Text            =   "0"
      AllowEmpty      =   0   'False
      TXTAlign        =   2
   End
   Begin VB.CommandButton B_Ok 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "&OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5520
      TabIndex        =   3
      Top             =   2160
      Width           =   1245
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "When entering details, ignore any accessories which were required for disabled use."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   405
      Index           =   1
      Left            =   120
      TabIndex        =   7
      Top             =   1680
      Width           =   6705
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Price of accessories bought later (new accessories). Ignore any accessories         fitted before 1st August 1993"
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
      Height          =   435
      Index           =   5
      Left            =   60
      TabIndex        =   5
      Top             =   600
      Width           =   5595
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Price of accessories when the car was first registered"
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
      Index           =   4
      Left            =   60
      TabIndex        =   4
      Top             =   150
      Width           =   3750
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   $"Caracc.frx":0054
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
      Height          =   405
      Index           =   0
      Left            =   60
      TabIndex        =   6
      Top             =   1080
      Width           =   5250
   End
End
Attribute VB_Name = "F_CompanyCarAcc"
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

Private Sub B_OK_Click()
  Call CheckValidityAndBenefitDirty(benefit, Me)
End Sub

Private Function CheckChanged(ctl As Control) As Boolean
  Dim mdi As MDIForm
  Dim ben As IBenefitClass
  Dim i As Long
  On Error GoTo CheckChanged_Err
  Call xSet("CheckChanged")
  Set mdi = MDIMain
  Set ben = Me.benefit
  Select Case ctl.Name
    Case "TB_Acc"
      Select Case ctl.Index
        Case 0
          i = StrComp(ctl.Text, ben.value(car_AccessoriesOriginal_db), vbBinaryCompare)
          If i <> 0 Then ben.value(car_AccessoriesOriginal_db) = ctl.Text
        Case 1
          i = StrComp(ctl.Text, ben.value(car_AccessoriesNew_db), vbBinaryCompare)
          If i <> 0 Then ben.value(car_AccessoriesNew_db) = ctl.Text
        Case 2
          i = StrComp(ctl.Text, ben.value(car_CheapAccessories_db), vbBinaryCompare)
          If i <> 0 Then ben.value(car_CheapAccessories_db) = ctl.Text
        Case Else
          ECASE "Unknown control"
      End Select
    Case Else
      ECASE "Unknown control"
  End Select
  
  If i <> 0 Then ben.Dirty = True
  
CheckChanged_End:
  Call xReturn("CheckChanged")
  Exit Function
CheckChanged_Err:
  Call ErrorMessage(ERR_ERROR, Err, "CheckChanged", "ERR_CHECKCHANGED", "This function has failed for the form " & Me.Name & ".")
  Resume CheckChanged_End
End Function

Private Sub IBenefitForm2_AddBenefit()

End Sub

Private Function IBenefitForm2_AddBenefitSetDefaults(ben As IBenefitClass) As Long

End Function

Private Property Get IBenefitForm2_benefit() As IBenefitClass
  Set IBenefitForm2_benefit = benefit
End Property

Private Property Set IBenefitForm2_benefit(NewValue As IBenefitClass)
  Set benefit = NewValue
End Property

Private Function IBenefitForm2_BenefitFormState(ByVal fState As BENEFIT_FORM_STATE) As Boolean
  'not used
End Function

Private Function IBenefitForm2_BenefitOff() As Boolean
  'note used
End Function

Private Function IBenefitForm2_BenefitOn() As Boolean

End Function

Private Function IBenefitForm2_BenefitsToListView() As Long
  'not used
End Function

Private Function IBenefitForm2_BenefitToListView(ben As IBenefitClass, ByVal lBenefitIndex As Long) As Long

End Function

Private Function IBenefitForm2_BenefitToScreen(Optional ByVal BenefitIndex As Long = -1&, Optional ByVal UpdateBenefit As Boolean = True) As Boolean
  Dim ben As IBenefitClass
  Dim ibf As IBenefitForm2
On Error GoTo BenefitToScreen_Err
    
  Call xSet("BenefitToScreen")
  
  If BenefitIndex <> -1 Then
    Set ben = p11d32.CurrentEmployer.CurrentEmployee.benefits(BenefitIndex)
    Set ibf = Me
    If benefit.BenefitClass <> ibf.benclass Then Call Err.Raise(ERR_INVALIDBENCLASS, "BenefitToScreen", "Benefit type invalid.")
    
    Set benefit = ben
    TB_Acc(0).Text = ben.value(car_AccessoriesOriginal_db)
    TB_Acc(1).Text = ben.value(car_AccessoriesNew_db)
    TB_Acc(2).Text = ben.value(car_CheapAccessories_db)
  End If
  
'  Me.Show vbModal
  Call p11d32.Help.ShowForm(Me, vbModal)
  
BenefitToScreen_End:
  Set ben = Nothing
  Call xReturn("BenefitToScreen")
  Exit Function
BenefitToScreen_Err:
  
  Call ErrorMessage(ERR_ERROR, Err, "BenefitToScreen", "Benefit To Screen", "Unlable to place the fuel benefit to the screen.")
  Resume BenefitToScreen_End
  

End Function

Private Property Let IBenefitForm2_benclass(ByVal NewValue As BEN_CLASS)
End Property

Private Property Get IBenefitForm2_benclass() As BEN_CLASS
  IBenefitForm2_benclass = BC_COMPANY_CARS_F
End Property

'Private Property Get IBenefitForm2_ControlDefault() As Control
'   Set IBenefitForm2_ControlDefault = TB_Acc(0)
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
      Case "TB_Acc"
        Select Case .Index
          Case 0
            bDirty = CheckTextInput(.Text, benefit, car_AccessoriesOriginal_db)
          Case 1
            bDirty = CheckTextInput(.Text, benefit, car_AccessoriesNew_db)
          Case 2
            bDirty = CheckTextInput(.Text, benefit, car_CheapAccessories_db)
          Case Else
            ECASE "Unknown control"
        End Select
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

Private Sub TB_Acc_FieldInvalid(Index As Integer, Valid As Boolean, Message As String)
  Call SetPanel2(Message)
End Sub

Private Sub TB_Acc_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  benefit.Dirty = True
End Sub


Private Sub TB_Acc_Validate(Index As Integer, Cancel As Boolean)
  Call IFrmGeneral_CheckChanged(TB_Acc(Index))
End Sub
