VERSION 5.00
Object = "{8D988532-0F0C-460C-B00E-7B5637E97680}#1.0#0"; "atc2vtext.ocx"
Begin VB.Form F_CompanyCarFuel 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Car Fuel Benefit"
   ClientHeight    =   2625
   ClientLeft      =   1710
   ClientTop       =   4830
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
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2625
   ScaleWidth      =   6915
   StartUpPosition =   1  'CenterOwner
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
      Height          =   360
      Left            =   5535
      TabIndex        =   5
      Top             =   2160
      Width           =   1275
   End
   Begin atc2valtext.ValText TB_Fuel 
      Height          =   315
      Left            =   5550
      TabIndex        =   0
      Top             =   120
      Width           =   1215
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
      AllowEmpty      =   0   'False
      TXTAlign        =   2
   End
   Begin VB.CheckBox Bn_Fuel 
      Alignment       =   1  'Right Justify
      Caption         =   "Is this a diesel car?"
      DataField       =   "Diesel"
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
      Height          =   285
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   6645
   End
   Begin VB.CheckBox Bn_Fuel 
      Alignment       =   1  'Right Justify
      Caption         =   "If so, was the total cost actually made good?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   3
      Left            =   120
      TabIndex        =   4
      Top             =   1560
      Width           =   6645
   End
   Begin VB.CheckBox Bn_Fuel 
      Alignment       =   1  'Right Justify
      Caption         =   "If so, was the total cost required to be made good?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   2
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   6645
   End
   Begin VB.CheckBox Bn_Fuel 
      Alignment       =   1  'Right Justify
      Caption         =   "Was fuel provided for private use?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   810
      Width           =   6645
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Engine cc"
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
      Index           =   0
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   2475
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
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
      Index           =   1
      Left            =   60
      TabIndex        =   7
      Top             =   840
      Width           =   45
   End
End
Attribute VB_Name = "F_CompanyCarFuel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'MP DB ToDo - remove this form ? in use?
Option Explicit
Public benefit As IBenefitClass
Implements IFrmGeneral
Implements IBenefitForm2
Private m_InvalidVT As Control

Public Parentibf As IBenefitForm2

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
    Case "Bn_Fuel"
      Select Case ctl.Index
        Case 0
'AM To be removed
'          i = (IIf(ctl.value = vbChecked, True, False) <> ben.value(car_Diesel))
'          If i <> 0 Then ben.value(car_Diesel) = IIf(ctl.value = vbChecked, True, False)
        Case 1
          i = (IIf(ctl.value = vbChecked, True, False) <> ben.value(car_privatefuel_db))
          If i <> 0 Then ben.value(car_privatefuel_db) = IIf(ctl.value = vbChecked, True, False)
        Case 2
          i = (IIf(ctl.value = vbChecked, True, False) <> ben.value(car_requiredmakegood_db))
          If i <> 0 Then ben.value(car_requiredmakegood_db) = IIf(ctl.value = vbChecked, True, False)
        Case 3
          i = (IIf(ctl.value = vbChecked, True, False) <> ben.value(car_actualmadegood_db))
          If i <> 0 Then ben.value(car_actualmadegood_db) = IIf(ctl.value = vbChecked, True, False)
        Case Else
          ECASE "Unknown control"
      End Select
    Case "TB_Fuel"
      i = StrComp(ctl.Text, ben.value(car_enginesize_db), vbBinaryCompare)
      If i <> 0 Then ben.value(car_enginesize_db) = ctl.Text
    Case Else
      ECASE "Unknown control"
  End Select
  If i <> 0 Then
    ben.Dirty = True
  End If
CheckChanged_End:
  Set ben = Nothing
  Set mdi = Nothing
  Call xReturn("CheckChanged")
  Exit Function
CheckChanged_Err:
  Call ErrorMessage(ERR_ERROR, Err, "CheckChanged", "ERR_CHECKCHANGED", "This function has failed for the form " & Me.Name & ".")
  Resume CheckChanged_End
End Function

Private Sub Bn_Fuel_Click(Index As Integer)
  Call IFrmGeneral_CheckChanged(Bn_Fuel(Index))
End Sub

Private Sub Bn_Fuel_KeyPress(Index As Integer, KeyAscii As Integer)
  'Check for return key - tab to next field
  If KeyAscii = 13 Then Call SendKeys(vbTab)
End Sub


Private Sub IBenefitForm2_AddBenefit()
  'NOT USED
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
  'NOT USED
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
    
    TB_Fuel.Text = ben.value(car_enginesize_db)
'AM    Bn_Fuel(0) = IIf(ben.value(car_Diesel), vbChecked, vbUnchecked)
    Bn_Fuel(1) = IIf(ben.value(car_privatefuel_db), vbChecked, vbUnchecked)
    Bn_Fuel(2) = IIf(ben.value(car_requiredmakegood_db), vbChecked, vbUnchecked)
    Bn_Fuel(3) = IIf(ben.value(car_actualmadegood_db), vbChecked, vbUnchecked)
  Else
    TB_Fuel.Text = ""
    Bn_Fuel(0) = vbGrayed
    Bn_Fuel(1) = vbGrayed
    Bn_Fuel(2) = vbGrayed
    Bn_Fuel(3) = vbGrayed
  End If
  
'  Me.Show vbModal
  Call p11d32.Help.ShowForm(Me, vbModal)
  
BenefitToScreen_End:
  Set ibf = Nothing
  Set ben = Nothing
  Call xReturn("BenefitToScreen")
  Exit Function
BenefitToScreen_Err:
  
  Call ErrorMessage(ERR_ERROR, Err, "BenefitToScreen", "Benefit To Screen", "Unable to place the fuel benefit to the screen.")
  Resume BenefitToScreen_End
  Resume
End Function

Private Property Let IBenefitForm2_benclass(ByVal NewValue As BEN_CLASS)
End Property

Private Property Get IBenefitForm2_benclass() As BEN_CLASS
  IBenefitForm2_benclass = BC_COMPANY_CARS_F
End Property

'Private Property Get IBenefitForm2_ControlDefault() As Control
'  Set IBenefitForm2_ControlDefault = TB_Fuel
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
  Dim i As Long
  
  On Error GoTo CheckChanged_Err
  Call xSet("CheckChanged")
  
  If p11d32.CurrentEmployeeIsNothing Then GoTo CheckChanged_End
  If benefit Is Nothing Then GoTo CheckChanged_End
  
  With c
    Select Case .Name
      Case "Bn_Fuel"
        Select Case .Index
          Case 0
'AM To be removed            bDirty = CheckCheckBoxInput(.value, benefit, car_Diesel)
            ' If p11d32.AppYear = 2001 And
            If bDirty Then Call SetFuelTypeAsDiesel                   'RH
          Case 1
            bDirty = CheckCheckBoxInput(.value, benefit, car_privatefuel_db)
          Case 2
            bDirty = CheckCheckBoxInput(.value, benefit, car_requiredmakegood_db)
          Case 3
            bDirty = CheckCheckBoxInput(.value, benefit, car_actualmadegood_db)
          Case Else
            ECASE "Unknown control"
        End Select
      Case "TB_Fuel"
        bDirty = CheckTextInput(.Text, benefit, car_enginesize_db)
      Case Else
        ECASE "Unknown control"
    End Select
  End With
  
  IFrmGeneral_CheckChanged = AfterCheckChanged(c, Parentibf, bDirty)
   
CheckChanged_End:
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

Private Sub TB_Fuel_FieldInvalid(Valid As Boolean, Message As String)
  Call SetPanel2(Message)
End Sub
Private Sub TB_Fuel_Validate(Cancel As Boolean)
  Call IFrmGeneral_CheckChanged(TB_Fuel)
End Sub

Private Sub SetFuelTypeAsDiesel()
 
'RH
  On Error GoTo SetFuelTypeAsDiesel_Err
  Call xSet("SetFuelTypeAsDiesel")
  
'AM To be removed
'    Select Case benefit.value(car_Diesel)
'     Case True
'        benefit.value(car_p46FuelType) = 1
'     Case False
'       If benefit.value(car_p46FuelType) = 1 Then benefit.value(car_p46FuelType) = 0
'     End Select
     'If p11d32.AppYear = 2000 Then  'AM
      'F_CompanyCar.CB_FuelType(0).ListIndex = benefit.value(car_p46FuelType)
     'Else
     F_CompanyCar.CB_FuelType(1).ListIndex = benefit.value(car_p46FuelType_db)
     'End If

SetFuelTypeAsDiesel_End:
  Call xReturn("SetFuelTypeAsDiesel")
  Exit Sub
  
SetFuelTypeAsDiesel_Err:
  Call ErrorMessage(ERR_ERROR, Err, "SetFuelTypeAsDiesel", "Set Fuel Type as Diesel", "Error setting the P46 fuel type to match the diesel check box.")
  Resume SetFuelTypeAsDiesel_End
  Resume

End Sub
