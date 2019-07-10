VERSION 5.00
Object = "{E297AE83-F913-4A8C-873C-EDEAC00CB9AC}#2.1#0"; "ATC3UBGRD.OCX"
Object = "{8D988532-0F0C-460C-B00E-7B5637E97680}#1.0#0"; "ATC2VTEXT.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form F_EmployeeCar 
   Caption         =   " "
   ClientHeight    =   5685
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8325
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   ScaleHeight     =   5685
   ScaleWidth      =   8325
   WindowState     =   2  'Maximized
   Begin VB.Frame fmeCar 
      Height          =   3510
      Left            =   0
      TabIndex        =   10
      Top             =   2115
      Width           =   8295
      Begin atc3ubgrd.UBGRD UBGRD 
         Height          =   2055
         Left            =   3840
         TabIndex        =   18
         Top             =   480
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   3625
      End
      Begin VB.CommandButton cmdExtras 
         Caption         =   "Extras / (Made Good)"
         Height          =   330
         Left            =   1800
         TabIndex        =   4
         Tag             =   "FREE,FONT"
         Top             =   1935
         Width           =   1815
      End
      Begin VB.CheckBox Op_Data 
         Alignment       =   1  'Right Justify
         Caption         =   "Use company car scheme for mileage allowance"
         DataField       =   "BusMilesActual"
         DataSource      =   "DB"
         ForeColor       =   &H00800000&
         Height          =   390
         Index           =   0
         Left            =   90
         TabIndex        =   5
         Tag             =   "free,font"
         Top             =   2295
         Width           =   3510
      End
      Begin VB.CheckBox Op_Data 
         Alignment       =   1  'Right Justify
         Caption         =   "Average IR rates?"
         DataField       =   "BusMilesActual"
         DataSource      =   "DB"
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   1
         Left            =   135
         TabIndex        =   7
         Tag             =   "free,font"
         Top             =   3105
         Visible         =   0   'False
         Width           =   3450
      End
      Begin VB.ComboBox CB_FPCS 
         Appearance      =   0  'Flat
         DataField       =   "FPCS"
         DataSource      =   "DB"
         ForeColor       =   &H00000080&
         Height          =   315
         Left            =   1980
         TabIndex        =   6
         Tag             =   "free,font"
         Text            =   "CB_FPCS"
         Top             =   2700
         Width           =   1635
      End
      Begin VB.CheckBox Op_Data 
         Alignment       =   1  'Right Justify
         Caption         =   "Did the employee own and use any other car (which they claimed mileage on) at the same time?"
         DataField       =   "BusMilesActual"
         DataSource      =   "DB"
         ForeColor       =   &H00800000&
         Height          =   525
         Index           =   2
         Left            =   3840
         TabIndex        =   9
         Tag             =   "free,font"
         Top             =   2835
         Visible         =   0   'False
         Width           =   4320
      End
      Begin atc2valtext.ValText TB_Data 
         Height          =   330
         Index           =   2
         Left            =   2520
         TabIndex        =   2
         Tag             =   "FREE,FONT"
         Top             =   1125
         Width           =   1065
         _ExtentX        =   1879
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
         MaxLength       =   10
         Text            =   "0"
         Minimum         =   "0"
         AllowEmpty      =   0   'False
         TXTAlign        =   2
      End
      Begin atc2valtext.ValText TB_Data 
         Height          =   315
         Index           =   0
         Left            =   600
         TabIndex        =   0
         Tag             =   "free,font"
         Top             =   240
         Width           =   3015
         _ExtentX        =   5318
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
         Text            =   ""
         TypeOfData      =   3
         AllowEmpty      =   0   'False
      End
      Begin atc2valtext.ValText TB_Data 
         Height          =   330
         Index           =   1
         Left            =   2520
         TabIndex        =   1
         Tag             =   "FREE,FONT"
         Top             =   675
         Width           =   1080
         _ExtentX        =   1905
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
         MaxLength       =   10
         Text            =   "0"
         Minimum         =   "0"
         AllowEmpty      =   0   'False
         TXTAlign        =   2
      End
      Begin VB.Label lblTotalOfExtras 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "total of extras"
         Height          =   345
         Left            =   135
         TabIndex        =   3
         Tag             =   "FREE,FONT"
         Top             =   1890
         Width           =   1485
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblTotalMiles 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "total miles for car"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   5940
         TabIndex        =   8
         Tag             =   "FREE,FONT"
         Top             =   2520
         Width           =   2160
         WordWrap        =   -1  'True
      End
      Begin VB.Label L_Data 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Mileage"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   12
         Left            =   3825
         TabIndex        =   17
         Tag             =   "FREE,FONT"
         Top             =   270
         Width           =   1260
         WordWrap        =   -1  'True
      End
      Begin VB.Label L_Data 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Total miles"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   5
         Left            =   3825
         TabIndex        =   16
         Tag             =   "FREE,FONT"
         Top             =   2520
         Width           =   1845
         WordWrap        =   -1  'True
      End
      Begin VB.Label L_Data 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Company mileage scheme"
         ForeColor       =   &H00800000&
         Height          =   375
         Index           =   3
         Left            =   135
         TabIndex        =   15
         Tag             =   "FREE,FONT"
         Top             =   2700
         Width           =   1800
         WordWrap        =   -1  'True
      End
      Begin VB.Label L_Data 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Engine size"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   1
         Left            =   135
         TabIndex        =   14
         Tag             =   "FREE,FONT"
         Top             =   720
         Width           =   810
      End
      Begin VB.Label L_Data 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Car"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   13
         Tag             =   "free,font"
         Top             =   270
         Width           =   240
      End
      Begin VB.Label L_Data 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Amount received by employee"
         ForeColor       =   &H00800000&
         Height          =   420
         Index           =   2
         Left            =   135
         TabIndex        =   12
         Tag             =   "FREE,FONT"
         Top             =   1170
         Width           =   2340
         WordWrap        =   -1  'True
      End
   End
   Begin MSComctlLib.ListView LB 
      Height          =   2085
      Left            =   0
      TabIndex        =   11
      Tag             =   "free,font"
      Top             =   45
      Width           =   8280
      _ExtentX        =   14605
      _ExtentY        =   3678
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
         Text            =   "Car Reference"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Benefit"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "F_EmployeeCar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Implements IFrmGeneral
Implements IBenefitForm2

Public benefit As IBenefitClass

Private mclsResize As New clsFormResize
Private Const L_DES_HEIGHT  As Long = 6090
Private Const L_DES_WIDTH  As Long = 8445
Private m_InvalidVT As Control

Private Const S_AMOUNT_RECEIVED_DIRECT As String = "Amount received by employee"
Private Const S_AMOUNT_RECEIVED_PER_FPCS As String = "Amount received by employee per AMAP rates"

Private Sub CB_FPCS_Click()
  Call IFrmGeneral_CheckChanged(CB_FPCS)
End Sub

Private Sub CB_FPCS_KeyDown(KeyCode As Integer, Shift As Integer)
  CB_FPCS.Tag = SetChanged
End Sub

Private Sub cmdExtras_Click()
  Call DialogToScreen(F_EmployeeCarExtras, lblTotalOfExtras, eecar_TotalExtras, Me, p11d32.CurrentEmployer.CurrentEmployee.benefits.ItemIndex(benefit))
End Sub

Private Sub Form_Resize()
  mclsResize.Resize
    Call ColumnWidths(lb, 70, 30)
End Sub
Private Sub Form_Load()

  On Error GoTo F_Employee_Car_Load_ERR
    
  Call xSet("F_Employee_Car_Load")
  
  If Not (mclsResize.InitResize(Me, L_DES_HEIGHT, L_DES_WIDTH, DESIGN, , , MDIMain)) Then
    Err.Raise ERR_Application
  End If
  
  Call FPCSToCombo(False)
  Call InitMilesGrid(UBGRD.Grid)
  
F_Employee_Car_Load_END:
  Call xReturn("F_Employee_Car_Load")
  Exit Sub
F_Employee_Car_Load_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "F_Employee_Car_Load", "F_Employee_Car_Load", "Error loading the employee car form.")
  Resume F_Employee_Car_Load_END
  Resume
End Sub


Private Sub IBenefitForm2_AddBenefit()
  Dim ben As IBenefitClass
  
On Error GoTo AddBenefit_Err
  
  Call xSet("AddBenefit")
  
  Set ben = New EmployeeCar
  Call AddBenefitHelper(Me, ben)
  TB_Data(0).SetFocus
  
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
    
    .value(eecar_Item_db) = "Please enter a description..."
    
    .value(eecar_UseCompanyCarScheme_db) = False
'MP DB    .value(eecar_AverageIRRate_db) = False
    'EK No longer needed 2003 .value(eecar_AlternativeMethod) = False
    
    .value(eecar_CarMadeGood_db) = 0
    .value(eecar_AmountReceived_db) = 0
    .value(eecar_LumpSum_db) = 0
    .value(eecar_HireCost_db) = 0
    .value(eecar_HireCostMadeGood_db) = 0
    .value(eecar_EngineSize_db) = 0
    .value(eecar_FPCS_db) = 1 'rates.fpcs index 'default to IR
  End With
End Function

Private Property Set IBenefitForm2_benefit(NewValue As IBenefitClass)
  Set benefit = NewValue
End Property

Private Property Get IBenefitForm2_benefit() As IBenefitClass
  Set IBenefitForm2_benefit = benefit
End Property


Private Function IBenefitForm2_BenefitFormState(ByVal fState As BENEFIT_FORM_STATE) As Boolean

  On Error GoTo BenefitFormState_err
  
  Call xSet("IBenefitForm2_EnableBenefitForm")
  
  If (fState = FORM_ENABLED) Or (fState = FORM_CDB) Then
    If fState = FORM_ENABLED Then
      fmeCar.Enabled = True
    End If
    UBGRD.Grid.Enabled = True
    Call SetLVEnabled(lb, True)
    Call MDIMain.SetDelete
  ElseIf fState = FORM_DISABLED Then
    Set benefit = Nothing
    fmeCar.Enabled = False
    Call SetLVEnabled(lb, False)
    UBGRD.Grid.Enabled = False
    Call MDIMain.ClearDelete
    Call MDIMain.ClearConfirmUndo
  End If
  IBenefitForm2_BenefitFormState = True
    
BenefitFormState_end:
  Call xReturn("BenefitFormState")
  Exit Function
  
BenefitFormState_err:
  IBenefitForm2_BenefitFormState = False
  Call ErrorMessage(ERR_ERROR, Err, "BenefitFormState", "Benefit Form State", "Error setting the benefit form state for the employee owned car form.")
  Resume BenefitFormState_end

End Function


Private Function IBenefitForm2_BenefitOff() As Boolean
  Dim o As ObjectList
  
  On Error GoTo BenefitOff_ERR
  
  Call xSet("BenefitOff")
  
    TB_Data(0).Text = ""
    TB_Data(1).Text = ""
    TB_Data(2).Text = ""
    
    Op_Data(0).Enabled = False
    Op_Data(1).Enabled = False
    Op_Data(2).Enabled = False
    
    lblTotalOfExtras.Caption = ""
    lblTotalMiles = ""
    
    'UBGRD.Grid.ClearFields
    Me.lb.Enabled = True
    Set Me.UBGRD.ObjectList = Nothing
    'ubgrd.Grid.ReBind
    'Call grd.Refresh
    
    CB_FPCS.Text = ""
    CB_FPCS.Enabled = False

BenefitOff_END:
  Call xReturn("BenefitOff")
  Exit Function
BenefitOff_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "BenefitOff", "Benefit Off", "Error turning an employee owned car benefit off.")
  Resume BenefitOff_END
  Resume
End Function

Private Function IBenefitForm2_BenefitOn() As Boolean
  Dim eecar As EmployeeCar
  Dim FPCS As FPCS
  
  TB_Data(0).Text = benefit.value(eecar_Item_db)
  TB_Data(1).Text = benefit.value(eecar_EngineSize_db)
  
  'CAD review
  If Not benefit.value(eecar_UseCompanyCarScheme_db) Then
    TB_Data(2).Text = benefit.value(eecar_AmountReceived_db)
  Else
    TB_Data(2).Text = benefit.value(eecar_CompanyFPCSValue)
  End If
  
  Call AmountReceivedText(True)
  
  Set eecar = benefit
  
  lblTotalOfExtras = FormatWN(eecar.GetTotalOfExtrasPostMadeGood, "£")
  
  Op_Data(0).Enabled = True
  Op_Data(0) = IIf(benefit.value(eecar_UseCompanyCarScheme_db), vbChecked, vbUnchecked)
  Call UseFPCS(Op_Data(0))
'MP DB - Op_Data(1 & 2).visible are set to false below. And is not made visible anywhere
'  Op_Data(1).Enabled = True
'  Op_Data(1) = IIf(benefit.value(eecar_AverageIRRate_db), vbChecked, vbUnchecked)
  ' Op_Data(2).Enabled = True
  'EK No longer needed 2003 Op_Data(2) = IIf(benefit.value(eecar_AlternativeMethod), vbChecked, vbUnchecked)
    
  lblTotalMiles = benefit.value(eecar_TotalMiles)
  Set FPCS = p11d32.CurrentEmployer.FPCSchemes(benefit.value(eecar_FPCS_db))
  
  CB_FPCS.Text = FPCS.Name
  Set UBGRD.ObjectList = eecar.mileage
  'ubgrd.Grid.ReBind

  Set eecar = Nothing
  Set FPCS = Nothing
End Function

Private Function IBenefitForm2_BenefitsToListView() As Long
  Dim l As Long
  
  
  On Error GoTo BenefitsToListView_err
  
  Call xSet("BenefitsToListView")
  
  IBenefitForm2_BenefitsToListView = BenefitsToListView(Me)
  
BenefitsToListView_end:
  
  Call xSet("BenefitsToListView")
  Exit Function
BenefitsToListView_err:
  Call ErrorMessage(ERR_ERROR, Err, "BenefitsToListView", "Benefits To List View", "Unable to place the benefits to the list view.")
  Resume BenefitsToListView_end
End Function

Private Function IBenefitForm2_BenefitToListView(ben As IBenefitClass, ByVal lBenefitIndex As Long) As Long
  IBenefitForm2_BenefitToListView = BenefitToListView(ben, Me, lBenefitIndex)
End Function

Private Function IBenefitForm2_BenefitToScreen(Optional ByVal BenefitIndex As Long = -1&, Optional ByVal UpdateBenefit As Boolean = True) As Boolean
  IBenefitForm2_BenefitToScreen = BenefitToScreenHelper(Me, BenefitIndex, UpdateBenefit)
End Function

Private Property Let IBenefitForm2_benclass(ByVal NewValue As BEN_CLASS)
End Property

Private Property Get IBenefitForm2_benclass() As BEN_CLASS
  IBenefitForm2_benclass = BC_EMPLOYEE_CAR_E
End Property

'Private Property Get IBenefitForm2_ControlDefault() As Control
'  Set IBenefitForm2_ControlDefault = TB_Data(0)
'End Property

Private Property Get IBenefitForm2_lv() As MSComctlLib.IListView
  Set IBenefitForm2_lv = lb
End Property

Private Function IBenefitForm2_RemoveBenefit(ByVal BenefitIndex As Long) As Boolean
  
  IBenefitForm2_RemoveBenefit = p11d32.CurrentEmployer.CurrentEmployee.RemoveBenefit(Me, benefit, BenefitIndex)
End Function

Private Function IBenefitForm2_UpdateBenefitListViewItem(li As MSComctlLib.IListItem, benefit As IBenefitClass, Optional ByVal BenefitIndex As Long = 0&, Optional ByVal SelectItem As Boolean = False) As Long
  On Error GoTo UpdateBenefitListViewItemEmpoyeeCar_ERR
  
  Call xSet("UpdateBenefitListViewItemEmpoyeeCar")
  
  If Not li Is Nothing And Not benefit Is Nothing Then
    If BenefitIndex > 0 Then li.Tag = BenefitIndex
    li.SmallIcon = benefit.ImageListKey
    li.Text = benefit.Name
    li.SubItems(1) = FormatWN(benefit.Calculate)
    
    If SelectItem Then li.Selected = SelectItem
    IBenefitForm2_UpdateBenefitListViewItem = li.Index
  End If

UpdateBenefitListViewItemEmpoyeeCar_END:
  Call xReturn("UpdateBenefitListViewItemEmpoyeeCar")
  Exit Function
UpdateBenefitListViewItemEmpoyeeCar_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "UpdateBenefitListViewItemEmpoyeeCar", "Update Benefit List View Item Employee Car", "Error updating the benefit list view item for an employee owned car.")
  Resume UpdateBenefitListViewItemEmpoyeeCar_END


End Function

Private Function IBenefitForm2_ValididateBenefit(ben As IBenefitClass) As Boolean
  Select Case ben.BenefitClass
    Case BC_EMPLOYEE_CAR_E
      IBenefitForm2_ValididateBenefit = True
  End Select
End Function

Private Function IFrmGeneral_CheckChanged(c As Control) As Boolean
  Dim eecar As EmployeeCar
  Dim Band As FPCSBand
  Dim s As String
  Dim lst As ListItem
  Dim i As Long
  Dim bDirty As Boolean, bDirtyUseCompanyCarScheme As Boolean
  
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
      Case "TB_Data"
        Select Case .Index
          Case 0
            bDirty = CheckTextInput(.Text, benefit, eecar_Item_db)
          Case 1
            bDirty = CheckTextInput(.Text, benefit, eecar_EngineSize_db)
          Case 2
            bDirty = CheckTextInput(.Text, benefit, eecar_AmountReceived_db)
          Case Else
            ECASE "Unknown control"
        End Select
      Case "Op_Data"
        Select Case .Index
          Case 0
            bDirty = CheckCheckBoxInput(.value, benefit, eecar_UseCompanyCarScheme_db)
            Call UseFPCS(.value)
            bDirtyUseCompanyCarScheme = bDirty
'MP DB - commented case 1 and 2 - controls are never made visible
'          Case 1
'            bDirty = CheckCheckBoxInput(.value, benefit, eecar_AverageIRRate_db)
'            If bDirty And benefit.value(eecar_UseCompanyCarScheme_db) > 0 Then
'              bDirtyUseCompanyCarScheme = True
'            End If
'          Case 2
'            'EK No longer needed 2003 bDirty = CheckCheckBoxInput(.value, benefit, eecar_AlternativeMethod)
          Case Else
            ECASE "Unknown control"
        End Select
      Case "CB_FPCS"
        If StrComp(.Text, p11d32.CurrentEmployer.FPCSchemes(benefit.value(eecar_FPCS_db)).Name) Then
          benefit.value(eecar_FPCS_db) = CB_FPCS.ItemData(CB_FPCS.ListIndex)
          bDirty = True
        End If
      Case "lblTotalMilesForCar"
        bDirty = CheckTextInput(.Caption, benefit, eecar_TotalMiles)
      Case "UBGRD"
        bDirty = True
      Case Else
        ECASE "Unknown control"
    End Select
    
    'must be required in all check changed
    IFrmGeneral_CheckChanged = AfterCheckChanged(c, Me, bDirty)
    Call AmountReceivedText(bDirtyUseCompanyCarScheme)
    
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
  If Not (lb.SelectedItem Is Nothing) Then
    IBenefitForm2_BenefitToScreen (Item.Tag)
  End If
End Sub

Private Sub lb_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
  Call SetSortOrder(lb, ColumnHeader)
End Sub

Private Sub LB_KeyDown(KeyCode As Integer, Shift As Integer)
  Call LVKeyDown(KeyCode, Shift)
End Sub

Private Sub Op_Data_Click(Index As Integer)
  Call IFrmGeneral_CheckChanged(Op_Data(Index))
End Sub

Private Sub Op_Data_KeyPress(Index As Integer, KeyAscii As Integer)
  'Check for return key - tab to next field
  If KeyAscii = 13 Then Call SendKeys(vbTab)
End Sub


Private Sub TB_Data_FieldInvalid(Index As Integer, Valid As Boolean, Message As String)
  Call SetPanel2(Message)
End Sub


Private Function CheckCarAlternativeMethod(cB As CheckBoxConstants) As Boolean

On Error GoTo CheckCarAlternativeMethod_Err
  
  Call xSet("CheckCarAlternativeMethod")

  If cB = vbChecked Then
    L_Data(3).Visible = True
    L_Data(4).Visible = True
    L_Data(5).Visible = True
    L_Data(6).Visible = True
    L_Data(8).Visible = True
    TB_Data(3).Visible = True
    TB_Data(4).Visible = True
    TB_Data(5).Visible = True
    TB_Data(6).Visible = True
    TB_Data(8).Visible = True
  Else
    L_Data(3).Visible = False
    L_Data(4).Visible = False
    L_Data(5).Visible = False
    L_Data(6).Visible = False
    L_Data(8).Visible = False
    TB_Data(3).Visible = False
    TB_Data(4).Visible = False
    TB_Data(5).Visible = False
    TB_Data(6).Visible = False
    TB_Data(8).Visible = False
  End If

  CheckCarAlternativeMethod = True
  
CheckCarAlternativeMethod_End:
  Call xReturn("CheckCarAlternativeMethod")
  Exit Function

CheckCarAlternativeMethod_Err:
  CheckCarAlternativeMethod = False
  Call ErrorMessage(ERR_ERROR, Err, "CheckCarAlternativeMethod", "Check Car Alternative Method", "Unable to set the car alternative method form display.")
  Resume CheckCarAlternativeMethod_End
End Function

Private Sub TB_Data_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  TB_Data(Index).Tag = SetChanged
End Sub

Private Function UseFPCS(cbc As CheckBoxConstants) As Boolean
  Dim FPCS As FPCS
  
  On Error GoTo UseFPCS_Err
  Call xSet("UseFPCS")

  If cbc = vbChecked Then
    CB_FPCS.Enabled = True
    Set FPCS = p11d32.CurrentEmployer.FPCSchemes(benefit.value(eecar_FPCS_db))
    CB_FPCS.Text = FPCS.Name
    TB_Data(2).Enabled = False
    benefit.value(eecar_AmountReceived_db) = benefit.value(eecar_CompanyFPCSValue)
  Else
    Set FPCS = p11d32.CurrentEmployer.FPCSchemes(1)
    benefit.value(eecar_FPCS_db) = 1
    CB_FPCS.Text = FPCS.Name
    CB_FPCS.Enabled = False
    TB_Data(2).Enabled = True
  End If

UseFPCS_End:
  Set FPCS = Nothing
  Call xReturn("UseFPCS")
  Exit Function

UseFPCS_Err:
  Call ErrorMessage(ERR_ERROR, Err, "UseFPCS", "Use FPCS", "Error setting the employee car form when selecting the 'Use FPCS' option.")
  Resume UseFPCS_End
End Function


Public Function FPCSToCombo(bCheckCurrentFPCS As Boolean) As Boolean
  Dim l As Long
  Dim FPCS As FPCS
  
  On Error GoTo FPCSToCombo_Err
  Call xSet("FPCSToCombo")

  CB_FPCS.Clear
  
  With p11d32.CurrentEmployer
    For l = 1 To .FPCSchemes.Count
     Set FPCS = .FPCSchemes(l)
     If Not FPCS Is Nothing Then
       Call CB_FPCS.AddItem(FPCS.Name)
       CB_FPCS.ItemData(CB_FPCS.ListCount - 1) = l
     End If
    Next
   
   If bCheckCurrentFPCS Then
    Set FPCS = .FPCSchemes(benefit.value(eecar_FPCS_db))
    If Not FPCS Is Nothing Then
      CB_FPCS.Text = FPCS.Name
    Else
      Call ECASE("FPCS is nothing in FPCSToCombo")
    End If
   End If
  End With
  
  FPCSToCombo = True
FPCSToCombo_End:
  Set FPCS = Nothing
  Call xReturn("FPCSToCombo")
  Exit Function

FPCSToCombo_Err:
  Call ErrorMessage(ERR_ERROR, Err, "FPCSToCombo", "FPCS To Combo", "Error placing a FPCS to the employee car 'FPCS' combo box.")
  Resume FPCSToCombo_End
  Resume
End Function

Private Sub TB_Data_Validate(Index As Integer, Cancel As Boolean)
  Call IFrmGeneral_CheckChanged(TB_Data(Index))
End Sub

Private Sub ubgrd_DeleteData(ObjectList As ObjectList, ObjectListIndex As Long)
  Call MilesDelete(lblTotalMiles, benefit, eecar_TotalMiles, ObjectList, ObjectListIndex)
  Call IFrmGeneral_CheckChanged(UBGRD)
End Sub



Private Sub UBGRD_ReadData(RowBuf As TrueDBGrid60.RowBuffer, ByVal RowBufRowIndex As Long, ObjectList As ObjectList, ByVal ObjectListIndex As Long)
  Call MilesRead(RowBuf, RowBufRowIndex, ObjectList, ObjectListIndex)

End Sub

Private Sub UBGRD_ValidateTCS(FirstColIndexInError As Long, ValidateMessage As String, ByVal RowBuf As TrueDBGrid60.RowBuffer, ByVal RowBufRowIndex As Long, ByVal ObjectListIndex As Long)
  Call MilesValidate(FirstColIndexInError, ValidateMessage, RowBuf, RowBufRowIndex, ObjectListIndex)
End Sub

Private Sub ubgrd_WriteData(ByVal RowBuf As TrueDBGrid60.RowBuffer, ByVal RowBufRowIndex As Long, ObjectList As ObjectList, ObjectListIndex As Long)
  Call MilesWrite(lblTotalMiles, eecar_TotalMiles, RowBuf, RowBufRowIndex, ObjectList, ObjectListIndex, benefit)
  Call IFrmGeneral_CheckChanged(UBGRD)
End Sub


Public Function AmountReceivedText(bDirtyUseCompanyCarScheme As Boolean) As Boolean
  Dim eecar As EmployeeCar
  
  On Error GoTo AmountReceivedText_Err
  Call xSet("AmountReceivedText")

  If benefit.value(eecar_UseCompanyCarScheme_db) Then
    TB_Data(2).Text = benefit.value(eecar_CompanyFPCSValue)
    Set eecar = benefit
    If Not eecar.TopBand Is Nothing Then
      L_Data(2).Caption = S_AMOUNT_RECEIVED_PER_FPCS & " (top band: " & eecar.TopBand.Name & ", rate: £" & eecar.TopBand.Rate & ")"
    Else
      L_Data(2).Caption = S_AMOUNT_RECEIVED_PER_FPCS & " (top band: " & "no band)"
    End If
  Else
    'If bDirtyUseCompanyCarScheme Then TB_Data(2).Text = 0
    TB_Data(2).Text = benefit.value(eecar_AmountReceived_db)
    L_Data(2).Caption = S_AMOUNT_RECEIVED_DIRECT
  End If
  
  
AmountReceivedText_End:
  Set eecar = Nothing
  Call xReturn("AmountReceivedText")
  Exit Function

AmountReceivedText_Err:
  Call ErrorMessage(ERR_ERROR, Err, "AmountReceivedText", "Amount Received Text", "Error setting the amount received text per the selected FPCS.")
  Resume AmountReceivedText_End
  Resume
End Function

