VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.2#0"; "COMCTL32.OCX"
Object = "{4582CA9E-1A45-11D2-8D2F-00C04FA9DD6F}#1.0#0"; "ATC2VTEXT.OCX"
Begin VB.Form F_BenEECar0 
   Caption         =   "Employee Owned Cars"
   ClientHeight    =   5475
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8145
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   5475
   ScaleWidth      =   8145
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   120
      TabIndex        =   25
      Top             =   4440
      Width           =   3615
      Begin VB.CommandButton B_Actual 
         Appearance      =   0  'Flat
         Caption         =   "Change &Mileage.."
         Height          =   255
         Left            =   1800
         TabIndex        =   7
         Top             =   240
         Width           =   1455
      End
      Begin atc2valtext.ValText TB_Data 
         Height          =   285
         Index           =   9
         Left            =   120
         TabIndex        =   6
         Tag             =   "FREE,FONT"
         Top             =   600
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   503
         Enabled         =   0   'False
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
      Begin VB.Label L_Data 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Enter the mileage for each car whichever method is used."
         ForeColor       =   &H000000FF&
         Height          =   390
         Index           =   13
         Left            =   1320
         TabIndex        =   30
         Tag             =   "FREE,FONT"
         Top             =   480
         Width           =   2160
         WordWrap        =   -1  'True
      End
      Begin VB.Label L_Data 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Annual Mileage in this car"
         ForeColor       =   &H00800000&
         Height          =   390
         Index           =   9
         Left            =   120
         TabIndex        =   27
         Tag             =   "FREE,FONT"
         Top             =   120
         Width           =   1215
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Frame fmeCar 
      Height          =   3375
      Left            =   0
      TabIndex        =   15
      Top             =   2160
      Width           =   8115
      Begin VB.CheckBox Op_Data 
         Alignment       =   1  'Right Justify
         Caption         =   "Average IR Rates?"
         DataField       =   "BusMilesActual"
         DataSource      =   "DB"
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Tag             =   "free,font"
         Top             =   600
         Width           =   1755
      End
      Begin VB.ComboBox CB_FPCS 
         Appearance      =   0  'Flat
         DataField       =   "FPCS"
         DataSource      =   "DB"
         ForeColor       =   &H00000080&
         Height          =   315
         Left            =   6360
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Tag             =   "free,font"
         Top             =   1440
         Width           =   1635
      End
      Begin VB.CheckBox Op_Data 
         Alignment       =   1  'Right Justify
         Caption         =   "Alternative Method?"
         DataField       =   "BusMilesActual"
         DataSource      =   "DB"
         ForeColor       =   &H00800000&
         Height          =   525
         Index           =   1
         Left            =   3840
         TabIndex        =   8
         Tag             =   "free,font"
         Top             =   240
         Width           =   1125
      End
      Begin atc2valtext.ValText TB_Data 
         Height          =   285
         Index           =   1
         Left            =   2520
         TabIndex        =   5
         Tag             =   "FREE,FONT"
         Top             =   1920
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
         Text            =   "0"
         Minimum         =   "0"
         AllowEmpty      =   0   'False
         TXTAlign        =   2
      End
      Begin atc2valtext.ValText TB_Data 
         Height          =   285
         Index           =   4
         Left            =   6960
         TabIndex        =   11
         Tag             =   "FREE,FONT"
         Top             =   1920
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
         Text            =   "0"
         Minimum         =   "0"
         AllowEmpty      =   0   'False
         TXTAlign        =   2
      End
      Begin atc2valtext.ValText TB_Data 
         Height          =   285
         Index           =   0
         Left            =   2520
         TabIndex        =   4
         Tag             =   "FREE,FONT"
         Top             =   1440
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
         Text            =   "0"
         Minimum         =   "0"
         AllowEmpty      =   0   'False
         TXTAlign        =   2
      End
      Begin atc2valtext.ValText TB_Data 
         Height          =   315
         Index           =   7
         Left            =   600
         TabIndex        =   1
         Tag             =   "free,font"
         Top             =   240
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   556
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
      Begin atc2valtext.ValText TB_Data 
         Height          =   285
         Index           =   2
         Left            =   3000
         TabIndex        =   3
         Tag             =   "FREE,FONT"
         Top             =   600
         Width           =   585
         _ExtentX        =   1032
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
         Text            =   "0"
         Minimum         =   "0"
         AllowEmpty      =   0   'False
         TXTAlign        =   2
      End
      Begin atc2valtext.ValText TB_Data 
         Height          =   285
         Index           =   3
         Left            =   6960
         TabIndex        =   9
         Tag             =   "FREE,FONT"
         Top             =   960
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
         Text            =   "0"
         Minimum         =   "0"
         AllowEmpty      =   0   'False
         TXTAlign        =   2
      End
      Begin atc2valtext.ValText TB_Data 
         Height          =   285
         Index           =   5
         Left            =   6960
         TabIndex        =   12
         Tag             =   "FREE,FONT"
         Top             =   2280
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
         Text            =   "0"
         Minimum         =   "0"
         AllowEmpty      =   0   'False
         TXTAlign        =   2
      End
      Begin atc2valtext.ValText TB_Data 
         Height          =   285
         Index           =   6
         Left            =   6960
         TabIndex        =   13
         Tag             =   "FREE,FONT"
         Top             =   2640
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
         Text            =   "0"
         Minimum         =   "0"
         AllowEmpty      =   0   'False
         TXTAlign        =   2
      End
      Begin atc2valtext.ValText TB_Data 
         Height          =   285
         Index           =   8
         Left            =   6960
         TabIndex        =   14
         Tag             =   "FREE,FONT"
         Top             =   3000
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
         Caption         =   "Either enter the MA paid to the employee for this car or choose the company scheme under which payment was made."
         ForeColor       =   &H000000FF&
         Height          =   585
         Index           =   12
         Left            =   5160
         TabIndex        =   29
         Tag             =   "FREE,FONT"
         Top             =   240
         Width           =   2925
         WordWrap        =   -1  'True
      End
      Begin VB.Label L_Data 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Enter the Total Mileage done in all cars and the amount paid to the employee."
         ForeColor       =   &H000000FF&
         Height          =   390
         Index           =   11
         Left            =   120
         TabIndex        =   28
         Tag             =   "FREE,FONT"
         Top             =   960
         Width           =   3390
         WordWrap        =   -1  'True
      End
      Begin VB.Label L_Data 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Company Car Scheme"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   10
         Left            =   3840
         TabIndex        =   26
         Tag             =   "free,font"
         Top             =   1560
         Width           =   1695
         WordWrap        =   -1  'True
      End
      Begin VB.Label L_Data 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Any hire cost amount made good"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   8
         Left            =   3840
         TabIndex        =   24
         Tag             =   "FREE,FONT"
         Top             =   3000
         Width           =   2325
      End
      Begin VB.Label L_Data 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Hire Cost"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   6
         Left            =   3840
         TabIndex        =   23
         Tag             =   "FREE,FONT"
         Top             =   2640
         Width           =   645
      End
      Begin VB.Label L_Data 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "LumpSum"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   5
         Left            =   3840
         TabIndex        =   22
         Tag             =   "FREE,FONT"
         Top             =   2280
         Width           =   705
      End
      Begin VB.Label L_Data 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Amount made good/ subject to PAYE"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   3
         Left            =   3840
         TabIndex        =   21
         Tag             =   "FREE,FONT"
         Top             =   1920
         Width           =   2655
      End
      Begin VB.Label L_Data 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Engine Size"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   2
         Left            =   2040
         TabIndex        =   20
         Tag             =   "FREE,FONT"
         Top             =   600
         Width           =   840
      End
      Begin VB.Label L_Data 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Car"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   7
         Left            =   120
         TabIndex        =   19
         Tag             =   "free,font"
         Top             =   240
         Width           =   240
      End
      Begin VB.Label L_Data 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Mileage Allowance override"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   4
         Left            =   3840
         TabIndex        =   18
         Tag             =   "FREE,FONT"
         Top             =   960
         Width           =   1950
      End
      Begin VB.Label L_Data 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Total Mileage Allowance and Lumpsum (override)"
         ForeColor       =   &H00800000&
         Height          =   390
         Index           =   1
         Left            =   120
         TabIndex        =   17
         Tag             =   "FREE,FONT"
         Top             =   1800
         Width           =   2160
         WordWrap        =   -1  'True
      End
      Begin VB.Label L_Data 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Total Annual Mileage in all cars"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   16
         Tag             =   "FREE,FONT"
         Top             =   1440
         Width           =   2205
      End
   End
   Begin ComctlLib.ListView LB 
      Height          =   2175
      Left            =   0
      TabIndex        =   0
      Tag             =   "free,font"
      Top             =   0
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   3836
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
      NumItems        =   3
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Car Reference"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Alignment       =   1
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Benefit"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Alignment       =   1
         SubItemIndex    =   2
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "P/Y Value"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "F_BenEECar0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Implements IBenefitForm


Public benefit As IBenefitClass
Private mclsResize As New clsFormResize
Private Const L_DES_HEIGHT = 6090
Private Const L_DES_WIDTH = 8445

Private Sub B_BenFuel_Click()

End Sub

Private Sub B_Actual_Click()
  LoadingMiles = True
End Sub
Private Sub B_Actual_LostFocus()
  
  If Not LoadingMiles Then
    Call CheckChanged(B_Actual)
  End If
  
End Sub

Private Sub CB_FPCS_LostFocus()
  Call CheckChanged(CB_FPCS)
End Sub


Private Sub Form_Resize()
  mclsResize.Resize
    Call ColumnWidths(LB, 50, 25, 20)
End Sub


Private Sub IBenefitForm_AddBenefit()
  Dim beneecar As clsBenEECar
  Dim ben As IBenefitClass
  Dim lst As ListItem
  Dim I As Long
  On Error GoTo AddBenefit_Err
  Call xSet("AddBenefit")
  
  Set beneecar = New clsBenEECar
  Set ben = beneecar
  
  'Put in defaults for benefit ( zzzz Use defaults from the database?? - Good idea
  Set ben.Parent = CurrentEmployee
  Call ben.SetItem(eecar_EmployeeNumber, CurrentEmployee.PersonelNo)
  Call ben.SetItem(eecar_Item, "Please enter registration...")
  Call ben.SetItem(eecar_MadeGood, 0)
  Call ben.SetItem(eecar_AmountReceived, 0)
  Call ben.SetItem(eecar_LumpSum, 0)
  Call ben.SetItem(eecar_HireCost, 0)
  Call ben.SetItem(eecar_HireCostMadeGood, 0)
  Call ben.SetItem(eecar_EngineSize, 0)
  Call ben.SetItem(eecar_totalExtras, 0)
  
  Call MDIMain.SetConfirmUndo
  ben.Dirty = True
  ben.ReadFromDB = True
  I = CurrentEmployee.benefits.Add(ben)
  
  Set lst = LB.ListItems.Add(, , ben.Name)
  With lst
    .Tag = I
    .Text = ben.Name
    .SubItems(1) = formatworkingnumber(ben.Calculate, "£")
  End With
  
  Call EECarDetails(lst.Tag)
  Set LB.SelectedItem = lst
  'Me.fmeCar.Enabled = True
  Me.LB.Enabled = True
  Call MDIMain.SetDelete
  TB_Data(0).SetFocus
AddBenefit_End:
  Set ben = Nothing
  Call xReturn("AddBenefit")
  Exit Sub
AddBenefit_Err:
  Call ErrorMessage(ERR_ERROR, Err, "AddBenefit", "ERR_ADDBENEFIT", "Error in AddBenefit function, called from the form " & Me.Name & ".")
  Resume AddBenefit_End
End Sub

Private Function IBenefitForm_BenefitToScreen(Optional ByVal lTag As Long = -1, Optional ByVal lIndex As Long = -1&) As IBenefitClass

End Function

Private Property Let IBenefitForm_bentype(NewValue As benClass)
  ECASE "F_BenCar_BenType"
End Property

Private Property Get IBenefitForm_bentype() As benClass
  ECASE "F_BenCar_BenType"
End Property

Private Sub IBenefitForm_ClearFields()
  On Error GoTo ClearFields_Err
  Call xSet("ClearFields")
  With Me
    .TB_Data(0).Text = ""
    .TB_Data(1).Text = ""
    .TB_Data(2).Text = ""
    .TB_Data(3).Text = ""
    .TB_Data(4).Text = ""
    .TB_Data(5).Text = ""
    .TB_Data(6).Text = ""
    .TB_Data(7).Text = ""
    .TB_Data(8).Text = ""
    .TB_Data(9).Text = ""
    .Op_Data(0).value = vbUnchecked
    .Op_Data(1).value = vbUnchecked
    .CB_FPCS.Text = S_FPCS_DEFAULT
    
  End With
ClearFields_End:
  Call xReturn("ClearFields")
  Exit Sub
ClearFields_Err:
  Call ErrorMessage(ERR_ERROR, Err, "ClearFields", "ERR_UNDEFINED", "Undefined error.")
  Resume ClearFields_End
End Sub

Private Function IBenefitForm_Remove(I As Long) As Boolean
  On Error GoTo KillBenefit_Err
  Call xSet("KillBenefit")
  If Not benefit.CompanyDefined Then
    Call benefit.DeleteDB
    Call CurrentEmployee.benefits.Remove(I)
    Call IBenefitForm_ClearFields
    Call IBenefitForm_ListBenefits
  End If
KillBenefit_End:
  Call xReturn("KillBenefit")
  Exit Function
KillBenefit_Err:
  Call ErrorMessage(ERR_ERROR, Err, "KillBenefit", "ERR_UNDEFINED", "Undefined error.")
  Resume KillBenefit_End
End Function

Private Function IBenefitForm_ListBenefits(Optional ByVal Index As Long = 0&) As Boolean
  Dim ben As IBenefitClass
  Dim benfrm As IBenefitForm
  Dim lst As ListItem
  Dim I As Long, j As Long
  Dim lben As Variant
  
  On Error GoTo F_BenEECar_ListBenefits_Err
  Call xSet("F_BenEECar_ListBenefits")
  I = 0
  Call LockWindowUpdate(LB.hWnd)
  Set benfrm = Me
  Me.LB.ListItems.Clear
  benfrm.ClearFields
  Call MDIMain.SetAdd
  For I = 1 To CurrentEmployee.benefits.count
    Set ben = CurrentEmployee.benefits(I)
    If Not (ben Is Nothing) Then
      If ben.BenefitClass = BC_EECAR Then
        lben = ben.Calculate
        Set lst = LB.ListItems.Add(, , ben.Name)
        lst.Tag = I
        If VarType(lben) = vbString Then
          lst.SubItems(1) = lben
        Else
          lst.SubItems(1) = formatworkingnumber(lben, "£")
        End If
        j = 1
      End If
    End If
  Next I
  If j = 0 Then
    'Me.fmeCar.Enabled = False
    Me.LB.Enabled = False
    Call MDIMain.ClearDelete
    Call MDIMain.ClearConfirmUndo
    Set Me.benefit = Nothing
  Else
    'Me.fmeCar.Enabled = True
    Me.LB.Enabled = True
    Call MDIMain.SetDelete
    LB.SelectedItem = LB.ListItems(1)
    Call EECarDetails(LB.SelectedItem.Tag)
  End If
F_BenEECar_ListBenefits_End:
  Call LockWindowUpdate(0)
  Call xReturn("F_BenEECar_ListBenefits")
  Exit Function
  
F_BenEECar_ListBenefits_Err:
  Call ErrorMessage(ERR_ERROR, Err, "F_BenEECar_ListBenefits", "EECar Benefits", "Error listing benefits.")
  Resume F_BenEECar_ListBenefits_End
End Function

Private Sub Form_Load()
  
  Dim I As Long
  Dim CurScheme As String
  
  If Not (mclsResize.InitResize(Me, L_DES_HEIGHT, L_DES_WIDTH, DESIGN, , , MDIMain)) Then
    Err.Raise ERR_Application
  End If
  
  CurScheme = ""
  For I = 1 To CurrentEmployer.FPCSBands.count
    If CurrentEmployer.GetFPCSSchemeName(I) <> CurScheme Then
      CB_FPCS.AddItem CurrentEmployer.GetFPCSSchemeName(I)
      CurScheme = CurrentEmployer.GetFPCSSchemeName(I)
    End If
  Next I
  
  
End Sub

Private Function CheckChanged(ctl As Control) As Boolean
  Dim lst As ListItem
  Dim I As Long, d0 As Variant
  On Error GoTo CheckChanged_Err
  Call xSet("CheckChanged")
  
  If CurrentEmployee Is Nothing Then GoTo CheckChanged_End
  If benefit Is Nothing Then GoTo CheckChanged_End
  Select Case ctl.Name
    Case "TB_Data"
      Select Case ctl.Index
        Case 0
          I = StrComp(ctl.Text, CurrentEmployee.TotalEECarMiles, vbBinaryCompare)
          If I <> 0 Then CurrentEmployee.TotalEECarMiles = CLng(ctl.Text)
        Case 1
          I = StrComp(ctl.Text, CurrentEmployee.TotalEEcarAllowance, vbBinaryCompare)
          'Debug.Print "CheckChanged acted on " & ctl.Text
          If I <> 0 Then CurrentEmployee.TotalEEcarAllowance = CLng(ctl.Text)
        Case 2
          I = StrComp(ctl.Text, benefit.GetItem(eecar_EngineSize), vbBinaryCompare)
          If I <> 0 Then Call benefit.SetItem(eecar_EngineSize, ctl.Text)
        Case 3
          I = StrComp(ctl.Text, benefit.GetItem(eecar_MadeGood), vbBinaryCompare)
          If I <> 0 Then Call benefit.SetItem(eecar_MadeGood, ctl.Text)
        Case 4
          I = StrComp(ctl.Text, benefit.GetItem(eecar_AmountReceived), vbBinaryCompare)
          If I <> 0 Then Call benefit.SetItem(eecar_AmountReceived, ctl.Text)
        Case 5
          I = StrComp(ctl.Text, benefit.GetItem(eecar_LumpSum), vbBinaryCompare)
          If I <> 0 Then Call benefit.SetItem(eecar_LumpSum, ctl.Text)
        Case 6
          I = StrComp(ctl.Text, benefit.GetItem(eecar_HireCost), vbBinaryCompare)
          If I <> 0 Then Call benefit.SetItem(eecar_HireCost, ctl.Text)
        Case 7
          I = StrComp(ctl.Text, benefit.GetItem(eecar_Item), vbBinaryCompare)
          If I <> 0 Then Call benefit.SetItem(eecar_Item, ctl.Text)
        Case 8
          I = StrComp(ctl.Text, benefit.GetItem(eecar_HireCostMadeGood), vbBinaryCompare)
          If I <> 0 Then Call benefit.SetItem(eecar_HireCostMadeGood, ctl.Text)
        
        Case Else
          ECASE "Unknown control"
      End Select
    Case "Op_Data"
      Select Case ctl.Index
        Case 0
          I = (IIf(ctl.value = vbChecked, True, False) <> benefit.GetItem(eecar_AverageIRRate))
          If I <> 0 Then Call benefit.SetItem(eecar_AverageIRRate, IIf(ctl.value = vbChecked, True, False))
        Case 1
          I = (IIf(ctl.value = vbChecked, True, False) <> benefit.GetItem(eecar_AlternativeMethod))
          If I <> 0 Then
            Call benefit.SetItem(eecar_AlternativeMethod, IIf(ctl.value = vbChecked, True, False))
            With F_BenEECar
              If ctl.value = vbChecked Then
                .L_Data(3).Visible = True
                .L_Data(4).Visible = True
                .L_Data(5).Visible = True
                .L_Data(6).Visible = True
                .L_Data(8).Visible = True
                .TB_Data(3).Visible = True
                .TB_Data(4).Visible = True
                .TB_Data(5).Visible = True
                .TB_Data(6).Visible = True
                .TB_Data(8).Visible = True
              Else
                .L_Data(3).Visible = False
                .L_Data(4).Visible = False
                .L_Data(5).Visible = False
                .L_Data(6).Visible = False
                .L_Data(8).Visible = False
                .TB_Data(3).Visible = False
                .TB_Data(4).Visible = False
                .TB_Data(5).Visible = False
                .TB_Data(6).Visible = False
                .TB_Data(8).Visible = False
              End If
            End With
          End If
        Case Else
          ECASE "Unknown control"
      End Select
    Case "CB_FPCS"
      I = StrComp(ctl.Text, benefit.GetItem(eecar_FPCS), vbBinaryCompare)
      If I <> 0 Then Call benefit.SetItem(eecar_FPCS, ctl.Text)
      Me.CB_FPCS = benefit.GetItem(eecar_FPCS)
    Case "B_Actual"
      I = StrComp(F_BenEECar.TB_Data(9).Text, benefit.GetItem(eecar_TotalMiles), vbBinaryCompare)
      If I <> 0 Then Call benefit.SetItem(eecar_TotalMiles, F_BenEECar.TB_Data(9).Text)
    Case Else
      ECASE "Unknown control"
  End Select
  If I <> 0 Then benefit.InvalidFields = InvalidFields(Me)
  If benefit.InvalidFields > 0 Then
    Call MDIMain.sts.SetStatus(0, "", S_NOSAVE)
    Call MDIMain.SetUndo
    benefit.Dirty = False
    Set lst = LB.SelectedItem
    If Not lst Is Nothing Then
      With lst
        .Text = benefit.Name
        .SubItems(1) = formatworkingnumber(benefit.Calculate, "£")
      End With
    End If
  ElseIf I <> 0 Then
    Call MDIMain.sts.SetStatus(0, "", "")
    Call MDIMain.SetConfirmUndo
    benefit.Dirty = True
    Set lst = LB.SelectedItem
    With lst
      .Text = benefit.Name
      .SubItems(1) = formatworkingnumber(benefit.Calculate, "£")
    End With
  ElseIf benefit.Dirty Then
    Call MDIMain.sts.SetStatus(0, "", "")
    Call MDIMain.SetConfirmUndo
  End If
  
CheckChanged_End:
  Call xReturn("CheckChanged")
  Exit Function

CheckChanged_Err:
  Call ErrorMessage(ERR_ERROR, Err, "CheckChanged", "ERR_CHECKCHANGED", "This function has failed for the form " & Me.Name & ".")
  Resume CheckChanged_End
End Function

Private Sub LB_ItemClick(ByVal Item As ComctlLib.ListItem)
  If Not (LB.SelectedItem Is Nothing) Then
    Call EECarDetails(Item.Tag)
  End If
End Sub

Private Sub LB_ColumnClick(ByVal ColumnHeader As ComctlLib.ColumnHeader)
  LB.SortKey = ColumnHeader.Index - 1
  LB.SelectedItem.EnsureVisible
End Sub

Private Sub LB_KeyPress(KeyAscii As Integer)
  'Check for return key - tab to next field
  If KeyAscii = 13 Then Call SendKeys(vbTab)
End Sub

Private Sub Op_Data_Click(Index As Integer)
  Call CheckChanged(Op_Data(Index))
End Sub

Private Sub Op_Data_KeyPress(Index As Integer, KeyAscii As Integer)
  'Check for return key - tab to next field
  If KeyAscii = 13 Then Call SendKeys(vbTab)
End Sub


Private Sub TB_Data_FieldInvalid(Index As Integer, Valid As Boolean, Message As String)
  Call MDIMain.sts.SetStatus(0, Message)
End Sub

Private Sub TB_data_Lostfocus(Index As Integer)
  'Debug.Print "CheckChanged called"
  Call CheckChanged(TB_Data(Index))
End Sub

Private Sub Op_Data_LostFocus(Index As Integer)
  Call CheckChanged(Op_Data(Index))
End Sub


