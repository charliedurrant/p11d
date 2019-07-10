VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{8D988532-0F0C-460C-B00E-7B5637E97680}#1.0#0"; "ATC2VTEXT.OCX"
Begin VB.Form F_NonSharedVans 
   Caption         =   "Vans"
   ClientHeight    =   5925
   ClientLeft      =   3375
   ClientTop       =   555
   ClientWidth     =   8460
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   ScaleHeight     =   5925
   ScaleWidth      =   8460
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Frame fmeSharedVans 
      Caption         =   "Shared vans"
      ForeColor       =   &H8000000D&
      Height          =   1710
      Left            =   45
      TabIndex        =   11
      Tag             =   "FREE,FONT"
      Top             =   45
      Width           =   8385
      Begin VB.CheckBox chkbx 
         Alignment       =   1  'Right Justify
         Caption         =   "Report daily calculation benefit to employee?"
         DataField       =   "DailyCalc"
         DataSource      =   "DB_Share"
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   1
         Left            =   4215
         TabIndex        =   4
         Tag             =   "FREE,FONT"
         Top             =   300
         Width           =   3885
      End
      Begin VB.CheckBox chkbx 
         Alignment       =   1  'Right Justify
         Caption         =   "Were one or more shared vans available?"
         DataField       =   "SharedVans"
         DataSource      =   "DB_Share"
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Tag             =   "FREE,FONT"
         Top             =   300
         Width           =   3765
      End
      Begin VB.CheckBox chkbx 
         Alignment       =   1  'Right Justify
         Caption         =   "Was a non-shared van available at the same time as any shared van?"
         DataField       =   "TwoPlusVans"
         DataSource      =   "DB_Share"
         ForeColor       =   &H00800000&
         Height          =   405
         Index           =   2
         Left            =   120
         TabIndex        =   2
         Tag             =   "FREE,FONT"
         Top             =   700
         Width           =   3765
      End
      Begin VB.CommandButton B_Van 
         Height          =   288
         Left            =   7125
         MaskColor       =   &H8000000F&
         Picture         =   "Benvan.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   6
         Tag             =   "FREE,FONT"
         Top             =   1250
         Width           =   975
      End
      Begin atc2valtext.ValText TxtBx 
         Height          =   315
         Index           =   1
         Left            =   2895
         TabIndex        =   3
         Tag             =   "FREE,FONT"
         Top             =   1250
         Width           =   1005
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
         MouseIcon       =   "Benvan.frx":106A
         Text            =   "0"
         Minimum         =   "0"
         AllowEmpty      =   0   'False
         TXTAlign        =   2
      End
      Begin atc2valtext.ValText TxtBx 
         Height          =   315
         Index           =   0
         Left            =   7125
         TabIndex        =   5
         Tag             =   "FREE,FONT"
         Top             =   700
         Width           =   975
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
         MouseIcon       =   "Benvan.frx":1086
         Text            =   "0"
         Maximum         =   "365"
         Minimum         =   "0"
         AllowEmpty      =   0   'False
         TXTAlign        =   2
      End
      Begin VB.Label Lab 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "(0 vans available for share)"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   11
         Left            =   4275
         TabIndex        =   15
         Tag             =   "FREE,FONT"
         Top             =   1400
         Width           =   2535
      End
      Begin VB.Label Lab 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "If so, no. of relevant days"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   6
         Left            =   4245
         TabIndex        =   14
         Tag             =   "FREE,FONT"
         Top             =   700
         Width           =   2655
      End
      Begin VB.Label Lab 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Actual amount made good, or amount subjected to PAYE"
         ForeColor       =   &H00800000&
         Height          =   390
         Index           =   10
         Left            =   120
         TabIndex        =   13
         Tag             =   "FREE,FONT"
         Top             =   1200
         Width           =   2535
         WordWrap        =   -1  'True
      End
      Begin VB.Label Lab 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Click here for shared vans schedule:"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   9
         Left            =   4245
         TabIndex        =   12
         Tag             =   "FREE,FONT"
         Top             =   1200
         Width           =   2910
      End
   End
   Begin VB.Frame fmeDivider 
      Height          =   45
      Index           =   1
      Left            =   90
      TabIndex        =   10
      Tag             =   "free"
      Top             =   5400
      Width           =   8070
   End
   Begin VB.Frame fmeInput 
      Caption         =   "Non shared vans"
      ForeColor       =   &H8000000D&
      Height          =   3630
      Left            =   45
      TabIndex        =   9
      Tag             =   "FREE,FONT"
      Top             =   1785
      Width           =   8370
      Begin VB.Frame frms155a 
         Caption         =   "Vans for which the restricted private use condition is met, s155(2)(a)"
         ForeColor       =   &H8000000D&
         Height          =   495
         Left            =   120
         TabIndex        =   30
         Tag             =   "FREE,FONT"
         Top             =   3000
         Width           =   8055
         Begin VB.CheckBox chkbx 
            Alignment       =   1  'Right Justify
            Caption         =   "Mainly available for business travel only"
            DataField       =   "TwoPlusVans"
            DataSource      =   "DB_Share"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   6
            Left            =   4560
            TabIndex        =   32
            Tag             =   "FREE,FONT"
            Top             =   240
            Width           =   3240
         End
         Begin VB.CheckBox chkbx 
            Alignment       =   1  'Right Justify
            Caption         =   "Substantially available for Ordinary Commuting only?"
            DataField       =   "TwoPlusVans"
            DataSource      =   "DB_Share"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   4
            Left            =   240
            TabIndex        =   31
            Tag             =   "FREE,FONT"
            Top             =   240
            Width           =   3960
         End
      End
      Begin VB.CheckBox chkbx 
         Alignment       =   1  'Right Justify
         Caption         =   "Is the amount above an amount subjected to PAYE?"
         ForeColor       =   &H00800000&
         Height          =   350
         Index           =   5
         Left            =   180
         TabIndex        =   24
         Tag             =   "free,font"
         Top             =   2610
         Width           =   3855
      End
      Begin MSComctlLib.ListView lb 
         Height          =   2730
         Left            =   4275
         TabIndex        =   0
         Tag             =   "free,font"
         Top             =   180
         Width           =   3915
         _ExtentX        =   6906
         _ExtentY        =   4815
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
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Van Reference"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Benefit"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.CheckBox chkbx 
         Alignment       =   1  'Right Justify
         Caption         =   "Was the van registered on or after 06/04/95"
         DataField       =   "TwoPlusVans"
         DataSource      =   "DB_Share"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   3
         Left            =   180
         TabIndex        =   20
         Tag             =   "FREE,FONT"
         Top             =   1810
         Width           =   3840
      End
      Begin atc2valtext.ValText TxtBx 
         Height          =   315
         Index           =   7
         Left            =   2925
         TabIndex        =   19
         Tag             =   "FREE,FONT"
         Top             =   1410
         Width           =   1125
         _ExtentX        =   0
         _ExtentY        =   0
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
         MouseIcon       =   "Benvan.frx":10A2
         Text            =   ""
         Maximum         =   "365"
         Minimum         =   "0"
         AllowEmpty      =   0   'False
         TXTAlign        =   2
      End
      Begin atc2valtext.ValText TxtBx 
         Height          =   315
         Index           =   6
         Left            =   2940
         TabIndex        =   22
         Tag             =   "FREE,FONT"
         Top             =   2210
         Width           =   1125
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
         MouseIcon       =   "Benvan.frx":10BE
         Text            =   "0"
         AllowEmpty      =   0   'False
         TXTAlign        =   2
      End
      Begin atc2valtext.ValText TxtBx 
         Height          =   315
         Index           =   4
         Left            =   2925
         TabIndex        =   17
         Tag             =   "FREE,FONT"
         Top             =   610
         Width           =   1125
         _ExtentX        =   0
         _ExtentY        =   0
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
         MouseIcon       =   "Benvan.frx":10DA
         Text            =   ""
         TypeOfData      =   2
         Maximum         =   "5/4/99"
         Minimum         =   "6/4/98"
         AllowEmpty      =   0   'False
      End
      Begin atc2valtext.ValText TxtBx 
         Height          =   315
         Index           =   5
         Left            =   2925
         TabIndex        =   18
         Tag             =   "FREE,FONT"
         Top             =   1010
         Width           =   1125
         _ExtentX        =   0
         _ExtentY        =   0
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
         MouseIcon       =   "Benvan.frx":10F6
         Text            =   ""
         TypeOfData      =   2
         Maximum         =   "5/4/99"
         Minimum         =   "6/4/98"
         AllowEmpty      =   0   'False
      End
      Begin atc2valtext.ValText TxtBx 
         Height          =   315
         Index           =   3
         Left            =   1845
         TabIndex        =   16
         Tag             =   "FREE,FONT"
         Top             =   210
         Width           =   2205
         _ExtentX        =   3889
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
         MouseIcon       =   "Benvan.frx":1112
         Text            =   ""
         TypeOfData      =   3
         AllowEmpty      =   0   'False
      End
      Begin atc2valtext.ValText TxtBx 
         Height          =   315
         Index           =   21
         Left            =   2925
         TabIndex        =   21
         Tag             =   "FREE,FONT"
         Top             =   1810
         Width           =   1125
         _ExtentX        =   0
         _ExtentY        =   0
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
         MouseIcon       =   "Benvan.frx":112E
         Text            =   ""
         TypeOfData      =   2
         AllowEmpty      =   0   'False
      End
      Begin VB.Label Lab 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Van reference"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   14
         Left            =   180
         TabIndex        =   29
         Tag             =   "FREE,FONT"
         Top             =   210
         Width           =   1005
      End
      Begin VB.Label Lab 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Actual amount made good, or amount subjected to PAYE"
         ForeColor       =   &H00800000&
         Height          =   390
         Index           =   18
         Left            =   180
         TabIndex        =   28
         Tag             =   "FREE,FONT"
         Top             =   2210
         Width           =   2535
         WordWrap        =   -1  'True
      End
      Begin VB.Label Lab 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Qualifying days unavailable or shared"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   19
         Left            =   180
         TabIndex        =   27
         Tag             =   "FREE,FONT"
         Top             =   1410
         Width           =   2985
      End
      Begin VB.Label Lab 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Available to"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   17
         Left            =   180
         TabIndex        =   26
         Tag             =   "FREE,FONT"
         Top             =   1010
         Width           =   2655
      End
      Begin VB.Label Lab 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Available from"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   16
         Left            =   180
         TabIndex        =   25
         Tag             =   "FREE,FONT"
         Top             =   610
         Width           =   990
      End
      Begin VB.Label Lab 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date of Registration"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   20
         Left            =   180
         TabIndex        =   23
         Tag             =   "FREE,FONT"
         Top             =   1810
         Width           =   1410
      End
   End
   Begin VB.Label Lab 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "£00000"
      DataField       =   "BenvanTotal"
      DataSource      =   "DB_Share"
      ForeColor       =   &H00800000&
      Height          =   285
      Index           =   1
      Left            =   6840
      TabIndex        =   7
      Tag             =   "FREE,FONT"
      Top             =   5520
      Width           =   1305
   End
   Begin VB.Label Lab 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Van benefit (shared and non-shared)"
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   2
      Left            =   120
      TabIndex        =   8
      Tag             =   "FREE,FONT"
      Top             =   5520
      Width           =   2580
   End
End
Attribute VB_Name = "F_NonSharedVans"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Implements IBenefitForm
Implements IBenefitForm2
Implements IFrmGeneral

Public benefit As IBenefitClass
Private mclsResize As New clsFormResize
Private m_InvalidVT As Control
Private m_BenClass As BEN_CLASS

Private m_SettingFormState As Boolean
Private Const L_DES_HEIGHT  As Long = 6090
Private Const L_DES_WIDTH  As Long = 8445

Private Sub B_Van_Click()
  Call ToolBarButton(TBR_SHAREDVANS, F_Employees.lb.SelectedItem.Tag)
End Sub

Private Sub ChkBx_Click(Index As Integer)
  Call IFrmGeneral_CheckChanged(chkbx(Index))
End Sub

Private Sub Form_Resize()
  mclsResize.Resize
  Call ColumnWidths(lb, 70, 30)
End Sub

Private Sub Form_Load()
'MP DB - sets visible=false below, so no need to change caption
'MP DB  chkbx(3).Caption = "Was the van registered on or after " & Format$(p11d32.Rates.value(vanOldDate), "dd/mm/yyyy")
  If Not (mclsResize.InitResize(Me, L_DES_HEIGHT, L_DES_WIDTH, DESIGN, , , MDIMain)) Then
    Err.Raise ERR_Application
  End If
  Call SetDefaultVTDate(TxtBx(4))
  Call SetDefaultVTDate(TxtBx(5))
  TxtBx(0).Maximum = p11d32.Rates.value(DaysInYearLeap)
  TxtBx(7).Maximum = p11d32.Rates.value(DaysInYearLeap)
  
'  If p11d32.AppYear = 2000 Then 'km
'    TxtBx(21).Visible = False
'    Lab(20).Visible = False
'  Else
'MP DB ToDo - should delete chkbx(3) control from form & remove above caption stmnt
    chkbx(3).Visible = False
'  End If
  
End Sub
Private Property Let IBenefitForm_benclass(NewValue As BEN_CLASS)
  ECASE "F_CompanyCar_benclass"
End Property

Private Property Get IBenefitForm_benclass() As BEN_CLASS
  ECASE "F_CompanyCar_benclass"
End Property
Private Sub IBenefitForm2_AddBenefit()
  Dim ben As IBenefitClass
  Dim lst As ListItem, i As Long
  Dim ibf As IBenefitForm2
  Dim NonSharedVans As NonSharedVans
  
  On Error GoTo AddBenefit_Err
  Call xSet("AddBenefit")
  
  Set ben = New NonSharedVan
  Set ibf = Me
  With ben
    .BenefitClass = BC_nonSHAREDVAN_G
    Set .Parent = p11d32.CurrentEmployer.CurrentEmployee.NonSharedVans
    Call ibf.AddBenefitSetDefaults(ben)
  End With
  
  Call MDIMain.SetConfirmUndo
  
  ben.ReadFromDB = True
  
  Set NonSharedVans = p11d32.CurrentEmployer.CurrentEmployee.NonSharedVans
  i = NonSharedVans.Vans.Add(ben)
  Set lst = lb.listitems.Add(, , ben.Name)
  Call ibf.UpdateBenefitListViewItem(lst, ben, i, True)
  ben.Dirty = True
  
  Call SelectBenefitByListItem(ibf, lst)
  
AddBenefit_End:
  Set NonSharedVans = Nothing
  Set ibf = Nothing
  Set ben = Nothing
  Call xReturn("AddBenefit")
  Exit Sub
AddBenefit_Err:
  Call ErrorMessage(ERR_ERROR, Err, "AddBenefit", "ERR_ADDBENEFIT", "Error in the AddBenefit function, called from the form " & Me.Name & ".")
  Resume AddBenefit_End
  Resume

End Sub

Private Function IBenefitForm2_AddBenefitSetDefaults(ben As IBenefitClass) As Long
  Dim benParent As IBenefitClass
  
  
  With ben
    Call StandardReadData(ben)
    Set benParent = ben.Parent
    Call SetAvaialbleRange(ben, benParent.Parent, nsvan_AvailableFrom_db, nsvan_Availableto_db)
    
    .value(nsvan_item_db) = "Please enter description..."
    .value(nsvan_DaysUnavailable_db) = 0
    .value(nsvan_madegood_db) = 0
'MP DB    .value(nsvan_LessThanOrEqualT4YearsOldAtEndOfTaxYear) = True
    .value(nsvan_RegistrationDate_db) = p11d32.Rates.value(VanRegDateNew)
  End With
End Function

Private Property Set IBenefitForm2_benefit(NewValue As IBenefitClass)
  Set benefit = NewValue
End Property

Private Property Get IBenefitForm2_benefit() As IBenefitClass
  If benefit Is Nothing And (Not m_SettingFormState) Then
    Set IBenefitForm2_benefit = p11d32.CurrentEmployer.CurrentEmployee.NonSharedVans
  Else
    Set IBenefitForm2_benefit = benefit
  End If
End Property

Private Function IBenefitForm2_BenefitFormState(ByVal fState As BENEFIT_FORM_STATE) As Boolean
  IBenefitForm2_BenefitFormState = BenefitFormStateEx(fState, benefit, fmeInput)
End Function
Private Function IBenefitForm2_BenefitOff() As Boolean
  TxtBx(3).Text = ""
  TxtBx(4).Text = ""
  TxtBx(5).Text = ""
  TxtBx(6).Text = ""
  TxtBx(7).Text = ""
'  If p11d32.AppYear > 2000 Then
  TxtBx(21).Text = "" 'km
End Function
Private Function IBenefitForm2_BenefitOn() As Boolean
  With benefit
    TxtBx(3).Text = .value(nsvan_item_db)
    TxtBx(4).Text = DateValReadToScreen(.value(nsvan_AvailableFrom_db))
    TxtBx(5).Text = DateValReadToScreen(.value(nsvan_Availableto_db))
    TxtBx(6).Text = DateValReadToScreen(.value(nsvan_madegood_db))
    TxtBx(7).Text = .value(nsvan_DaysUnavailable_db)
    
'    If p11d32.AppYear > 2000 Then 'km
     'cad review 28/06/2002 no need to do as defaults to value in readdb
     'If Not IsNull(.value(nsvan_RegistrationDate)) Then
        TxtBx(21).Text = DateValReadToScreen(.value(nsvan_RegistrationDate_db))
      'End If
'    End If
    chkbx(5).value = BoolToChkBox(.value(nsvan_MadeGoodIsTaxDeducted_db))
    chkbx(4).value = BoolToChkBox(.value(nsvan_commuter_use_req_db))
    chkbx(6).value = BoolToChkBox(.value(nsvan_availablle_for_bus_use_only))
    
  End With
End Function
Private Function IBenefitForm2_BenefitsToListView() As Long
  Dim i As Long
  Dim ibf As IBenefitForm2
  Dim ben As IBenefitClass
  Dim NonSharedVans As NonSharedVans
  Dim bc As BEN_CLASS
  
On Error GoTo BenefitsToListView_err
  
  Call xSet("BenefitsToListView")
  
  Set ibf = Me
  Call ClearForm(ibf)
  Call MDIMain.SetAdd
  
  Set NonSharedVans = p11d32.CurrentEmployer.CurrentEmployee.benefits(L_VANS_BENINDEX)
  bc = BC_nonSHAREDVAN_G
  
  For i = 1 To NonSharedVans.Vans.Count
    Set ben = NonSharedVans.Vans(i)
    IBenefitForm2_BenefitsToListView = IBenefitForm2_BenefitsToListView + ibf.BenefitToListView(ben, i)
  Next i
  
  Set ben = p11d32.CurrentEmployer.CurrentEmployee
  TxtBx(0).Text = ben.value(ee_RelevantDaysForDailySharedVanCalc_db)
  TxtBx(1).Text = ben.value(ee_PaymentsForPrivateUseOfSharedVans_db)
  chkbx(0) = IIf(ben.value(ee_OneOrMoreSharedVanAvailable_db), vbChecked, vbUnchecked)
  chkbx(1) = IIf(ben.value(ee_ReportyDailyCalculationOfSharedVans_db), vbChecked, vbUnchecked)
  chkbx(2) = IIf(ben.value(ee_NonSharedVanAvailableAtSameTimeAsSharedVan_db), vbChecked, vbUnchecked)
  
BenefitsToListView_end:
  Set ben = Nothing
  Set ibf = Nothing
  Set NonSharedVans = Nothing
  Call xReturn("BenefitsToListView")
  Exit Function
BenefitsToListView_err:
  Call ErrorMessage(ERR_ERROR, Err, "BenefitsToListView", "Adding Benefits to ListView", "Error in adding benefits to listview.")
  Resume BenefitsToListView_end
  Resume
End Function

Private Function IBenefitForm2_BenefitToListView(ben As IBenefitClass, ByVal lBenefitIndex As Long) As Long
  IBenefitForm2_BenefitToListView = BenefitToListView(ben, Me, lBenefitIndex)
End Function

Private Function IBenefitForm2_BenefitToScreen(Optional ByVal BenefitIndex As Long = -1&, Optional ByVal UpdateBenefit As Boolean = True) As Boolean
  Dim ibf As IBenefitForm2
  Dim NonSharedVans As NonSharedVans
  Dim ben As IBenefitClass
  Dim ee As Employee
  
  On Error GoTo NonSharedVanBenefitToScreen_Err
  Call xSet("NonSharedVanBenefitToScreen")
  
  Set ibf = Me
  
  If UpdateBenefit Then Call UpdateBenefitFromTags
  
  If BenefitIndex <> -1 Then
    Set NonSharedVans = p11d32.CurrentEmployer.CurrentEmployee.NonSharedVans
    Set ben = NonSharedVans.Vans(BenefitIndex)
    Set ibf.benefit = ben
    Call ibf.BenefitOn
  Else
    Set ibf.benefit = Nothing
    Call ibf.BenefitOff
  End If
  Call UpdateTotalVansBenefit
  m_SettingFormState = True
    
  Call SetBenefitFormState(ibf)
  IBenefitForm2_BenefitToScreen = True
  
NonSharedVanBenefitToScreen_End:
  m_SettingFormState = True
  Call xReturn("NonSharedVanBenefitToScreen")
  Exit Function
NonSharedVanBenefitToScreen_Err:
  Call ErrorMessage(ERR_ERROR, Err, "NonSharedVanBenefitToScreen", "Non Shared Van Benefit To Screen", "Unable to place the non shared van benefit onto the screen. Benefit index = " & BenefitIndex & ".")
  Resume NonSharedVanBenefitToScreen_End
  Resume
End Function


Private Property Let IBenefitForm2_benclass(ByVal NewValue As BEN_CLASS)
  
End Property

Private Property Get IBenefitForm2_benclass() As BEN_CLASS
  IBenefitForm2_benclass = BC_NONSHAREDVANS_G
End Property


Private Property Get IBenefitForm2_lv() As MSComctlLib.IListView
  Set IBenefitForm2_lv = lb
End Property

Private Function IBenefitForm2_RemoveBenefit(ByVal BenefitIndex As Long) As Boolean
  Dim NonSharedVans As NonSharedVans
  Dim ben As IBenefitClass
  Dim NextBenefitIndex As Long
  Dim ibf As IBenefitForm2
  
  Call xSet("NonSharedVan_RemoveBenefit")
  
  On Error GoTo NonSharedVan_RemoveBenefit_ERR
  
  On Error GoTo RemoveBenefit_ERR
  Call xSet("RemoveBenefit")
  
  Set ibf = Me
  Set NonSharedVans = p11d32.CurrentEmployer.CurrentEmployee.NonSharedVans
  Set ben = NonSharedVans.Vans(BenefitIndex)
  Call NonSharedVans.Vans.Remove(BenefitIndex)
  
  If Not ben.CompanyDefined Then
    Call ben.DeleteDB
    NextBenefitIndex = GetNextBestListItemBenefitIndex(ibf, BenefitIndex)
    Call ibf.BenefitsToListView
    Call SelectBenefitByBenefitIndex(Me, NextBenefitIndex)
    IBenefitForm2_RemoveBenefit = True
  End If
    
    
RemoveBenefit_END:
  Set ben = Nothing
  Set ibf = Nothing
  Call xReturn("RemoveBenefit")
  Exit Function
  
RemoveBenefit_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "RemoveBenefit", "Remove Benefit", "Error removing benefit.")
  Resume RemoveBenefit_END
  Resume
NonSharedVan_RemoveBenefit_END:
  Call xReturn("NonSharedVan_RemoveBenefit")
  Exit Function
NonSharedVan_RemoveBenefit_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "NonSharedVan_RemoveBenefit", "NonSharedVan RemoveBenefit", "Error removing a non shared van.")
  Resume NonSharedVan_RemoveBenefit_END
  Resume
End Function
Private Function IBenefitForm2_UpdateBenefitListViewItem(li As MSComctlLib.IListItem, benefit As IBenefitClass, Optional ByVal BenefitIndex As Long = 0&, Optional ByVal SelectItem As Boolean = False) As Long
  
  IBenefitForm2_UpdateBenefitListViewItem = UpdateBenefitListViewItem(li, benefit, BenefitIndex, SelectItem)
  
End Function

Private Function IBenefitForm2_ValididateBenefit(ben As IBenefitClass) As Boolean
  If ben.BenefitClass = BC_nonSHAREDVAN_G Or ben.BenefitClass = BC_NONSHAREDVANS_G Then IBenefitForm2_ValididateBenefit = True
End Function

Private Function IFrmGeneral_CheckChanged(c As Control) As Boolean
  Dim bDirty As Boolean
  Dim ben As IBenefitClass
  
On Error GoTo CheckChanged_Err

  Call xSet("CheckChanged")
  
  With c
    If p11d32.CurrentEmployeeIsNothing Then
      GoTo CheckChanged_End
    End If

    Set ben = p11d32.CurrentEmployer.CurrentEmployee
    'we are asking if the value has changed and if it is valid thus save
    Select Case UCASE$(.Name)
      
      Case "TXTBX"
        Select Case .Index
          Case 0
            bDirty = CheckTextInput(.Text, ben, ee_RelevantDaysForDailySharedVanCalc_db)
            ben.Dirty = ben.Dirty Or bDirty
            ben.InvalidFields = InvalidFields(Me, fmeSharedVans)
          Case 1
            bDirty = CheckTextInput(.Text, ben, ee_PaymentsForPrivateUseOfSharedVans_db)
            ben.Dirty = ben.Dirty Or bDirty
            ben.InvalidFields = InvalidFields(Me, fmeSharedVans)
          Case 3
            If benefit Is Nothing Then GoTo CheckChanged_End
            bDirty = CheckTextInput(.Text, benefit, nsvan_item_db)
          Case 4
            If benefit Is Nothing Then GoTo CheckChanged_End
            bDirty = CheckTextInput(.Text, benefit, nsvan_AvailableFrom_db)
          Case 5
            If benefit Is Nothing Then GoTo CheckChanged_End
            bDirty = CheckTextInput(.Text, benefit, nsvan_Availableto_db)
          Case 6
            If benefit Is Nothing Then GoTo CheckChanged_End
            bDirty = CheckTextInput(.Text, benefit, nsvan_madegood_db)
          Case 7
            If benefit Is Nothing Then GoTo CheckChanged_End
            bDirty = CheckTextInput(.Text, benefit, nsvan_DaysUnavailable_db)
          Case 21   'km
'            If p11d32.AppYear > 2000 Then 'km
              If benefit Is Nothing Then GoTo CheckChanged_End
              bDirty = CheckTextInput(.Text, benefit, nsvan_RegistrationDate_db)
 '           End If
          Case Else
            ECASE "Unknown control index"
        End Select
      Case UCASE$("CHKBX")
        Select Case .Index
          Case 0
            bDirty = CheckCheckBoxInput(.value, ben, ee_OneOrMoreSharedVanAvailable_db)
            ben.Dirty = ben.Dirty Or bDirty
            ben.InvalidFields = InvalidFields(Me, fmeSharedVans)
          Case 1
            bDirty = CheckCheckBoxInput(.value, ben, ee_ReportyDailyCalculationOfSharedVans_db)
            ben.Dirty = ben.Dirty Or bDirty
            ben.InvalidFields = InvalidFields(Me, fmeSharedVans)
          Case 2
            bDirty = CheckCheckBoxInput(.value, ben, ee_NonSharedVanAvailableAtSameTimeAsSharedVan_db)
            ben.Dirty = ben.Dirty Or bDirty
            ben.InvalidFields = InvalidFields(Me, fmeSharedVans)
          Case 3
          Case 4
            bDirty = CheckCheckBoxInput(.value, benefit, nsvan_commuter_use_req_db)
          Case 5
            If benefit Is Nothing Then GoTo CheckChanged_End
            bDirty = CheckCheckBoxInput(.value, benefit, nsvan_MadeGoodIsTaxDeducted_db)
          Case 6
            bDirty = CheckCheckBoxInput(.value, benefit, nsvan_availablle_for_bus_use_only)
          Case Else
            ECASE "Unknown control index"
        End Select
      Case Else
        ECASE "Unknown control"
    End Select
    
    
    If Not benefit Is Nothing Then
      IFrmGeneral_CheckChanged = AfterCheckChanged(c, Me, bDirty, fmeInput)
      Call UpdateTotalVansBenefit
    ElseIf bDirty Then
      Call UpdateTotalVansBenefit
      Call MDIMain.SetConfirmUndo
      'dont move me to below as AfterCheckChagned sets me  *********************8
    End If
    
    
  End With

CheckChanged_End:
  Call xReturn("CheckChanged")
  Exit Function

CheckChanged_Err:
  IFrmGeneral_CheckChanged = False
  Call ErrorMessage(ERR_ERROR, Err, "CheckChanged", "ERR_CHECKCHANGED", "This function has failed for the form " & Me.Name & ".")
  Resume CheckChanged_End
  Resume
End Function

Private Property Get IFrmGeneral_InvalidVT() As Control
  IFrmGeneral_InvalidVT = m_InvalidVT
End Property

Private Property Set IFrmGeneral_InvalidVT(NewValue As Control)
  Set m_InvalidVT = NewValue
End Property

Private Sub lb_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
  Call SetSortOrder(lb, ColumnHeader)
End Sub

Private Sub LB_ItemClick(ByVal Item As MSComctlLib.ListItem)
  Call SetLastListItemSelected(Item)
  If Not (lb.SelectedItem Is Nothing) Then
    IBenefitForm2_BenefitToScreen (Item.Tag)
  End If
End Sub

Private Sub TxtBx_FieldInvalid(Index As Integer, Valid As Boolean, Message As String)
  Call SetPanel2(Message)
End Sub

Private Sub TxtBx_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  TxtBx(Index).Tag = SetChanged
End Sub

Public Function UpdateTotalVansBenefit() As Boolean
  Dim ben As IBenefitClass
  
  On Error GoTo UpdateTotalVansBenefit_Err
  Call xSet("UpdateTotalVansBenefit")

  Set ben = p11d32.CurrentEmployer.CurrentEmployee.NonSharedVans
  If Not ben Is Nothing Then
    Lab(1) = FormatWN(ben.Calculate)
  Else
    Lab(1) = FormatWN(0)
  End If
  'no of shared vans availablie
  'AM
  Select Case p11d32.CurrentEmployer.SharedVans.Vans.CountValid
    Case 1
      Lab(11) = "(" & p11d32.CurrentEmployer.SharedVans.Vans.CountValid & " van available.)"
    Case Else
      Lab(11) = "(" & p11d32.CurrentEmployer.SharedVans.Vans.CountValid & " vans available.)"
  End Select
  
UpdateTotalVansBenefit_End:
  Set ben = Nothing
  Call xReturn("UpdateTotalVansBenefit")
  Exit Function

UpdateTotalVansBenefit_Err:
  Call ErrorMessage(ERR_ERROR, Err, "UpdateTotalVansBenefit", "Update Total Vans Benefit", "Error updating the total non shared vans benefit.")
  Resume UpdateTotalVansBenefit_End
End Function

Private Sub TxtBx_Validate(Index As Integer, Cancel As Boolean)
  Call IFrmGeneral_CheckChanged(TxtBx(Index))
End Sub
