VERSION 5.00
Object = "{8D988532-0F0C-460C-B00E-7B5637E97680}#1.0#0"; "atc2vtext.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form F_Relocation 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   " "
   ClientHeight    =   5985
   ClientLeft      =   1080
   ClientTop       =   2130
   ClientWidth     =   8340
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   ScaleHeight     =   5985
   ScaleWidth      =   8340
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.CheckBox Op_Data 
      Alignment       =   1  'Right Justify
      Caption         =   "Is the amount above an amount subjected to PAYE?"
      DataSource      =   "DB"
      ForeColor       =   &H00800000&
      Height          =   405
      Index           =   1
      Left            =   4560
      TabIndex        =   9
      Tag             =   "free,font"
      Top             =   4250
      Width           =   3570
   End
   Begin VB.Frame Frame3 
      Height          =   495
      Left            =   0
      TabIndex        =   28
      Top             =   0
      Width           =   8295
      Begin VB.Label lab 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total benefit from non-qualifying relocation expenses:"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   30
         Tag             =   "free,font"
         Top             =   240
         Width           =   3750
      End
      Begin VB.Label lab 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         DataField       =   "NetNQ"
         DataSource      =   "DB_Reloc"
         ForeColor       =   &H00000080&
         Height          =   345
         Index           =   1
         Left            =   7110
         TabIndex        =   29
         Tag             =   "free,font"
         Top             =   120
         Width           =   1095
      End
   End
   Begin VB.Frame fmeDetails 
      Height          =   3075
      Left            =   0
      TabIndex        =   23
      Top             =   2790
      Width           =   8325
      Begin VB.CheckBox Op_Data 
         Alignment       =   1  'Right Justify
         Caption         =   "Is this a qualifying relocation expense?"
         DataField       =   "Qualify"
         DataSource      =   "DB"
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   0
         Left            =   4560
         TabIndex        =   10
         Tag             =   "free,font"
         Top             =   1850
         Width           =   3570
      End
      Begin VB.CommandButton B_Delete 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   3870
         Picture         =   "Benreloc.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   5
         Tag             =   "free,font"
         Top             =   2550
         Width           =   420
      End
      Begin VB.CommandButton B_Add 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   3870
         Picture         =   "Benreloc.frx":0342
         Style           =   1  'Graphical
         TabIndex        =   4
         Tag             =   "free,font"
         Top             =   2070
         Width           =   420
      End
      Begin atc2valtext.ValText TB_Data 
         Height          =   315
         Index           =   5
         Left            =   5520
         TabIndex        =   11
         Tag             =   "free,font"
         Top             =   2250
         Width           =   2645
         _ExtentX        =   4657
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
         MaxLength       =   255
         Text            =   ""
         TypeOfData      =   3
      End
      Begin atc2valtext.ValText TB_Data 
         Height          =   315
         Index           =   4
         Left            =   7155
         TabIndex        =   8
         Tag             =   "free,font"
         Top             =   1050
         Width           =   1020
         _ExtentX        =   1799
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
         Text            =   "0"
         AllowEmpty      =   0   'False
         TXTAlign        =   2
      End
      Begin atc2valtext.ValText TB_Data 
         Height          =   315
         Index           =   3
         Left            =   7155
         TabIndex        =   7
         Tag             =   "free,font"
         Top             =   645
         Width           =   1020
         _ExtentX        =   1799
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
         Text            =   "0"
         AllowEmpty      =   0   'False
         TXTAlign        =   2
      End
      Begin atc2valtext.ValText TB_Data 
         Height          =   315
         Index           =   2
         Left            =   5445
         TabIndex        =   6
         Tag             =   "free,font"
         Top             =   250
         Width           =   2715
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   255
         ForeColor       =   128
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxLength       =   255
         Text            =   ""
         TypeOfData      =   3
         AllowEmpty      =   0   'False
      End
      Begin MSComctlLib.ListView lbItems 
         Height          =   2730
         Left            =   90
         TabIndex        =   15
         Tag             =   "free,font"
         Top             =   270
         Width           =   3705
         _ExtentX        =   6535
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
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Relocation Item"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   1
            Text            =   "Amount"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Qualifying?"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label Label1 
         Caption         =   "To enter non-qualifying relocation subject to Class 1A NIC, enter under Section M - Other Items (Class 1A NIC)"
         ForeColor       =   &H00800000&
         Height          =   420
         Left            =   4590
         TabIndex        =   31
         Tag             =   "free,font"
         Top             =   2610
         Width           =   3570
      End
      Begin VB.Label lab 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Item"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   14
         Left            =   4605
         TabIndex        =   27
         Tag             =   "free,font"
         Top             =   250
         Width           =   735
      End
      Begin VB.Label lab 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Amount"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   15
         Left            =   4590
         TabIndex        =   26
         Tag             =   "free,font"
         Top             =   650
         Width           =   2610
      End
      Begin VB.Label lab 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Actual amount made good, or amount subjected to PAYE"
         ForeColor       =   &H00800000&
         Height          =   390
         Index           =   16
         Left            =   4605
         TabIndex        =   25
         Tag             =   "free,font"
         Top             =   1035
         Width           =   2550
         WordWrap        =   -1  'True
      End
      Begin VB.Label lab 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Comments"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   18
         Left            =   4560
         TabIndex        =   24
         Tag             =   "free,font"
         Top             =   2250
         Width           =   975
      End
   End
   Begin VB.Frame fmeMain 
      Height          =   2280
      Left            =   0
      TabIndex        =   16
      Top             =   495
      Width           =   8320
      Begin VB.ComboBox CboBx 
         DataField       =   "TaxYear"
         DataSource      =   "DB_Reloc"
         Height          =   315
         Index           =   0
         Left            =   2295
         TabIndex        =   1
         Tag             =   "free,font"
         Top             =   810
         Width           =   1650
      End
      Begin atc2valtext.ValText TB_Data 
         Height          =   315
         Index           =   0
         Left            =   1095
         TabIndex        =   0
         Tag             =   "free,font"
         Top             =   360
         Width           =   2835
         _ExtentX        =   14208
         _ExtentY        =   3836
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
         MaxLength       =   255
         Text            =   ""
         TypeOfData      =   3
         AllowEmpty      =   0   'False
      End
      Begin atc2valtext.ValText TB_Data 
         Height          =   315
         Index           =   1
         Left            =   1320
         TabIndex        =   2
         Tag             =   "free,font"
         Top             =   1860
         Width           =   975
         _ExtentX        =   14208
         _ExtentY        =   3836
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
         TXTAlign        =   2
      End
      Begin MSComctlLib.ListView LB 
         Height          =   1815
         Left            =   4140
         TabIndex        =   3
         Tag             =   "free,font"
         Top             =   360
         Width           =   4080
         _ExtentX        =   7197
         _ExtentY        =   3201
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
            Text            =   "Relocation address"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   1
            Text            =   "Benefit"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label lab 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Qualifying relocation expenses"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   7
         Left            =   120
         TabIndex        =   22
         Tag             =   "free,font"
         Top             =   1260
         Width           =   2145
      End
      Begin VB.Label lab 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Used last year"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   10
         Left            =   1320
         TabIndex        =   21
         Tag             =   "free,font"
         Top             =   1635
         Width           =   1215
      End
      Begin VB.Label lab 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Used this year"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   9
         Left            =   120
         TabIndex        =   20
         Tag             =   "free,font"
         Top             =   1635
         Width           =   1215
      End
      Begin VB.Label lab 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   19
         Tag             =   "free,font"
         Top             =   360
         Width           =   885
      End
      Begin VB.Label lab 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tax year of the move"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   5
         Left            =   135
         TabIndex        =   18
         Tag             =   "free,font"
         Top             =   810
         Width           =   1500
      End
      Begin VB.Label lab 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Remaining relief"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   8
         Left            =   2520
         TabIndex        =   17
         Tag             =   "free,font"
         Top             =   1650
         Width           =   1185
      End
      Begin VB.Label lab 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         DataField       =   "QualTotal"
         DataSource      =   "DB_Reloc"
         ForeColor       =   &H00000080&
         Height          =   315
         Index           =   6
         Left            =   2950
         TabIndex        =   12
         Tag             =   "free,font"
         Top             =   1260
         Width           =   975
      End
      Begin VB.Label lab 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         DataField       =   "ThisYear"
         DataSource      =   "DB_Reloc"
         ForeColor       =   &H00000080&
         Height          =   315
         Index           =   12
         Left            =   120
         TabIndex        =   13
         Tag             =   "free,font"
         Top             =   1860
         Width           =   1005
      End
      Begin VB.Label lab 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         DataField       =   "Available"
         DataSource      =   "DB_Reloc"
         ForeColor       =   &H00000080&
         Height          =   315
         Index           =   11
         Left            =   2520
         TabIndex        =   14
         Tag             =   "free,font"
         Top             =   1860
         Width           =   960
      End
   End
End
Attribute VB_Name = "F_Relocation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Implements IFrmGeneral
Implements IBenefitForm2

Public benefit As IBenefitClass
Private m_InvalidVT As Control


Private mclsResize As New clsFormResize
Private Const L_DES_HEIGHT As Long = 6090
Private Const L_DES_WIDTH As Long = 8445
Private Function RemoveDetail(lst As ListItem) As Boolean
  Dim benReloc As Relocation
  Dim li As ListItem
  Dim ibf As IBenefitForm2
  Dim bGotNext As Boolean
  
On Error GoTo RemoveDetail_ERR
  
  Call xSet("RemoveDetail")
  
  Set benReloc = benefit
  benReloc.RelocDetails.Remove (lst.Tag)
  bGotNext = GetNextBestListItem(li, lbItems, lst, True)
  lbItems.listitems.Remove (lst.Index)
  benefit.Dirty = True
  Call MDIMain.SetConfirmUndo
  
  Set ibf = Me
  If bGotNext Then
    Set lbItems.SelectedItem = li
    Call RelocationDetailToScreen(li.Tag)
  End If
  
  
  Call ibf.UpdateBenefitListViewItem(ibf.lv.SelectedItem, benefit)
  Call SetDetailsDelete
  
RemoveDetail_END:
  Set ibf = Nothing
  Call xReturn("RemoveDetail")
  Set li = Nothing
  Set benReloc = Nothing
  Exit Function
RemoveDetail_ERR:
  RemoveDetail = False
  Call ErrorMessage(ERR_ERROR, Err, "RemoveDetail", "Remove Detail", "Error removing a relocation detail.")
  Resume RemoveDetail_END
  Resume
End Function
Private Function SetDetailsDelete() As Boolean

On Error GoTo SetDetailsDelete_ERR

  Call xSet("SetDetailsDelete")

  If lbItems.listitems.Count Then
    If Not B_Delete.Enabled Then
      B_Delete.Enabled = True
      lbItems.Enabled = True
      TB_Data(2).Enabled = True
      TB_Data(3).Enabled = True
      TB_Data(4).Enabled = True
      
      TB_Data(2).AllowEmpty = False
      TB_Data(3).AllowEmpty = False
      TB_Data(4).AllowEmpty = False
      
      
      TB_Data(5).Enabled = True
      Op_Data(0).Enabled = True
      Op_Data(1).Enabled = True
      B_Add.Enabled = True
      
    End If
  Else
    TB_Data(2).AllowEmpty = True
    TB_Data(3).AllowEmpty = True
    TB_Data(4).AllowEmpty = True
    
    B_Add.Enabled = True
    B_Delete.Enabled = False
    lbItems.Enabled = False
    TB_Data(2).Text = ""
    TB_Data(3).Text = ""
    TB_Data(4).Text = ""
    TB_Data(5).Text = ""
   ' TB_Data(6).Text = ""
    
    TB_Data(2).Enabled = False
    TB_Data(3).Enabled = False
    TB_Data(4).Enabled = False
    TB_Data(5).Enabled = False
    Op_Data(0).Enabled = False
    Op_Data(1).Enabled = False
  End If
  
  SetDetailsDelete = True
  
SetDetailsDelete_END:
  Call xReturn("SetDetailsDelete")
  Exit Function
SetDetailsDelete_ERR:
  SetDetailsDelete = False
  Call ErrorMessage(ERR_ERROR, Err, "SetDetailsDelete", "Set Delete", "Error setting the state of the relocation items listview and the delete buttons enabled state.")
  Resume SetDetailsDelete_END
  Resume
End Function

Private Sub B_Add_Click()
  Call AddDetail
End Sub

Private Sub B_Delete_Click()
  If Not lbItems.SelectedItem Is Nothing Then
    Call RemoveDetail(lbItems.SelectedItem)
  End If
End Sub
Private Function AddDetail()
  Dim Relocation As Relocation
  Dim detail As RelocationDetail
  Dim lst As ListItem
  Dim ibf As IBenefitForm2
  Static bInFunc As Boolean
  
On Error GoTo AddDetail_Err
  If bInFunc Then
    GoTo AddDetail_Err
  Else
    bInFunc = True
  End If
  
  Call xSet("AddDetail")
  
  Set detail = New RelocationDetail
  Set Relocation = benefit
  
  'Put in defaults for benefit
  detail.Item = "Please enter description"
  detail.value = 0
  detail.MadeGood = 0
  detail.Comments = ""
  detail.Qualify = True
  detail.IsTaxDeducted = False
  detail.Key = Relocation.RelocDetails.Add(detail)
  Set lst = lbItems.listitems.Add(, , detail.Item)
  
  With lst
    .Tag = detail.Key
    .SubItems(1) = FormatWN(detail.benefit, "£")
    .SubItems(2) = IIf(detail.Qualify, "Yes", "No")
  End With
  
  Set lbItems.SelectedItem = lst
  
  TB_Data(2).Text = detail.Item
  TB_Data(3).Text = detail.value
  TB_Data(4).Text = detail.MadeGood
  TB_Data(5).Text = detail.Comments
  Op_Data(0) = IIf(detail.Qualify, 1, 0)
  Op_Data(1) = IIf(detail.IsTaxDeducted, 1, 0)
  
  benefit.Dirty = True
  
  Call MDIMain.SetConfirmUndo
  Call RelocationDetailToScreen(lst.Tag)
  Set ibf = Me
  'always pass the LB listitems tag
  Call ibf.UpdateBenefitListViewItem(ibf.lv.SelectedItem, benefit, ibf.lv.SelectedItem.Tag)
  Call TB_Data(2).SetFocus
  bInFunc = False
AddDetail_End:
  Set ibf = Nothing
  Set lst = Nothing
  Set Relocation = Nothing
  Set detail = Nothing
  Call xReturn("AddDetail")
  Exit Function
AddDetail_Err:
  bInFunc = False
  Call ErrorMessage(ERR_ERROR, Err, "AddDetail", "ERR_AddDetail", "Error in AddDetail function, called from the form " & Me.Name & ".")
  Resume AddDetail_End
  Resume
End Function
Private Sub CboBx_Validate(Index As Integer, Cancel As Boolean)
  Call IFrmGeneral_CheckChanged(CboBx(Index))
End Sub

Private Sub Form_Resize()
  mclsResize.Resize
  Call ColumnWidths(LB, 75, 25)
  Call ColumnWidths(lbItems, 50, 25, 25)
End Sub


Private Sub Form_Load()
  If Not (mclsResize.InitResize(Me, L_DES_HEIGHT, L_DES_WIDTH, DESIGN, , , MDIMain)) Then
    Err.Raise ERR_Application
  End If
  'asasdasknhdakjdhaskjdhkasjhdaskjhsk need to go to pd file
  CboBx(0).AddItem (p11d32.Rates.value(RelocationThisYear))
  CboBx(0).AddItem (p11d32.Rates.value(RelocationLastYear))
End Sub


Private Sub IBenefitForm2_AddBenefit()
  Dim ben As Relocation
  
On Error GoTo AddBenefit_Err

  Call xSet("AddBenefit")
  
  Set ben = New Relocation
  Call lbItems.listitems.Clear
  Call AddBenefitHelper(Me, ben)
  Call ben.AddNonQualifyingRelocation
  Call StandardReadData(ben.NonQualifyingRelocation)
  
 
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
    
    .value(reloc_Item_db) = "Please enter address..."
    
    .value(reloc_TaxYear_db) = p11d32.Rates.value(RelocationThisYear)
    .value(reloc_UsedLastyear_db) = 0
    .value(reloc_RemainingRelief) = L_RELOCEXEMPT
    .value(reloc_UsedThisYear_db) = 0
    
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
    Call SetDetailsDelete
    TB_Data(0).Enabled = True
    TB_Data(1).Enabled = True
    Lab(6).Enabled = True
    Lab(11).Enabled = True
    Lab(12).Enabled = True
    CboBx(0).Enabled = True
    fmeDetails.Enabled = True
    
    'fmeMain.Enabled = True
    Call MDIMain.SetDelete
  ElseIf fState = FORM_DISABLED Then
    Set benefit = Nothing
    lbItems.Enabled = False
    fmeDetails.Enabled = False
    
    TB_Data(0).Enabled = False
    TB_Data(1).Enabled = False
    Lab(6).Enabled = False
    Lab(11).Enabled = False
    Lab(12).Enabled = False
    CboBx(0).Enabled = False
    
    
    'fmeMain.Enabled = False
    B_Add.Enabled = False
    B_Delete.Enabled = False
    Call MDIMain.ClearDelete
    Call MDIMain.ClearConfirmUndo
  End If
  IBenefitForm2_BenefitFormState = True
    
BenefitFormState_end:
  Call xReturn("BenefitFormState")
  Exit Function
BenefitFormState_err:
  IBenefitForm2_BenefitFormState = False
  
  Call ErrorMessage(ERR_ERROR, Err, "BenefitFormState", "Benefit Form State", "Error setting the benefit form state.")
  Resume BenefitFormState_end
End Function

Private Function IBenefitForm2_BenefitOff() As Boolean
  TB_Data(0).Text = ""
  TB_Data(1).Text = ""
  Lab(6).Caption = ""
  Lab(11).Caption = ""
  Lab(12).Caption = ""
  CboBx(0).Text = ""
  Lab(1).Caption = ""
  Call RelocationDetails
End Function

Private Function IBenefitForm2_BenefitOn() As Boolean
  TB_Data(0).Text = benefit.value(reloc_Item_db)
  TB_Data(1).Text = benefit.value(reloc_UsedLastyear_db)
  Lab(6).Caption = benefit.value(reloc_QualifyTotal)
  Call UpdateReliefLabels
  CboBx(0).Text = benefit.value(reloc_TaxYear_db)
  Call RelocationDetails
End Function
Private Sub RelocationDetails()
  On Error GoTo RelocationDetails_ERR
   
  Call xSet("RelocationDetails")
  
  If RelocationDetailsToListView(benefit) Then
    Set lbItems.SelectedItem = lbItems.listitems(1)
    Call RelocationDetailToScreen(lbItems.SelectedItem.Tag)
  End If
  
RelocationDetails_END:
  Call xReturn("RelocationDetails")
  Exit Sub
RelocationDetails_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "RelocationDetails", "Relocation Details", "Error placing the relocation details to the screen.")
  Resume RelocationDetails_END
  Resume
End Sub

Private Function IBenefitForm2_BenefitsToListView() As Long
  IBenefitForm2_BenefitsToListView = BenefitsToListView(Me)
End Function

Private Function IBenefitForm2_BenefitToListView(ben As IBenefitClass, ByVal lBenefitIndex As Long) As Long
  Dim v As Variant
  Dim lst As ListItem
  Dim ibf As IBenefitForm2
  
  On Error GoTo BenefitToListView_Err
  Call xSet("BenefitToListView")
  
  If Not ben Is Nothing Then
    Set ibf = Me
    If ibf.ValididateBenefit(ben) Then
      Set lst = ibf.lv.listitems.Add(, , ben.Name)
      Call RelocationDetailsToListView(ben)
      IBenefitForm2_BenefitToListView = ibf.UpdateBenefitListViewItem(lst, ben, lBenefitIndex)
    End If
  End If

BenefitToListView_End:
  Set lst = Nothing
  Call xReturn("BenefitToListView")
  Exit Function
BenefitToListView_Err:
  Call ErrorMessage(ERR_ERROR, Err, "BenefitToListView", "Benefit To List View", "Error placing a relocation benefit to a list view.")
  Resume BenefitToListView_End
End Function

Private Function IBenefitForm2_BenefitToScreen(Optional ByVal BenefitIndex As Long = -1&, Optional ByVal UpdateBenefit As Boolean = True) As Boolean
  IBenefitForm2_BenefitToScreen = BenefitToScreenHelper(Me, BenefitIndex, UpdateBenefit)
End Function

Private Property Let IBenefitForm2_benclass(ByVal NewValue As BEN_CLASS)
End Property

Private Property Get IBenefitForm2_benclass() As BEN_CLASS
  IBenefitForm2_benclass = BC_QUALIFYING_RELOCATION_J
End Property

'Private Property Get IBenefitForm2_ControlDefault() As Control
'  Set IBenefitForm2_ControlDefault = TB_Data(0)
'End Property

Private Property Get IBenefitForm2_lv() As MSComctlLib.IListView
  Set IBenefitForm2_lv = LB
End Property

Private Function IBenefitForm2_RemoveBenefit(ByVal BenefitIndex As Long) As Boolean
  Dim reloc As Relocation
  
  On Error GoTo RemoveBenefit_ERR
  
  Call xSet("RemoveBenefit")
  
  Set reloc = benefit
  
  IBenefitForm2_RemoveBenefit = p11d32.CurrentEmployer.CurrentEmployee.RemoveBenefitWithLinks(Me, benefit, BenefitIndex, reloc.NonQualifyingRelocation)
  
  If LB.listitems.Count = 0 Then lbItems.listitems.Clear
  
RemoveBenefit_END:
  Call xReturn("RemoveBenefit")
  Exit Function
RemoveBenefit_ERR:
  IBenefitForm2_RemoveBenefit = False
  Call ErrorMessage(ERR_ERROR, Err, "RemoveBenefit", "Remove Benefit", "Error removing the selected benefit.")
  Resume RemoveBenefit_END
End Function

Private Function IBenefitForm2_UpdateBenefitListViewItem(li As MSComctlLib.IListItem, benefit As IBenefitClass, Optional ByVal BenefitIndex As Long = 0&, Optional ByVal SelectItem As Boolean = False) As Long
  Dim b As Boolean
  Dim detail As RelocationDetail
  Dim reloc As Relocation
  
On Error GoTo UpdateBenefitListViewItem_ERR
  
  Call xSet("UpdateBenefitListViewItem")
  
  b = UpdateBenefitListViewItem(li, benefit, BenefitIndex, SelectItem) 'calculate updated here
  
  If b And (lbItems.listitems.Count > 0) And Not lbItems.SelectedItem Is Nothing Then
    Set reloc = benefit
    Set detail = reloc.RelocDetails(lbItems.SelectedItem.Tag)
    Call DetailToListItem(lbItems.SelectedItem, detail)
  End If
  
  Call UpdateReliefLabels
  
  IBenefitForm2_UpdateBenefitListViewItem = li.Index
  
UpdateBenefitListViewItem_END:
  Call xReturn("UpdateBenefitListViewItem")
  Exit Function
  
UpdateBenefitListViewItem_ERR:
  IBenefitForm2_UpdateBenefitListViewItem = False
  Call ErrorMessage(ERR_ERROR, Err, "UpdateBenefitListViewItem", "Updating a benefits list view text", "Unable to update the benefits list view text.")
  Resume UpdateBenefitListViewItem_END
  Resume
End Function

Private Function IBenefitForm2_ValididateBenefit(ben As IBenefitClass) As Boolean
  If ben.BenefitClass = BC_QUALIFYING_RELOCATION_J Then IBenefitForm2_ValididateBenefit = True
End Function

Private Function IFrmGeneral_CheckChanged(c As Control) As Boolean
  Dim i As Long
  Dim bDirty As Boolean, bRefreshDetail As Boolean
  Dim detail As RelocationDetail
  Dim benReloc As Relocation
  Dim bGotDetail As Boolean
  
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
    'get the current relocation item
    bGotDetail = GetCurrentDetail(detail)
    
    Select Case .Name
      Case "TB_Data"
        Select Case .Index
          Case 0
              bDirty = CheckTextInput(.Text, benefit, reloc_Item_db)
           Case 1
              bDirty = CheckTextInput(.Text, benefit, reloc_UsedLastyear_db)
          Case 2
            If bGotDetail Then
              bDirty = StrComp(.Text, detail.Item)
              If bDirty Then detail.Item = .Text
            End If
            bRefreshDetail = bDirty
          Case 3
            If bGotDetail Then
              bDirty = StrComp(.Text, detail.value)
              If bDirty Then detail.value = .Text
            End If
            bRefreshDetail = bDirty
          Case 4
            If bGotDetail Then
              
              bDirty = StrComp(.Text, detail.MadeGood)
              If bDirty Then detail.MadeGood = .Text
            End If
            bRefreshDetail = bDirty
          Case 5
            If bGotDetail Then
              bDirty = StrComp(.Text, detail.Comments)
              If bDirty Then detail.Comments = .Text
            End If
            bRefreshDetail = bDirty
          Case 6
            'not used CAD ZZZZZZ
          Case Else
            ECASE "Unknown control"
        
        End Select
    Case "Op_Data"
      Select Case .Index
        Case 0
          If bGotDetail Then
            bDirty = (IIf(.value = vbChecked, True, False) <> detail.Qualify)
            If bDirty Then detail.Qualify = IIf(.value = vbChecked, True, False)
          End If
          bRefreshDetail = bDirty
        Case 1
          If bGotDetail Then
            bDirty = (IIf(.value = vbChecked, True, False) <> detail.IsTaxDeducted)
            If bDirty Then detail.IsTaxDeducted = IIf(.value = vbChecked, True, False)
          End If
          bRefreshDetail = bDirty
        Case Else
          ECASE "Unknown control"
      End Select
    Case "CboBx"
      bDirty = CheckTextInput(.Text, benefit, reloc_TaxYear_db)
    Case Else
      ECASE "Unknown control"
    End Select
    
    Call RefreshDetail(bRefreshDetail)
    'must be required in all check changed
    IFrmGeneral_CheckChanged = AfterCheckChanged(c, Me, bDirty)
    
  End With
  
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

Private Sub LB_ItemClick(ByVal Item As MSComctlLib.ListItem)
  Call SetLastListItemSelected(Item)
  If Not (LB.SelectedItem Is Nothing) Then
    Call IBenefitForm2_BenefitToScreen(Item.Tag)
  End If
End Sub

Private Sub lb_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
  Call SetSortOrder(LB, ColumnHeader)
End Sub

Private Sub LB_KeyDown(KeyCode As Integer, Shift As Integer)
  Call LVKeyDown(KeyCode, Shift)
End Sub

Private Sub lb_KeyPress(KeyAscii As Integer)
  'Check for return key - tab to next field
  If KeyAscii = 13 Then Call SendKeys(vbTab)
End Sub

Private Function RelocationDetailsToListView(reloc As Relocation) As Long
  Dim detail As RelocationDetail
  Dim i As Long
  Dim lst As ListItem
  
  On Error GoTo RelocationDetailsToListView_ERR
  
  Call xSet("RelocationDetailsToListView")
  
  lbItems.listitems.Clear
  If Not reloc Is Nothing Then
    If reloc.RelocDetails.Count Then
      For i = 1 To reloc.RelocDetails.Count
        Set detail = reloc.RelocDetails(i)
        If Not detail Is Nothing Then
          Set lst = lbItems.listitems.Add
          Call DetailToListItem(lst, detail)
          lst.Tag = i
          RelocationDetailsToListView = RelocationDetailsToListView + 1
        End If
      Next i
    End If
  End If
  
  Call SetDetailsDelete
  
RelocationDetailsToListView_END:
  Call xReturn("RelocationDetailsToListView")
  Exit Function
RelocationDetailsToListView_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "RelocationDetailsToListView", "Relocation Details To List View", "Error placing the relocation details to the listview.")
  Resume RelocationDetailsToListView_END
End Function

Private Function RelocationDetailToScreen(DetailIndex As Long) As Boolean
  Dim detail As RelocationDetail
  Dim benReloc As Relocation
  
On Error GoTo RelocationDetailToScreen_ERR
  
  Call xSet("RelocationDetailToScreen")
  
  If DetailIndex <> -1 Then
    Set benReloc = benefit
    Set detail = benReloc.RelocDetails(DetailIndex)
    If Not detail Is Nothing Then
      TB_Data(2).Text = detail.Item
      TB_Data(3).Text = detail.value
      TB_Data(4).Text = detail.MadeGood
      Op_Data(0).Enabled = True
      Op_Data(0) = IIf(detail.Qualify, vbChecked, vbUnchecked)
      Op_Data(1).Enabled = True
      Op_Data(1) = IIf(detail.IsTaxDeducted, vbChecked, vbUnchecked)
      TB_Data(5).Text = detail.Comments
    End If
  Else
    TB_Data(2).Text = ""
    TB_Data(3).Text = ""
    TB_Data(4).Text = ""
    Op_Data(0).Enabled = False
    Op_Data(1).Enabled = False
    TB_Data(5).Text = ""
  End If
  Call SetDetailsDelete
  RelocationDetailToScreen = True
    
RelocationDetailToScreen_END:
  Set detail = Nothing
  Set benReloc = Nothing
  Call xReturn("RelocationDetailToScreen")
  Exit Function
RelocationDetailToScreen_ERR:
  RelocationDetailToScreen = False
  Call ErrorMessage(ERR_ERROR, Err, "RelocationDetailToScreen", "Relocation Detail To Screen", "Error placing a relocation detail to the screen.")
  Resume RelocationDetailToScreen_END
  Resume
End Function

Private Sub lbItems_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
  Call SetSortOrder(lbItems, ColumnHeader)
End Sub

Private Sub LBItems_ItemClick(ByVal Item As MSComctlLib.ListItem)
 
  If Not (lbItems.SelectedItem Is Nothing) And Not (LB.SelectedItem Is Nothing) Then
    Call RelocationDetailToScreen(Item.Tag)
  End If
End Sub

Private Sub lbItems_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Call TestChangedControls(Me)
End Sub


Private Sub Op_Data_Click(Index As Integer)
  Call IFrmGeneral_CheckChanged(Op_Data(Index))
End Sub

Private Sub TB_Data_FieldInvalid(Index As Integer, Valid As Boolean, Message As String)
  Call SetPanel2(Message)
End Sub

Private Sub TB_Data_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  TB_Data(Index).Tag = SetChanged
End Sub


Private Function ScreenToRelocationDetail(DetailIndex As Long) As RelocationDetail
  Dim benReloc As Relocation

  On Error GoTo ScreenToRelocationDetail_Err
  
  Call xSet("ScreenToRelocationDetail")
  
  If Not benefit Is Nothing Then
    Set benReloc = benefit
    Set ScreenToRelocationDetail = benReloc.RelocDetails(DetailIndex)
    If Not ScreenToRelocationDetail Is Nothing Then
      With ScreenToRelocationDetail
        .Item = TB_Data(2).Text
        .value = CorrectBenValue(BC_QUALIFYING_RELOCATION_J, reloc_Value, TB_Data(3).Text)
        .MadeGood = CorrectBenValue(BC_QUALIFYING_RELOCATION_J, reloc_MadeGood, TB_Data(4).Text)
        .Qualify = IIf(Op_Data(0) = vbChecked, True, False)
        .IsTaxDeducted = IIf(Op_Data(1) = vbChecked, True, False)
        .Comments = TB_Data(5).Text
      End With
    End If
  End If

ScreenToRelocationDetail_End:
  Call xReturn("ScreenToRelocationDetail")
  Exit Function

ScreenToRelocationDetail_Err:
  Call ErrorMessage(ERR_ERROR, Err, "ScreenToRelocationDetail", "Screen To Relocation Detail", "Error copying the screen information to a relocation detail.")
  Resume ScreenToRelocationDetail_End
  Resume
End Function


Public Function RefreshDetail(Dirty As Boolean) As Boolean
  Dim RelocationDetail As RelocationDetail
  
  On Error GoTo RefreshDetail_Err
  Call xSet("RefreshDetail")

  If Dirty Then
    If Not lbItems.SelectedItem Is Nothing Then
      With lbItems.SelectedItem
        Set RelocationDetail = ScreenToRelocationDetail(.Tag)
        'update the listview item
        .Text = RelocationDetail.Item
        .SubItems(1) = FormatWN(RelocationDetail.Calculate)
        .SubItems(2) = IIf(RelocationDetail.Qualify, "Yes", "No")
      End With
    End If
  End If

RefreshDetail_End:
  Call xReturn("RefreshDetail")
  Exit Function

RefreshDetail_Err:
  Call ErrorMessage(ERR_ERROR, Err, "RefreshDetail", "Refresh Detail", "Error refreshing a relocation detail to the screen.")
  Resume RefreshDetail_End
  Resume
End Function


Private Function GetCurrentDetail(detail As RelocationDetail) As Boolean
  Dim benReloc As Relocation
  
  On Error GoTo GetCurrentDetail_Err
  Call xSet("GetCurrentDetail")

  If Not lbItems.SelectedItem Is Nothing Then
    Set benReloc = benefit
    Set detail = benReloc.RelocDetails(lbItems.SelectedItem.Tag)
    GetCurrentDetail = True
  End If

GetCurrentDetail_End:
  Call xReturn("GetCurrentDetail")
  Exit Function

GetCurrentDetail_Err:
  Call ErrorMessage(ERR_ERROR, Err, "GetCurrentDetail", "Get Current Detail", "Error obtaining the selected relocation detail.")
  Resume GetCurrentDetail_End
End Function




Private Function DetailToListItem(li As ListItem, detail As RelocationDetail) As Boolean
  Dim vBenefit As Variant
  
  On Error GoTo DetailToListItem_Err
  
  Call xSet("DetailToListItem")

  
  If Not li Is Nothing Or Not detail Is Nothing Then
    With li
      .Text = detail.Item
      .SubItems(1) = FormatWN(detail.Calculate, "£")
      .SubItems(2) = IIf(detail.Qualify, "Yes", "No")
    End With
  Else
    Call Err.Raise(ERR_DETAIL_TO_LISTITEM, "DetailToListItem", "Either the detail or listitem are set to nothing.")
  End If
  
  DetailToListItem = True
  
DetailToListItem_End:
  Call xReturn("DetailToListItem")
  Exit Function

DetailToListItem_Err:
  DetailToListItem = False
  Call ErrorMessage(ERR_ERROR, Err, "DetailToListItem", "Detail To List Item", "Error placing a reloaction detail to a list item.")
  Resume DetailToListItem_End
End Function


Private Function UpdateReliefLabels() As Boolean

  On Error GoTo UpdateReliefLabels_Err
  Call xSet("UpdateReliefLabels")

  If Not benefit Is Nothing Then
    Lab(11).Caption = FormatWN(benefit.value(reloc_RemainingRelief))
    Lab(12).Caption = FormatWN(benefit.value(reloc_UsedThisYear_db))
    Lab(1).Caption = FormatWN(benefit.value(reloc_NQBenefit))
    Lab(6).Caption = benefit.value(reloc_QualifyTotal)
  
  End If
  
  UpdateReliefLabels = True

UpdateReliefLabels_End:
  Call xReturn("UpdateReliefLabels")
  Exit Function

UpdateReliefLabels_Err:
  Call ErrorMessage(ERR_ERROR, Err, "UpdateReliefLabels", "Update Relief Labels", "Error updating relief labels for a relocation benefit.")
  Resume UpdateReliefLabels_End
End Function

Private Sub TB_Data_Validate(Index As Integer, Cancel As Boolean)
  Call IFrmGeneral_CheckChanged(TB_Data(Index))
End Sub
