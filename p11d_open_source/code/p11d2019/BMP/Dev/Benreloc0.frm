VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.2#0"; "comctl32.ocx"
Object = "{4582CA9E-1A45-11D2-8D2F-00C04FA9DD6F}#1.0#0"; "ATC2VTEXT.OCX"
Begin VB.Form F_BenReloc0 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Relocation Expenses"
   ClientHeight    =   5715
   ClientLeft      =   1080
   ClientTop       =   2130
   ClientWidth     =   8355
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   ScaleHeight     =   5715
   ScaleWidth      =   8355
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame3 
      Height          =   495
      Left            =   0
      TabIndex        =   27
      Top             =   0
      Width           =   8295
      Begin VB.Label Lab 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total benefit from non-qualifying relocation expenses for all relocations:"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   29
         Tag             =   "free,font"
         Top             =   240
         Width           =   6615
      End
      Begin VB.Label Lab 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         DataField       =   "NetNQ"
         DataSource      =   "DB_Reloc"
         ForeColor       =   &H00000080&
         Height          =   345
         Index           =   1
         Left            =   7110
         TabIndex        =   28
         Tag             =   "free,font"
         Top             =   120
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      Height          =   2895
      Left            =   0
      TabIndex        =   14
      Top             =   2760
      Width           =   8415
      Begin VB.CheckBox Op_Data 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Is this a qualifying relocation expense?"
         DataField       =   "Qualify"
         DataSource      =   "DB"
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   0
         Left            =   4605
         TabIndex        =   20
         Tag             =   "free,font"
         Top             =   1635
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
         Picture         =   "Benreloc0.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   17
         Tag             =   "free,font"
         Top             =   2415
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
         Picture         =   "Benreloc0.frx":0342
         Style           =   1  'Graphical
         TabIndex        =   16
         Tag             =   "free,font"
         Top             =   1935
         Width           =   420
      End
      Begin atc2valtext.ValText TB_Data 
         Height          =   315
         Index           =   5
         Left            =   5520
         TabIndex        =   15
         Tag             =   "free,font"
         Top             =   2100
         Width           =   2715
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
         Text            =   ""
         TypeOfData      =   3
      End
      Begin atc2valtext.ValText TB_Data 
         Height          =   315
         Index           =   4
         Left            =   6930
         TabIndex        =   18
         Tag             =   "free,font"
         Top             =   1155
         Width           =   1245
         _ExtentX        =   2196
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
         TXTAlign        =   2
      End
      Begin atc2valtext.ValText TB_Data 
         Height          =   315
         Index           =   3
         Left            =   6930
         TabIndex        =   19
         Tag             =   "free,font"
         Top             =   720
         Width           =   1245
         _ExtentX        =   2196
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
         TXTAlign        =   2
      End
      Begin atc2valtext.ValText TB_Data 
         Height          =   315
         Index           =   2
         Left            =   5445
         TabIndex        =   21
         Tag             =   "free,font"
         Top             =   240
         Width           =   2715
         _ExtentX        =   0
         _ExtentY        =   0
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
         Text            =   ""
         TypeOfData      =   3
      End
      Begin ComctlLib.ListView lbItems 
         Height          =   2595
         Left            =   120
         TabIndex        =   22
         Tag             =   "free,font"
         Top             =   240
         Width           =   3705
         _ExtentX        =   6535
         _ExtentY        =   4577
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
            Text            =   "Relocation Item"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            SubItemIndex    =   1
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Amount"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            SubItemIndex    =   2
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Qualifying?"
            Object.Width           =   2540
         EndProperty
      End
      Begin atc2valtext.ValText TB_Data 
         Height          =   315
         Index           =   6
         Left            =   7320
         TabIndex        =   30
         Tag             =   "free,font"
         Top             =   2520
         Width           =   675
         _ExtentX        =   1191
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
      Begin VB.Label Lab 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Item"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   14
         Left            =   4605
         TabIndex        =   26
         Tag             =   "free,font"
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Lab 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Amount"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   15
         Left            =   4605
         TabIndex        =   25
         Tag             =   "free,font"
         Top             =   735
         Width           =   2610
      End
      Begin VB.Label Lab 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Amount made good"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   16
         Left            =   4605
         TabIndex        =   24
         Tag             =   "free,font"
         Top             =   1185
         Width           =   2550
      End
      Begin VB.Label Lab 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Comments"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   18
         Left            =   4605
         TabIndex        =   23
         Tag             =   "free,font"
         Top             =   2130
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2415
      Left            =   0
      TabIndex        =   0
      Top             =   480
      Width           =   8415
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
         TabIndex        =   2
         Tag             =   "free,font"
         Top             =   360
         Width           =   2835
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
         Text            =   ""
         TypeOfData      =   3
      End
      Begin atc2valtext.ValText TB_Data 
         Height          =   315
         Index           =   1
         Left            =   1320
         TabIndex        =   3
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
      Begin ComctlLib.ListView LB 
         Height          =   1815
         Left            =   4140
         TabIndex        =   4
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
            Text            =   "Relocation address"
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
      Begin VB.Label Lab 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Qualifying Relocation Expenses"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   7
         Left            =   120
         TabIndex        =   13
         Tag             =   "free,font"
         Top             =   1260
         Width           =   2715
      End
      Begin VB.Label Lab 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Used last year"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   10
         Left            =   1320
         TabIndex        =   12
         Tag             =   "free,font"
         Top             =   1635
         Width           =   1215
      End
      Begin VB.Label Lab 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Used this year"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   9
         Left            =   120
         TabIndex        =   11
         Tag             =   "free,font"
         Top             =   1635
         Width           =   1215
      End
      Begin VB.Label Lab 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   10
         Tag             =   "free,font"
         Top             =   360
         Width           =   885
      End
      Begin VB.Label Lab 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tax year of the move"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   5
         Left            =   120
         TabIndex        =   9
         Tag             =   "free,font"
         Top             =   735
         Width           =   1500
      End
      Begin VB.Label Lab 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Remaining relief"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   8
         Left            =   2520
         TabIndex        =   8
         Tag             =   "free,font"
         Top             =   1650
         Width           =   1185
      End
      Begin VB.Label Lab 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         DataField       =   "QualTotal"
         DataSource      =   "DB_Reloc"
         ForeColor       =   &H00000080&
         Height          =   315
         Index           =   6
         Left            =   2970
         TabIndex        =   7
         Tag             =   "free,font"
         Top             =   1260
         Width           =   975
      End
      Begin VB.Label Lab 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         DataField       =   "ThisYear"
         DataSource      =   "DB_Reloc"
         ForeColor       =   &H00000080&
         Height          =   315
         Index           =   12
         Left            =   120
         TabIndex        =   6
         Tag             =   "free,font"
         Top             =   1860
         Width           =   1005
      End
      Begin VB.Label Lab 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         DataField       =   "Available"
         DataSource      =   "DB_Reloc"
         ForeColor       =   &H00000080&
         Height          =   315
         Index           =   11
         Left            =   2520
         TabIndex        =   5
         Tag             =   "free,font"
         Top             =   1860
         Width           =   960
      End
   End
End
Attribute VB_Name = "F_BenReloc0"
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

Private Sub B_Add_Click()
  Call AddItem
End Sub
Private Sub CboBx_Lostfocus(Index As Integer)
  Call CheckChanged(CboBx(Index))
End Sub
Private Sub Form_Resize()
  mclsResize.Resize
  Call ColumnWidths(LB, 30, 25, 25)
  Call ColumnWidths(lbItems, 30, 25, 25)
End Sub
Private Sub IBenefitForm_AddBenefit()
  Dim benReloc As clsBenRelocation
  Dim ben As IBenefitClass
  Dim lst As ListItem
  Dim i As Long
  On Error GoTo AddBenefit_Err
  Call xSet("AddBenefit")

  Set benReloc = New clsBenRelocation
  Set ben = benReloc

  'Put in defaults for benefit ( zzzz Use defaults from the database?? - Good idea
  Set ben.Parent = CurrentEmployee
  Call ben.SetItem(reloc_employeeref, CurrentEmployee.PersonelNo)
  Call ben.SetItem(reloc_item, "Please enter address...")
  Call ben.SetItem(reloc_DateRelocated, rates.GetItem(taxyearstart))
  Call ben.SetItem(reloc_TaxYear, S_THISYEAR)
  Call ben.SetItem(reloc_Lastyear, 0)
  
  Call MDIMain.SetConfirmUndo
  ben.Dirty = True
  ben.ReadFromDB = True
  i = CurrentEmployee.benefits.Add(ben)

  Set lst = LB.ListItems.Add(, , ben.name)
  With lst
    .Tag = i
    .Text = ben.name
    .SubItems(1) = formatworkingnumber(ben.Calculate, "£")
  End With

  Call LoadRelocDetails(lst.Tag)
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
  Call ErrorMessage(ERR_ERROR, Err, "AddBenefit", "ERR_ADDBENEFIT", "Error in AddBenefit function, called from the form " & Me.name & ".")
  Resume AddBenefit_End
End Sub

Private Function IBenefitForm_BenefitToScreen(Optional ByVal lTag As Long = -1, Optional ByVal lIndex As Long = -1&) As IBenefitClass

End Function

Private Property Let IBenefitForm_bentype(NewValue As benClass)
'  ECASE "F_BenCar_BenType"
End Property

Private Property Get IBenefitForm_bentype() As benClass
'  ECASE "F_BenCar_BenType"
End Property

Private Sub IBenefitForm_ClearFields()
'  On Error GoTo ClearFields_Err
'  Call xSet("ClearFields")
'  With Me
'    .TB_Data(0).Text = ""
'    .TB_Data(1).Text = ""
'    .TB_Data(2).Text = ""
'    .TB_Data(3).Text = ""
'    .TB_Data(4).Text = ""
'    .TB_Data(5).Text = ""
'    .TB_Data(6).Text = ""
'    .TB_Data(7).Text = ""
'    .TB_Data(8).Text = ""
'    '.TB_Data(9).Text = ""
'    .Op_Data(0).Value = vbUnchecked
'    .Op_Data(1).Value = vbUnchecked
'    .Op_Data(2).Value = vbUnchecked
'    .Op_Data(3).Value = vbUnchecked
'    '.Op_Data(4).value  = vbUnchecked
'    .Op_Data(5).Value = vbUnchecked
'    .CB_BusMiles.Text = S_UNDER2499
'    .lblAccessories = ""
'  End With
'ClearFields_End:
'  Call xReturn("ClearFields")
'  Exit Sub
'ClearFields_Err:
'  Call ErrorMessage(ERR_ERROR, Err, "ClearFields", "ERR_UNDEFINED", "Undefined error.")
'  Resume ClearFields_End
End Sub

Private Function IBenefitForm_Remove(i As Long) As Boolean
'  On Error GoTo KillBenefit_Err
'  Call xSet("KillBenefit")
'  If Not benefit.CompanyDefined Then
'    Call benefit.DeleteDB
'    Call CurrentEmployee.benefits.Remove(i)
'    Call IBenefitForm_ClearFields
'    Call IBenefitForm_ListBenefits
'  End If
'KillBenefit_End:
'  Call xReturn("KillBenefit")
'  Exit Function
'KillBenefit_Err:
'  Call ErrorMessage(ERR_ERROR, Err, "KillBenefit", "ERR_UNDEFINED", "Undefined error.")
'  Resume KillBenefit_End
End Function

Private Function IBenefitForm_ListBenefits(Optional ByVal Index As Long = 0&) As Boolean
  Dim ben As IBenefitClass
  Dim benfrm As IBenefitForm
  Dim lst As ListItem
  Dim i As Long, j As Long
  Dim lben As Variant

  On Error GoTo F_BenReloc_ListBenefits_Err
  Call xSet("F_BenReloc_ListBenefits")
  i = 0
  Call LockWindowUpdate(LB.hWnd)
  Set benfrm = Me
  Me.LB.ListItems.Clear
  benfrm.ClearFields
  Call MDIMain.SetAdd
  For i = 1 To CurrentEmployee.benefits.count
    Set ben = CurrentEmployee.benefits(i)
    If Not (ben Is Nothing) Then
      If ben.BenefitClass = BC_RELOC Then
        lben = ben.Calculate
        Set lst = LB.ListItems.Add(, , ben.name)
        lst.Tag = i
        If VarType(lben) = vbString Then
          lst.SubItems(1) = lben
        Else
          lst.SubItems(1) = formatworkingnumber(lben, "£")
       End If
        j = 1
      End If
    End If
  Next i
  If j = 0 Then
    Me.LB.Enabled = False
    Call MDIMain.ClearDelete
    Call MDIMain.ClearConfirmUndo
    Set Me.benefit = Nothing
  Else
    Me.LB.Enabled = True
    Call MDIMain.SetDelete
    LB.SelectedItem = LB.ListItems(1)
    Call LoadRelocDetails(LB.SelectedItem.Tag)
  End If
  
F_BenReloc_ListBenefits_End:
  Call LockWindowUpdate(0)
  Call xReturn("F_BenReloc_ListBenefits")
  Exit Function

F_BenReloc_ListBenefits_Err:
  Call ErrorMessage(ERR_ERROR, Err, "F_BenReloc_ListBenefits", "Relocation benefits", "Error listing benefits.")
  Resume F_BenReloc_ListBenefits_End
End Function

Private Sub Form_Load()
  If Not (mclsResize.InitResize(Me, L_DES_HEIGHT, L_DES_WIDTH, DESIGN, , , MDIMain)) Then
    Err.Raise ERR_Application
  End If
  Me.CboBx(0).AddItem (S_THISYEAR)
  Me.CboBx(0).AddItem (S_LASTYEAR)
End Sub

Private Function CheckChanged(ctl As Control) As Boolean
  Dim lst As ListItem
  Dim Itemlst As ListItem
  Dim i As Long, d0 As Variant
  Dim detail As New clsRelocDetail
  Dim RefreshDetail As Boolean
  
  On Error GoTo CheckChanged_Err
  Call xSet("CheckChanged")

  RefreshDetail = False
  Set detail = New clsRelocDetail
  If CurrentEmployee Is Nothing Then GoTo CheckChanged_End
  If benefit Is Nothing Then GoTo CheckChanged_End
  Select Case ctl.name
    Case "TB_Data"
      Select Case ctl.Index
        Case 0
          i = StrComp(ctl.Text, benefit.GetItem(reloc_item), vbBinaryCompare)
          If i <> 0 Then
            Call benefit.SetItem(reloc_item, ctl.Text)
            RefreshDetail = True
          End If
         Case 1
          i = StrComp(ctl.Text, benefit.GetItem(reloc_Lastyear), vbBinaryCompare)
          If i <> 0 Then
            Call benefit.SetItem(reloc_Lastyear, ctl.Text)
            RefreshDetail = True
          End If
        Case 2
          i = StrComp(ctl.Text, benefit.GetItem(reloc_currentrelocitem), vbBinaryCompare)
          If i <> 0 Then
            Call benefit.SetItem(reloc_currentrelocitem, ctl.Text)
            Set lst = LB.SelectedItem
            Set Itemlst = lbItems.SelectedItem
            If Not lst Is Nothing And Not Itemlst Is Nothing Then
              Call RelocItemDetails(lst.Tag, F_BenReloc.TB_Data(6).Text)
            End If
          End If
        Case 3
          i = StrComp(ctl.Text, benefit.GetItem(reloc_currentrelocgr), vbBinaryCompare)
          If i <> 0 Then
            Call benefit.SetItem(reloc_currentrelocgr, ctl.Text)
            Set lst = LB.SelectedItem
            Set Itemlst = lbItems.SelectedItem
            If Not lst Is Nothing And Not Itemlst Is Nothing Then
              Call RelocItemDetails(lst.Tag, F_BenReloc.TB_Data(6).Text)
            End If
          End If
        Case 4
          i = StrComp(ctl.Text, benefit.GetItem(reloc_currentrelocgr), vbBinaryCompare)
          If i <> 0 Then
            Call benefit.SetItem(reloc_currentrelocgr, ctl.Text)
            
            Set lst = LB.SelectedItem
            If Not lst Is Nothing Then
              Call RelocItemDetails(lst.Tag, F_BenReloc.TB_Data(6).Text)
            End If
          End If
        Case 5
          i = StrComp(ctl.Text, benefit.GetItem(reloc_currentreloccomments), vbBinaryCompare)
          If i <> 0 Then
            Call benefit.SetItem(reloc_currentreloccomments, ctl.Text)
            Set lst = LB.SelectedItem
            Set Itemlst = lbItems.SelectedItem
            If Not lst Is Nothing And Not Itemlst Is Nothing Then
              Call RelocItemDetails(lst.Tag, Itemlst.Tag)
            End If
          End If
        Case Else
          ECASE "Unknown control"
      End Select
    Case "Op_Data"
      Select Case ctl.Index
        Case 0
          i = (IIf(ctl.value = vbChecked, True, False) <> benefit.GetItem(reloc_currentrelocqual))
          If i <> 0 Then
            Call benefit.SetItem(reloc_currentrelocqual, IIf(ctl.value = vbChecked, True, False))
            Set lst = LB.SelectedItem
            Set Itemlst = lbItems.SelectedItem
            If Not lst Is Nothing And Not Itemlst Is Nothing Then
              Call RelocItemDetails(lst.Tag, Itemlst.Tag)
              
            End If
          End If
        Case 1
        Case Else
          ECASE "Unknown control"
      End Select
    Case "CboBx"
      i = StrComp(ctl.Text, benefit.GetItem(reloc_TaxYear), vbBinaryCompare)
      If i <> 0 Then
        Call benefit.SetItem(reloc_TaxYear, ctl.Text)
        RefreshDetail = True
      End If
    Case Else
      ECASE "Unknown control"
  End Select
  If i <> 0 Then benefit.InvalidFields = InvalidFields(Me)
  If benefit.InvalidFields > 0 Then
    Call MDIMain.sts.SetStatus(0, "", S_NOSAVE)
    Call MDIMain.SetUndo
    benefit.Dirty = False
    Set lst = LB.SelectedItem
    If Not lst Is Nothing Then
      With lst
        .Text = benefit.name
        .SubItems(1) = formatworkingnumber(benefit.Calculate, "£")
      End With
    End If
    If RefreshDetail Then
      Set lst = lbItems.SelectedItem
      If Not lst Is Nothing Then
        With lst
          .Text = F_BenReloc.TB_Data(2).Text
          .SubItems(1) = formatworkingnumber(CDbl(F_BenReloc.TB_Data(3).Text) - CDbl(F_BenReloc.TB_Data(3).Text), "£")
          .SubItems(2) = IIf(Op_Data(0).value, "Yes", "No")
        End With
      End If
    End If
  ElseIf i <> 0 Then
    Call MDIMain.sts.SetStatus(0, "", "")
    Call MDIMain.SetConfirmUndo
    benefit.Dirty = True
    Set lst = LB.SelectedItem
    With lst
      .Text = benefit.name
      .SubItems(1) = formatworkingnumber(benefit.Calculate, "£")
    End With
    Lab(1).Caption = formatworkingnumber(benefit.GetItem(reloc_NQBenefit), "£")
    If RefreshDetail Then
      Set lst = lbItems.SelectedItem
      If Not lst Is Nothing Then
        With lst
          .Text = F_BenReloc.TB_Data(2).Text
          .SubItems(1) = formatworkingnumber(CDbl(F_BenReloc.TB_Data(3).Text) - CDbl(F_BenReloc.TB_Data(4).Text), "£")
          .SubItems(2) = IIf(Op_Data(0).value, "Yes", "No")
        End With
      End If
    End If
  ElseIf benefit.Dirty Then
    Call MDIMain.sts.SetStatus(0, "", "")
    Call MDIMain.SetConfirmUndo
  End If

CheckChanged_End:
  Call xReturn("CheckChanged")
  Exit Function

CheckChanged_Err:
  Call ErrorMessage(ERR_ERROR, Err, "CheckChanged", "ERR_CHECKCHANGED", "This function has failed for the form " & Me.name & ".")
  Resume CheckChanged_End
End Function

Private Sub LB_Click()
  If Not (LB.SelectedItem Is Nothing) Then
    LB.SelectedItem.Selected = True
  End If
End Sub

Private Sub LB_ItemClick(ByVal Item As ComctlLib.ListItem)
  If Not (LB.SelectedItem Is Nothing) Then
    Call LoadRelocDetails(Item.Tag)
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

Public Sub AddItem()
  Dim relocation As clsBenRelocation
  Dim detail As clsRelocDetail
  Dim lst As ListItem, i As Long
  On Error GoTo AddItem_Err
  Call xSet("AddItem")
  Set detail = New clsRelocDetail
  Set relocation = benefit
  
  'Put in defaults for benefit
  detail.Item = "Please enter description"
  detail.value = 0
  detail.MadeGood = 0
  detail.Comments = ""
  detail.Qualify = True
  
  F_BenReloc.TB_Data(2).Text = detail.Item
  F_BenReloc.TB_Data(3).Text = detail.value
  F_BenReloc.TB_Data(4).Text = detail.MadeGood
  F_BenReloc.TB_Data(5).Text = detail.Comments
  F_BenReloc.Op_Data(0).value = IIf(detail.Qualify, 1, 0)
  
  Call benefit.SetItem(reloc_currentrelocitem, detail.Item)
  Call benefit.SetItem(reloc_currentrelocgr, detail.value)
  Call benefit.SetItem(reloc_currentrelocmd, detail.MadeGood)
  Call benefit.SetItem(reloc_currentrelocqual, detail.Qualify)
  Call benefit.SetItem(reloc_currentreloccomments, detail.Comments)
  
  Call MDIMain.SetConfirmUndo
  benefit.Dirty = True
  i = relocation.RelocDetails.Add(detail)
  Set lst = lbItems.ListItems.Add(, , detail.Item)
  Call benefit.Calculate
  With lst
    .Tag = i
    .SubItems(1) = formatworkingnumber(detail.benefit, "£")
  End With
  
  Set LB.SelectedItem = lst
  Call MDIMain.SetDelete
  
AddItem_End:
  Call xReturn("AddItem")
  Exit Sub
AddItem_Err:
  Call ErrorMessage(ERR_ERROR, Err, "AddItem", "ERR_AddItem", "Error in AddItem function, called from the form " & Me.name & ".")
  Resume AddItem_End
End Sub

Private Sub LBItems_Click()
  If Not (lbItems.SelectedItem Is Nothing) Then
    lbItems.SelectedItem.Selected = True
  End If
End Sub
Private Sub LBItems_ItemClick(ByVal Item As ComctlLib.ListItem)
  If Not (lbItems.SelectedItem Is Nothing) And Not (LB.SelectedItem Is Nothing) Then
    Call DisplayRelocDetailsEx(LB.SelectedItem.Tag, Item.Tag)
  End If
End Sub

Private Sub Op_Data_Click(Index As Integer)
  Call CheckChanged(Op_Data(Index))
End Sub

Private Sub Op_Data_LostFocus(Index As Integer)
  Call CheckChanged(Op_Data(Index))
End Sub

Private Sub TB_Data_FieldInvalid(Index As Integer, Valid As Boolean, Message As String)
  Call MDIMain.sts.SetStatus(0, Message)
End Sub

Private Sub TB_data_Lostfocus(Index As Integer)
  'Debug.Print "CheckChanged called"
  Call CheckChanged(TB_Data(Index))
End Sub
