VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form F_AllBenefits 
   Caption         =   " "
   ClientHeight    =   5685
   ClientLeft      =   780
   ClientTop       =   2040
   ClientWidth     =   8325
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
   Icon            =   "F_AllBenefits.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5685
   ScaleWidth      =   8325
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ListView lb 
      Height          =   5130
      Left            =   0
      TabIndex        =   0
      Tag             =   "free,font"
      Top             =   45
      Width           =   8220
      _ExtentX        =   14499
      _ExtentY        =   9049
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
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Benefit Description"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "P11d Category"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Category Description"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Benefit"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label lblTotalClass1a 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Total Class 1A"
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
      Height          =   240
      Left            =   2630
      TabIndex        =   4
      Tag             =   "FREE,FONT"
      Top             =   5295
      Width           =   1185
   End
   Begin VB.Label lblTotalNIC 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
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
      Height          =   255
      Left            =   3690
      TabIndex        =   3
      Tag             =   "FREE,FONT"
      Top             =   5288
      Width           =   1725
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Total benefit"
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
      Height          =   240
      Left            =   5280
      TabIndex        =   1
      Tag             =   "FREE,FONT"
      Top             =   5295
      Width           =   1125
   End
   Begin VB.Label lblTotal 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
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
      Height          =   255
      Left            =   6480
      TabIndex        =   2
      Tag             =   "FREE,FONT"
      Top             =   5288
      Width           =   1725
   End
End
Attribute VB_Name = "F_AllBenefits"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Implements IBenefitForm2
Implements IFrmGeneral

Private mclsResize As New clsFormResize
Private Const L_DES_HEIGHT  As Long = 6090
Private Const L_DES_WIDTH  As Long = 8445


Private Sub Form_GotFocus()

  'Call MDIMain.ClearAdd 'EK
  'Call MDIMain.ClearDelete 'EK

End Sub

Private Sub Form_Resize()
  mclsResize.Resize
  Call ColumnWidths(lb, 35, 15, 35, 10)
End Sub
Private Sub IBenefitForm2_AddBenefit()
  ECASE ("AddBenefit")
End Sub

Private Function IBenefitForm2_AddBenefitSetDefaults(ben As IBenefitClass) As Long

End Function

Private Property Get IBenefitForm2_benefit() As IBenefitClass
  Set IBenefitForm2_benefit = p11d32.CurrentEmployer.CurrentEmployee
End Property

Private Property Set IBenefitForm2_benefit(NewValue As IBenefitClass)
  'not required
End Property

Private Function IBenefitForm2_BenefitFormState(ByVal fState As BENEFIT_FORM_STATE) As Boolean
  If fState = FORM_DISABLED Then
    Call SetLVEnabled(lb, False)
  Else
    Call SetLVEnabled(lb, True)
  End If
  Call MDIMain.ClearConfirmUndo
  Call MDIMain.ClearAdd
  Call MDIMain.ClearDelete
  IBenefitForm2_BenefitFormState = True
End Function

Private Function IBenefitForm2_BenefitOff() As Boolean

End Function

Private Function IBenefitForm2_BenefitOn() As Boolean

End Function

Private Function IBenefitForm2_BenefitsToListView() As Long
  Dim benEmployee As IBenefitClass
  Dim i As Long, j As Long
  Dim ben As IBenefitClass
  Dim ibf As IBenefitForm2
  Dim loans As loans
  
On Error GoTo BenefitsToListView_err
  
  Call xSet("BenefitsToListView")

  Set ibf = Me
  
  Set benEmployee = p11d32.CurrentEmployer.CurrentEmployee
  lblTotal.Caption = FormatWN(benEmployee.Calculate)
  lblTotalNIC.Caption = FormatWN(benEmployee.value(ITEM_NIC_CLASS1A_BENEFIT))
  
  
  With ibf.lv
    .SmallIcons = MDIMain.imlListViewBenefits
    .listitems.Clear
    For i = 1 To p11d32.CurrentEmployer.CurrentEmployee.benefits.Count
      Set ben = p11d32.CurrentEmployer.CurrentEmployee.benefits(i)
      
      If ben Is Nothing Then GoTo NEXT_ITEM
      If (ben.BenefitClass = BC_EMPLOYEE_CAR_E) Then
        ben.BenefitClass = BC_EMPLOYEE_CAR_E
      End If
      Select Case ben.BenefitClass
        Case BC_NONSHAREDVANS_G
          If p11d32.CurrentEmployer.CurrentEmployee.AnyVanBenefit Then
            IBenefitForm2_BenefitsToListView = IBenefitForm2_BenefitsToListView + 1
            Call ibf.BenefitToListView(ben, i)
          Else
            GoTo NEXT_ITEM
          End If
        Case BC_LOANS_H
          Set loans = ben
          For j = 1 To loans.loans.Count
            Set ben = loans.loans(j)
            If Not ben Is Nothing Then
              IBenefitForm2_BenefitsToListView = IBenefitForm2_BenefitsToListView + 1
              Call ibf.BenefitToListView(ben, j)
            End If
          Next
       Case Else
          Call ibf.BenefitToListView(ben, i)
          IBenefitForm2_BenefitsToListView = IBenefitForm2_BenefitsToListView + 1
          
      End Select
NEXT_ITEM:
      Next
    End With
  

  
  
  
BenefitsToListView_end:
  
  
  Call xReturn("BenefitsToListView")
  Exit Function
BenefitsToListView_err:
  Call ErrorMessage(ERR_ERROR, Err, "BenefitsToListView", "Benefits To List View", "Unable to place all the benefits to the list view.")
  Resume BenefitsToListView_end
  Resume
End Function
Private Function IBenefitForm2_BenefitToListView(ben As IBenefitClass, ByVal lBenefitIndex As Long) As Long
  Dim ibf As IBenefitForm2
  Dim li As ListItem
  Dim vBen As Variant
  
  Set ibf = Me
  Set li = lb.listitems.Add()
  vBen = ben.Calculate
  If IsNumeric(vBen) Then IBenefitForm2_BenefitToListView = vBen
  Call ibf.UpdateBenefitListViewItem(li, ben, lBenefitIndex)
  
End Function

Private Function IBenefitForm2_BenefitToScreen(Optional ByVal BenefitIndex As Long = -1&, Optional ByVal UpdateBenefit As Boolean = True) As Boolean
  'no entry in here
End Function

Private Property Let IBenefitForm2_benclass(ByVal RHS As BEN_CLASS)
  
End Property

Private Property Get IBenefitForm2_benclass() As BEN_CLASS
  Call MDIMain.ClearAdd 'EK
  Call MDIMain.ClearDelete 'EK
  IBenefitForm2_benclass = BC_ALL
End Property


'Private Property Get IBenefitForm2_ControlDefault() As Control''
'
'End Property

Private Property Get IBenefitForm2_lv() As MSComctlLib.IListView
  Set IBenefitForm2_lv = lb
End Property

Private Function IBenefitForm2_RemoveBenefit(ByVal BenefitIndex As Long) As Boolean
  Dim IBen As IBenefitClass

On Error GoTo RemoveBenefit_ERR
    
  Call xSet("RemoveBenefit")
  
  Set IBen = p11d32.CurrentEmployer.CurrentEmployee.benefits(BenefitIndex)
  IBenefitForm2_RemoveBenefit = p11d32.CurrentEmployer.CurrentEmployee.RemoveBenefit(Me, IBen, BenefitIndex)
  
RemoveBenefit_END:
  Set IBen = Nothing
  Call xReturn("RemoveBenefit")
  Exit Function
RemoveBenefit_ERR:
  IBenefitForm2_RemoveBenefit = False
  Call ErrorMessage(ERR_ERROR, Err, "RemoveBenefit", "Remove Benefit", "Error removing benefit from all benefits screen.")
  Resume RemoveBenefit_END
  Resume
End Function

Private Function IBenefitForm2_UpdateBenefitListViewItem(li As MSComctlLib.IListItem, benefit As IBenefitClass, Optional ByVal BenefitIndex As Long = 0&, Optional ByVal SelectItem As Boolean = False) As Long
  If BenefitIndex > 0 Then li.Tag = BenefitIndex
  li.SmallIcon = benefit.ImageListKey
  li.SubItems(3) = FormatWN(benefit.Calculate)
  li.Text = benefit.Name
  li.SubItems(1) = p11d32.Rates.BenClassTo(benefit.BenefitClass, BCT_HMIT_SECTION_STRING)
  li.SubItems(2) = p11d32.Rates.BenClassTo(benefit.BenefitClass, BCT_FORM_CAPTION)
  IBenefitForm2_UpdateBenefitListViewItem = li.Index
End Function
Private Function IBenefitForm2_ValididateBenefit(ben As IBenefitClass) As Boolean
  'not used
End Function

Private Function IFrmGeneral_CheckChanged(c As Control) As Boolean
  'not used
End Function

Private Property Get IFrmGeneral_InvalidVT() As Control
  'not used
End Property

Private Property Set IFrmGeneral_InvalidVT(RHS As Control)
  'not used
End Property

Private Sub lb_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
  Call SetSortOrder(lb, ColumnHeader)
End Sub
Private Sub lb_DblClick()
  Dim ben As IBenefitClass
  Dim loans As loans
  
  If Not (lb.SelectedItem Is Nothing) Then
    If StrComp(lb.SelectedItem.SubItems(1), p11d32.Rates.BenClassTo(BC_LOAN_OTHER_H, BCT_HMIT_SECTION_STRING)) = 0 Then
      Set loans = p11d32.CurrentEmployer.CurrentEmployee.benefits(p11d32.CurrentEmployer.CurrentEmployee.GetLoansBenefitIndex)
      Set ben = loans.loans(lb.SelectedItem.Tag)
      Set loans = Nothing
    Else
      Set ben = p11d32.CurrentEmployer.CurrentEmployee.benefits(lb.SelectedItem.Tag)
    End If
    Call BenScreenSwitch(ben.BenefitClass)
  End If
  
End Sub

Private Sub Form_Load()

  On Error GoTo Form_Load_ERR
  Call xSet("Form_Load")
  
  If Not (mclsResize.InitResize(Me, L_DES_HEIGHT, L_DES_WIDTH, DESIGN, , , MDIMain)) Then
    Err.Raise ERR_Application
  End If
  
'
Form_Load_END:
  Call xReturn("Form_Load")
  Exit Sub

Form_Load_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "Form_Load", "Form Load", "Error loading the All Benefits form.")
  Resume Form_Load_END
End Sub

Private Sub lb_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then 'Return key
    Call lb_DblClick
  End If
End Sub

