VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.2#0"; "COMCTL32.OCX"
Begin VB.Form F_Ben 
   Caption         =   "All Benefits"
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
   Icon            =   "Benefits.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5685
   ScaleWidth      =   8325
   WindowState     =   2  'Maximized
   Begin ComctlLib.ListView lb 
      Height          =   4920
      Left            =   45
      TabIndex        =   0
      Tag             =   "free,font"
      Top             =   45
      Width           =   8220
      _ExtentX        =   14499
      _ExtentY        =   8678
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
      NumItems        =   2
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Benefit description"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Alignment       =   1
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Value"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Total benefit for this employee"
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
      Left            =   240
      TabIndex        =   1
      Tag             =   "FREE,FONT"
      Top             =   5280
      Width           =   2535
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
      Left            =   6240
      TabIndex        =   2
      Tag             =   "FREE,FONT"
      Top             =   5265
      Width           =   1725
   End
End
Attribute VB_Name = "F_Ben"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Implements IBenefitForm
Implements IBenefitForm2

Private mclsResize As New clsFormResize
Private Const L_DES_HEIGHT = 6090
Private Const L_DES_WIDTH = 8445

Private Sub Form_Resize()
  mclsResize.Resize
  Call ColumnWidths(lb, 70, 15)
End Sub
Private Sub IBenefitForm2_AddBenefit()
  Ecase ("AddBenefit")
End Sub

Private Property Get IBenefitForm2_benefit() As IBenefitClass
  'not required
End Property

Private Property Let IBenefitForm2_benefit(NewValue As IBenefitClass)
  'not required
End Property

Private Function IBenefitForm2_BenefitFormState(ByVal fState As BENEFIT_FORM_STATE) As Boolean
On Error GoTo BenefitFormState_err

  Call xSet("IBenefitForm2_EnableBenefitForm")
  
  If fState = FORM_ENABLED Then
    lb.Enabled = True
    Call MDIMain.ClearAdd
    Call MDIMain.SetDelete
  ElseIf fState = FORM_DISABLED Then
    lb.Enabled = False
    Call MDIMain.ClearAdd
    Call MDIMain.ClearDelete
  End If
  Call MDIMain.ClearConfirmUndo
  
  IBenefitForm2_BenefitFormState = True
    
BenefitFormState_end:
  Call xReturn("BenefitFormState")
  Exit Function
  
BenefitFormState_err:
  IBenefitForm2_BenefitFormState = False
  Call ErrorMessage(ERR_ERROR, Err, "BenefitFormState", "ERR_UNDEFINED", "Undefined error.")
  Resume BenefitFormState_end

End Function

Private Function IBenefitForm2_BenefitsToListView() As Long
  Dim l As Long, lTotal As Long
  Dim li As ListItem
  Dim vBen As Variant
  Dim vancol As clsVansCollection
  Dim ben As IBenefitClass
  Dim ibf As IBenefitForm2
  
On Error GoTo BenefitsToListView_err
  
  Call xSet("BenefitsToListView")
  
  With lb
    .ListItems.Clear
    Set vancol = CurrentEmployee.benefits(1)
    If vancol.count > 0 Or CurrentEmployee.SharedVan = True Then
      Set ben = CurrentEmployee.benefits(1)
      vBen = ben.Calculate
      If VarType(vBen) <> vbString Then
        lTotal = lTotal + vBen
      End If
      Set li = .ListItems.Add(, , ben.Name)
      lb.Tag = 1
      li.SubItems(1) = formatworkingnumber(vBen, "£")
      IBenefitForm2_BenefitsToListView = IBenefitForm2_BenefitsToListView + 1
    End If
    
    For l = 2 To CurrentEmployee.benefits.count
      Set ben = CurrentEmployee.benefits(l)
      If Not ben Is Nothing Then
        vBen = ben.Calculate
        If VarType(vBen) <> vbString Then
          lTotal = lTotal + vBen
        End If
        Set li = .ListItems.Add(, , ben.Name)
        li.Tag = l
        li.SubItems(1) = formatworkingnumber(vBen, "£")
        IBenefitForm2_BenefitsToListView = IBenefitForm2_BenefitsToListView + 1
      End If
    Next
  End With
  
  Set ibf = Me
  
  If IBenefitForm2_BenefitsToListView > 0 Then
    ibf.BenefitFormState (FORM_ENABLED)
  Else
    ibf.BenefitFormState (FORM_DISABLED)
  End If
  
  lblTotal.Caption = formatworkingnumber(lTotal, "£") & " "
  
  
BenefitsToListView_end:
  Set ben = Nothing
  Set ibf = Nothing
  Set vancol = Nothing
  Call xReturn("BenefitsToListView")
  Exit Function
BenefitsToListView_err:
  Call ErrorMessage(ERR_ERROR, Err, "BenefitsToListView", "Benefits To List View", "Unable to place all the befits to the list view.")
  Resume BenefitsToListView_end
End Function

Private Function IBenefitForm2_BenefitToScreen(Optional ByVal BenefitIndex As Long = -1&, Optional ByVal UpdateBenefit As Boolean = True) As IBenefitClass
  'no entry in here
End Function

Private Property Let IBenefitForm2_bentype(ByVal RHS As benClass)
  Ecase ("bentype")
End Property

Private Property Get IBenefitForm2_bentype() As benClass
  Dim ben As IBenefitClass
  
On Error GoTo bentype_ERR
  
  Call xSet("bentype")
  
  If Not lb.SelectedItem Is Nothing Then
    Set ben = CurrentEmployee.benefits(lb.SelectedItem.Tag)
    IBenefitForm2_bentype = ben.BenefitClass
  End If
  
bentype_END:
  Set ben = Nothing
  Call xSet("bentype")
  Exit Function
bentype_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "bentype", "ben type", "Error obtaining the benfit type on the all benefits screen.")
  Resume bentype_END
End Property

Private Property Get IBenefitForm2_lv() As ComctlLib.IListView
  Set IBenefitForm2_lv = lb
End Property

Private Function IBenefitForm2_RemoveBenefit(ByVal BenefitIndex As Long) As Boolean
  Dim IBen As IBenefitClass

On Error GoTo RemoveBenefit_ERR
    
  Call xSet("RemoveBenefit")
  
  Set IBen = CurrentEmployee.benefits(BenefitIndex)
  IBenefitForm2_RemoveBenefit = RemoveBenefit(Me, IBen, BenefitIndex)
  
RemoveBenefit_END:
  Set IBen = Nothing
  Call xReturn("RemoveBenefit")
  Exit Function
RemoveBenefit_ERR:
  IBenefitForm2_RemoveBenefit = False
  Call ErrorMessage(ERR_ERROR, Err, "RemoveBenefit", "Remove Benefit", "Error removing benefit from all benefits screen.")
  Resume RemoveBenefit_END
End Function

Private Function IBenefitForm2_UpdateBenefitListViewItem(li As ComctlLib.IListItem, benefit As IBenefitClass, Optional ByVal BenefitIndex As Long = 0&, Optional ByVal SelectItem As Boolean = False) As Boolean
  Ecase ("UpadateBenefitListViewItem")
End Function

Private Sub LB_ColumnClick(ByVal ColumnHeader As ComctlLib.ColumnHeader)
  Me.lb.SortKey = ColumnHeader.Index - 1
  lb.SelectedItem.EnsureVisible
End Sub

Private Sub LB_DblClick()
  If Not (lb.SelectedItem Is Nothing) Then
    Call ViewBenefit(lb.SelectedItem.Tag)
  End If
End Sub

Private Sub Form_Load()

  On Error GoTo Form_Load_Err
  Call xSet("Form_Load")
  
  If Not (mclsResize.InitResize(Me, L_DES_HEIGHT, L_DES_WIDTH, DESIGN, , , MDIMain)) Then
    Err.Raise ERR_Application
  End If
  
Form_Load_End:
  Call xReturn("Form_Load")
  Exit Sub

Form_Load_Err:
  Call ErrorMessage(ERR_ERROR, Err, "Form_Load", "ERR_UNDEFINED", "Undefined error.")
  Resume Form_Load_End
End Sub

Private Sub LB_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then 'Return key
    Call LB_DblClick
  End If
End Sub
