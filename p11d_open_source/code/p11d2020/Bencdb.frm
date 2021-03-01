VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.2#0"; "COMCTL32.OCX"
Begin VB.Form F_BenCDB 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Company Defined Benefits"
   ClientHeight    =   6090
   ClientLeft      =   1275
   ClientTop       =   2745
   ClientWidth     =   8445
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   ScaleHeight     =   6090
   ScaleWidth      =   8445
   Begin VB.Frame FMECDB 
      BorderStyle     =   0  'None
      Height          =   1845
      Left            =   7020
      TabIndex        =   5
      Top             =   240
      Width           =   1215
      Begin VB.CommandButton B_Add 
         Caption         =   "&Add"
         Height          =   315
         Left            =   0
         TabIndex        =   1
         Tag             =   "free,font"
         Top             =   30
         Width           =   1200
      End
      Begin VB.CommandButton B_Delete 
         Caption         =   "&Delete"
         Height          =   315
         Left            =   0
         TabIndex        =   2
         Tag             =   "free,font"
         Top             =   510
         Width           =   1200
      End
      Begin VB.CommandButton B_Edit 
         Caption         =   "&Edit"
         Height          =   315
         Left            =   0
         TabIndex        =   3
         Tag             =   "free,font"
         Top             =   990
         Width           =   1200
      End
      Begin VB.CommandButton B_Apply 
         Caption         =   "A&pply"
         Height          =   315
         Left            =   0
         TabIndex        =   4
         Tag             =   "free,font"
         Top             =   1470
         Width           =   1200
      End
   End
   Begin ComctlLib.ListView LB 
      Height          =   5355
      Left            =   60
      TabIndex        =   0
      Tag             =   "free,font"
      Top             =   240
      Width           =   6705
      _ExtentX        =   11827
      _ExtentY        =   9446
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
      NumItems        =   5
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Code"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Category"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   2
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Description"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Alignment       =   1
         SubItemIndex    =   3
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Value"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Alignment       =   1
         SubItemIndex    =   4
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Made Good"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "F_BenCDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Implements IBenefitForm

Public benefit As IBenefitClass
Private mclsResize As New clsFormResize
Private Const L_DES_HEIGHT = 6090
Private Const L_DES_WIDTH = 8445

Private Sub B_Apply_Click()
  Call LoadApply
End Sub

Private Sub B_Edit_Click()
  Call LoadCDB(False)
End Sub

Private Sub Form_Load()
  If Not (mclsResize.InitResize(Me, L_DES_HEIGHT, L_DES_WIDTH, DESIGN, , , MDIMain)) Then
    Err.Raise ERR_Application
  End If
  Call ColumnWidths(lb, 20, 20, 20, 20, 20)
End Sub

Private Sub Form_Resize()
  mclsResize.Resize
  Call ColumnWidths(lb, 20, 20, 20, 20, 20)
End Sub
  

Private Sub IBenefitForm_AddBenefit()
'  Dim benother As Other
'  Dim ben As IBenefitClass
'  Dim lst As ListItem, i As Long
'  On Error GoTo AddBenefit_Err
'  Call xSet("AddBenefit")
'
'  Set benother = New Other
'  Set ben = benother
'  Set benefit = ben
'  Call LoadCDB(True)
'  Call MDIMain.SetConfirmUndo
'  ben.CompanyDefined = True
'  'Set ben.Parent = P11d32.CurrentEmployer.SysEmployee
'  Call ben.SetItem(Oth_Category, P11d32.Rates.ClassStrToCategory(ben.GetItem(Oth_Class)))
'  Call ben.SetItem(Oth_UDBCode, P11d32.Rates.ClassStrToCode(ben.GetItem(Oth_Class)))
'  ben.Dirty = True
'  ben.ReadFromDB = True
'  i = P11d32.CurrentEmployer.SysEmployee.benefits.Add(ben)
'
'  Set lst = lb.ListItems.Add(, , Right$(ben.GetItem(Oth_EmployeeReference), Len(ben.GetItem(Oth_EmployeeReference)) - 4))
'  lst.SubItems(1) = ben.GetItem(Oth_Class)
'  lst.SubItems(2) = ben.GetItem(Oth_item)
'  lst.SubItems(3) = ben.GetItem(Oth_Value)
'  lst.SubItems(4) = ben.GetItem(Oth_MadeGood)
'  lst.Tag = i
'
'  Call CDBDetails(lst.Tag)
'  Set lb.SelectedItem = lst
'
'  Me.lb.Enabled = True
'  Me.fmeCDB.Enabled = True
'  Call MDIMain.SetDelete
'
'AddBenefit_End:
'  Set ben = Nothing
'  Call xReturn("AddBenefit")
'  Exit Sub
'AddBenefit_Err:
'  Call ErrorMessage(ERR_ERROR, Err, "AddBenefit", "ERR_ADDBENEFIT", "Error in AddBenefit function, called from the form " & Me.Name & ".")
'  Resume AddBenefit_End


End Sub

Private Function IBenefitForm_BenefitToScreen(Optional ByVal lTag As Long = -1, Optional ByVal lIndex As Long = -1&) As IBenefitClass

End Function

Private Property Let IBenefitForm_bentype(NewValue As benClass)

End Property

Private Property Get IBenefitForm_bentype() As benClass

End Property

Private Sub IBenefitForm_ClearFields()

End Sub

Private Function IBenefitForm_ListBenefits(Optional ByVal Index As Long = 0&) As Boolean
'  Dim ben As IBenefitClass
'  Dim benfrm As IBenefitForm
'  Dim lst As ListItem
'  Dim i As Long, j As Long
'  Dim lben As Variant
'
'  On Error GoTo F_CDB_ListBenefits_Err
'  Call xSet("F_CDB_ListBenefits")
'  i = 0
'  Call LockWindowUpdate(lb.hWnd)
'  Set benfrm = Me
'  Me.lb.ListItems.Clear
'  benfrm.ClearFields
'  Call MDIMain.SetAdd
'  For i = 1 To P11d32.CurrentEmployer.SysEmployee.benefits.count
'    Set ben = P11d32.CurrentEmployer.SysEmployee.benefits(i)
'    If Not (ben Is Nothing) Then
'      Set lst = lb.ListItems.Add(, , Right$(ben.GetItem(Oth_EmployeeReference), Len(ben.GetItem(Oth_EmployeeReference)) - 4))
'      lst.SubItems(1) = ben.GetItem(Oth_Class)
'      lst.SubItems(2) = ben.GetItem(Oth_item)
'      lst.SubItems(3) = ben.GetItem(Oth_Value)
'      lst.SubItems(4) = ben.GetItem(Oth_MadeGood)
'      lst.Tag = i
'      j = 1
'    End If
'  Next i
'  If j = 0 Then
'    Me.fmeCDB.Enabled = False
'    Me.lb.Enabled = False
'    Call MDIMain.ClearDelete
'    Call MDIMain.ClearConfirmUndo
'    Set Me.benefit = Nothing
'  Else
'    Me.fmeCDB.Enabled = True
'    Me.lb.Enabled = True
'    Call MDIMain.SetDelete
'    lb.SelectedItem = lb.ListItems(1)
'  End If
'F_cdb_ListBenefits_End:
'  Call LockWindowUpdate(0)
'  Call xReturn("F_CDB_ListBenefits")
'  Exit Function
'F_CDB_ListBenefits_Err:
'  Call ErrorMessage(ERR_ERROR, Err, "F_SharedVans_ListBenefits", "Car Benefits", "Error listing benefits.")
'  Resume F_cdb_ListBenefits_End
End Function

Private Function IBenefitForm_Remove(i As Long) As Boolean
'  Call benefit.DeleteDB
'  Call P11d32.CurrentEmployer.SysEmployee.benefits.Remove(i)
'  Call IBenefitForm_ClearFields
'  Call IBenefitForm_ListBenefits
End Function

Private Sub LB_ColumnClick(ByVal ColumnHeader As ComctlLib.ColumnHeader)
  lb.SortKey = ColumnHeader.Index - 1
  lb.SelectedItem.EnsureVisible
End Sub

Private Sub LB_KeyPress(KeyAscii As Integer)
  'Check for return key - tab to next field
  If KeyAscii = 13 Then Call SendKeys(vbTab)
End Sub

Private Sub lb_DblClick()
  Call LoadCDB(False)
End Sub

Private Sub LB_ItemClick(ByVal Item As ComctlLib.ListItem)
  If Not (lb.SelectedItem Is Nothing) Then
    Call CDBDetails(Item.Tag)
  End If
End Sub

