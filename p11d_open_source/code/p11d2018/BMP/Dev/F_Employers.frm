VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form F_Employers 
   Appearance      =   0  'Flat
   Caption         =   "Employer Details"
   ClientHeight    =   5655
   ClientLeft      =   840
   ClientTop       =   1725
   ClientWidth     =   9165
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
   ForeColor       =   &H80000002&
   Icon            =   "F_Employers.frx":0000
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5655
   ScaleWidth      =   9165
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ListView lb 
      Height          =   4095
      Left            =   0
      TabIndex        =   0
      Tag             =   "free,font"
      Top             =   1560
      Width           =   9120
      _ExtentX        =   16087
      _ExtentY        =   7223
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
         Text            =   "Employer name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "PAYE reference"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Filename"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "No of employees"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   1800
      Left            =   0
      ScaleHeight     =   1800
      ScaleWidth      =   10995
      TabIndex        =   1
      Tag             =   "free,font"
      Top             =   0
      Width           =   10995
      Begin VB.Label lblEmployersDirectory 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Employer files in this directory:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00996633&
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Tag             =   "FREE,FONT"
         Top             =   1350
         Width           =   8490
      End
      Begin VB.Label lblYear 
         BackStyle       =   0  'Transparent
         Caption         =   "97"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   21.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00996633&
         Height          =   495
         Left            =   5640
         TabIndex        =   4
         Tag             =   "FREE,FONT"
         Top             =   240
         Width           =   1605
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "P11D"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   39.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00996633&
         Height          =   855
         Left            =   3000
         TabIndex        =   3
         Tag             =   "FREE,FONT"
         Top             =   0
         Width           =   2655
      End
      Begin VB.Label lblTitle 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "abatec, Deloitte && Touche, (020) 7438 3491"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00996633&
         Height          =   210
         Left            =   2400
         TabIndex        =   2
         Tag             =   "FREE,FONT"
         Top             =   840
         Width           =   4140
      End
   End
End
Attribute VB_Name = "F_Employers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Implements IBenefitForm2
Implements IFrmGeneral

Private mclsResize As New clsFormResize

Private Const L_DES_HEIGHT  As Long = 6165
Private Const L_DES_WIDTH  As Long = 9285
Private benefit As IBenefitClass


Private Sub IBenefitForm2_AddBenefit()
   
  On Error GoTo AddBenefit_Err
  
  Call xSet("AddBenefit")
  
  Call p11d32.EditEmployer(-1)
  
AddBenefit_End:
  Call xReturn("AddBenefit")
  Exit Sub
AddBenefit_Err:
  Call ErrorMessage(ERR_ERROR, Err, "AddBenefit", "Add Benefit", "Error adding an employer.")
  Resume AddBenefit_Err
  Resume
End Sub

Private Function IBenefitForm2_AddBenefitSetDefaults(ben As IBenefitClass) As Long
  'nothing to do as no defaults for an employer
End Function

Private Property Let IBenefitForm2_benclass(ByVal RHS As BEN_CLASS)
  'no set alwats BC_EMPLOYER
End Property

Private Property Get IBenefitForm2_benclass() As BEN_CLASS
  IBenefitForm2_benclass = BC_EMPLOYER
End Property

Private Property Set IBenefitForm2_benefit(NewValue As IBenefitClass)
  Set benefit = NewValue
End Property

Private Property Get IBenefitForm2_benefit() As IBenefitClass
  Set IBenefitForm2_benefit = benefit
End Property

Private Function IBenefitForm2_BenefitFormState(ByVal fState As BENEFIT_FORM_STATE) As Boolean
  IBenefitForm2_BenefitFormState = BenefitFormStateEx(fState, benefit)
  
End Function

Private Function IBenefitForm2_BenefitOff() As Boolean
  Call DisplayEx(D_EMPLOYER_OFF)
  IBenefitForm2_BenefitOff = True
End Function

Private Function IBenefitForm2_BenefitOn() As Boolean
  Call DisplayEx(D_EMPLOYER_ON)
  IBenefitForm2_BenefitOn = True
End Function

Private Function IBenefitForm2_BenefitsToListView() As Long
  Dim i As Long
  Dim ben As IBenefitClass
  Dim lst As ListItem
  Dim ibf As IBenefitForm2
    
  On Error GoTo BenefitsToListView_err
  Call xSet("BenefitsToListView")
  
  Set ibf = Me
  
  Call ClearForm(ibf)
  Call MDIMain.SetAdd
  
  ibf.lv.Sorted = False
  
  For i = 1 To p11d32.Employers.Count
    Set ben = p11d32.Employers(i)
    IBenefitForm2_BenefitsToListView = IBenefitForm2_BenefitsToListView + ibf.BenefitToListView(ben, i)
  Next
  
  ibf.lv.Sorted = True

BenefitsToListView_end:
  Set ben = Nothing
  Set lst = Nothing
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



Private Function IBenefitForm2_UpdateBenefitListViewItem(li As MSComctlLib.IListItem, benefit As IBenefitClass, Optional ByVal BenefitIndex As Long = 0&, Optional ByVal SelectItem As Boolean = False) As Long
    
  On Error GoTo UpdateBenefitListViewItemEmployer_ERR
  
  Call xSet("UpdateBenefitListViewItemEmployer")
  
  If Not li Is Nothing And Not benefit Is Nothing Then
    If BenefitIndex > 0 Then li.Tag = BenefitIndex
    li.SmallIcon = benefit.ImageListKey
    li.Text = benefit.Name
    li.SubItems(1) = benefit.value(employer_Payeref_db)
    li.SubItems(2) = benefit.value(employer_FileName)
    li.SubItems(3) = benefit.value(employer_EmployeesCount)
    If p11d32.FixLevelsShow Then li.SubItems(4) = benefit.value(employer_FixLevel_db)
    IBenefitForm2_UpdateBenefitListViewItem = li.Index
  End If

UpdateBenefitListViewItemEmployer_END:
  Call xReturn("UpdateBenefitListViewItemEmployer")
  Exit Function
UpdateBenefitListViewItemEmployer_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "UpdateBenefitListViewItemEmployer", "Update Benefit List View Item Employer", "Error updating the benefit list view item for an employer.")
  Resume UpdateBenefitListViewItemEmployer_END
  
End Function
Private Function IBenefitForm2_BenefitToScreen(Optional ByVal BenefitIndex As Long = -1&, Optional ByVal UpdateBenefit As Boolean = True) As Boolean
  Dim ben As IBenefitClass
  Dim ibf As IBenefitForm2
  
  On Error GoTo BenefitToScreen_Err
  Call xSet("BenefitToScreen")
  
  Set ibf = Me
  If BenefitIndex <> -1 Then
    Set ben = p11d32.Employers(BenefitIndex)
    If Not ibf.ValididateBenefit(ben) Then Call Err.Raise(ERR_INVALIDBENCLASS, "BenefitToScreen", "Benefit To Screen Helper", "Invalid benefit type.")
    Set ibf.benefit = ben
    Call ibf.BenefitOn
  Else
    Set ibf.benefit = Nothing
    Call ibf.BenefitOff
  End If
  Call SetBenefitFormState(ibf, False)
  IBenefitForm2_BenefitToScreen = True
  
BenefitToScreen_End:
  Set ben = Nothing
  Call xReturn("BenefitToScreen")
  Exit Function
BenefitToScreen_Err:
  Call ErrorMessage(ERR_ERROR, Err, "BenefitToScreen", "Benefit To Screen", "Unable to place an employer to the screen. Benefit index = " & BenefitIndex & ".")
  Resume BenefitToScreen_End
  Resume

  
End Function

Private Property Get IBenefitForm2_lv() As MSComctlLib.IListView
  Set IBenefitForm2_lv = lb
End Property

Private Function IBenefitForm2_RemoveBenefit(ByVal BenefitIndex As Long) As Boolean
  
  Dim lst As ListItem
  Dim ben As IBenefitClass
  Dim NextBenefitIndex As Long
  Dim ibf As IBenefitForm2
  
  
  On Error GoTo RemoveBenefit_ERR
  Call xSet("RemoveBenefit")
  
  Set ibf = Me
  NextBenefitIndex = GetNextBestListItemBenefitIndex(ibf, BenefitIndex)
  
  Set ben = p11d32.Employers(BenefitIndex)
  
  ben.Kill
  Call p11d32.Employers.Remove(BenefitIndex)
  Call BackupEmployer(ben)
  
  Call ibf.BenefitsToListView
  'select an item
  Call SelectBenefitByBenefitIndex(ibf, NextBenefitIndex)
  IBenefitForm2_RemoveBenefit = True
  
RemoveBenefit_END:
  Call xReturn("RemoveBenefit")
  Exit Function

RemoveBenefit_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "RemoveBenefit", "", "Error while removing an employer.")
  Resume RemoveBenefit_END
  Resume
End Function


Private Function IBenefitForm2_ValididateBenefit(ben As IBenefitClass) As Boolean
  If ben.BenefitClass = BC_EMPLOYER Then IBenefitForm2_ValididateBenefit = True
End Function
Private Sub Form_Load()
  If Not (mclsResize.InitResize(Me, L_DES_HEIGHT, L_DES_WIDTH, DESIGN, , , MDIMain)) Then
    Err.Raise ERR_Application
  End If
    lblYear = p11d32.AppYear
    
    lblTitle.Caption = "abatec, Deloitte, " & S_TELEPHONE
    
End Sub
Private Sub Form_Resize()
  mclsResize.Resize
  Call FixLevelShowFunction(LEF_SIZE_COLUMNS, p11d32.FixLevelsShow)
End Sub
'Private Function GetContactTelephone() As String
'  'Parse generic contact string for telephone number RK hack for telephone number change
'  Dim iStartPoint As Integer, sContact As String
'  sContact = p11d32.Contact
'  iStartPoint = InStr(1, sContact, "(020)", vbTextCompare)
'  GetContactTelephone = Mid$(sContact, iStartPoint)
'End Function
Private Function IFrmGeneral_CheckChanged(c As Control) As Boolean

End Function

Private Property Get IFrmGeneral_InvalidVT() As Control

End Property

Private Property Set IFrmGeneral_InvalidVT(RHS As Control)

End Property

Private Sub L_Title_Click()

End Sub

Private Sub lb_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
  Call SetSortOrder(lb, ColumnHeader)
End Sub
Private Sub lb_DblClick()
  If Not lb.SelectedItem Is Nothing Then Call ToolBarButton(TBR_EMPLOYEESCREEN, lb.SelectedItem.Tag)
    
End Sub
Private Sub lb_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
  Dim li As ListItem
  
  If Button And vbRightButton Then
    Set li = lb.HitTest(x, y)
    If Not li Is Nothing Then
      Call p11d32.EditEmployer(li.Tag)
    End If
  End If
End Sub

Private Sub Picture1_GotFocus()
  SendKeys (vbTab)
End Sub

Private Sub lb_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then 'Return key
    Call lb_DblClick
  End If
End Sub


