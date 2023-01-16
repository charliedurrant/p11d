VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.2#0"; "COMCTL32.OCX"
Object = "{E297AE83-F913-4A8C-873C-EDEAC00CB9AC}#2.1#0"; "atc3ubgrd.ocx"
Begin VB.Form F_ErFPCS 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Company-Defined rates"
   ClientHeight    =   4050
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6525
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4050
   ScaleWidth      =   6525
   StartUpPosition =   3  'Windows Default
   Begin ComctlLib.Toolbar tbar 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   6525
      _ExtentX        =   11509
      _ExtentY        =   741
      Appearance      =   1
      _Version        =   327682
   End
   Begin atc3ubgrd.UBGRD ubgrd 
      Height          =   2940
      Left            =   1755
      TabIndex        =   4
      Top             =   495
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   5186
   End
   Begin VB.TextBox txtFPCS 
      Height          =   375
      Left            =   45
      TabIndex        =   3
      Text            =   "txtFPCS"
      Top             =   3465
      Width           =   2175
   End
   Begin ComctlLib.ListView LB 
      Height          =   2940
      Left            =   45
      TabIndex        =   1
      ToolTipText     =   "Car schemes"
      Top             =   495
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   5186
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Schemes"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   375
      Left            =   5220
      TabIndex        =   0
      Top             =   3645
      Width           =   1245
   End
End
Attribute VB_Name = "F_ErFPCS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Implements IBenefitForm2
Implements IFrmGeneral

Private m_FPCS As FPCS
Private m_InvalidVt As atc2valtext.ValText

Private Sub B_Cancel_Click()
 Call F_ErFPCS.Hide
End Sub

Private Sub B_OK_Click()
  
End Sub
Private Function ValidateFPCSName(ByVal sFPCSName As String)
  Dim CS As FPCS
  Dim l As Long
  
On Error GoTo ValidateFPCSName_Err

  Call xSet("ValidateFPCSName")
  
  For l = 1 To rates.FPCShemes.count
    Set CS = rates.FPCShemes(l)
    If Not CS Is Nothing Then
      If StrComp(CS.Name, sFPCSName) = 0 Then
        ValidateFPCSName = False
        GoTo ValidateFPCSName_End
      End If
    End If
  Next
  
  ValidateFPCSName = True
  
ValidateFPCSName_End:
  Set CS = Nothing
  Call xReturn("ValidateFPCSName")
  Exit Function

ValidateFPCSName_Err:
  Call ErrorMessage(ERR_ERROR, Err, "ValidateFPCSName", "ERR_UNDEFINED", "Undefined error.")
  Resume ValidateFPCSName_End
  Resume
End Function
Private Function AddNewFPCS(sNewFPCS As String) As Boolean
  Dim l As Long
  Dim CS As FPCS
  Dim li As ListItem
  Dim ibf As IBenefitForm2
  
  On Error GoTo AddNewFPCS_Err
  Call xSet("AddNewFPCS")
  
  If ValidateFPCSName(sNewFPCS) = False Then GoTo AddNewFPCS_End
  
  Set CS = New FPCS
  CS.Dirty = True
  CS.Name = sNewFPCS
  CS.ObjectListIndex = rates.FPCShemes.Add(CS)
  Set li = LB.ListItems.Add(, , sNewFPCS)
  li.Tag = CS.ObjectListIndex
  Set ibf = Me
  Set ibf.lv.SelectedItem = li
  ibf.BenefitToScreen (li.Tag)
  AddNewFPCS = True
  
AddNewFPCS_End:
  Set li = Nothing
  Set CS = Nothing
  Call xReturn("AddNewFPCS")
  Exit Function

AddNewFPCS_Err:
  Call ErrorMessage(ERR_ERROR, Err, "AddNewFPCS", "ERR_UNDEFINED", "Undefined error.")
  Resume AddNewFPCS_End
End Function

Private Sub cmdClose_Click()
  CurrentEmployer.WriteFPCS
  Unload Me
End Sub

Private Sub Form_Load()
  Dim b As Button
  Dim c As TrueDBGrid50.Column
  Dim grd As TrueDBGrid50.TDBGrid
  
  'set up toolbar
  Set tbar.ImageList = MDIMain.ImgToolbar(0)
  Set b = tbar.Buttons.Add(1, , , , 19)
  b.ToolTipText = "Add car scheme"
  Set b = tbar.Buttons.Add(2, , , , 18)
  b.ToolTipText = "Delete scheme"
  
  Set grd = ubgrd.Grid
  
  Call AddUBGRDStandardColumn(grd, 0, 1244.976, "Band Name", "")
  Call AddUBGRDStandardColumn(grd, 1, 1019.906, "Miles above", "")
  Call AddUBGRDStandardColumn(grd, 2, 945.0709, "CC above", "")
  Call AddUBGRDStandardColumn(grd, 3, 1244.976, "Band Name", "")
  
  grd.AllowUpdate = True
End Sub


'Private Sub grd_UnboundAddData(ByVal RowBuf As TrueDBGrid50.RowBuffer, NewRowBookmark As Variant)
'  Dim Band As FPCSBand
'
'  If ValidateBand(RowBuf) Then
'    Set Band = New FPCSBand
'    Call AddNewRow(Band, ByVal RowBuf)
'    Band.ObjectListIndex = m_FPCS.Bands.Add(Band)
'    NewRowBookmark = Band.ObjectListIndex
'  End If
'
'  Set Band = Nothing
'
'End Sub

'Private Sub grd_UnboundDeleteRow(Bookmark As Variant)
'  If Not m_FPCS Is Nothing Then
'    m_FPCS.Bands.Remove (Bookmark)
'  End If
'End Sub

'Private Sub grd_UnboundReadDataEx(ByVal RowBuf As TrueDBGrid50.RowBuffer, StartLocation As Variant, ByVal Offset As Long, ApproximatePosition As Long)
'  If Not m_FPCS Is Nothing Then
'    Call mUBData.UnboundObjectListRead(m_FPCS.Bands, RowBuf, StartLocation, Offset, ApproximatePosition)
'  End If
'End Sub

Private Function IBenefitForm2_UpdateBenefitListViewItem(li As ComctlLib.IListItem, benefit As IBenefitClass, Optional ByVal BenefitIndex As Long = 0&, Optional ByVal SelectItem As Boolean = False) As Boolean
  ECASE "UpdateBenefitListViewItem"
End Function

Private Function IFrmGeneral_CheckChanged(c As Control, ByVal UpdateCurrentListItem As Boolean) As Boolean
  
  Dim ibf As IBenefitForm2
  
  On Error GoTo CheckChanged_Err
  Call xSet("CheckChanged")
  
  With c
    If m_FPCS Is Nothing Then
      GoTo CheckChanged_End
    End If
    If LB.SelectedItem Is Nothing Then
      GoTo CheckChanged_End
    End If
    'we are asking if the value has changed and if it is valid thus save
    
    Select Case .Name
      Case "txtFPCS"
        If StrComp(LB.SelectedItem.Text, .Text) Then
         
          If ValidateFPCSName(txtFPCS) = False Then
            Call ErrorMessage(ERR_INFO, Err, "txtFPCS_KeyDown", "txtFPCS KeyDown", "The fixed profit car scheme name you chosen is already in use. Press escape to cancel change.")
            txtFPCS.SetFocus
            txtFPCS.SelLength = Len(txtFPCS)
          Else
            m_FPCS.Name = txtFPCS
            LB.SelectedItem.Text = txtFPCS
          End If
          
          
        End If
      Case Else
         ECASE "Unknown"
     End Select
    
    'must be required in all check changed
    Set ibf = Me
    IFrmGeneral_CheckChanged = True 'ibf.UpdateBenefitListViewItem(ibf.lv.SelectedItem, ibf.benefit)

  End With
  
CheckChanged_End:
  Set ibf = Nothing
  Call xReturn("CheckChanged")
  Exit Function
  
CheckChanged_Err:
  IFrmGeneral_CheckChanged = False
  Call ErrorMessage(ERR_ERROR, Err, "CheckChanged", "ERR_CHECKCHANGED", "This function has failed for the form " & Me.Name & ".")
  Resume CheckChanged_End

End Function

Private Property Get IFrmGeneral_InvalidVT() As atc2valtext.ValText
  IFrmGeneral_InvalidVT = m_InvalidVt
End Property

Private Property Set IFrmGeneral_InvalidVT(NewValue As atc2valtext.ValText)
  Set m_InvalidVt = NewValue
  
End Property

Private Sub LB_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Call TestChangedControls(Me, True)
End Sub

Private Sub txtFPCS_KeyDown(KeyCode As Integer, Shift As Integer)
  txtFPCS.Tag = SetChanged(False)
  If KeyCode = vbKeyEscape Then
    txtFPCS = m_FPCS.Name
    txtFPCS.SelLength = Len(txtFPCS)
  End If
End Sub

Private Sub txtFPCS_LostFocus()
  Call IFrmGeneral_CheckChanged(txtFPCS, True)
End Sub

Private Sub IBenefitForm2_AddBenefit()
  Dim sNewScheme As String
  Dim sPrompt As String
  
  
  sPrompt = "Please enter a name for a company car scheme."
RETRY:
  sNewScheme = InputBox(sPrompt, "New Scheme", "")
  
  If Len(sNewScheme) Then
    'try and add then scheme
    If Not AddNewFPCS(sNewScheme) Then
      sPrompt = "The name you chose is already present, please try again."
      GoTo RETRY
    Else
      
    End If
  End If
  
  
  
End Sub

Private Property Let IBenefitForm2_benefit(RHS As IBenefitClass)

End Property

Private Property Get IBenefitForm2_benefit() As IBenefitClass
  ECASE "IBenefitForm2_benefit"
End Property

Private Function IBenefitForm2_BenefitFormState(ByVal fState As BENEFIT_FORM_STATE) As Boolean
  'do nothing here as is first selected then disble cross / done in benefit to screen
  
End Function

Private Function IBenefitForm2_BenefitsToListView() As Long
  Dim CS As FPCS
  Dim l As Long
  Dim li As ListItem
  
  
  LB.ListItems.Clear
  For l = 1 To rates.FPCShemes.count
    Set CS = rates.FPCShemes(l)
    If Not CS Is Nothing Then
      Set li = LB.ListItems.Add(, , CS.Name)
      li.Tag = CS.ObjectListIndex
    End If
  Next
  
  Set CS = Nothing
  Set li = Nothing
End Function

Private Function IBenefitForm2_BenefitToScreen(Optional ByVal BenefitIndex As Long = -1&, Optional ByVal UpdateBenefit As Boolean = True) As IBenefitClass
  Dim CS As FPCS
  Dim Band As FPCSBand
  Dim l As Long
  Dim r As RowBuffer
  Dim lCount As Long
  
  If BenefitIndex <> -1 Then
    Set m_FPCS = rates.FPCShemes(BenefitIndex)
    If Not m_FPCS Is Nothing Then
      ubgrd.ObjectList = m_FPCS.Bands
      If BenefitIndex = 1 Then
        txtFPCS.Visible = False
        ubgrd.Grid.Enabled = False
        tbar.Buttons(2).Enabled = False
      Else
        txtFPCS.Visible = True
        ubgrd.Grid.Enabled = True
        tbar.Buttons(2).Enabled = True
      End If
      Call ubgrd.Grid.ReBind 'paint me / fills the grid
      txtFPCS = m_FPCS.Name
    Else
      'err.raise ZZZZZ
    End If
  Else
    'err.raise zzzz
  End If
  
  
End Function

Private Property Let IBenefitForm2_bentype(ByVal RHS As benClass)

End Property

Private Property Get IBenefitForm2_bentype() As benClass

End Property

Private Property Get IBenefitForm2_lv() As ComctlLib.IListView
  Set IBenefitForm2_lv = LB
End Property

Private Function IBenefitForm2_RemoveBenefit(ByVal BenefitIndex As Long) As Boolean
  Dim NextBenefitIndex As Long
  Dim ibf As IBenefitForm2
  
On Error GoTo RemoveBenefit_ERR
  Call xSet("RemoveBenefit")
  
  NextBenefitIndex = GetNextBestListItemBenefitIndex(ibf, BenefitIndex)
  rates.FPCShemes.Remove (BenefitIndex)
  Set ibf = Me
  'Call ibf.BenefitsToListView
  ibf.lv.ListItems.Remove (ibf.lv.SelectedItem.Index)
  Call SelectBenefit(ibf, NextBenefitIndex)
  IBenefitForm2_RemoveBenefit = True
  
    
RemoveBenefit_END:
  Set ibf = Nothing
  Call xReturn("RemoveBenefit")
  Exit Function
  
RemoveBenefit_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "RemoveBenefit", "Remove Benefit", "Error removing benefit.")
  Resume RemoveBenefit_END
  Resume
End Function



Private Sub LB_ItemClick(ByVal Item As ComctlLib.ListItem)
  If Not Item Is Nothing Then
    Call IBenefitForm2_BenefitToScreen(Item.Tag)
  End If
End Sub

Private Sub tbar_ButtonClick(ByVal Button As ComctlLib.Button)
  Dim ibf As IBenefitForm2
  
  Set ibf = Me
  Select Case Button.Index
    Case 1 'tick
        Call ibf.AddBenefit
        
        
    Case 2 'cross
        Call ibf.RemoveBenefit(ibf.lv.SelectedItem.Tag)
  End Select
End Sub





Private Sub UBGRD_Click()

End Sub

Private Sub ubgrd_DeleteData(ObjectList As TCSCOREDLL.ObjectList, ObjectListIndex As Long)
  Call ObjectList.Remove(ObjectListIndex)
End Sub

Private Sub ubgrd_ReadData(RowBuf As TrueDBGrid50.RowBuffer, ByVal RowBufRowIndex As Long, ObjectList As TCSCOREDLL.ObjectList, ByVal ObjectListIndex As Long)
  Dim I As Long
  Dim Band As FPCSBand

  Set Band = ObjectList(ObjectListIndex)

  For I = 0 To (RowBuf.ColumnCount - 1)
    Select Case I
      Case 0
        RowBuf.value(RowBufRowIndex, I) = Band.Name
      Case 1
        RowBuf.value(RowBufRowIndex, I) = IIf(Band.GreaterThanMiles = DBL_GREATERTHANZERO, 0, Band.GreaterThanMiles)
      Case 2
        RowBuf.value(RowBufRowIndex, I) = IIf(Band.GreaterThanCC = DBL_GREATERTHANZERO, 0, Band.GreaterThanCC)
      Case 3
        RowBuf.value(RowBufRowIndex, I) = Band.Rate
      Case Else
        ECASE ("Invalid column if get user data")
    End Select
  Next I

  Set Band = Nothing

End Sub

Private Sub ubgrd_Validate(FirstColIndexInError As Long, ValidateMessage As String, ByVal RowBuf As TrueDBGrid50.RowBuffer, ByVal RowBufRowIndex As Long)
  Dim I As Long

  With RowBuf
    For I = 0 To RowBuf.ColumnCount - 1
      If Not IsNull(RowBuf.value(RowBufRowIndex, I)) Then 'has it changed / 0 for 1st row as there is only one
        Select Case I
          Case 0
            If Len(RowBuf.value(RowBufRowIndex, I)) = 0 Then
              ValidateMessage = "Zero length strings are no allowed."
              FirstColIndexInError = I
              Exit Sub
            End If
          Case 1, 2, 3
            If Not IsNumeric(RowBuf.value(RowBufRowIndex, I)) Then
              ValidateMessage = "The value must be numeric."
              FirstColIndexInError = I
              Exit Sub
            End If
        End Select
      End If
    Next
  End With
End Sub

Private Sub ubgrd_WriteData(ByVal RowBuf As TrueDBGrid50.RowBuffer, ByVal RowBufRowIndex As Long, ObjectList As TCSCOREDLL.ObjectList, ObjectListIndex As Long)
  Dim Band As FPCSBand
  
  If ObjectListIndex = -1 Then
    Set Band = New FPCSBand
    Band.ObjectListIndex = ObjectList.Add(Band)
    ObjectListIndex = Band.ObjectListIndex
  Else
    Set Band = ObjectList(ObjectListIndex)
  End If
  
  With Band
    If Not IsNull(RowBuf.value(RowBufRowIndex, 0)) Then .Name = RowBuf.value(RowBufRowIndex, 0)
    If Not IsNull(RowBuf.value(RowBufRowIndex, 1)) Then .GreaterThanMiles = IIf(RowBuf.value(RowBufRowIndex, 1) <= 0, DBL_GREATERTHANZERO, RowBuf.value(RowBufRowIndex, 1))
    If Not IsNull(RowBuf.value(RowBufRowIndex, 2)) Then .GreaterThanCC = IIf(RowBuf.value(RowBufRowIndex, 2) <= 0, DBL_GREATERTHANZERO, RowBuf.value(RowBufRowIndex, 2))
    If Not IsNull(RowBuf.value(RowBufRowIndex, 3)) Then .Rate = RowBuf.value(RowBufRowIndex, 3)
  End With
  
End Sub
