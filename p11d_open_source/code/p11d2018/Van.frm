VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{8D988532-0F0C-460C-B00E-7B5637E97680}#1.0#0"; "atc2vtext.ocx"
Begin VB.Form F_SharedVans 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Shared vans"
   ClientHeight    =   6795
   ClientLeft      =   1860
   ClientTop       =   2250
   ClientWidth     =   5310
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6795
   ScaleWidth      =   5310
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.Toolbar tbar 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   5310
      _ExtentX        =   9366
      _ExtentY        =   741
      ButtonWidth     =   635
      ButtonHeight    =   582
      Appearance      =   1
      _Version        =   393216
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Default         =   -1  'True
      Height          =   375
      Left            =   4095
      TabIndex        =   6
      Top             =   6240
      Width           =   1140
   End
   Begin VB.Frame fmeInput 
      Height          =   3105
      Left            =   45
      TabIndex        =   8
      Top             =   3060
      Width           =   5190
      Begin VB.CheckBox chkbx 
         Alignment       =   1  'Right Justify
         Caption         =   "Is the van electric?"
         ForeColor       =   &H00800000&
         Height          =   350
         Index           =   2
         Left            =   90
         TabIndex        =   16
         Tag             =   "free,font"
         Top             =   2640
         Width           =   5025
      End
      Begin VB.CheckBox chkbx 
         Alignment       =   1  'Right Justify
         Caption         =   "Was fuel available?"
         ForeColor       =   &H00800000&
         Height          =   350
         Index           =   1
         Left            =   90
         TabIndex        =   15
         Tag             =   "free,font"
         Top             =   2280
         Width           =   5025
      End
      Begin VB.CheckBox chkbx 
         Alignment       =   1  'Right Justify
         Caption         =   "First registered on or after 6/4/94?"
         DataField       =   "RegAfter"
         DataSource      =   "DB"
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Tag             =   "FREE,FONT"
         Top             =   585
         Width           =   4980
      End
      Begin atc2valtext.ValText TxtBx 
         Height          =   315
         Index           =   3
         Left            =   4005
         TabIndex        =   5
         Tag             =   "FREE,FONT"
         Top             =   1935
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
         MaxLength       =   10
         Text            =   "0"
         Maximum         =   "365"
         Minimum         =   "0"
         AllowEmpty      =   0   'False
         TXTAlign        =   2
      End
      Begin atc2valtext.ValText TxtBx 
         Height          =   315
         Index           =   1
         Left            =   4005
         TabIndex        =   3
         Tag             =   "FREE,FONT"
         Top             =   1075
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
         Text            =   ""
         TypeOfData      =   2
         Maximum         =   "5/4/99"
         Minimum         =   "6/4/98"
         AllowEmpty      =   0   'False
      End
      Begin atc2valtext.ValText TxtBx 
         Height          =   315
         Index           =   0
         Left            =   2070
         TabIndex        =   0
         Tag             =   "FREE,FONT"
         Top             =   180
         Width           =   3060
         _ExtentX        =   5398
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
      Begin atc2valtext.ValText TxtBx 
         Height          =   315
         Index           =   4
         Left            =   4005
         TabIndex        =   2
         Tag             =   "FREE,FONT"
         Top             =   650
         Width           =   1125
         _ExtentX        =   1984
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
         Text            =   ""
         TypeOfData      =   2
         AllowEmpty      =   0   'False
      End
      Begin atc2valtext.ValText TxtBx 
         Height          =   315
         Index           =   2
         Left            =   4005
         TabIndex        =   4
         Tag             =   "FREE,FONT"
         Top             =   1500
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
         Text            =   ""
         TypeOfData      =   2
         Maximum         =   "5/4/99"
         Minimum         =   "6/4/98"
         AllowEmpty      =   0   'False
      End
      Begin VB.Label Lab 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Available to"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   14
         Tag             =   "FREE,FONT"
         Top             =   1560
         Width           =   2535
      End
      Begin VB.Label Lab 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date of Registration"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   6
         Left            =   135
         TabIndex        =   13
         Tag             =   "FREE,FONT"
         Top             =   630
         Width           =   1410
      End
      Begin VB.Label Lab 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Van reference"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   12
         Tag             =   "FREE,FONT"
         Top             =   225
         Width           =   1005
      End
      Begin VB.Label Lab 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Days unavailable"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   5
         Left            =   120
         TabIndex        =   11
         Tag             =   "FREE,FONT"
         Top             =   1980
         Width           =   1215
      End
      Begin VB.Label Lab 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Available from"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   10
         Tag             =   "FREE,FONT"
         Top             =   1135
         Width           =   990
      End
   End
   Begin MSComctlLib.ListView LB 
      Height          =   2655
      Left            =   45
      TabIndex        =   7
      Tag             =   "free,font"
      Top             =   405
      Width           =   5205
      _ExtentX        =   9181
      _ExtentY        =   4683
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
         Text            =   "Van Reference"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   1
         Text            =   "Value"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "F_SharedVans"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Implements IBenefitForm2
Implements IFrmGeneral

Public benefit As IBenefitClass
Private m_BenClass As BEN_CLASS
Private m_bDirty As Boolean

Private m_InvalidVT As Control


Private Sub ChkBx_Click(Index As Integer)
  Call IFrmGeneral_CheckChanged(chkbx(Index))
End Sub

Private Sub ChkBx_KeyPress(Index As Integer, KeyAscii As Integer)
  'Check for return key - tab to next field
  If KeyAscii = 13 Then Call SendKeys(vbTab)
End Sub
Private Sub cmdClose_Click()
 
  
 On Error GoTo cmdClose_ERR
    
     
    
    Call xSet("cmdClose")
    Unload Me
  
cmdClose_END:
  ' Unload Me
  Call xReturn("cmdClose")
  Exit Sub
cmdClose_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "cmdClose", "cmd Close", "Error writing or finishing the shared van information.")
  Resume cmdClose_END
End Sub


Private Sub Form_Load()
  Call AddAddDelete(tbar)
  Call SetDefaultVTDate(TxtBx(1))
  Call SetDefaultVTDate(TxtBx(2))
  
  chkbx(0).Visible = False
End Sub


Private Sub Form_Terminate()

  ' EK Call MDIMain.ClearAdd

End Sub


Private Sub IBenefitForm2_AddBenefit()
  Dim ben As IBenefitClass
  Dim lst As ListItem, i As Long
  Dim ibf As IBenefitForm2
  
  On Error GoTo AddBenefit_Err
  Call xSet("AddBenefit")
  
  Set ben = New SharedVan
  'Put in defaults for benefit
  Set ibf = Me
  Call ibf.AddBenefitSetDefaults(ben)
  With ben
    Set .Parent = p11d32.CurrentEmployer.SharedVans
    .ReadFromDB = True
    i = p11d32.CurrentEmployer.SharedVans.Vans.Add(ben)
    Set lst = LB.listitems.Add(, , .Name)
    Call ibf.UpdateBenefitListViewItem(lst, ben, i, True)
    .Dirty = True
    Call ibf.BenefitToScreen(i)
    Call MDIMain.SetDelete
    
  End With
  
AddBenefit_End:
  Set ibf = Nothing
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
    .value(shvan_item_db) = "Please enter description..."
    .value(shvan_AvailableFrom_db) = p11d32.Rates.value(TaxYearStart)
    .value(shvan_Availableto_db) = p11d32.Rates.value(TaxYearEnd)
    .value(shvan_DaysUnavailable_db) = 0
    .value(shvan_fuel_available_db) = True
    .value(shvan_RegistrationDate_db) = p11d32.Rates.value(VanRegDateNew)
    .value(shvan_is_electric_db) = False 'cad 2010
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
      fmeInput.Enabled = True
      Call MDIMain.SetDelete
    Else
      fmeInput.Enabled = False
      Call MDIMain.ClearDelete
    End If
    Call SetLVEnabled(LB, True)
  ElseIf fState = FORM_DISABLED Then
    Set benefit = Nothing
    fmeInput.Enabled = False
    Call SetLVEnabled(LB, False)
    Call MDIMain.ClearDelete
    Call MDIMain.ClearConfirmUndo
  End If
  IBenefitForm2_BenefitFormState = True
    
BenefitFormState_end:
  Call xReturn("BenefitFormState")
  Exit Function
  
BenefitFormState_err:
  IBenefitForm2_BenefitFormState = False
  Call ErrorMessage(ERR_ERROR, Err, "BenefitFormState", "Benefit Form State", "Undefined error.")
  Resume BenefitFormState_end
  Resume
End Function

Private Function IBenefitForm2_BenefitOff() As Boolean
  TxtBx(0).text = ""
  TxtBx(1).text = ""
  TxtBx(2).text = ""
  TxtBx(3).text = ""
'  If p11d32.AppYear > 2000 Then
  TxtBx(4).text = "" 'km
'  If p11d32.AppYear = 2000 Then ChkBx(0) = vbUnchecked
  tbar.Buttons(2).Enabled = False 'cross
End Function

Private Function IBenefitForm2_BenefitOn() As Boolean
  Dim X As Date
  
  With benefit
    TxtBx(0).text = .value(shvan_item_db)
    TxtBx(1).text = DateValReadToScreen(.value(shvan_AvailableFrom_db))
    TxtBx(2).text = DateValReadToScreen(.value(shvan_Availableto_db))
    TxtBx(3).text = .value(shvan_DaysUnavailable_db)
    TxtBx(4).text = IIf(.value(shvan_RegistrationDate_db) = UNDATED, "", DateValReadToScreen(.value(shvan_RegistrationDate_db)))
    chkbx(1).value = BoolToChkBox(.value(shvan_fuel_available_db))
    chkbx(2).value = BoolToChkBox(.value(shvan_is_electric_db)) 'CAD 2010
    tbar.Buttons(2).Enabled = True 'minus button
  End With
End Function

Private Function IBenefitForm2_BenefitsToListView() As Long
  Dim i As Long
  Dim ben As IBenefitClass
  Dim lst As ListItem
  Dim bc As BEN_CLASS
  Dim ibf As IBenefitForm2
  
  On Error GoTo BenefitsToListView_err
  Call xSet("BenefitsToListView")
  
  Set ibf = Me
  
  Call ClearForm(ibf)
  Call MDIMain.SetAdd
  
  bc = ibf.benclass
  
  For i = 1 To p11d32.CurrentEmployer.SharedVans.Vans.Count
    Set ben = p11d32.CurrentEmployer.SharedVans.Vans(i)
    If Not ben Is Nothing Then
      IBenefitForm2_BenefitsToListView = IBenefitForm2_BenefitsToListView + ibf.BenefitToListView(ben, i)
    End If
  Next
  
  'Call SetBenefitFormState(Me)
  
BenefitsToListView_end:
  Set ibf = Nothing
  Set ben = Nothing
  Set lst = Nothing
  Call xReturn("BenefitsToListView")
  Exit Function
  
BenefitsToListView_err:
  Call ErrorMessage(ERR_ERROR, Err, "BenefitsToListView", "Adding Benefits to ListView", "Error in adding benefits to listview.")
  Resume BenefitsToListView_end
End Function

Private Function IBenefitForm2_BenefitToListView(ben As IBenefitClass, ByVal lBenefitIndex As Long) As Long
  IBenefitForm2_BenefitToListView = BenefitToListView(ben, Me, lBenefitIndex)
End Function

Private Function IBenefitForm2_BenefitToScreen(Optional ByVal BenefitIndex As Long = -1&, Optional ByVal UpdateBenefit As Boolean = True) As Boolean
  Dim ben As IBenefitClass
  Dim ibf As IBenefitForm2
  
  On Error GoTo BenefitToScreenHelper_Err
  Call xSet("BenefitToScreenHelper")
  
  Set ibf = Me
  
  If UpdateBenefit Then
    Call UpdateBenefitFromTags
  End If
  
  If BenefitIndex <> -1 Then
    Set ben = p11d32.CurrentEmployer.SharedVans.Vans(BenefitIndex)
    If Not ben Is Nothing Then
      If ben.BenefitClass <> ibf.benclass Then Call Err.Raise(ERR_INVALIDBENCLASS, "BenefitToScreen", "Benefit type invalid")
      Set ibf.benefit = ben
      Call ibf.BenefitOn
    End If
  Else
    Set ibf.benefit = Nothing
    Call ibf.BenefitOff
  End If
  Call SetBenefitFormState(ibf)
  IBenefitForm2_BenefitToScreen = True
  
BenefitToScreenHelper_End:
  Set ibf = Nothing
  Set ben = Nothing
  Call xReturn("BenefitToScreenHelper")
  Exit Function

BenefitToScreenHelper_Err:
  Call ErrorMessage(ERR_ERROR, Err, "BenefitToScreenHelper", "Benefit To Screen Helper", "Unable to place the chosen benefit to the screen. Benefit index = " & BenefitIndex & ".")
  Resume BenefitToScreenHelper_End
  Resume

End Function

Private Property Let IBenefitForm2_benclass(ByVal NewValue As BEN_CLASS)
  m_BenClass = NewValue
End Property

Private Property Get IBenefitForm2_benclass() As BEN_CLASS
  IBenefitForm2_benclass = m_BenClass
End Property

'Private Property Get IBenefitForm2_ControlDefault() As Control
'  Set IBenefitForm2_ControlDefault = TxtBx(0)
'End Property

Private Property Get IBenefitForm2_lv() As MSComctlLib.IListView
  Set IBenefitForm2_lv = LB
End Property

Private Function IBenefitForm2_RemoveBenefit(ByVal BenefitIndex As Long) As Boolean
  Dim NextBenefitIndex As Long
  Dim ibf As IBenefitForm2
  
  On Error GoTo RemoveBenefit_ERR
  
  Call xSet("RemoveBenefit")
  
  If Not benefit.CompanyDefined Then
    Call benefit.DeleteDB
    NextBenefitIndex = GetNextBestListItemBenefitIndex(ibf, BenefitIndex)
    Call p11d32.CurrentEmployer.SharedVans.Vans.Remove(BenefitIndex)
    're-add the benefits to the list view
    Set ibf = Me
    Call ibf.BenefitsToListView
    'select an item
    Call SelectBenefitByBenefitIndex(ibf, NextBenefitIndex)
    IBenefitForm2_RemoveBenefit = True
  End If
    
RemoveBenefit_END:
  Set ibf = Nothing
  Call xReturn("RemoveBenefit")
  Exit Function
  
RemoveBenefit_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "RemoveBenefit", "Remove Benefit", "Error removing benefit.")
  Resume RemoveBenefit_END
  Resume
  
End Function

Private Function IBenefitForm2_UpdateBenefitListViewItem(li As MSComctlLib.IListItem, benefit As IBenefitClass, Optional ByVal BenefitIndex As Long = 0&, Optional ByVal SelectItem As Boolean = False) As Long
  IBenefitForm2_UpdateBenefitListViewItem = UpdateBenefitListViewItem(li, benefit, BenefitIndex, SelectItem)
End Function

Private Function IBenefitForm2_ValididateBenefit(ben As IBenefitClass) As Boolean
  If m_BenClass = BC_SHAREDVAN_G Then IBenefitForm2_ValididateBenefit = True
End Function

Private Function IFrmGeneral_CheckChanged(c As Control) As Boolean
  Dim bDirty As Boolean
  Dim ben As IBenefitClass
  
On Error GoTo CheckChanged_Err

  Call xSet("CheckChanged")
  
  With c
    If benefit Is Nothing Then
      GoTo CheckChanged_End
    End If
    
    Select Case LCase(.Name)
      Case "txtbx"
        Select Case .Index
          Case 0
            bDirty = CheckTextInput(.text, benefit, shvan_item_db)
          Case 1
            bDirty = CheckTextInput(.text, benefit, shvan_AvailableFrom_db)
          Case 2
            bDirty = CheckTextInput(.text, benefit, shvan_Availableto_db)
          Case 3
            bDirty = CheckTextInput(.text, benefit, shvan_DaysUnavailable_db)
          Case 4
'            If p11d32.AppYear > 2000 Then 'km
              bDirty = CheckTextInput(.text, benefit, shvan_RegistrationDate_db)
'            End If
          Case Else
            ECASE "Unknown control index"
        End Select
      Case "chkbx"
        Select Case .Index
          Case 1
            bDirty = CheckCheckBoxInput(chkbx(1).value, benefit, shvan_fuel_available_db)
          Case 2
            bDirty = CheckCheckBoxInput(chkbx(2).value, benefit, shvan_is_electric_db)
          Case Else
            ECASE "Unknown control index"
        End Select
      Case Else
        ECASE "Unknown control"
    End Select
    
    'must be required in all check changed
    IFrmGeneral_CheckChanged = AfterCheckChanged(c, Me, bDirty)
    MDIMain.ClearConfirmUndo
    Set ben = benefit.Parent
    ben.Dirty = ben.Dirty Or bDirty
    Call CheckValidity(Me, , False)
    
  End With
  
CheckChanged_End:
  Set ben = Nothing
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

  'Dim bValid As Boolean
  'bValid = van_valid()'
 
   ' If bValid = True Then
      Call SetLastListItemSelected(Item)
      If Not (LB.SelectedItem Is Nothing) Then Call IBenefitForm2_BenefitToScreen(Item.Tag)
    'Else
     ' Call MsgBox("Invalid data, all entries must be valid.  Check data")
      ' EK Need to add line to get back to previously selected lb list item.
      
   ' End If
  
  
  
End Sub

Private Sub lb_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
  Call SetSortOrder(LB, ColumnHeader)
End Sub

Private Sub lb_KeyPress(KeyAscii As Integer)
  'Check for return key - tab to next field
  If KeyAscii = 13 Then Call SendKeys(vbTab)
End Sub

Private Sub tbar_ButtonClick(ByVal Button As MSComctlLib.Button)
  Call AddDeleteClick(Button.Index, Me)
End Sub

Private Sub TxtBx_FieldInvalid(Index As Integer, Valid As Boolean, Message As String)
  Call SetPanel2(Message)
End Sub

Private Sub TxtBx_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  TxtBx(Index).Tag = SetChanged(False)
End Sub

Private Sub TxtBx_Validate(Index As Integer, Cancel As Boolean)
 Call IFrmGeneral_CheckChanged(TxtBx(Index))
End Sub
