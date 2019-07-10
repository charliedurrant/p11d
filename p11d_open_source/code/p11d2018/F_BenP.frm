VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{412521B9-9CBB-4049-9E66-2AA0112EC306}#1.0#0"; "ATC2VTEXT.OCX"
Begin VB.Form F_P 
   Caption         =   "P - Other Expenses"
   ClientHeight    =   5655
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8190
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5655
   ScaleWidth      =   8190
   WindowState     =   2  'Maximized
   Begin VB.Frame P_NoBenefits 
      ForeColor       =   &H00FF0000&
      Height          =   5580
      Left            =   4020
      TabIndex        =   9
      Top             =   0
      Width           =   4125
      Begin VB.Frame fmeCDB 
         BorderStyle     =   0  'None
         Height          =   975
         Left            =   90
         TabIndex        =   18
         Top             =   2820
         Width           =   3975
         Begin VB.CommandButton BN_PushPull 
            Caption         =   "Copy"
            Enabled         =   0   'False
            Height          =   375
            Left            =   2550
            TabIndex        =   3
            Tag             =   "FREE,FONT"
            Top             =   540
            Width           =   1335
         End
         Begin VB.Label lblCDB 
            Caption         =   "Company defined benefit"
            ForeColor       =   &H80000008&
            Height          =   270
            Left            =   0
            TabIndex        =   20
            Tag             =   "FREE,FONT"
            Top             =   240
            Width           =   2655
         End
         Begin VB.Label PushPullText 
            Appearance      =   0  'Flat
            BackStyle       =   0  'Transparent
            Caption         =   "Copy the benefit to the individual"
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   0
            TabIndex        =   19
            Tag             =   "FREE,FONT"
            Top             =   540
            Width           =   2535
         End
      End
      Begin VB.Frame fmeInput 
         BorderStyle     =   0  'None
         Height          =   2625
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   3855
         Begin atc2valtext.ValText TB_Data 
            Height          =   285
            Index           =   1
            Left            =   2520
            TabIndex        =   1
            Tag             =   "FREE,FONT"
            Top             =   1560
            Width           =   1305
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
            Text            =   ""
            TXTAlign        =   2
         End
         Begin atc2valtext.ValText TB_Data 
            Height          =   285
            Index           =   2
            Left            =   2520
            TabIndex        =   2
            Tag             =   "FREE,FONT"
            Top             =   2160
            Width           =   1305
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
            Text            =   ""
            TXTAlign        =   2
         End
         Begin atc2valtext.ValText TB_Data 
            Height          =   285
            Index           =   0
            Left            =   1440
            TabIndex        =   0
            Tag             =   "FREE,FONT"
            Top             =   960
            Width           =   2385
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
            Text            =   ""
            TypeOfData      =   3
         End
         Begin VB.Label lblCategory 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            DataField       =   "UDBCode"
            DataSource      =   "DB"
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   1440
            TabIndex        =   8
            Tag             =   "FREE,FONT"
            Top             =   405
            Width           =   2385
         End
         Begin VB.Label Label6 
            Appearance      =   0  'Flat
            BackStyle       =   0  'Transparent
            Caption         =   "Actual amount made good, or amount subjected to PAYE"
            ForeColor       =   &H80000008&
            Height          =   495
            Index           =   0
            Left            =   0
            TabIndex        =   17
            Tag             =   "FREE,FONT"
            Top             =   2160
            Width           =   2535
         End
         Begin VB.Label Label3 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "P11D Class"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   0
            Left            =   0
            TabIndex        =   16
            Tag             =   "FREE,FONT"
            Top             =   0
            Width           =   825
         End
         Begin VB.Label Label7 
            Appearance      =   0  'Flat
            BackStyle       =   0  'Transparent
            Caption         =   "Gross annual amount paid by the employer "
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   0
            Left            =   0
            TabIndex        =   15
            Tag             =   "FREE,FONT"
            Top             =   1560
            Width           =   2415
         End
         Begin VB.Label Label3 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Description"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   1
            Left            =   0
            TabIndex        =   14
            Tag             =   "FREE,FONT"
            Top             =   960
            Width           =   1095
         End
         Begin VB.Label lblClass 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            DataField       =   "UDBCode"
            DataSource      =   "DB"
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   2520
            TabIndex        =   7
            Tag             =   "FREE,FONT"
            Top             =   0
            Width           =   1305
         End
      End
      Begin VB.Frame fmeApportion 
         Caption         =   "Note: Only annualised values require apportionment."
         Height          =   1065
         Left            =   120
         TabIndex        =   10
         Tag             =   "free,font"
         Top             =   3930
         Width           =   3915
         Begin atc2valtext.ValText TB_Data 
            Height          =   285
            Index           =   4
            Left            =   2550
            TabIndex        =   5
            Tag             =   "FREE,FONT"
            Top             =   660
            Width           =   1305
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
            Maximum         =   "5/4/1999"
            Minimum         =   "6/4/1998"
            AllowEmpty      =   0   'False
         End
         Begin atc2valtext.ValText TB_Data 
            Height          =   285
            Index           =   3
            Left            =   2520
            TabIndex        =   4
            Tag             =   "FREE,FONT"
            Top             =   315
            Width           =   1305
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
            Maximum         =   "5/4/1999"
            Minimum         =   "6/4/1998"
            AllowEmpty      =   0   'False
         End
         Begin VB.Label Label7 
            Appearance      =   0  'Flat
            BackStyle       =   0  'Transparent
            Caption         =   "Available From"
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   1
            Left            =   60
            TabIndex        =   12
            Tag             =   "FREE,FONT"
            Top             =   330
            Width           =   1950
         End
         Begin VB.Label Label7 
            Appearance      =   0  'Flat
            BackStyle       =   0  'Transparent
            Caption         =   "Available To"
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   2
            Left            =   60
            TabIndex        =   11
            Tag             =   "FREE,FONT"
            Top             =   660
            Width           =   1950
         End
      End
   End
   Begin MSComctlLib.ListView lb 
      Height          =   5505
      Left            =   0
      TabIndex        =   6
      Tag             =   "free,font"
      Top             =   90
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   9710
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Item"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Benefit Value"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "P/Y Value"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "F_P"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Implements IFrmGeneral
Implements IBenefitForm2

Public benefit As IBenefitClass

Public OtherDefault As String  ' Determines the default when adding benefits

Private mclsResize As New clsFormResize
Private Const L_DES_HEIGHT  As Long = 6090
Private Const L_DES_WIDTH  As Long = 8445
Private m_InvalidVT As atc2valtext.ValText


Private Sub Form_Resize()
  mclsResize.Resize
  Call ColumnWidths(lb, 50, 25, 20)
End Sub
Private Sub Form_Load()
  If Not (mclsResize.InitResize(Me, L_DES_HEIGHT, L_DES_WIDTH, DESIGN, , , MDIMain)) Then
    Err.Raise ERR_Application
  End If
End Sub

Private Sub IBenefitForm2_AddBenefit()
  Dim ben As IBenefitClass
  
  On Error GoTo AddBenefit_Err
  Call xSet("AddBenefit")
  
  Set ben = New P
  'Put in defaults for benefit
  Call AddBenefitHelper(Me, ben)
    
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
    Call .SetItem(P_CDCItem, S_P)
    Call .SetItem(P_CDCKey, "1")
    Call .SetItem(P_UDBCode, S_P_UDB)
    Call .SetItem(P_EmployeeReference, P11d32.CurrentEmployer.CurrentEmployee.PersonnelNumber)
    Call .SetItem(P_availablefrom, P11d32.Rates.GetItem(TaxYearStart))
    Call .SetItem(P_availableto, P11d32.Rates.GetItem(TaxYearEnd))
    Call .SetItem(P_item, "Please enter description...")
    Call .SetItem(P_Value, 0)
    Call .SetItem(P_MadeGood, 0)
  End With
End Function

Private Property Let IBenefitForm2_benefit(NewValue As IBenefitClass)
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
      fmeCDB.Visible = False
      fmeCDB.Enabled = False
      fmeInput.Enabled = True
      fmeApportion.Enabled = True
    Else
      fmeCDB.Visible = True
      fmeCDB.Enabled = True
      fmeInput.Enabled = False
      fmeApportion.Enabled = False
    End If
    lb.Enabled = True
    Call MDIMain.SetDelete
  ElseIf fState = FORM_DISABLED Then
    Set benefit = Nothing
    fmeInput.Enabled = False
    fmeApportion.Enabled = False
    fmeCDB.Enabled = False
    lb.Enabled = False
    Call MDIMain.ClearDelete
    Call MDIMain.ClearConfirmUndo
  End If
  IBenefitForm2_BenefitFormState = True
    
BenefitFormState_end:
  Call xReturn("BenefitFormState")
  Exit Function
  
BenefitFormState_err:
  IBenefitForm2_BenefitFormState = False
  Call ErrorMessage(ERR_ERROR, Err, "BenefitFormState", "Benefit Form State", "Error setting the 'P' benefit form state.")
  Resume BenefitFormState_end
  Resume
  
End Function

Private Function IBenefitForm2_BenefitOff() As Boolean
  TB_Data(0).Text = ""
  TB_Data(1).Text = ""
  TB_Data(2).Text = ""
  TB_Data(3).Text = ""
  TB_Data(4).Text = ""
  lblCategory = ""
  lblClass.Caption = ""
End Function

Private Function IBenefitForm2_BenefitOn() As Boolean
  TB_Data(0).Text = benefit.GetItem(P_item)
  TB_Data(1).Text = benefit.GetItem(P_Value)
  TB_Data(2).Text = benefit.GetItem(P_MadeGood)
  TB_Data(3).Text = DateStringEx(benefit.GetItem(P_availablefrom), benefit.GetItem(P_availablefrom))
  TB_Data(4).Text = DateStringEx(benefit.GetItem(P_availableto), benefit.GetItem(P_availablefrom))
  lblCategory = benefit.GetItem(P_CDCItem)
    
  lblClass.Caption = benefit.GetItem(P_UDBCode)
End Function

Private Function IBenefitForm2_BenefitsToListView() As Long
  IBenefitForm2_BenefitsToListView = BenefitsToListView(Me)
End Function

Private Function IBenefitForm2_BenefitToScreen(Optional ByVal BenefitIndex As Long = -1&, Optional ByVal UpdateBenefit As Boolean = True) As Boolean
  IBenefitForm2_BenefitToScreen = BenefitToScreenHelper(Me, BenefitIndex, UpdateBenefit)
End Function

Private Property Let IBenefitForm2_benclass(ByVal NewValue As BEN_CLASS)
End Property

Private Property Get IBenefitForm2_benclass() As BEN_CLASS
  IBenefitForm2_benclass = BC_POTHER_P
End Property

Private Property Get IBenefitForm2_lv() As MSComCtlLib.IListView
  Set IBenefitForm2_lv = lb
End Property
Private Function IBenefitForm2_RemoveBenefit(ByVal BenefitIndex As Long) As Boolean
  IBenefitForm2_RemoveBenefit = P11d32.CurrentEmployer.CurrentEmployee.RemoveBenefit(Me, benefit, BenefitIndex)
End Function
Private Function IBenefitForm2_UpdateBenefitListViewItem(li As MSComCtlLib.IListItem, benefit As IBenefitClass, Optional ByVal BenefitIndex As Long = 0&, Optional ByVal SelectItem As Boolean = False) As Boolean
  IBenefitForm2_UpdateBenefitListViewItem = UpdateBenefitListViewItem(li, benefit, BenefitIndex, SelectItem)
End Function

Private Function IBenefitForm2_ValididateBenefit(ben As IBenefitClass) As Boolean

End Function

Private Function IFrmGeneral_CheckChanged(C As Control, ByVal UpdateCurrentListItem As Boolean) As Boolean
  Dim bDirty As Boolean
  
On Error GoTo CheckChanged_Err

  Call xSet("CheckChanged")
  
  With C
    If P11d32.CurrentEmployeeIsNothing Then
      GoTo CheckChanged_End
    End If
    If benefit Is Nothing Then
      GoTo CheckChanged_End
    End If
    'we are asking if the value has changed and if it is valid thus save
    Select Case .Name
      Case "TB_Data"
        Select Case .Index
          Case 0
            bDirty = CheckTextInput(.Text, benefit, P_item)
          Case 1
            bDirty = CheckTextInput(.Text, benefit, P_Value)
          Case 2
            bDirty = CheckTextInput(.Text, benefit, P_MadeGood)
          Case 3
            bDirty = CheckTextInput(.Text, benefit, P_availablefrom)
          Case 4
            bDirty = CheckTextInput(.Text, benefit, P_availableto)
          Case Else
            ECASE "Unknown control"
        End Select
      Case Else
        ECASE "Unknown control"
    End Select
    'must be required in all check changed
    IFrmGeneral_CheckChanged = AfterCheckChanged(C, Me, bDirty, UpdateCurrentListItem)
  End With
  
CheckChanged_End:
  Call xReturn("CheckChanged")
  Exit Function
  
CheckChanged_Err:
  IFrmGeneral_CheckChanged = False
  Call ErrorMessage(ERR_ERROR, Err, "CheckChanged", "ERR_CHECKCHANGED", "This function has failed for the form " & Me.Name & ".")
  Resume CheckChanged_End












End Function

Private Property Get IFrmGeneral_InvalidVT() As atc2valtext.ValText
  IFrmGeneral_InvalidVT = m_InvalidVT
End Property

Private Property Set IFrmGeneral_InvalidVT(NewValue As atc2valtext.ValText)
  Set m_InvalidVT = NewValue
End Property

Private Sub LB_ItemClick(ByVal Item As MSComCtlLib.ListItem)
  Call SetLastListItemSelected(Item)
  If Not (lb.SelectedItem Is Nothing) Then
    Call IBenefitForm2_BenefitToScreen(Item.Tag)
  End If
End Sub

Private Sub LB_ColumnClick(ByVal ColumnHeader As MSComCtlLib.ColumnHeader)
  lb.SortKey = ColumnHeader.Index - 1
  lb.SelectedItem.EnsureVisible
End Sub

Private Sub LB_KeyPress(KeyAscii As Integer)
  'Check for return key - tab to next field
  If KeyAscii = 13 Then Call SendKeys(vbTab)
End Sub


Private Sub TB_Data_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  TB_Data(Index).Tag = SetChanged
End Sub

Private Sub TB_data_Lostfocus(Index As Integer)
  Call IFrmGeneral_CheckChanged(TB_Data(Index), True)
End Sub



