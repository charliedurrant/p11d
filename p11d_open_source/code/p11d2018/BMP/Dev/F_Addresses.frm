VERSION 5.00
Object = "{8D988532-0F0C-460C-B00E-7B5637E97680}#1.0#0"; "atc2vtext.ocx"
Begin VB.Form F_Addresses 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Address"
   ClientHeight    =   3735
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5370
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3735
   ScaleWidth      =   5370
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   375
      Left            =   4005
      TabIndex        =   7
      Top             =   3285
      Width           =   1275
   End
   Begin atc2valtext.ValText txt 
      Height          =   375
      Index           =   0
      Left            =   1440
      TabIndex        =   0
      Top             =   90
      Width           =   3840
      _ExtentX        =   6773
      _ExtentY        =   661
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
      AutoSelect      =   0
   End
   Begin atc2valtext.ValText txt 
      Height          =   375
      Index           =   1
      Left            =   1440
      TabIndex        =   1
      Top             =   540
      Width           =   3840
      _ExtentX        =   6773
      _ExtentY        =   661
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
      AutoSelect      =   0
   End
   Begin atc2valtext.ValText txt 
      Height          =   375
      Index           =   2
      Left            =   1440
      TabIndex        =   2
      Top             =   990
      Width           =   3840
      _ExtentX        =   6773
      _ExtentY        =   661
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
      AutoSelect      =   0
   End
   Begin atc2valtext.ValText txt 
      Height          =   375
      Index           =   3
      Left            =   1440
      TabIndex        =   3
      Top             =   1440
      Width           =   2580
      _ExtentX        =   4551
      _ExtentY        =   661
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
      AutoSelect      =   0
   End
   Begin atc2valtext.ValText txt 
      Height          =   375
      Index           =   4
      Left            =   1440
      TabIndex        =   4
      Top             =   1890
      Width           =   2580
      _ExtentX        =   4551
      _ExtentY        =   661
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
      AutoSelect      =   0
   End
   Begin atc2valtext.ValText txt 
      Height          =   375
      Index           =   5
      Left            =   1440
      TabIndex        =   5
      Top             =   2340
      Width           =   1770
      _ExtentX        =   3122
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxLength       =   15
      Text            =   ""
      TypeOfData      =   3
      AutoSelect      =   0
   End
   Begin atc2valtext.ValText txt 
      Height          =   375
      Index           =   6
      Left            =   1440
      TabIndex        =   6
      Top             =   2790
      Width           =   1770
      _ExtentX        =   3122
      _ExtentY        =   661
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
      AutoSelect      =   0
   End
   Begin VB.Label lbl 
      Caption         =   "Country"
      ForeColor       =   &H8000000D&
      Height          =   240
      Index           =   6
      Left            =   180
      TabIndex        =   14
      Top             =   2835
      Width           =   1095
   End
   Begin VB.Label lbl 
      Caption         =   "Address line 1"
      ForeColor       =   &H8000000D&
      Height          =   240
      Index           =   5
      Left            =   180
      TabIndex        =   13
      Top             =   180
      Width           =   1095
   End
   Begin VB.Label lbl 
      Caption         =   "Postcode"
      ForeColor       =   &H8000000D&
      Height          =   240
      Index           =   4
      Left            =   180
      TabIndex        =   12
      Top             =   2430
      Width           =   1095
   End
   Begin VB.Label lbl 
      Caption         =   "City"
      ForeColor       =   &H8000000D&
      Height          =   240
      Index           =   3
      Left            =   180
      TabIndex        =   11
      Top             =   1530
      Width           =   1095
   End
   Begin VB.Label lbl 
      Caption         =   "Address line 3"
      ForeColor       =   &H8000000D&
      Height          =   240
      Index           =   2
      Left            =   180
      TabIndex        =   10
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label lbl 
      Caption         =   "Address line 2"
      ForeColor       =   &H8000000D&
      Height          =   240
      Index           =   1
      Left            =   180
      TabIndex        =   9
      Top             =   630
      Width           =   1095
   End
   Begin VB.Label lbl 
      Caption         =   "County"
      ForeColor       =   &H8000000D&
      Height          =   240
      Index           =   0
      Left            =   180
      TabIndex        =   8
      Top             =   1980
      Width           =   1095
   End
End
Attribute VB_Name = "F_Addresses"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Implements IFrmGeneral
Implements IBenefitForm2

Public benefit As IBenefitClass
Public Parentibf As IBenefitForm2

Private m_InvalidVT As Control
Private Sub cmdClose_Click()
  Call CheckValidityAndBenefitDirty(benefit, Me)
End Sub

Private Sub IBenefitForm2_AddBenefit()

End Sub

Private Function IBenefitForm2_AddBenefitSetDefaults(ben As IBenefitClass) As Long

End Function

Private Property Set IBenefitForm2_benefit(NewValue As IBenefitClass)
  Set benefit = NewValue
End Property

Private Property Get IBenefitForm2_benefit() As IBenefitClass
  Set IBenefitForm2_benefit = benefit
End Property

Private Function IBenefitForm2_BenefitFormState(ByVal fState As BENEFIT_FORM_STATE) As Boolean

End Function

Private Function IBenefitForm2_BenefitOff() As Boolean
End Function

Private Function IBenefitForm2_BenefitOn() As Boolean
  txt(0).Text = benefit.value(ee_AddressLine1_db)
  txt(1).Text = benefit.value(ee_AddressLine2_db)
  txt(2).Text = benefit.value(ee_AddressLine3_db)
  txt(3).Text = benefit.value(ee_City_db)
  txt(4).Text = benefit.value(ee_County_db)
  txt(5).Text = benefit.value(ee_PostCode_db)
  txt(6).Text = benefit.value(ee_Country_db)
End Function

Private Function IBenefitForm2_BenefitsToListView() As Long

End Function

Private Function IBenefitForm2_BenefitToListView(ben As IBenefitClass, ByVal lBenefitIndex As Long) As Long

End Function

Private Function IBenefitForm2_BenefitToScreen(Optional ByVal BenefitIndex As Long = -1&, Optional ByVal UpdateBenefit As Boolean = True) As Boolean
  Dim ibf As IBenefitForm2
  
On Error GoTo BenefitToScreen_Err
    
  Call xSet("BenefitToScreen")
  
  If BenefitIndex <> -1 Then
    Set ibf = Me
    Set benefit = p11d32.CurrentEmployer.CurrentEmployee
    If benefit.BenefitClass <> ibf.benclass Then Call Err.Raise(ERR_INVALIDBENCLASS, "BenefitToScreen", "Benefit type invalid.")
    Call ibf.BenefitOn
'    Me.Show vbModal
    Call p11d32.Help.ShowForm(Me, vbModal)
    IBenefitForm2_BenefitToScreen = True
  End If
  
BenefitToScreen_End:
  Set ibf = Nothing
  Call xReturn("BenefitToScreen")
  Exit Function
BenefitToScreen_Err:
  Call ErrorMessage(ERR_ERROR, Err, "BenefitToScreen", "Benefit To Screen", "Unlable to place the address data to the screen.")
  Resume BenefitToScreen_End
  Resume
End Function

Private Property Let IBenefitForm2_benclass(ByVal RHS As BEN_CLASS)

End Property

Private Property Get IBenefitForm2_benclass() As BEN_CLASS
  IBenefitForm2_benclass = BC_EMPLOYEE
End Property

'Private Property Get IBenefitForm2_ControlDefault() As Control
'  Set IBenefitForm2_ControlDefault = txt(0)
'End Property

Private Property Get IBenefitForm2_lv() As MSComctlLib.IListView

End Property

Private Function IBenefitForm2_RemoveBenefit(ByVal BenefitIndex As Long) As Boolean

End Function

Private Function IBenefitForm2_UpdateBenefitListViewItem(li As MSComctlLib.IListItem, benefit As IBenefitClass, Optional ByVal BenefitIndex As Long = 0&, Optional ByVal SelectItem As Boolean = False) As Long

End Function

Private Function IBenefitForm2_ValididateBenefit(ben As IBenefitClass) As Boolean

End Function

Private Function IFrmGeneral_CheckChanged(c As Control) As Boolean
  Dim bDirty As Boolean
  Dim ee As Employee
  
  On Error GoTo CheckChanged_Err
  Call xSet("CheckChanged")
  
  
  If p11d32.CurrentEmployeeIsNothing Then GoTo CheckChanged_End
  If benefit Is Nothing Then GoTo CheckChanged_End
  
  With c
    Select Case .Name
      Case "txt"
        Select Case .Index
          Case 0
            bDirty = CheckTextInput(.Text, benefit, ee_AddressLine1_db)
          Case 1
            bDirty = CheckTextInput(.Text, benefit, ee_AddressLine2_db)
          Case 2
            bDirty = CheckTextInput(.Text, benefit, ee_AddressLine3_db)
          Case 3
            bDirty = CheckTextInput(.Text, benefit, ee_City_db)
          Case 4
            bDirty = CheckTextInput(.Text, benefit, ee_County_db)
          Case 5
            bDirty = CheckTextInput(.Text, benefit, ee_PostCode_db)
          Case 6
            bDirty = CheckTextInput(.Text, benefit, ee_Country_db)
          Case Else
            ECASE "Unknown control"
        End Select
      Case Else
        ECASE "Unknown control"
    End Select
  End With
  
  Set ee = benefit
  ee.HasAddress = ee.HasAddress Or bDirty
  IFrmGeneral_CheckChanged = AfterCheckChanged(c, Me.Parentibf, bDirty)
   
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

Private Sub txt_Validate(Index As Integer, Cancel As Boolean)
  Call IFrmGeneral_CheckChanged(txt(Index))
End Sub
