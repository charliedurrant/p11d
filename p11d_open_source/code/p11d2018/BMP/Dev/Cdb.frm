VERSION 5.00
Object = "{412521B9-9CBB-4049-9E66-2AA0112EC306}#1.0#0"; "ATC2VTEXT.OCX"
Begin VB.Form F_CDB 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Company Defined Benefits"
   ClientHeight    =   1935
   ClientLeft      =   630
   ClientTop       =   5370
   ClientWidth     =   8100
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1935
   ScaleWidth      =   8100
   Begin VB.CommandButton B_Ok 
      Caption         =   "&OK"
      Height          =   350
      Left            =   5760
      TabIndex        =   5
      Top             =   1500
      Width           =   1065
   End
   Begin VB.CommandButton B_Cancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   350
      Left            =   6870
      TabIndex        =   6
      Top             =   1500
      Width           =   1065
   End
   Begin VB.ComboBox CboBx 
      ForeColor       =   &H00000080&
      Height          =   315
      Index           =   0
      Left            =   2040
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   960
      Width           =   2865
   End
   Begin atc2valtext.ValText TxtBx 
      Height          =   315
      Index           =   3
      Left            =   6900
      TabIndex        =   4
      Top             =   870
      Width           =   1005
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
      Text            =   "0"
      AllowEmpty      =   0   'False
      TXTAlign        =   2
   End
   Begin atc2valtext.ValText TxtBx 
      Height          =   315
      Index           =   1
      Left            =   6900
      TabIndex        =   3
      Top             =   480
      Width           =   1005
      _ExtentX        =   1773
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
   Begin atc2valtext.ValText TxtBx 
      Height          =   285
      Index           =   2
      Left            =   2010
      TabIndex        =   1
      Top             =   510
      Width           =   2865
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
      Text            =   ""
      TypeOfData      =   3
      AllowEmpty      =   0   'False
   End
   Begin atc2valtext.ValText TxtBx 
      Height          =   285
      Index           =   0
      Left            =   2010
      TabIndex        =   0
      Top             =   90
      Width           =   1005
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
      Text            =   ""
      TypeOfData      =   3
      AllowEmpty      =   0   'False
   End
   Begin VB.Label Lab 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Amount made good"
      ForeColor       =   &H00800000&
      Height          =   192
      Index           =   3
      Left            =   5136
      TabIndex        =   11
      Top             =   960
      Width           =   1716
   End
   Begin VB.Label Lab 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Value of Benefit"
      ForeColor       =   &H00800000&
      Height          =   192
      Index           =   2
      Left            =   5136
      TabIndex        =   10
      Top             =   576
      Width           =   1716
   End
   Begin VB.Label Lab 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Description of benefit"
      ForeColor       =   &H00800000&
      Height          =   192
      Index           =   1
      Left            =   36
      TabIndex        =   8
      Top             =   540
      Width           =   1896
   End
   Begin VB.Label Lab 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Benefit Catagory"
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   4
      Left            =   90
      TabIndex        =   9
      Top             =   960
      Width           =   1890
   End
   Begin VB.Label Lab 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Unique code for benefit"
      ForeColor       =   &H00800000&
      Height          =   192
      Index           =   0
      Left            =   36
      TabIndex        =   7
      Top             =   120
      Width           =   1908
   End
End
Attribute VB_Name = "F_CDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public benefit As IBenefitClass

Private Sub B_OK_Click()
  If benefit.Dirty = True Then
    If InvalidFields(Me) > 0 Then
      Call MDIMain.sts.SetStatus(0, "", "There are invalid fields on this dialog")
      Beep
    Else
      Call MDIMain.sts.SetStatus(0, "", "")
      MDIMain.SetConfirmUndo
      Me.Hide
    End If
  Else
    Me.Hide
  End If
  

End Sub

Private Function CheckChanged(ctl As Control) As Boolean
  Dim mdi As MDIForm
  Dim ben As IBenefitClass
  Dim i As Long
  On Error GoTo CheckChanged_Err
  Call xSet("CheckChanged")
  Set mdi = MDIMain
  Set ben = Me.benefit
  Select Case ctl.Name
    Case "TxtBx"
      Select Case ctl.Index
        Case 0
          'i = StrComp(S_CDB & ctl.Text, ben.GetItem(Oth_EmployeeReference), vbTextCompare)
          'If i <> 0 Then Call ben.SetItem(Oth_EmployeeReference, S_CDB & ctl.Text)
        Case 1
          i = StrComp(ctl.Text, ben.GetItem(Oth_Value), vbBinaryCompare)
          If i <> 0 Then Call ben.SetItem(Oth_Value, ctl.Text)
        Case 2
          i = StrComp(ctl.Text, ben.GetItem(Oth_item), vbBinaryCompare)
          If i <> 0 Then Call ben.SetItem(Oth_item, ctl.Text)
        Case 3
          i = StrComp(ctl.Text, ben.GetItem(Oth_MadeGood), vbBinaryCompare)
          If i <> 0 Then Call ben.SetItem(Oth_MadeGood, ctl.Text)
        Case Else
          ECASE "Unknown"
      End Select
    Case "CboBx"
      Select Case ctl.Index
        Case 0
          i = StrComp(ctl.Text, ben.GetItem(Oth_Class), vbBinaryCompare)
          If i <> 0 Then Call ben.SetItem(Oth_Class, ctl.Text)
        Case Else
          ECASE "Unknown"
      End Select
    Case Else
      ECASE "Unknown"
  End Select
  If i <> 0 Then
    ben.Dirty = True
  End If
CheckChanged_End:
  Set ben = Nothing
  Set mdi = Nothing
  Call xReturn("CheckChanged")
  Exit Function
CheckChanged_Err:
  Call ErrorMessage(ERR_ERROR, Err, "CheckChanged", "ERR_CHECKCHANGED", "This function has failed for the form " & Me.Name & ".")
  Resume CheckChanged_End
End Function

Private Sub CboBx_Lostfocus(Index As Integer)
  Call CheckChanged(CboBx(Index))
End Sub

Private Sub Form_Load()
  CboBx(0).AddItem (S_MEDICAL)
  CboBx(0).AddItem (S_CREDIT)
  CboBx(0).AddItem (S_EDUCATION)
  CboBx(0).AddItem (S_ENTERTAINMENT)
  CboBx(0).AddItem (S_GENERAL)
  CboBx(0).AddItem (S_NOTIONAL)
  CboBx(0).AddItem (S_NURSERY)
  CboBx(0).AddItem (S_PAYMENTS)
  CboBx(0).AddItem (S_SHARES)
  CboBx(0).AddItem (S_SUBSCRIPTION)
  CboBx(0).AddItem (S_TRAVEL)
  CboBx(0).AddItem (S_TAXPAID)
End Sub

Private Sub TxtBx_FieldInvalid(Index As Integer, Valid As Boolean, Message As String)
  Call MDIMain.sts.SetStatus(0, Message)
End Sub

Private Sub TxtBx_LostFocus(Index As Integer)
  Call CheckChanged(TxtBx(Index))
End Sub
