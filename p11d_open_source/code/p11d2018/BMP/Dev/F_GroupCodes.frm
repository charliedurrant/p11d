VERSION 5.00
Object = "{8D988532-0F0C-460C-B00E-7B5637E97680}#1.0#0"; "ATC2VTEXT.OCX"
Begin VB.Form F_GroupCodes 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Group code display names"
   ClientHeight    =   2760
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2760
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Default         =   -1  'True
      Height          =   375
      Left            =   3330
      TabIndex        =   8
      Top             =   2295
      Width           =   1230
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   2025
      TabIndex        =   7
      Top             =   2295
      Width           =   1230
   End
   Begin atc2valtext.ValText vtGroupCode1 
      Height          =   375
      Left            =   2385
      TabIndex        =   0
      Top             =   900
      Width           =   2130
      _ExtentX        =   3757
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
      Text            =   ""
      TXTAlign        =   2
      AutoSelect      =   0
   End
   Begin atc2valtext.ValText vtGroupCode2 
      Height          =   375
      Left            =   2385
      TabIndex        =   3
      Top             =   1350
      Width           =   2130
      _ExtentX        =   3757
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
      Text            =   ""
      TXTAlign        =   2
      AutoSelect      =   0
   End
   Begin atc2valtext.ValText vtGroupCode3 
      Height          =   375
      Left            =   2385
      TabIndex        =   4
      Top             =   1800
      Width           =   2130
      _ExtentX        =   3757
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
      Text            =   ""
      TXTAlign        =   2
      AutoSelect      =   0
   End
   Begin VB.Label lblGroupCode3 
      Caption         =   "Group Code 3"
      Height          =   330
      Left            =   180
      TabIndex        =   6
      Top             =   1845
      Width           =   1770
   End
   Begin VB.Label lblGroupCode2 
      Caption         =   "Group Code 2"
      Height          =   330
      Left            =   180
      TabIndex        =   5
      Top             =   1395
      Width           =   1770
   End
   Begin VB.Label lblinfo 
      Caption         =   $"F_GroupCodes.frx":0000
      Height          =   645
      Left            =   180
      TabIndex        =   2
      Top             =   90
      Width           =   4245
   End
   Begin VB.Label lblGroupCode1 
      Caption         =   "Group Code 1"
      Height          =   330
      Left            =   180
      TabIndex        =   1
      Top             =   945
      Width           =   1770
   End
End
Attribute VB_Name = "F_GroupCodes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub VTInvalid(ByVal vt As ValText, ByVal lbl As Label)
  vt.Text = Trim$(vt.Text)
  If vt.FieldInvalid Then
    Call Err.Raise(ERR_INVALID, "VTInvalid", "The field " & lbl.Caption & " is invalid")
  End If
End Sub
Private Sub VTSettings(ByVal vt As ValText)
  vt.AllowEmpty = False
  vt.TypeOfData = VT_STRING
  
End Sub

Public Sub SettingsToScreen()
End Sub

Private Sub cmdCancel_Click()
  Me.Hide
  Unload Me
End Sub

Private Sub cmdOK_Click()
  On Error GoTo err_Err
  
  Call VTInvalid(vtGroupCode1, lblGroupCode1)
  Call VTInvalid(vtGroupCode2, lblGroupCode2)
  Call VTInvalid(vtGroupCode3, lblGroupCode3)
  
  With p11d32
  
    .GroupCode1Alias = vtGroupCode1.Text
    .GroupCode2Alias = vtGroupCode2.Text
    .GroupCode3Alias = vtGroupCode3.Text
    .BenDataLinkFieldDetails(BC_EMPLOYEE, ee_Group1_db).Description = .GroupCode1Alias
    .BenDataLinkFieldDetails(BC_EMPLOYEE, ee_Group2_db).Description = .GroupCode2Alias
    .BenDataLinkFieldDetails(BC_EMPLOYEE, ee_Group3_db).Description = .GroupCode3Alias
  End With
  
  Call F_Employees.UpdateGroupCodeLables
  Call Me.Hide
  Unload Me
err_End:
  
  Exit Sub
err_Err:
  Call ErrorMessage(ERR_ERROR, Err, "OK", "OK", Err.Description)
  Resume err_End
End Sub

Private Sub F_GroupCode_Click()

End Sub

