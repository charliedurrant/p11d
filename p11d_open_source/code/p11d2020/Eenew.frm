VERSION 5.00
Object = "{8D988532-0F0C-460C-B00E-7B5637E97680}#1.0#0"; "ATC2VTEXT.OCX"
Begin VB.Form F_EmployeeNew 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Add Employee"
   ClientHeight    =   2985
   ClientLeft      =   2820
   ClientTop       =   3795
   ClientWidth     =   5565
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2985
   ScaleWidth      =   5565
   StartUpPosition =   1  'CenterOwner
   Begin atc2valtext.ValText TxtBx 
      Height          =   285
      Index           =   2
      Left            =   1215
      TabIndex        =   2
      Tag             =   "0"
      Top             =   810
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   503
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxLength       =   20
      Text            =   ""
      TypeOfData      =   3
   End
   Begin VB.CommandButton B_Add 
      Caption         =   "&Add"
      Height          =   375
      Left            =   2925
      TabIndex        =   6
      Tag             =   "4"
      Top             =   2565
      Width           =   1245
   End
   Begin VB.CommandButton B_Cancel 
      Caption         =   "&Close"
      Height          =   375
      Left            =   4230
      TabIndex        =   7
      Tag             =   "5"
      Top             =   2565
      Width           =   1245
   End
   Begin atc2valtext.ValText TxtBx 
      Height          =   285
      Index           =   0
      Left            =   1215
      TabIndex        =   0
      Tag             =   "0"
      Top             =   90
      Width           =   1605
      _ExtentX        =   2831
      _ExtentY        =   503
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
      Height          =   285
      Index           =   5
      Left            =   1215
      TabIndex        =   5
      Tag             =   "0"
      Top             =   1890
      Width           =   1605
      _ExtentX        =   2831
      _ExtentY        =   503
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
   End
   Begin atc2valtext.ValText TxtBx 
      Height          =   285
      Index           =   3
      Left            =   1215
      TabIndex        =   3
      Tag             =   "0"
      Top             =   1170
      Width           =   4260
      _ExtentX        =   7514
      _ExtentY        =   503
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
   End
   Begin atc2valtext.ValText TxtBx 
      Height          =   285
      Index           =   1
      Left            =   1215
      TabIndex        =   1
      Tag             =   "0"
      Top             =   450
      Width           =   4260
      _ExtentX        =   7514
      _ExtentY        =   503
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
      Height          =   285
      Index           =   4
      Left            =   1215
      TabIndex        =   4
      Tag             =   "0"
      Top             =   1530
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   503
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
   End
   Begin VB.Label Lab 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Initials"
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   6
      Left            =   135
      TabIndex        =   13
      Top             =   1575
      Width           =   435
   End
   Begin VB.Label Lab 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "First name"
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   5
      Left            =   135
      TabIndex        =   12
      Top             =   1215
      Width           =   720
   End
   Begin VB.Label Lab 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Title"
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   4
      Left            =   135
      TabIndex        =   11
      Top             =   855
      Width           =   300
   End
   Begin VB.Label Lab 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Surname"
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   0
      Left            =   135
      TabIndex        =   8
      Top             =   495
      Width           =   630
   End
   Begin VB.Label Lab 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NI number"
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   2
      Left            =   135
      TabIndex        =   10
      Top             =   1935
      Width           =   735
   End
   Begin VB.Label Lab 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Personnel ID"
      ForeColor       =   &H00800000&
      Height          =   240
      Index           =   1
      Left            =   135
      TabIndex        =   9
      Top             =   135
      Width           =   1275
   End
End
Attribute VB_Name = "F_EmployeeNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Implements IFrmGeneral
Public m_ok As Boolean
Private m_InvalidVT As Control

Private Function IFrmGeneral_CheckChanged(c As Control) As Boolean

End Function

Private Property Get IFrmGeneral_InvalidVT() As Control
  Set IFrmGeneral_InvalidVT = m_InvalidVT
End Property

Private Property Set IFrmGeneral_InvalidVT(NewValue As Control)
  Set m_InvalidVT = NewValue
End Property

Public Sub ClearFields()

  On Error GoTo ClearFields_Err
  Call xSet("ClearFields")
  
  TxtBx(0).Text = ""
  TxtBx(1).Text = ""
  TxtBx(2).Text = ""
  
ClearFields_End:
  Call xReturn("ClearFields")
  Exit Sub

ClearFields_Err:
  Call ErrorMessage(ERR_ERROR, Err, "ClearFields", "Clear Fields", "Error clearing the forms fields.")
  Resume ClearFields_End
End Sub


Private Sub B_Add_Click()
  m_ok = True
  Call CheckValidity(Me)
End Sub

Private Sub B_Cancel_Click()
  m_ok = False
  Me.Hide
End Sub

Private Sub TxtBx_FieldInvalid(Index As Integer, Valid As Boolean, Message As String)
  Call SetPanel2(Message)
End Sub
Private Sub TxtBx_GotFocus(Index As Integer)
  TxtBx(Index).lValidate
End Sub
