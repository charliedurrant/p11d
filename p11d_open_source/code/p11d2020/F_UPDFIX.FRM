VERSION 5.00
Object = "{8D988532-0F0C-460C-B00E-7B5637E97680}#1.0#0"; "atc2vtext.ocx"
Begin VB.Form F_UpdateFixLevel 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Update Fix Level"
   ClientHeight    =   1350
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3180
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1350
   ScaleWidth      =   3180
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   1980
      TabIndex        =   3
      Top             =   900
      Width           =   1140
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   810
      TabIndex        =   2
      Top             =   900
      Width           =   1095
   End
   Begin atc2valtext.ValText vtFixLevel 
      Height          =   330
      Left            =   1665
      TabIndex        =   0
      Top             =   450
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   582
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
   Begin VB.Label lblCurrentFixLEvel 
      Caption         =   "lblCurrentFixLevel"
      Height          =   285
      Left            =   90
      TabIndex        =   1
      Top             =   90
      Width           =   2985
   End
End
Attribute VB_Name = "F_UpdateFixLevel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public OK As Boolean

Private Sub cmdOk_Click()
  If Not vtFixLevel.FieldInvalid Then
    OK = True
    Me.Hide
  Else
    Call ErrorMessage(ERR_ERROR, Err, "cmdOK_Click", "cmd OK Click", "Invalid fix level, must be a value between " & vtFixLevel.Minimum & " and " & vtFixLevel.Maximum & ".")
  End If
End Sub
