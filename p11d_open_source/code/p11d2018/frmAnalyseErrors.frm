VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.1#0"; "RICHTX32.OCX"
Begin VB.Form frmAnalyseErrors 
   Caption         =   "Form1"
   ClientHeight    =   5100
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7530
   LinkTopic       =   "Form1"
   ScaleHeight     =   5100
   ScaleWidth      =   7530
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      Height          =   375
      Left            =   6435
      TabIndex        =   2
      Tag             =   "LOCKBR"
      Top             =   4680
      Width           =   1050
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Default         =   -1  'True
      Height          =   375
      Left            =   5265
      TabIndex        =   1
      Tag             =   "LOCKBR"
      Top             =   4680
      Width           =   1095
   End
   Begin RichTextLib.RichTextBox RTErrors 
      Height          =   4515
      Left            =   45
      TabIndex        =   0
      Tag             =   "EQUALISE"
      Top             =   45
      Width           =   7440
      _ExtentX        =   13123
      _ExtentY        =   7964
      _Version        =   327681
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"frmAnalyseErrors.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmAnalyseErrors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cR As clsFormResize
Private Const L_HEIGHT As Long = 5505
Private Const L_WIDTH As Long = 7650

Private Sub cmdClose_Click()
  Me.Hide
End Sub

Private Sub cmdPrint_Click()
  Call Importing.PrintErrors(RTErrors.Text)
End Sub

Private Sub Form_Load()
  Set cR = New clsFormResize
  If cR.InitResize(Me, L_HEIGHT, L_WIDTH) = False Then
    ECASE ("Analyse Errors Resize Init Failed")
  End If
End Sub

Private Sub Form_Resize()
  Call cR.Resize
End Sub
