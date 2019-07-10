VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmOutPut 
   Caption         =   "Info"
   ClientHeight    =   5790
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5700
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5790
   ScaleWidth      =   5700
   WindowState     =   2  'Maximized
   Begin RichTextLib.RichTextBox txtInfo 
      Height          =   5730
      Left            =   45
      TabIndex        =   0
      Tag             =   "EQUALISE"
      Top             =   45
      Width           =   5640
      _ExtentX        =   9948
      _ExtentY        =   10107
      _Version        =   393217
      ScrollBars      =   2
      TextRTF         =   $"frmOutPut.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmOutPut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_RS As clsFormResize


Private Sub Form_Load()
  Set m_RS = New clsFormResize
  
  Call m_RS.InitResize(Me, 6195, 5820, DESIGN, , , frmMain)
End Sub

Private Sub Form_Resize()
  Call m_RS.Resize
End Sub

