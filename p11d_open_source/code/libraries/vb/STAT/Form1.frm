VERSION 5.00
Object = "{8B84FB9A-24FA-11D3-8C27-00508B2FF337}#1.1#0"; "TCSSTAT.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3630
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5070
   LinkTopic       =   "Form1"
   ScaleHeight     =   3630
   ScaleWidth      =   5070
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1845
      Top             =   1575
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0000
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin TCSStat.TCSStatus TCSStatus1 
      Align           =   2  'Align Bottom
      Height          =   420
      Left            =   0
      TabIndex        =   0
      Top             =   3210
      Width           =   5070
      _ExtentX        =   8943
      _ExtentY        =   741
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private p As TCSPANEL

Private Sub Form_Load()
  Set p = Me.TCSStatus1.AddPanel(50, "Fred")
  p.ForeColor = 400000
  p.Font.Bold = True
  p.Font.Italic = True
  p.ToolTipText = "ToopTip0"
  Set p.Picture = Me.ImageList1.ListImages(1).Picture
  
  Set p = Me.TCSStatus1.AddPanel(50, "Jim")
  p.ForeColor = 800000
  p.Font.Name = "Times New Roman"
  p.Font.Size = 11
  
  p.Font.Bold = True
  p.ToolTipText = "ToopTip1"
  'Set p.Picture = Me.ImageList1.ListImages(1).Picture
  
End Sub
Private Sub TCSStatus1_PanelMouseDown(ByVal p As TCSStat.TCSPANEL, Button As Integer, Shift As Integer, x As Single, y As Single)
  MsgBox p.Caption
  'p.Font.Name = "WingDings"
End Sub
Private Sub TCSStatus1_PictureMouseDown(ByVal p As TCSStat.TCSPANEL, Button As Integer, Shift As Integer, x As Single, y As Single)
  MsgBox p.Caption
End Sub
