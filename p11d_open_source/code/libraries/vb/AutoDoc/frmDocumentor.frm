VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDocument 
   Caption         =   "Documentor"
   ClientHeight    =   5835
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8400
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5835
   ScaleWidth      =   8400
   Begin VB.CommandButton cmdProcess 
      Caption         =   "&Process Files"
      Height          =   375
      Left            =   6930
      TabIndex        =   1
      Tag             =   "LOCKBR"
      Top             =   5400
      Width           =   1455
   End
   Begin MSComctlLib.ListView lvFiles 
      Height          =   5280
      Left            =   45
      TabIndex        =   0
      Tag             =   "EQUALISE"
      Top             =   45
      Width           =   8340
      _ExtentX        =   14711
      _ExtentY        =   9313
      View            =   2
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
End
Attribute VB_Name = "frmDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_CR As clsFormResize
Private Sub cmdProcess_Click()
  Dim sFiles() As String
  Call ProcessFiles(sFiles, GetFiles(sFiles))
End Sub

Private Sub Form_Load()
  Set m_CR = New clsFormResize
  Call m_CR.InitResize(Me, 6240, 8520, DESIGN, , , frmMain)
  Set Me.lvFiles.SmallIcons = frmMain.imlTickCross
  frmMain.drive.drive = AppPath
  frmMain.dir.Path = AppPath
End Sub
Private Sub Form_Resize()
  m_CR.Resize
End Sub

Private Sub lvFiles_ItemClick(ByVal Item As MSComctlLib.ListItem)
  If Item.SmallIcon = IMG_CROSS Then
    Item.SmallIcon = IMG_TICK
  Else
    Item.SmallIcon = IMG_CROSS
  End If
End Sub
