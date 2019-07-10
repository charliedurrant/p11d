VERSION 5.00
Object = "{9095D21F-26FB-4046-99E9-6211396F83DF}#1.0#0"; "atc2dmenu.OCX"
Begin VB.Form frmTestMenus 
   Caption         =   "Test Menus"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin atc2DMenu.DMenu dmenu 
      Left            =   3930
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin VB.CommandButton cmdAddMenus 
      Caption         =   "Add menus"
      Height          =   525
      Left            =   90
      TabIndex        =   0
      Top             =   360
      Width           =   1455
   End
   Begin VB.Menu menuFile 
      Caption         =   "&File"
   End
End
Attribute VB_Name = "frmTestMenus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_Menu As TCSDMenu.VBMenu
Private Sub cmdAddMenus_Click()
  Dim vbmi As TCSDMenu.VBMenuItem
  
  On Error GoTo AddMenus_ERR
  
  Set m_Menu = Me.dmenu.Add("MainMenu")
  
  Set vbmi = m_Menu.Add("m1", "m1", "")
  Call m_Menu.Add("m2", "m2", "m1")
  dmenu.hWnd = Me.hWnd
  
AddMenus_END:
  Exit Sub
AddMenus_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "AddMenus", "Add Menus", "error in AddMenus")
  Resume AddMenus_END
  Resume
End Sub
