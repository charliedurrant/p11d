VERSION 5.00
Object = "{00028C4A-0000-0000-0000-000000000046}#5.0#0"; "TDBG5.OCX"
Begin VB.UserControl AutoCombo_DAO 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin VB.Data Data_i 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   915
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2310
      Visible         =   0   'False
      Width           =   1905
   End
   Begin TrueDBGrid50.TDBDropDown Combo_i 
      Bindings        =   "ACombo.ctx":0000
      Height          =   1410
      Left            =   210
      OleObjectBlob   =   "ACombo.ctx":0011
      TabIndex        =   0
      Top             =   390
      Width           =   4080
   End
End
Attribute VB_Name = "AutoCombo_DAO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Combo As TDBDropDown
Public ComboDataControl As Data

Private Sub UserControl_Initialize()
  Set Combo = Combo_i
  Set ComboDataControl = Data_i
End Sub
