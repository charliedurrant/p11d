VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form datest 
   Caption         =   "Form1"
   ClientHeight    =   6375
   ClientLeft      =   1545
   ClientTop       =   1455
   ClientWidth     =   7440
   LinkMode        =   1  'Source
   LinkTopic       =   "control"
   ScaleHeight     =   6375
   ScaleWidth      =   7440
   Begin VB.CommandButton dummy 
      Caption         =   "Dummy"
      Height          =   375
      Left            =   5520
      TabIndex        =   4
      Top             =   1680
      Width           =   975
   End
   Begin VB.CommandButton Touch 
      Caption         =   "Touch"
      Height          =   375
      Left            =   5520
      TabIndex        =   3
      Top             =   1080
      Width           =   975
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4680
      Width           =   2535
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "dummy.frx":0000
      Height          =   3255
      Left            =   360
      OleObjectBlob   =   "dummy.frx":0014
      TabIndex        =   2
      Top             =   1200
      Width           =   4575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Go"
      Height          =   375
      Left            =   5520
      TabIndex        =   1
      Top             =   480
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   480
      Width           =   2415
   End
End
Attribute VB_Name = "datest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim dadb As DADatabase
Dim WithEvents dars As DARecordset
Attribute dars.VB_VarHelpID = -1
Dim WithEvents interestdars As DARecordset
Attribute interestdars.VB_VarHelpID = -1

Public WithEvents dacon  As DAConnection
Attribute dacon.VB_VarHelpID = -1
Private Sub Command1_Click()
  Set dadb = dacon.OpenDatabase(Text1.text, "xx")
  Set dars = dadb.OpenRecordset("Select * from output_d3")
  Set interestdars = dadb.OpenRecordset("Select * from output_a22")
  Set Data1.Recordset = dars.Recordset
'MsgBox "ok"
End Sub

Private Sub dacon_DAConnectionError(errno As Long, text As String)
  MsgBox "error " + text
End Sub

Private Sub dars_DARSRefreshRequired()
'  MsgBox "dars event"
  dars.Refresh
  Set Data1.Recordset = dars.Recordset
End Sub

Private Sub dummy_Click()
  Set dadb = dacon.OpenDatabase(Text1.text, "xx")
  Set dars = dadb.OpenRecordset("dummy")
  Set Data1.Recordset = dars.Recordset

End Sub


Private Sub Form_LinkExecute(CmdStr As String, Cancel As Integer)
  Dim pos As String
  Dim lp As String
  Dim rp As String
  On Error GoTo err_linkexecute
  
  pos = InStr(CmdStr, " ")
  If pos > 0 Then
    lp = Left$(CmdStr, pos - 1)
    rp = Right$(CmdStr, Len(CmdStr) - pos)
  Else
    lp = CmdStr
    rp = ""
  End If
  
  Select Case lp
  Case "FILE"
    Set dadb = dacon.OpenDatabase(rp, "xx")
    Set dars = dadb.OpenRecordset("Select * from output_d3")
'    Set Data1.Recordset = dars.Recordset
  Case "DOWNLOAD"
    dadb.run "sys_addtotree"
    If dadb.mode = "FLAT" Then
      dadb.run "sys_FlatCreateSubScheds"
    End If
  Case "RECALC"
    dadb.Touch "abacustree"
    dadb.Touch "schedules"
    dadb.Touch "ProfitAndLoss"
    dadb.Touch "InvestmentIncome"
    dadb.Touch "LeaseCars"
    dadb.Touch "InvestmentIncome"
    dadb.Touch "Entertaining"
    dadb.Touch "InterestPaid"
    dadb.Touch "Fadds"
    dadb.Touch "FAdisps"
  Case Else
    MsgBox "Unknown dde command " + CmdStr
  End Select
  Cancel = 0
  Exit Sub
  
err_linkexecute:
    MsgBox "DDE error" + Err.Description
    
End Sub

Private Sub Form_Load()
  If CoreSetup("", VB.Global) Then
    Set dacon = New DAConnection
    dacon.debugmode = False
'  Text1.text = App.Path + "\dadev.mdb"
    Text1.text = "c:\c\devex\dadev.mdb"
  Else
    MsgBox "tcscore failure"
  End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
  If Not dars Is Nothing Then dars.closedars
  If Not dadb Is Nothing Then dadb.closedb
End Sub

Private Sub interestdars_DARSRefreshRequired()
  interestdars.Refresh
End Sub

Private Sub Touch_Click()
  dadb.Touch "ProfitAndLoss"
  dadb.Touch "InvestmentIncome"
  dadb.Touch "InterestPaid"
End Sub
