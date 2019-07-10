VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   6825
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10665
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6825
   ScaleWidth      =   10665
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "test nav unordered (a+)"
      Height          =   375
      Left            =   2760
      TabIndex        =   2
      Top             =   120
      Width           =   2055
   End
   Begin VB.TextBox TExt1 
      Height          =   6015
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Text            =   "frmMain.frx":0000
      Top             =   720
      Width           =   10455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "test nav ordered (a+)"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
  Dim xto As New XMLTreeview, sl As StringList, SLAttribs  As StringList
  Dim rs As Recordset, s As String
  Dim cmd As Command
  Dim cn As Connection
  Dim t0 As Long
  Dim i As Long
  Dim scn As String
  On Error GoTo Test_Nav_ERR
        
  'oracle
  scn = ADOOracleConnectString("APLUS.LONAP3004", "aplus_dev_ing", "ukcentral")
  Set cn = ADOConnect(scn, adUseClient)
  Set cmd = New Command
  Set cmd.ActiveConnection = cn
  cmd.CommandType = adCmdText
  cmd.CommandText = "SELECT id, parent, name, tooltip, children, editable, actual_value, nodeid, NULL AS default_value From (SELECT id, parent, name, tooltip, children, editable, actual_value, nodeid, NULL AS default_value, option_order From (SELECT 2 AS id, entity_pid AS nodeid, parent_id AS parent, entity_name AS name, entity_name AS tooltip, has_children AS children, entity_pid AS actual_value, 0 AS editable, entity_order As option_order From entity_values Union All SELECT 2 AS id, 0 AS nodeid, TO_NUMBER(NULL) AS parent, '(empty)'  AS name, TO_CHAR(NULL) AS tooltip, 0 AS children, TO_NUMBER(NULL) AS actual_value, 0 AS editable, 0 AS option_order FROM dual) ORDER BY UPPER(name))"
  Set rs = cmd.Execute
  t0 = GetTicks
  
  Set xto = New XMLTreeview
  Set sl = New StringList
  Call sl.Add("tooltip")
  Call sl.Add("name")
  Call sl.Add("default_value")
  Call sl.Add("actual_value")
  Call sl.Add("value")
  Call sl.Add("dependent_parent")
  Call sl.Add("dependent_value")
  Set SLAttribs = New StringList
  Call SLAttribs.Add("children")
  s = xto.RecordSetToXML2(rs, "", "nodeid", "parent", sl, SLAttribs)
  Call MsgBox(GetTicks() - t0 & "ms" & ": length " & Len(s))
  TExt1.Text = s
Test_Nav_END:
  Exit Sub
Test_Nav_ERR:
  MsgBox Err.Description
  Resume
End Sub

Private Sub Command2_Click()
  Dim xt As New XMLTreeviewUnordered, sl As StringList, SLAttribs  As StringList
  Dim rs As Recordset, s As String
  Dim cmd As Command
  Dim cn As Connection
  Dim t0 As Long
  Dim i As Long
  Dim scn As String
  On Error GoTo Test_Nav_ERR
        
  'oracle
  scn = ADOOracleConnectString("APLUS.LONAP3004", "aplus_dev_ing", "ukcentral")
  Set cn = ADOConnect(scn, adUseClient)
  Set cmd = New Command
  Set cmd.ActiveConnection = cn
  cmd.CommandType = adCmdText
  cmd.CommandText = "SELECT id, parent, name, tooltip, children, editable, actual_value, nodeid, NULL AS default_value From (SELECT id, parent, name, tooltip, children, editable, actual_value, nodeid, NULL AS default_value, option_order From (SELECT 2 AS id, entity_pid AS nodeid, parent_id AS parent, entity_name AS name, entity_name AS tooltip, has_children AS children, entity_pid AS actual_value, 0 AS editable, entity_order As option_order From entity_values Union All SELECT 2 AS id, 0 AS nodeid, TO_NUMBER(NULL) AS parent, '(empty)'  AS name, TO_CHAR(NULL) AS tooltip, 0 AS children, TO_NUMBER(NULL) AS actual_value, 0 AS editable, 0 AS option_order FROM dual) ORDER BY UPPER(name))"
  Set rs = cmd.Execute
  t0 = GetTicks
  
  Set xt = New XMLTreeviewUnordered
  Set sl = New StringList
  Call sl.Add("actual_value")
  Call sl.Add("tooltip")
  Call sl.Add("name")
  s = xt.RecordsetToXML(rs, "", "nodeid", "parent", sl)
            
  Call MsgBox(GetTicks() - t0 & "ms" & ": length " & Len(s))
  TExt1.Text = s
Test_Nav_END:
  Exit Sub
Test_Nav_ERR:
  MsgBox Err.Description
  Resume
End Sub
