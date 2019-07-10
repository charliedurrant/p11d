VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{FBD9E481-6BAC-4BAB-9C6F-A979A270ACBD}#1.0#0"; "atc2hook.OCX"
Begin VB.Form frmListViewTest 
   Caption         =   "Test ListView"
   ClientHeight    =   5025
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8910
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5025
   ScaleWidth      =   8910
   Begin atc2hook.HOOK HOOK 
      Left            =   4890
      Top             =   495
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin MSComctlLib.ListView lvTest 
      Height          =   2295
      Left            =   120
      TabIndex        =   4
      Top             =   1320
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   4048
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      Enabled         =   0   'False
      NumItems        =   0
   End
   Begin VB.TextBox txtSortElements 
      Height          =   375
      Left            =   2880
      TabIndex        =   2
      Text            =   "10000"
      Top             =   120
      Width           =   795
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Fill List View"
      Height          =   390
      Left            =   240
      TabIndex        =   0
      Top             =   720
      Width           =   1620
   End
   Begin MSComCtl2.UpDown UDSort 
      Height          =   375
      Left            =   3840
      TabIndex        =   1
      Top             =   120
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   661
      _Version        =   393216
      Value           =   1
      BuddyControl    =   "txtSortElements"
      BuddyDispid     =   196609
      OrigLeft        =   3720
      OrigTop         =   120
      OrigRight       =   3960
      OrigBottom      =   540
      Increment       =   1000
      Max             =   100000
      Min             =   1
      SyncBuddy       =   -1  'True
      BuddyProperty   =   0
      Enabled         =   -1  'True
   End
   Begin VB.Label lblResults 
      Caption         =   "Results:"
      Height          =   1050
      Left            =   120
      TabIndex        =   5
      Top             =   3720
      Width           =   8475
   End
   Begin VB.Label lblText 
      Caption         =   "Number of ListView elements"
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   120
      Width           =   2535
   End
End
Attribute VB_Name = "frmListViewTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const WM_PAINT = &HF
Private mKillPaint As Boolean


Private Sub cmdStart_Click()
  Dim nElements As Long
  
  nElements = CLng(Me.txtSortElements)
  Call SetCursor
  Call Me.lvTest.ListItems.Clear
  Call Me.lvTest.ColumnHeaders.Clear
  Call Me.lvTest.ColumnHeaders.Add(, , "Column 1")
  Call Me.lvTest.ColumnHeaders.Add(, , "Column 2")
  Call Me.lvTest.ColumnHeaders.Add(, , "Column 3")
  Call Me.lvTest.ColumnHeaders.Add(, , "Column 4")
  Call Me.lvTest.ColumnHeaders.Add(, , "Column 5")
  Call Me.lvTest.ColumnHeaders.Add(, , "Column 6")
  Call Me.lvTest.ColumnHeaders.Add(, , "Column 7")
  Me.lvTest.Enabled = False
  Me.lvTest.Visible = False
  Me.lvTest.Sorted = False
  'mKillPaint = True
  Call FillListView(nElements)
 ' mKillPaint = False
  Me.lvTest.Enabled = True
  Me.lvTest.Visible = True
  Call ClearCursor
End Sub

Private Sub FillListView(ByVal nElements As Long)
  Dim t0 As Long, t1 As Long
  Dim li As ListItem, lis As ListItems, hWnd As Long, lissub As ListSubItem
  Dim i As Long, s As String, s2 As String, s3 As String
  
  t0 = GetTicks
  Set lis = Me.lvTest.ListItems
  hWnd = Me.lvTest.hWnd
  For i = 1 To nElements
    s = "Main Item " & CStr(i)
    s2 = "Sub Item 1: " & CStr(i)
    s3 = "Sub Item 2: " & CStr(i)
    Set li = lis.Add(, s)
    li.Text = s
'    Call cSetLVItem(hWnd, li.Index, 1, s2)
'    Call cSetLVItem(hWnd, li.Index, 2, s3)
'    Call cSetLVItem(hWnd, li.Index, 3, s3)
'    Call cSetLVItem(hWnd, li.Index, 4, s3)
'    Call cSetLVItem(hWnd, li.Index, 5, s3)
'    Call cSetLVItem(hWnd, li.Index, 6, s3)
    Call li.ListSubItems.Add(, , s2)
    Call li.ListSubItems.Add(, , s3)
    Call li.ListSubItems.Add(, , s3)
    Call li.ListSubItems.Add(, , s3)
    Call li.ListSubItems.Add(, , s3)
    Call li.ListSubItems.Add(, , s3)
    'li.SubItems(2) = s3
    'li.SubItems(3) = s3
    'li.SubItems(4) = s3
    'li.SubItems(5) = s3
    'li.SubItems(6) = s3
    li.Tag = i
  Next i
  t1 = GetTicks
  Me.lblResults = "Results:" & vbCrLf & "Fill List View: " & Format$((t1 - t0) / 1000, "#,##0.00") & " seconds."
  t0 = GetTicks
  Set lis = Me.lvTest.ListItems
  For i = 1 To nElements
    s = "Main Item " & CStr(i)
    s2 = "Sub Item 1: " & CStr(i)
    s3 = "Sub Item 2: " & CStr(i)
    Set li = lis.Item(i)
    li.Text = s
    li.SubItems(1) = s2 & s3
    li.SubItems(2) = s3 & s3
    li.SubItems(3) = s3 & s3
    li.SubItems(4) = s3 & s3
    li.SubItems(5) = s3 & s3
    li.SubItems(6) = s3 & s3
    li.Tag = i
  Next i
  t1 = GetTicks
  Me.lblResults = Me.lblResults & vbCrLf & "Already Filled List View: " & Format$((t1 - t0) / 1000, "#,##0.00") & " seconds."
  
  t0 = GetTicks
  Set lis = Me.lvTest.ListItems
  For i = lis.Count To 1 Step -1
    Call lis.Remove(i)
  Next i
  t1 = GetTicks
  Me.lblResults = Me.lblResults & vbCrLf & "Remove: " & Format$((t1 - t0) / 1000, "#,##0.00") & " seconds."
End Sub

Private Sub Form_Load()
  Me.HOOK.hWnd = Me.lvTest.hWnd
  HOOK.Messages(WM_PAINT) = True
End Sub

Private Sub Hook_WndProc(Discard As Boolean, MsgReturn As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long)
  If mKillPaint Then
    Discard = True
    MsgReturn = 0
  End If
End Sub
