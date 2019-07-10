VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form F_Goto 
   Caption         =   "Goto Employee"
   ClientHeight    =   3555
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8445
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3555
   ScaleWidth      =   8445
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton B_Ok 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   350
      Left            =   6120
      TabIndex        =   1
      Tag             =   "LOCKBR"
      Top             =   3195
      Width           =   1095
   End
   Begin VB.CommandButton B_Cancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   350
      Left            =   7290
      TabIndex        =   2
      Tag             =   "LOCKBR"
      Top             =   3195
      Width           =   1095
   End
   Begin MSComctlLib.ListView lb 
      Height          =   3105
      Left            =   45
      TabIndex        =   0
      Tag             =   "EQUALISE"
      Top             =   45
      Width           =   8355
      _ExtentX        =   14737
      _ExtentY        =   5477
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
End
Attribute VB_Name = "F_Goto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public NewEmployeeListItem As ListItem
Private mclsResize As New clsFormResize
Private Const L_DES_HEIGHT  As Long = 4110
Private Const L_DES_WIDTH  As Long = 8565


Private m_LastSearch As String
Private Sub B_Cancel_Click()
  Set NewEmployeeListItem = Nothing
  Me.Hide
End Sub
Private Sub B_OK_Click()
  Me.Hide
End Sub

Private Sub Form_Load()
  If Not (mclsResize.InitResize(Me, L_DES_HEIGHT, L_DES_WIDTH, DESIGN, , , MDIMain)) Then ', DESIGN)) Then
    Err.Raise ERR_Application
  End If
End Sub

Private Sub Form_Resize()
  Call mclsResize.Resize
  Call ColumnWidths(Me.lb, L_NAME_COL, L_REFERENCE_COL, L_NINUMBER_COL&, L_STATUS_COL&, L_GROUP1_COL&, L_GROUP2_COL&, L_GROUP3_COL&)
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Set NewEmployeeListItem = Nothing
End Sub

Private Sub lb_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
  Call SetSortOrderEmployees(lb, ColumnHeader)
End Sub

Private Sub lb_DblClick()
  If Not lb.SelectedItem Is Nothing Then
    B_Ok.value = True
  End If
End Sub

Private Sub LB_ItemClick(ByVal Item As MSComctlLib.ListItem)
  Set NewEmployeeListItem = Item
End Sub

Private Sub LB_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF3 Then
    Call FoundEmployee(KeyCode, 0) 'EK 1/04 TTP#12
    KeyCode = 0
  End If
End Sub
Private Function FoundEmployee(ByVal KeyCode As Integer, ByVal KeyAscii As Integer) As Boolean
  On Error GoTo FoundEmployee_ERR
  
  If ListViewFastKey(lb, p11d32.EmployeeSortOrderColumn, KeyCode, KeyAscii, m_LastSearch) > 0 Then
    Set NewEmployeeListItem = lb.SelectedItem
 End If
  
FoundEmployee_END:

  Exit Function
FoundEmployee_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "Found Employee", "Found Employee", "Error finding employees from employee screen.")
  Resume FoundEmployee_END
End Function

Private Sub lb_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    B_Ok.value = True
  Else
    Call FoundEmployee(0, KeyAscii) 'EK 1/04 TTP#12
    KeyAscii = 0
  End If
End Sub

