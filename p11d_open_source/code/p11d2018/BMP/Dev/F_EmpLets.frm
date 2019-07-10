VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form F_EmployeeLetterSaveAs 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Save As..."
   ClientHeight    =   3000
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3840
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3000
   ScaleWidth      =   3840
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2655
      TabIndex        =   3
      Top             =   2520
      Width           =   1095
   End
   Begin VB.TextBox txtFile 
      Height          =   285
      Left            =   45
      TabIndex        =   2
      Top             =   2115
      Width           =   3705
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   1485
      TabIndex        =   1
      Top             =   2520
      Width           =   1095
   End
   Begin MSComctlLib.ListView lvfile 
      Height          =   2040
      Left            =   45
      TabIndex        =   0
      Top             =   45
      Width           =   3750
      _ExtentX        =   6615
      _ExtentY        =   3598
      View            =   2
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Letter"
         Object.Width           =   15875
      EndProperty
   End
End
Attribute VB_Name = "F_EmployeeLetterSaveAs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public FileName As String
Public NewFile As Boolean
Public OK As Boolean

Private Sub cmdCancel_Click()
  OK = False
  Me.Hide
End Sub

Private Sub cmdOK_Click()
  If ValidateFile(txtFile.Text) Then
    FileName = txtFile.Text & S_EMPLOYEE_LETTER_FILE_EXTENSION
    Me.Hide
    OK = True
  End If
End Sub

Private Function ValidateFile(ByVal sFile As String) As Boolean
  On Error GoTo ValidateFile_ERR
  
  Call xSet("ValidateFile")
  
  sFile = sFile & S_EMPLOYEE_LETTER_FILE_EXTENSION
  
  If P11d32.IsMasterLetterFile(sFile) Then
    Call ErrorMessage(ERR_ERROR, Err, "ValidateFile", "Validate File", "The file name is the same as the master file, please use another.")
    GoTo ValidateFile_END
  End If
  
  If StrComp(sFile, P11d32.EmployeeLetterFile) <> 0 Then
    If Not IsFileOpen(P11d32.EmployeeLetterPath & FileName) Then
      ValidateFile = True
      NewFile = True
    Else
      Call Err.Raise(ERR_FILE_OPEN, "ValidateFile", "Invalid file name, a file with the same name is already open.")
    End If
  Else
    NewFile = True
    ValidateFile = True
  End If
  
ValidateFile_END:
  Call xReturn("ValidateFile")
  Exit Function
ValidateFile_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "ValidateFile", "Validate File", "Error validating the new file name.")
  Resume ValidateFile_END
End Function

Public Sub FillFileList(ByVal sCurrentFileName As String)
  Dim sFiles() As String
  Dim sFile As String
  Dim i As Long
  
  
  If P11d32.GetLetterFiles(sFiles) Then
    For i = LBound(sFiles) To UBound(sFiles)
      If Not P11d32.IsMasterLetterFile(sFiles(i)) Then
        Call SplitPath(sFiles(i), , sFile)
        Call lvfile.ListItems.Add(, , sFile)
      End If
    Next
  End If
  
  If P11d32.IsMasterLetterFile(P11d32.EmployeeLetterPath & sCurrentFileName & S_EMPLOYEE_LETTER_FILE_EXTENSION) Then
    sCurrentFileName = "New file"
  End If
  
  txtFile = sCurrentFileName
  txtFile.SelStart = 0
  txtFile.SelLength = Len(sCurrentFileName)
  
End Sub
  
Private Sub lvfile_DblClick()
  If Not lvfile.SelectedItem Is Nothing Then txtFile = lvfile.SelectedItem.Text
  cmdOK.value = True
End Sub

Private Sub lvfile_ItemClick(ByVal Item As MSComctlLib.ListItem)
  txtFile = Item.Text
End Sub

Private Sub txtFile_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = Asc("\") Or KeyCode = vbKeyDivide Or KeyCode = vbKeyDecimal Or KeyCode = vbKeySpace Then
    KeyCode = 0
  End If
End Sub

