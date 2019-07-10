VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmConnection 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Connection"
   ClientHeight    =   3795
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4875
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3795
   ScaleWidth      =   4875
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtConnectionString 
      BackColor       =   &H8000000F&
      Height          =   735
      Left            =   90
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   2
      Top             =   2520
      Width           =   4695
   End
   Begin TabDlg.SSTab tabSource 
      Height          =   2175
      Left            =   90
      TabIndex        =   1
      Top             =   90
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   3836
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      TabCaption(0)   =   "Access"
      TabPicture(0)   =   "frmConnection.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblPasswordAccess"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblUserNameAccess"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblFileAccess"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblFileSystemDBAccess"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "txtFileAccess"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmdFileOpenAccess"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "txtPasswordAccess"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "txtUsernameAccess"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "txtFileSystemDBAccess"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "cmdFileOpenAccessSystem"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).ControlCount=   10
      TabCaption(1)   =   "SQL Server"
      TabPicture(1)   =   "frmConnection.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txtPasswordSQL"
      Tab(1).Control(1)=   "txtUserNameSQL"
      Tab(1).Control(2)=   "txtDataBaseSQL"
      Tab(1).Control(3)=   "txtServerSQL"
      Tab(1).Control(4)=   "lblPassWordSQL"
      Tab(1).Control(5)=   "lblUserNameSQL"
      Tab(1).Control(6)=   "lblDataBaseSQL"
      Tab(1).Control(7)=   "lblServerSQL"
      Tab(1).ControlCount=   8
      TabCaption(2)   =   "Oracle"
      TabPicture(2)   =   "frmConnection.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "txtPasswordOracle"
      Tab(2).Control(1)=   "txtUserNameOracle"
      Tab(2).Control(2)=   "txtDataSourceOracle"
      Tab(2).Control(3)=   "lblPasswordOracle"
      Tab(2).Control(4)=   "lblUserNameOracle"
      Tab(2).Control(5)=   "lblDataSource"
      Tab(2).ControlCount=   6
      Begin VB.CommandButton cmdFileOpenAccessSystem 
         Height          =   330
         Left            =   4050
         Picture         =   "frmConnection.frx":0054
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   1080
         Width           =   465
      End
      Begin VB.TextBox txtFileSystemDBAccess 
         Height          =   330
         Left            =   90
         TabIndex        =   24
         Top             =   1080
         Width           =   3930
      End
      Begin VB.TextBox txtUsernameAccess 
         Height          =   330
         Left            =   90
         TabIndex        =   23
         Top             =   1710
         Width           =   1590
      End
      Begin VB.TextBox txtPasswordAccess 
         Height          =   330
         IMEMode         =   3  'DISABLE
         Left            =   1980
         PasswordChar    =   "*"
         TabIndex        =   21
         Top             =   1710
         Width           =   1905
      End
      Begin VB.TextBox txtPasswordOracle 
         Height          =   330
         IMEMode         =   3  'DISABLE
         Left            =   -72795
         PasswordChar    =   "*"
         TabIndex        =   18
         Top             =   1350
         Width           =   2085
      End
      Begin VB.TextBox txtUserNameOracle 
         Height          =   330
         Left            =   -74865
         TabIndex        =   16
         Top             =   1350
         Width           =   1815
      End
      Begin VB.TextBox txtDataSourceOracle 
         Height          =   330
         Left            =   -74865
         TabIndex        =   15
         Top             =   675
         Width           =   1815
      End
      Begin VB.TextBox txtPasswordSQL 
         Height          =   330
         IMEMode         =   3  'DISABLE
         Left            =   -72795
         PasswordChar    =   "*"
         TabIndex        =   12
         Top             =   1350
         Width           =   2085
      End
      Begin VB.TextBox txtUserNameSQL 
         Height          =   330
         Left            =   -74865
         TabIndex        =   10
         Top             =   1350
         Width           =   1815
      End
      Begin VB.TextBox txtDataBaseSQL 
         Height          =   330
         Left            =   -72795
         TabIndex        =   8
         Top             =   675
         Width           =   2085
      End
      Begin VB.TextBox txtServerSQL 
         Height          =   330
         Left            =   -74865
         TabIndex        =   7
         Top             =   675
         Width           =   1815
      End
      Begin VB.CommandButton cmdFileOpenAccess 
         Height          =   330
         Left            =   4005
         Picture         =   "frmConnection.frx":0596
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   540
         Width           =   465
      End
      Begin VB.TextBox txtFileAccess 
         Height          =   330
         Left            =   90
         TabIndex        =   3
         Top             =   540
         Width           =   3930
      End
      Begin VB.Label lblFileSystemDBAccess 
         Caption         =   "System db"
         Height          =   195
         Left            =   90
         TabIndex        =   27
         Top             =   900
         Width           =   1275
      End
      Begin VB.Label lblFileAccess 
         Caption         =   "Database"
         Height          =   195
         Left            =   90
         TabIndex        =   26
         Top             =   360
         Width           =   1140
      End
      Begin VB.Label lblUserNameAccess 
         Caption         =   "Username"
         Height          =   240
         Left            =   90
         TabIndex        =   22
         Top             =   1485
         Width           =   1230
      End
      Begin VB.Label lblPasswordAccess 
         Caption         =   "Password"
         Height          =   195
         Left            =   1980
         TabIndex        =   20
         Top             =   1485
         Width           =   1140
      End
      Begin VB.Label lblPasswordOracle 
         Caption         =   "Password"
         Height          =   240
         Left            =   -72795
         TabIndex        =   19
         Top             =   1125
         Width           =   1095
      End
      Begin VB.Label lblUserNameOracle 
         Caption         =   "User name"
         Height          =   240
         Left            =   -74865
         TabIndex        =   17
         Top             =   1125
         Width           =   1725
      End
      Begin VB.Label lblDataSource 
         Caption         =   "Data source"
         Height          =   240
         Left            =   -74865
         TabIndex        =   14
         Top             =   450
         Width           =   1005
      End
      Begin VB.Label lblPassWordSQL 
         Caption         =   "Password"
         Height          =   240
         Left            =   -72795
         TabIndex        =   13
         Top             =   1125
         Width           =   1095
      End
      Begin VB.Label lblUserNameSQL 
         Caption         =   "User name"
         Height          =   240
         Left            =   -74865
         TabIndex        =   11
         Top             =   1125
         Width           =   1725
      End
      Begin VB.Label lblDataBaseSQL 
         Caption         =   "Database"
         Height          =   240
         Left            =   -72795
         TabIndex        =   9
         Top             =   450
         Width           =   1185
      End
      Begin VB.Label lblServerSQL 
         Caption         =   "Server"
         Height          =   240
         Left            =   -74820
         TabIndex        =   6
         Top             =   450
         Width           =   1005
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3690
      TabIndex        =   0
      Top             =   3375
      Width           =   1050
   End
   Begin VB.Label lblConnectionString 
      Caption         =   "Connection string"
      Height          =   195
      Left            =   90
      TabIndex        =   5
      Top             =   2295
      Width           =   1905
   End
End
Attribute VB_Name = "frmConnection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_SQLADO As SQLDebugADO
Public Function ShowConnection(Target As DATABASE_TARGET, sCurrentConnectionString As String, ByVal SQLADO As SQLDebugADO) As String
  Dim v As Variant
  Dim i As Long
  Dim j As Long
  Dim sName As String, sValue As String
  
  On Error GoTo ShowConnection_ERR
  
  frmConnection.tabSource.Tab = Target
  
  Set m_SQLADO = SQLADO
  
  frmConnection.txtConnectionString = sCurrentConnectionString
  i = GetDelimitedValues(v, sCurrentConnectionString, , , ";")
  
  For j = 1 To i
    If Not GetEqualsValue(sName, sValue, CStr(v(j))) Then GoTo NEXT_ITEM
    Select Case UCase$(sName)
      Case "DATA SOURCE"
        Select Case Target
          Case DB_TARGET_JET
            txtFileAccess = sValue
          Case DB_TARGET_ORACLE
            txtDataSourceOracle = sValue
          Case DB_TARGET_SQLSERVER
            txtServerSQL = sValue
        End Select
      Case "INITIAL CATALOG"
        If Target = DB_TARGET_SQLSERVER Then txtDataBaseSQL = sValue
      Case "USER ID"
        Select Case Target
          Case DB_TARGET_JET
            txtUsernameAccess = sValue
          Case DB_TARGET_ORACLE
            txtUserNameOracle = sValue
          Case DB_TARGET_SQLSERVER
            txtUserNameSQL = sValue
        End Select
      Case "PASSWORD"
        Select Case Target
          Case DB_TARGET_JET
            txtPasswordAccess = sValue
          Case DB_TARGET_ORACLE
            txtPasswordOracle = sValue
          Case DB_TARGET_SQLSERVER
            txtPasswordSQL = sValue
        End Select
      Case "SYSTEM DATABASE"
        If Target = DB_TARGET_SQLSERVER Then txtFileSystemDBAccess = sValue
    End Select
    
    
NEXT_ITEM:
  Next
  Select Case Target
    
  End Select
  txtConnectionString = sCurrentConnectionString
  Me.Show 1
  Target = Me.tabSource.Tab
  ShowConnection = txtConnectionString
  Unload Me
ShowConnection_END:
  Exit Function
ShowConnection_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "ShowConnection", "ShowConnection", "Error in ShowConnection")
  Resume ShowConnection_END
End Function
Private Function GetEqualsValue(sName As String, sValue As String, sString As String) As Boolean
  Dim i As Long
  sValue = ""
  sName = ""
  
  i = InStr(1, sString, "=", vbTextCompare)
  If i > 1 And i < Len(sString) Then
    sName = Left$(sString, i - 1)
    sValue = Right$(sString, Len(sString) - i)
  End If
  GetEqualsValue = Len(sValue) > 0
  
End Function
Private Sub cmdFileOpenAccess_Click()
  Dim sFileAndPath As String
  
  On Error GoTo OpenFile_ERR
  
  sFileAndPath = FileOpenDlg(m_SQLADO.FileCaption, m_SQLADO.FileFilter, m_SQLADO.CurrentDirectory)
  If Len(sFileAndPath) > 0 Then
    txtFileAccess = sFileAndPath
    Call SplitPath(sFileAndPath, sFileAndPath)
    m_SQLADO.CurrentDirectory = sFileAndPath
  End If
  
OpenFile_END:
  Exit Sub
  
OpenFile_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "OpenFile", "Open File", "Error opening a file.")
  Resume OpenFile_END
  
End Sub

Private Sub tab_DblClick()

End Sub

Private Sub SQlConnectChange()
  txtConnectionString = ADOSQLConnectString(txtServerSQL, txtUserNameSQL, txtPasswordSQL, txtDataBaseSQL)
  
End Sub
Private Function SubstitutePassword(ByVal txt As TextBox) As String
  Dim s As String
  SubstitutePassword = String$(Len(txt), "*")
End Function
  
Private Sub cmdFileOpenAccessSystem_Click()
  Dim sFileAndPath As String
  
  On Error GoTo OpenFile_ERR
  
  sFileAndPath = FileOpenDlg(m_SQLADO.FileCaptionSystem, m_SQLADO.FileFilterSystem, m_SQLADO.CurrentDirectory)
  If Len(sFileAndPath) > 0 Then
    txtFileSystemDBAccess = sFileAndPath
    Call SplitPath(sFileAndPath, sFileAndPath)
    m_SQLADO.CurrentDirectory = sFileAndPath
  End If
  
OpenFile_END:
  Exit Sub
  
OpenFile_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "OpenFile", "Open File", "Error opening a file.")
  Resume OpenFile_END
End Sub

Private Sub cmdOK_Click()
  Me.Hide
  
End Sub

Private Sub txtConnectionString_KeyDown(KeyCode As Integer, Shift As Integer)
  KeyCode = 0
End Sub

Private Sub txtDataBaseSQL_Change()
  Call SQlConnectChange
End Sub
Private Sub AccessConnectChange()
  txtConnectionString = ADOAccess4ConnectString(txtFileAccess, txtUsernameAccess, txtPasswordAccess, txtDataBaseSQL)
  
End Sub
Private Sub OracleConnectChange()
  txtConnectionString = ADOOracleConnectString(txtDataSourceOracle, txtUserNameOracle, txtPasswordOracle)
  
End Sub
Private Sub txtDataSourceOracle_Change()
  Call OracleConnectChange
End Sub

Private Sub txtFileAccess_Change()
  Call AccessConnectChange
End Sub

Private Sub txtFileSystemDBAccess_Change()
  Call AccessConnectChange
End Sub

Private Sub txtPasswordAccess_Change()
  Call AccessConnectChange
End Sub

Private Sub txtPasswordOracle_Change()
  Call OracleConnectChange
End Sub

Private Sub txtPasswordSQL_Change()
  Call SQlConnectChange
End Sub

Private Sub txtServerSQL_Change()
  Call SQlConnectChange
End Sub

Private Sub txtUsernameAccess_Change()
  Call AccessConnectChange
End Sub

Private Sub txtUserNameOracle_Change()
  Call OracleConnectChange
End Sub

Private Sub txtUserNameSQL_Change()
  Call SQlConnectChange
End Sub
