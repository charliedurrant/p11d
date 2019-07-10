VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "atecAuth Tester"
   ClientHeight    =   8085
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6270
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8085
   ScaleWidth      =   6270
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdTestServers 
      Caption         =   "&Test Servers"
      Height          =   375
      Left            =   2040
      TabIndex        =   22
      Top             =   4440
      Width           =   1125
   End
   Begin VB.TextBox txtFilter 
      Height          =   315
      Left            =   1500
      TabIndex        =   20
      Top             =   1560
      Width           =   3915
   End
   Begin VB.TextBox txtLogon 
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   1500
      TabIndex        =   19
      Top             =   2745
      Width           =   3915
   End
   Begin VB.TextBox txtContainer 
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   1500
      TabIndex        =   15
      Top             =   1200
      Width           =   3915
   End
   Begin VB.CommandButton cmdEnumerate 
      Caption         =   "&Enumerate Users"
      Default         =   -1  'True
      Height          =   375
      Left            =   4560
      TabIndex        =   14
      Top             =   4440
      Width           =   1605
   End
   Begin VB.TextBox txtPassword2 
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   1500
      PasswordChar    =   "*"
      TabIndex        =   11
      Top             =   3120
      Width           =   3915
   End
   Begin VB.ComboBox cboUsername2 
      Height          =   315
      Left            =   1500
      TabIndex        =   10
      Top             =   2385
      Width           =   3915
   End
   Begin VB.ComboBox cboServer 
      Height          =   315
      Left            =   1500
      TabIndex        =   8
      Top             =   3600
      Width           =   3915
   End
   Begin VB.ComboBox cboServerContext 
      Height          =   315
      Left            =   1500
      TabIndex        =   6
      Top             =   3990
      Width           =   3915
   End
   Begin VB.ComboBox cboUsername 
      Height          =   315
      Left            =   1500
      TabIndex        =   0
      Top             =   270
      Width           =   3915
   End
   Begin VB.TextBox txtResults 
      Height          =   2955
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   3
      Top             =   4920
      Width           =   6075
   End
   Begin VB.CommandButton cmdTest 
      Caption         =   "&Test Auth"
      Height          =   375
      Left            =   3360
      TabIndex        =   2
      Top             =   4440
      Width           =   1125
   End
   Begin VB.TextBox txtPassword 
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   1500
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   720
      Width           =   3915
   End
   Begin VB.Label Label2 
      Caption         =   "Filter"
      Height          =   255
      Left            =   240
      TabIndex        =   21
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "logon name"
      Height          =   285
      Left            =   270
      TabIndex        =   18
      Top             =   2790
      Width           =   945
   End
   Begin VB.Label lblCount 
      Height          =   375
      Left            =   120
      TabIndex        =   17
      Top             =   4440
      Width           =   1455
   End
   Begin VB.Label lsldkfj 
      BackStyle       =   0  'Transparent
      Caption         =   "Container"
      Height          =   285
      Left            =   240
      TabIndex        =   16
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label lblUsername2 
      BackStyle       =   0  'Transparent
      Caption         =   "Username"
      Height          =   285
      Left            =   270
      TabIndex        =   13
      Top             =   2430
      Width           =   945
   End
   Begin VB.Label lblassword2 
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      Height          =   285
      Left            =   240
      TabIndex        =   12
      Top             =   3120
      Width           =   885
   End
   Begin VB.Label lblServer 
      BackStyle       =   0  'Transparent
      Caption         =   "Server"
      Height          =   285
      Left            =   240
      TabIndex        =   9
      Top             =   3600
      Width           =   945
   End
   Begin VB.Label lblServerContext 
      BackStyle       =   0  'Transparent
      Caption         =   "Server Context"
      Height          =   285
      Left            =   240
      TabIndex        =   7
      Top             =   4080
      Width           =   1215
   End
   Begin VB.Label lblPassword 
      BackStyle       =   0  'Transparent
      Caption         =   "Test Password"
      Height          =   285
      Left            =   240
      TabIndex        =   5
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label lblUsername 
      BackStyle       =   0  'Transparent
      Caption         =   "Test Username"
      Height          =   285
      Left            =   270
      TabIndex        =   4
      Top             =   330
      Width           =   1245
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mLDAPHelp As LDAPHelper2
Private Const AUTH_USER As String = "CN=Options Test,CN=Users,DC=uk,DC=deloitte,DC=com"
Private Const AUTH_PASSWORD As String = "0pt10n5t3sT"


Private Sub cboUsername_Click()
  txtPassword.SetFocus
End Sub

Private Sub cboUsername2_Change()
  If cboUsername2.Text = AUTH_USER Then
    txtPassword2.Text = AUTH_PASSWORD
  End If
  txtPassword2.SetFocus
End Sub

Private Function IsServerAlive(ByVal ServerName As String) As Long
  Dim Root As IADs
  Dim t0 As Long
  
  On Error GoTo IsServerAlive_Err
  t0 = GetTicks
  Set Root = GetObject("LDAP://" & ServerName)
  IsServerAlive = GetTicks - t0
  Exit Function
  
IsServerAlive_Err:
  IsServerAlive = -1
End Function

Private Sub cmdTestServers_Click()
  Dim t0 As Long, i As Long, j As Long
  Dim sCount As Long, ldapServers As Variant, ldapServerTimes As Variant
  Dim tryServer As String, smsg As String
  Const ST_TESTS As Long = 10

  Me.cmdTestServers.Enabled = False
  sCount = GetDelimitedValues(ldapServers, cboServerContext.Text, , , ";")
  ReDim ldapServerTimes(1 To sCount) As Long
  For j = 1 To ST_TESTS
    For i = 1 To sCount
      tryServer = ldapServers(i)
      t0 = IsServerAlive(tryServer)
      ldapServerTimes(i) = ldapServerTimes(i) + t0
    Next i
  Next j
  Me.cmdTestServers.Enabled = True
  smsg = "Test Server connect time" & vbCrLf & ST_TESTS & " successful attempts at connection to each server" & vbCrLf
  For i = 1 To sCount
    tryServer = ldapServers(i)
    smsg = smsg & "Server: " & tryServer & " Time: " & ldapServerTimes(i) / ST_TESTS & "ms" & vbCrLf
  Next i
  txtResults.Text = smsg
End Sub

Private Sub Form_Load()

  Me.Show

  Call cboUsername.Clear
  Call cboUsername.AddItem("mpsharpe")

  Call cboUsername2.Clear
  Call cboUsername2.AddItem(AUTH_USER)
  cboUsername2.Text = cboUsername2.List(0)

  Call cboServerContext.Clear
  Call cboServerContext.AddItem("uklondc004.uk.deloitte.com;uklondc003.uk.deloitte.com;uklondc002.uk.deloitte.com;uklondc001.uk.deloitte.com")
  cboServerContext.Text = cboServerContext.List(0)

  Me.txtContainer.Text = "OU=Tax,OU=London,DC=uk,DC=deloitte,DC=com"
  Me.txtLogon.Text = "optionstest"
  Call cboServer.Clear
  Call cboServer.AddItem("DC=uk,DC=deloitte,DC=com")
  cboServer.Text = cboServer.List(0)

  cboUsername.SetFocus

End Sub


Private Sub cmdTest_Click()
  cmdTest.Enabled = False
  Call Test
  cmdTest.Enabled = True
End Sub


Private Sub Test()
  Dim ServerContext As String
  Dim Server As String
  Dim Username As String, Password As String
  Dim AuthUsername As String, AuthPassword As String
  Dim UserPath As String
  Dim EmpRef As String
  Dim Authenticated As Boolean
  Dim rs As Recordset, rsField As Field
  Dim Value As String
  Dim s As String
  Dim col As Collection, lp As LDAPProperty
  Dim t0 As Long

  On Error GoTo Test_Err
  s = "START"
  ServerContext = cboServerContext.Text
  Server = cboServer.Text
  Username = cboUsername.Text
  Password = txtPassword.Text

  Set mLDAPHelp = New LDAPHelper2
  Call mLDAPHelp.SetServerContext(ServerContext, Server)

  mLDAPHelp.S_AUTH_USER = cboUsername2.Text
  mLDAPHelp.S_AUTH_USER_SAM = Me.txtLogon.Text
  mLDAPHelp.S_AUTH_PASSWORD = txtPassword2.Text
  mLDAPHelp.S_USERNAME_DOMAIN = "UK\"
  mLDAPHelp.S_EMPLOYEEID_DOMAIN = "UK"
  t0 = GetTicks
  Authenticated = mLDAPHelp.Authenticate(Username, Password, True, UserPath, EmpRef)
  t0 = GetTicks - t0
  If Authenticated Then
    s = s & vbCrLf & vbCrLf & "Authentication SUCCESS"
    EmpRef = mLDAPHelp.EmployeeIDNoDomain(EmpRef)
    s = s & vbCrLf & vbCrLf & "EmpRef=" & EmpRef
    If Len(mLDAPHelp.ServerRoot) > 0 Then
      Set rs = mLDAPHelp.UserProperties(Username, "employeeID", "telephoneNumber", "mail")
      If Not rs Is Nothing Then
        s = s & vbCrLf
        For Each rsField In rs.Fields
          s = s & vbCrLf & rsField.Name & "=" & rsField.Value
        Next rsField
        Set rs = Nothing
      End If
      s = s & "BEGIN ALL PROPERTIES" & vbCrLf
      Set col = mLDAPHelp.GetAllProperties(Username)
      For Each lp In col
        s = s & lp.Name & "="
        If lp.MultiValued Then
          s = s & lp.Values(0)
        Else
          s = s & lp.Values
        End If
        s = s & vbCrLf
      Next lp
      s = s & "END ALL PROPERTIES" & vbCrLf
    End If
  Else
    s = s & vbCrLf & vbCrLf & "Authentication FAILURE"
  End If

Test_End:
  s = s & vbCrLf & vbCrLf & "END"
  txtResults.Text = "Time to authenticate: " & t0 & "ms" & vbCrLf & s
  Exit Sub

Test_Err:
  s = s & vbCrLf & vbCrLf & "ERROR"
  s = s & vbCrLf & vbCrLf & Err.Description
  s = s & vbCrLf & vbCrLf & Err.Source
  txtResults.Text = s
  'Call MsgBox(Err.Description, vbCritical + vbOKOnly, "Error in Test")
  Resume Test_End
  Resume
End Sub

Private Sub cmdEnumerate_Click()
  Dim ec As EnumerateClass
  Dim ServerContext As String
  Dim Server As String
  Dim Username As String, Password As String
  Dim s As String
  
  On Error GoTo Test_Err
  txtResults.Text = "Processing .. "
  ServerContext = cboServerContext.Text
  Server = cboServer.Text
  Username = cboUsername.Text
  Password = txtPassword.Text
  
  Set mLDAPHelp = New LDAPHelper2
  Call mLDAPHelp.SetServerContext(ServerContext, Server)
  mLDAPHelp.S_AUTH_USER = cboUsername2.Text
  mLDAPHelp.S_AUTH_USER_SAM = Me.txtLogon.Text
  mLDAPHelp.S_AUTH_PASSWORD = txtPassword2.Text
  mLDAPHelp.S_USERNAME_DOMAIN = "UK\"
  mLDAPHelp.S_EMPLOYEEID_DOMAIN = "UK"
  Set ec = New EnumerateClass
  Set ec.mLDAPHelp = mLDAPHelp
  Call mLDAPHelp.EnumeratePeople(ec, Me.txtContainer.Text, Me.txtFilter.Text)

Test_End:
  Exit Sub

Test_Err:
  s = "ERROR"
  s = s & vbCrLf & vbCrLf & Err.Description
  s = s & vbCrLf & vbCrLf & Err.Source
  txtResults.Text = s
  'Call MsgBox(Err.Description, vbCritical + vbOKOnly, "Error in Test")
  Resume Test_End
  Resume

End Sub

