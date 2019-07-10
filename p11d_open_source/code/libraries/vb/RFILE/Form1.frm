VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4275
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   4275
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Bye Bye"
      Height          =   495
      Left            =   1440
      TabIndex        =   1
      Top             =   3240
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&AddFile"
      Height          =   615
      Left            =   1320
      TabIndex        =   0
      Top             =   2160
      Width           =   1935
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open"
      End
      Begin VB.Menu mnuRecentFile 
         Caption         =   "-"
         Index           =   0
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Implements IRecentFileList

Public RFL As RecentFileList

Private Sub Command1_Click()
  Call RFL.AddRecentFile("c:\pm65\extras\text\dir3\deapp.exe")
  Call RFL.AddRecentFile("c:\program files\devstudio\vb5\dir\tcsaddin\addproc.dll")
  Call RFL.AddRecentFile("i:\tax\abacus16\dir2\abacus.exe")
  Call RFL.AddRecentFile("\\lonfs3002\vol1\aa\data\dir1\abacus\abwin32d.exe")
  Call RFL.AddRecentFile("\\lonfs3003\vol1\aa\apps\tax\dir\cortax\cortest.hlp")
End Sub

Private Sub Form_Load()
  Set RFL = New RecentFileList
  Call RFL.Setup(Form1.mnuRecentFile, "Project1", Me, 5, 30)
End Sub

Private Function IRecentFileList_Validate(ByVal FileName As String) As TCSRFILE.RFL_ACTION
 'do out testing on file return TCSRFILE.RFL_ACTION
 IRecentFileList_Validate = RFL_OK
 
End Function

Private Sub mnuRecentFile_Click(Index As Integer)
  Call RFL.RecentFileClick(Index)
End Sub

Private Sub command2_click()
Unload Me
End Sub
