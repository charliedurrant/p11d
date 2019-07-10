VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7095
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   16965
   LinkTopic       =   "Form1"
   ScaleHeight     =   7095
   ScaleWidth      =   16965
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Exec"
      Height          =   825
      Left            =   585
      TabIndex        =   5
      Top             =   3105
      Width           =   1770
   End
   Begin VB.TextBox Text2 
      Height          =   690
      Left            =   2025
      TabIndex        =   3
      Top             =   2250
      Width           =   14370
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Connect"
      Height          =   825
      Left            =   405
      TabIndex        =   1
      Top             =   1170
      Width           =   1770
   End
   Begin VB.TextBox Text1 
      Height          =   690
      Left            =   1890
      TabIndex        =   0
      Top             =   135
      Width           =   14370
   End
   Begin VB.Label Label2 
      Caption         =   "SQL "
      Height          =   375
      Left            =   225
      TabIndex        =   4
      Top             =   2295
      Width           =   1725
   End
   Begin VB.Label Label1 
      Caption         =   "Connection string"
      Height          =   375
      Left            =   405
      TabIndex        =   2
      Top             =   90
      Width           =   2040
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
  On Error GoTo err_err
  
  Dim con As ADODb.Connection
  
  Set con = New ADODb.Connection
  
  con.CommandTimeout = 2
   
  con.ConnectionString = Text1.Text
  Call con.Open
  
  Call MsgBox("Connected")
  
  Call con.Close
  
  Exit Sub
    

  
  
  
err_err:
  MsgBox Err.Description
  
  
End Sub

Private Sub Command2_Click()
  On Error GoTo err_err
  
  Dim con As ADODb.Connection
  
  Set con = New ADODb.Connection
  
  con.CommandTimeout = 2
   
  con.ConnectionString = Text1.Text
  Call con.Open
  
  Dim cmd As ADODb.Command
  
  Set cmd = New ADODb.Command
  cmd.CommandText = Text2.Text
  Dim rs As ADODb.Recordset
  Set rs = cmd.Execute()
  Call MsgBox("record count: " & rs.RecordCount)
  Call con.Close
  
  
  
   

  
  
  Exit Sub
    

  
  
  
err_err:
  MsgBox Err.Description
  
End Sub
