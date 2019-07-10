VERSION 5.00
Begin VB.Form frmExportToXML 
   Caption         =   "Export DB to XML"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   Begin VB.CommandButton cmdExport 
      Caption         =   "Export"
      Height          =   375
      Left            =   3000
      TabIndex        =   4
      Top             =   1440
      Width           =   1575
   End
   Begin VB.TextBox txtOutputPath 
      Height          =   375
      Left            =   1320
      TabIndex        =   2
      Top             =   840
      Width           =   3255
   End
   Begin VB.CommandButton cmdDB 
      Caption         =   "Database"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Output to:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label lblDBPath 
      Height          =   375
      Left            =   1440
      TabIndex        =   1
      Top             =   120
      Width           =   3135
   End
End
Attribute VB_Name = "frmExportToXML"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mDBPath As String

Private Sub cmdDB_Click()
  Dim s As String, sdir As String
  
  If Len(mDBPath) = 0 Then
    sdir = AppPath
  Else
    Call SplitPath(mDBPath, sdir)
  End If
  mDBPath = ""
  s = FileOpenDlg("Choose Database", "Access databases (*.mdb)|*.mdb|All files (*.*)|*.*", sdir)
  If Len(s) > 0 Then mDBPath = s
End Sub

Private Sub RefreshPath()
  Me.lblDBPath = mDBPath
End Sub

Private Sub cmdExport_Click()
  Call ExportDBtoXML(ADOAccessConnectString(mDBPath), Me.txtOutputPath)
End Sub
