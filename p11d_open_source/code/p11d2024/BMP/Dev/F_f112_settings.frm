VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form F_f12_Settings 
   Caption         =   "Settings"
   ClientHeight    =   5445
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5535
   LinkTopic       =   "Form1"
   ScaleHeight     =   5445
   ScaleWidth      =   5535
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab tab 
      Height          =   5100
      Left            =   135
      TabIndex        =   0
      Top             =   135
      Width           =   5235
      _ExtentX        =   9234
      _ExtentY        =   8996
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      Tab             =   1
      TabHeight       =   520
      TabCaption(0)   =   "User Ini File"
      TabPicture(0)   =   "F_f112_settings.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "rtUserInitFile"
      Tab(0).Control(1)=   "lblUserIni"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Global Ini File"
      TabPicture(1)   =   "F_f112_settings.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "lblGlobalIniFile"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "rtGlobalIni"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      Begin RichTextLib.RichTextBox rtUserInitFile 
         Height          =   3930
         Left            =   -74910
         TabIndex        =   1
         Top             =   855
         Width           =   5010
         _ExtentX        =   8837
         _ExtentY        =   6932
         _Version        =   393217
         ScrollBars      =   3
         TextRTF         =   $"F_f112_settings.frx":0038
      End
      Begin RichTextLib.RichTextBox rtGlobalIni 
         Height          =   3975
         Left            =   90
         TabIndex        =   4
         Top             =   945
         Width           =   5010
         _ExtentX        =   8837
         _ExtentY        =   7011
         _Version        =   393217
         ScrollBars      =   3
         TextRTF         =   $"F_f112_settings.frx":00C3
      End
      Begin VB.Label lblUserIni 
         Caption         =   "lblUserIni"
         Height          =   420
         Left            =   -74865
         TabIndex        =   3
         Top             =   405
         Width           =   3705
      End
      Begin VB.Label lblGlobalIniFile 
         Caption         =   "lblGlobalIniFile"
         Height          =   510
         Left            =   135
         TabIndex        =   2
         Top             =   360
         Width           =   5010
      End
   End
End
Attribute VB_Name = "F_f12_Settings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
  Dim s As String
  
  On Error GoTo err_Err
  
  lblUserIni = p11d32.IniPathAndFile
  rtUserInitFile.Text = TextFileLoad(p11d32.IniPathAndFile)
  If (FileExists(p11d32.IniPathAndFileGlobal)) Then
    s = TextFileLoad(p11d32.IniPathAndFileGlobal)
  Else
    s = "No Global File"
  End If
  lblGlobalIniFile = p11d32.IniPathAndFileGlobal
  rtGlobalIni.Text = s
  
  
err_End:
  Exit Sub
err_Err:
  Call ErrorMessage(ERR_ERROR, Err, "Load", "Load", Err.Description)
  Resume err_End
End Sub

