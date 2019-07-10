VERSION 5.00
Begin VB.Form Splash 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3900
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   8625
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   FontTransparent =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3900
   ScaleWidth      =   8625
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5850
      Left            =   15
      TabIndex        =   0
      Top             =   -240
      Width           =   8880
      Begin VB.PictureBox prgStartup 
         FillStyle       =   0  'Solid
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   285
         ScaleHeight     =   240
         ScaleWidth      =   7920
         TabIndex        =   7
         Top             =   3135
         Width           =   7980
      End
      Begin VB.PictureBox picDoors 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   930
         Left            =   240
         ScaleHeight     =   900
         ScaleWidth      =   900
         TabIndex        =   6
         Top             =   1440
         Width           =   930
      End
      Begin VB.Label lblDescription 
         BackColor       =   &H00000000&
         Caption         =   "lblProduct Description"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   750
         TabIndex        =   5
         Top             =   3675
         Width           =   6255
      End
      Begin VB.Line Line 
         BorderColor     =   &H00FFFFFF&
         X1              =   360
         X2              =   4320
         Y1              =   1320
         Y2              =   1320
      End
      Begin VB.Label lblMessage 
         BackColor       =   &H00000000&
         Caption         =   "lblMessage"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   720
         TabIndex        =   4
         Top             =   2520
         Width           =   4215
      End
      Begin VB.Label lblCompanyProduct 
         BackColor       =   &H00000000&
         Caption         =   "Deloitte && Touche"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   900
         Left            =   3240
         TabIndex        =   3
         Top             =   360
         Width           =   1800
      End
      Begin VB.Label lblProduct 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "lblProduct"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   420
         Left            =   3240
         TabIndex        =   2
         Top             =   1440
         Width           =   1755
      End
      Begin VB.Label lblVersion 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "lblVersion"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   345
         Left            =   6840
         TabIndex        =   1
         Top             =   2640
         Width           =   1395
      End
   End
End
Attribute VB_Name = "Splash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'* Do not change this code
Private Sub Form_Load()
  On Error Resume Next
  lblMessage = ""
  lblProduct = App.Title
  lblDescription = App.Comments
  lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
End Sub

Public Property Let Message(ByVal Msg As String)
  On Error Resume Next
  Me.lblMessage = Msg
End Property

Public Sub InitProgressBar()
  On Error Resume Next
  Me.prgStartup.Min = 0
  Me.prgStartup.Max = 10
  Me.prgStartup.Value = 1
End Sub

Public Sub IncrementProgressBar(Optional ByVal Finish As Boolean)
  On Error Resume Next
  If Finish Then
    Me.prgStartup.Value = 10
  ElseIf Me.prgStartup.Value < 10 Then
    Me.prgStartup.Value = Me.prgStartup.Value + 1
  End If
End Sub

Private Sub lblCompanyProduct_Click()

End Sub
