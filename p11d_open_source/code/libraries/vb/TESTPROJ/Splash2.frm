VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{AF27A9B5-A3F4-11D2-8DB7-00C04FA9DD6F}#1.2#0"; "tcsprog.ocx"
Begin VB.Form Splash2 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   7305
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   9195
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   FontTransparent =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7305
   ScaleWidth      =   9195
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3930
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9720
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1410
         Left            =   1500
         Picture         =   "Splash2.frx":0000
         ScaleHeight     =   1410
         ScaleWidth      =   6855
         TabIndex        =   8
         Top             =   -90
         Width           =   6855
      End
      Begin VB.Frame Frame2 
         Appearance      =   0  'Flat
         BackColor       =   &H00996699&
         BorderStyle     =   0  'None
         ForeColor       =   &H00FFFFFF&
         Height          =   1845
         Left            =   75
         TabIndex        =   3
         Top             =   2010
         Width           =   8505
         Begin TCSPROG.TCSProgressBar prgStartup 
            Height          =   300
            Left            =   240
            TabIndex        =   6
            Top             =   1035
            Width           =   8040
            _cx             =   14182
            _cy             =   529
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Min             =   0
            Max             =   100
            Value           =   0
            BarBackColor    =   16777215
            BarForeColor    =   33023
            Appearance      =   2
            Style           =   0
            CaptionColor    =   0
            CaptionInvertColor=   16777215
            FillStyle       =   0
            FadeFromColor   =   0
            FadeToColor     =   16777215
            Caption         =   ""
            InnerCircle     =   0   'False
            Percentage      =   2
            Skew            =   0
            PictureOffsetTop=   0
            PictureOffsetLeft=   0
            Enabled         =   -1  'True
            Increment       =   1
            TextAlignment   =   1
         End
         Begin VB.Label lblMessage 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "lblMessage"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   450
            Left            =   270
            TabIndex        =   7
            Top             =   630
            Width           =   6270
         End
         Begin VB.Label lblVersion 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "lblVersion"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   330
            Left            =   5550
            TabIndex        =   5
            Top             =   1485
            Width           =   2730
         End
         Begin VB.Label lblProduct 
            BackStyle       =   0  'Transparent
            Caption         =   "lblProduct"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   570
            Left            =   240
            TabIndex        =   4
            Top             =   120
            Width           =   4815
         End
      End
      Begin VB.Label lblSpacer 
         Height          =   450
         Left            =   2985
         TabIndex        =   2
         Top             =   1560
         Width           =   5595
      End
      Begin VB.Label lblTax 
         BackColor       =   &H000080FF&
         Caption         =   " Tax Services"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   450
         Left            =   75
         TabIndex        =   1
         Top             =   1560
         Width           =   2910
      End
      Begin VB.Line Line 
         BorderColor     =   &H00FFFFFF&
         X1              =   360
         X2              =   4320
         Y1              =   1320
         Y2              =   1320
      End
   End
   Begin MSComctlLib.ImageList iml 
      Left            =   1560
      Top             =   5520
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   45
      ImageHeight     =   47
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Splash2.frx":117F
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Splash2.frx":1EA1
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "Splash2"
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
  lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
End Sub

Public Property Let Message(ByVal msg As String)
  On Error Resume Next
  Me.lblMessage = msg
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

