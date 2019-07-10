VERSION 5.00
Object = "{AF27A9B5-A3F4-11D2-8DB7-00C04FA9DD6F}#1.2#0"; "TCSPROG.OCX"
Begin VB.Form F_Progress 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "F_Progress"
   ClientHeight    =   1020
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5805
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1020
   ScaleWidth      =   5805
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   375
      Left            =   4545
      TabIndex        =   4
      Top             =   585
      Width           =   1185
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "&Start"
      Height          =   375
      Left            =   3240
      TabIndex        =   3
      Top             =   585
      Width           =   1230
   End
   Begin TCSPROG.TCSProgressBar prg 
      Height          =   375
      Left            =   135
      TabIndex        =   2
      Top             =   90
      Width           =   5595
      _cx             =   4204173
      _cy             =   4194965
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
      BarBackColor    =   12632256
      BarForeColor    =   8388608
      Appearance      =   1
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
      Enabled         =   0   'False
      Increment       =   1
      TextAlignment   =   1
   End
   Begin VB.CommandButton cmdView 
      Caption         =   "&View"
      Height          =   375
      Left            =   90
      TabIndex        =   1
      Top             =   585
      Width           =   1230
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      Height          =   375
      Left            =   1395
      TabIndex        =   0
      Top             =   585
      Width           =   1230
   End
End
Attribute VB_Name = "F_Progress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mIPRG As IProgress
Private mProgressType As Long
Property Set IPrg(ProgressType As Long, NewValue As IProgress)
  Dim pr As Boolean, vw As Boolean
  
  Set mIPRG = NewValue
  mProgressType = ProgressType
  pr = True
  vw = True
  
  If mIPRG.PreProgress(vw, pr, prg, mProgressType) Then
    cmdView.Visible = vw
    cmdPrint.Visible = pr
'    Me.Show vbModal
    Call p11d32.Help.ShowForm(Me, vbModal)
  End If
  
End Property
Public Sub Kill()
  Set mIPRG = Nothing
End Sub
Private Sub cmdClose_Click()
  Unload Me
End Sub

Private Sub cmdOK_Click()
  
End Sub

Private Sub cmdPrint_Click()
  Call mIPRG.PrintPrg
End Sub

Private Sub cmdStart_Click()
  Dim vw As Boolean, pr As Boolean
  
  If Not mIPRG Is Nothing Then
    vw = cmdView.Visible
    pr = cmdPrint.Visible
    If mIPRG.Progress(vw, pr, prg, mProgressType) Then
      cmdView.Visible = vw
      cmdPrint.Visible = pr
      Call mIPRG.PostProgress(vw, pr, prg, mProgressType)
    End If
  End If
  
End Sub


Private Sub cmdView_Click()
  Call mIPRG.ViewPrg
End Sub

