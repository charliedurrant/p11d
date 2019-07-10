VERSION 5.00
Object = "{AF27A9B5-A3F4-11D2-8DB7-00C04FA9DD6F}#1.2#0"; "Tcsprog.ocx"
Object = "{4BA5AE86-C9BA-4B77-8E15-D04582204FDD}#1.0#0"; "atc2stat.OCX"
Begin VB.Form frmStatus 
   Caption         =   "Form1"
   ClientHeight    =   4620
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8910
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4620
   ScaleWidth      =   8910
   Begin atc2Stat.TCSStatus TCSStatus1 
      Align           =   2  'Align Bottom
      Height          =   645
      Left            =   0
      TabIndex        =   2
      Top             =   3975
      Width           =   8910
      _ExtentX        =   15716
      _ExtentY        =   1138
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin TCSPROG.TCSProgressBar TCSProgressBar1 
      Height          =   1455
      Left            =   1890
      TabIndex        =   1
      Top             =   1485
      Width           =   5655
      _cx             =   9975
      _cy             =   2566
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
      Value           =   50
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
      Percentage      =   0
      Skew            =   0
      PictureOffsetTop=   0
      PictureOffsetLeft=   0
      Enabled         =   -1  'True
      Increment       =   1
      TextAlignment   =   1
   End
   Begin VB.CommandButton cmdStatus 
      Caption         =   "Status bar"
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   2055
   End
End
Attribute VB_Name = "frmStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdStatus_Click()
  Dim i As Long
  frmMain.Status1.prg.Min = 0
  frmMain.Status1.prg.Max = 1
  For i = 0 To 1
    Call frmMain.Status1.StepCaption("Status: " & i)
    Call Sleep(5)
  Next i
  frmMain.Status1.prg.Min = 1
'  frmMain.Status1.prg.Max = 1
'  frmMain.Status1.prg.Value = 1
  Call frmMain.Status1.StopPrg
End Sub

