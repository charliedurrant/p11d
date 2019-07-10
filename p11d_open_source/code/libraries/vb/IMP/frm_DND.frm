VERSION 5.00
Object = "{AF27A9B5-A3F4-11D2-8DB7-00C04FA9DD6F}#1.2#0"; "tcsprog.ocx"
Begin VB.Form frm_DoNotDelete 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin TCSPROG.TCSProgressBar TCSProgressBar1 
      Height          =   780
      Left            =   360
      TabIndex        =   0
      Top             =   270
      Width           =   3345
      _cx             =   5900
      _cy             =   1376
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
      BarBackColor    =   -2147483633
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
      TextAlignment   =   2
   End
End
Attribute VB_Name = "frm_DoNotDelete"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

