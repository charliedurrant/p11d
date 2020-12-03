VERSION 5.00
Object = "{8D988532-0F0C-460C-B00E-7B5637E97680}#1.0#0"; "atc2vtext.ocx"
Begin VB.Form F_Input 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Caption"
   ClientHeight    =   1185
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3195
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1185
   ScaleWidth      =   3195
   StartUpPosition =   1  'CenterOwner
   Begin atc2valtext.ValText ValText 
      Height          =   330
      Left            =   45
      TabIndex        =   0
      Top             =   405
      Width           =   3120
      _ExtentX        =   5503
      _ExtentY        =   582
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   ""
      TXTAlign        =   2
      AutoSelect      =   0
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   330
      Left            =   2115
      TabIndex        =   2
      Top             =   810
      Width           =   1050
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   330
      Left            =   1080
      TabIndex        =   1
      Top             =   810
      Width           =   1005
   End
   Begin VB.Label lbl 
      Caption         =   "lbl"
      Height          =   285
      Left            =   90
      TabIndex        =   3
      Top             =   45
      Width           =   3030
   End
End
Attribute VB_Name = "F_Input"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_ok As Boolean
Private Sub cmdCancel_Click()
  Me.Hide
End Sub
Private Sub cmdm_OK_Click()
  If ValText.FieldInvalid Then Exit Sub
  m_ok = True
  Me.Hide
End Sub
Public Function Start(ByVal sFormCaption As String, sLabel As String, sStartText As String, Optional bSelectText As Boolean = True) As Boolean
  Me.Caption = sFormCaption
  lbl = sLabel
  ValText.Text = sStartText
  If bSelectText Then
    ValText.SelLength = Len(sStartText)
  Else
    ValText.SelStart = Len(sStartText)
  End If
  m_ok = False
'  Me.Show vbModal
  Call p11d32.Help.ShowForm(Me, vbModal)
  If m_ok Then
    Start = Not ValText.FieldInvalid
  End If
End Function

Private Sub cmdOK_Click()
  m_ok = True
  Me.Hide
End Sub

