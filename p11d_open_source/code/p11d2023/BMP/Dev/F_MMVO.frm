VERSION 5.00
Object = "{412521B9-9CBB-4049-9E66-2AA0112EC306}#1.0#0"; "ATC2VTEXT.OCX"
Begin VB.Form F_MMViewOptions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MM View options"
   ClientHeight    =   1005
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3840
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1005
   ScaleWidth      =   3840
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2745
      TabIndex        =   3
      Top             =   585
      Width           =   1005
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   1665
      TabIndex        =   2
      Top             =   585
      Width           =   1005
   End
   Begin atc2valtext.ValText vtRecordViewID 
      Height          =   375
      Left            =   2880
      TabIndex        =   0
      Top             =   90
      Width           =   870
      _ExtentX        =   1535
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxLength       =   2
      Text            =   ""
      TypeOfData      =   3
      AutoSelect      =   0
   End
   Begin VB.Label lbl 
      Caption         =   "Enter the record ID to view ie 2E for employee cars"
      Height          =   420
      Left            =   45
      TabIndex        =   1
      Top             =   90
      Width           =   2715
   End
End
Attribute VB_Name = "F_MMViewOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
  Unload Me
End Sub

Private Sub cmdOK_Click()
  If Validate Then
    
    Unload Me
  End If
End Sub
Private Function Validate() As Boolean
  Validate = Not vtRecordViewID.FieldInvalid
  If Validate = False Then Call ErrorMessage(ERR_ERROR, Err, "Validate", "Validate", "Some of the data entered is invalid, ,please ammend.")
End Function
Private Sub Form_Load()
  vtRecordViewID.Text = P11d32.MagneticMedia.RecordViewID
  vtRecordViewID.SelStart = 0
  vtRecordViewID.SelLength = Len(P11d32.MagneticMedia.RecordViewID)
End Sub

