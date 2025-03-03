VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.2#0"; "COMCTL32.OCX"
Begin VB.Form F_ErScreen 
   Appearance      =   0  'Flat
   Caption         =   "Employer details"
   ClientHeight    =   6165
   ClientLeft      =   840
   ClientTop       =   1725
   ClientWidth     =   9825
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   Icon            =   "Erscreen.frx":0000
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6165
   ScaleWidth      =   9825
   WindowState     =   2  'Maximized
   Begin ComctlLib.ListView lb 
      Height          =   4065
      Left            =   240
      TabIndex        =   0
      Tag             =   "free,font"
      Top             =   1755
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   7170
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   4
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Employer name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "PAYE reference"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   2
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Filename"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   3
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Fix Level"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.PictureBox Picture1 
      Height          =   1335
      Left            =   240
      ScaleHeight     =   1275
      ScaleWidth      =   9315
      TabIndex        =   2
      Tag             =   "free,font"
      Top             =   120
      Width           =   9375
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "32"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   495
         Left            =   5370
         TabIndex        =   6
         Tag             =   "FREE,FONT"
         Top             =   510
         Width           =   495
      End
      Begin VB.Label lblYear 
         Alignment       =   2  'Center
         Caption         =   "97"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   21.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   495
         Left            =   5775
         TabIndex        =   5
         Tag             =   "FREE,FONT"
         Top             =   105
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "P11D"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   39.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   855
         Left            =   3480
         TabIndex        =   4
         Tag             =   "FREE,FONT"
         Top             =   0
         Width           =   2175
      End
      Begin VB.Label L_Title 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Tax Computer Services, Arthur Andersen, (0171) 438 3491"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   2640
         TabIndex        =   3
         Tag             =   "FREE,FONT"
         Top             =   840
         Width           =   4290
      End
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Employer files in this directory:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   1
      Left            =   240
      TabIndex        =   1
      Tag             =   "FREE,FONT"
      Top             =   1515
      Width           =   2100
   End
End
Attribute VB_Name = "F_ErScreen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mclsResize As New clsFormResize
Private Const L_DES_HEIGHT = 6165
Private Const L_DES_WIDTH = 9825

Private Sub Form_Load()
  If Not (mclsResize.InitResize(Me, L_DES_HEIGHT, L_DES_WIDTH, DESIGN, , , MDIMain)) Then
    Err.Raise ERR_Application
  End If
  lblYear = g_sLabelYear
End Sub

Private Sub Form_Resize()
  mclsResize.Resize
  Call ColumnWidths(Me.LB, 50, 30, 10, 10)
End Sub


Private Sub LB_ColumnClick(ByVal ColumnHeader As ComctlLib.ColumnHeader)
  Me.LB.SortKey = ColumnHeader.Index - 1
  LB.SelectedItem.EnsureVisible
End Sub

Private Sub LB_DblClick()
  If Not LB.SelectedItem Is Nothing Then
    Call ToolBarButton(TBR_OPEN, CLng(LB.SelectedItem.Tag))
  End If
End Sub

Private Sub Picture1_GotFocus()
  SendKeys (vbTab)
End Sub

Private Sub LB_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then 'Return key
    Call LB_DblClick
  End If
End Sub


Public Sub NoEmployers()
  Dim i As Long
  On Error GoTo NoEmployers_Err
  Call xSet("NoEmployers")
  
  LB.Enabled = False
  For i = TBR_OPEN To TBR_KILLEMPLOYER
    If i <> TBR_ADDEMPLOYER Then
      MDIMain.tbrMain.Buttons(i).Enabled = False
    End If
  Next i
  For i = MNU_FILEOPEN To MNU_EDI 'MNU_FILEDELETE CAD ZZZZ
    If i <> MNU_SEPERATOR1 Then MDIMain.mnuFileItems(i).Enabled = False
  Next i
  
NoEmployers_End:
  Call xReturn("NoEmployers")
  Exit Sub

NoEmployers_Err:
  Call ErrorMessage(ERR_ERROR, Err, "NoEmployers", "ERR_UNDEFINED", "Undefined error.")
  Resume NoEmployers_End
  Resume
End Sub


Public Sub Employers()
  Dim i As Long
  On Error GoTo Employers_Err
  Call xSet("Employers")
  
  LB.Enabled = True
  For i = TBR_OPEN To TBR_KILLEMPLOYER
    If i <> TBR_ADDEMPLOYER Then
      MDIMain.tbrMain.Buttons(i).Enabled = True
    End If
  Next i
  For i = MNU_FILEOPEN To MNU_EDI 'MNU_FILEDELETE CAD ZZZZ
    If i <> MNU_SEPERATOR1 Then
      MDIMain.mnuFileItems(i).Enabled = True
      MDIMain.mnuFileItems(i).Visible = True
    End If
  Next i

Employers_End:
  Call xReturn("Employers")
  Exit Sub

Employers_Err:
  Call ErrorMessage(ERR_ERROR, Err, "Employers", "ERR_UNDEFINED", "Undefined error.")
  Resume Employers_End
End Sub

