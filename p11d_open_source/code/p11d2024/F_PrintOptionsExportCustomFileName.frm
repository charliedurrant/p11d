VERSION 5.00
Begin VB.Form F_PrintOptionsExportCustomFileName 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Custom export file name"
   ClientHeight    =   1245
   ClientLeft      =   150
   ClientTop       =   795
   ClientWidth     =   5265
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1245
   ScaleWidth      =   5265
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Default         =   -1  'True
      Height          =   300
      Left            =   4200
      TabIndex        =   2
      Top             =   840
      Width           =   855
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&OK"
      Height          =   300
      Left            =   3240
      TabIndex        =   1
      Top             =   840
      Width           =   855
   End
   Begin P11D2024.CodeEditor codeEditor 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   661
   End
   Begin VB.Label Label1 
      Caption         =   "Right click in the text box to insert control codes"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   4815
   End
   Begin VB.Menu mnuX 
      Caption         =   "Click ->"
   End
End
Attribute VB_Name = "F_PrintOptionsExportCustomFileName"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_Cancelled As Boolean

Private Sub cmdCancel_Click()
  m_Cancelled = True
  Me.Hide
End Sub

Private Sub cmdOk_Click()
  Me.Hide
End Sub

Private Sub codeEditor_MenuClick(ByVal vbm As atc2dmenu.VBMenu, ByVal vbmi As atc2dmenu.VBMenuItem)
  Call ControlCodeClick(vbmi.Tag)
End Sub

Private Sub Form_Load()
  m_Cancelled = False
  Call EmployeeLetterAddCodes(Me.codeEditor, True)
  codeEditor.FormHWnd = Me.hwnd
  mnuX.Visible = False
End Sub
Private Function ControlCodeClick(Index As Long) As String
  Dim sCode As String
  
  sCode = EmployeeLetterCode(Index, ELCT_LETTER_FILE_CODES, False)
  Call codeEditor.InsertCodeIntoText(sCode)
End Function
Public Function ShowDialog(ByVal formParent As Form) As VbMsgBoxResult
  Call p11d32.Help.ShowForm(Me, FormShowConstants.vbModal)
  ShowDialog = IIf(m_Cancelled, vbCancel, vbOK)
End Function
Public Property Get Text() As String
  Text = codeEditor.Text
End Property

Public Property Let Text(ByVal value As String)
  codeEditor.Text = value
End Property
