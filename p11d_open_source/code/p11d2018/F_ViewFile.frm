VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form F_ViewFile 
   Caption         =   "F_ViewFile"
   ClientHeight    =   4710
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   6300
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   4710
   ScaleWidth      =   6300
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdFindNext 
      Caption         =   "Find &Next"
      Height          =   375
      Left            =   2475
      TabIndex        =   4
      Tag             =   "LOCKB"
      Top             =   4320
      Width           =   1095
   End
   Begin VB.TextBox txtFindWhat 
      Height          =   360
      Left            =   45
      TabIndex        =   3
      Tag             =   "LOCKB"
      Top             =   4320
      Width           =   2400
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   5040
      TabIndex        =   1
      Tag             =   "LOCKBR"
      Top             =   4320
      Width           =   1230
   End
   Begin RichTextLib.RichTextBox rtViewFile 
      Height          =   3885
      Left            =   0
      TabIndex        =   0
      Tag             =   "EQUALISE"
      Top             =   360
      Width           =   6270
      _ExtentX        =   11060
      _ExtentY        =   6853
      _Version        =   393217
      HideSelection   =   0   'False
      ScrollBars      =   3
      RightMargin     =   65535
      TextRTF         =   $"F_ViewFile.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      Caption         =   "lblInfo"
      Height          =   195
      Left            =   3600
      TabIndex        =   5
      Tag             =   "LOCKB"
      Top             =   4410
      Width           =   420
   End
   Begin VB.Label lblRowColumn 
      Height          =   285
      Left            =   45
      TabIndex        =   2
      Tag             =   "LOCK"
      Top             =   45
      Width           =   2625
   End
End
Attribute VB_Name = "F_ViewFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_CR As clsFormResize
Private m_bInLoad As Boolean
Private m_lFindPos As Long

Public Function ViewFile(Optional ByVal sPathAndFile As String = "", Optional sFilter As String = "AllFiles (*.*)|*.*", Optional ByVal sStartDirectory As String = "C:\", Optional IViewFile As IViewFile = Nothing) As String
  Dim sLabelInfo As String
  
  On Error GoTo ViewFile_ERR
  
  Call xSet("ViewFile")
  
  If Len(sPathAndFile) = 0 Then
    sPathAndFile = FileOpenDlg("Open file", sFilter, sStartDirectory)
  End If
  If Len(sPathAndFile) = 0 Then GoTo ViewFile_END
  Call FileExistsAndNotOpenExclusive(sPathAndFile)
  Call SetCursor(vbArrowHourglass)
  m_bInLoad = True
  rtViewFile.FileName = sPathAndFile
  If Not IViewFile Is Nothing Then Call IViewFile.View(rtViewFile, sLabelInfo)
  lblInfo = sLabelInfo
  m_bInLoad = False
  rtViewFile.SelStart = 0
  Call ClearCursor
  
  Me.WindowState = vbMaximized
  Me.Caption = sPathAndFile
'  Me.Show vbModal
  Call p11d32.Help.ShowForm(Me, vbModal)
  ViewFile = sPathAndFile
  
ViewFile_END:
  
  Call xReturn("ViewFile")
  Exit Function
ViewFile_ERR:
  Call ClearCursor
  Unload Me
  Call ErrorMessage(ERR_ERROR, Err, "ViewFile", "View File", "Error viewing the file " & sPathAndFile & ".")
  Resume ViewFile_END
  Resume
End Function


Private Function FindString()
  
  
  m_lFindPos = rtViewFile.Find(txtFindWhat.Text, m_lFindPos)
  If m_lFindPos <> -1 Then
    rtViewFile.SelStart = m_lFindPos
    rtViewFile.SelLength = Len(txtFindWhat)
    m_lFindPos = m_lFindPos + Len(txtFindWhat)
  Else
    If MsgBox("No more matches found for " & txtFindWhat.Text & vbCrLf & "Continue searching form the top?", vbYesNo) = vbYes Then
      m_lFindPos = 0
      rtViewFile.SelStart = 0
      rtViewFile.SelLength = 0
      Call FindString
    End If
  End If

End Function

Private Sub cmdFindNext_Click()
  DoEvents
  Call FindString
End Sub

Private Sub cmdFindNext_GotFocus()
  m_lFindPos = rtViewFile.SelStart + rtViewFile.SelLength
End Sub

Private Sub cmdOK_Click()
  Unload Me
End Sub

Private Sub Form_Load()
  Set m_CR = New clsFormResize
  Call m_CR.InitResize(Me, 5115, 6420, DESIGN)
End Sub

Private Sub Form_Resize()
  Call m_CR.Resize
End Sub
Private Sub UpdateRowCol()
  Dim lRow As Long, lRow2 As Long
  Dim lCol As Long, lCol2 As Long
  
  
  lCol = rtViewFile.SelStart
  lRow = rtViewFile.GetLineFromChar(lCol)
  lCol2 = lCol
  
  If lRow > 0 Then
    Do
      lCol = lCol - 1
      lRow2 = rtViewFile.GetLineFromChar(lCol)
    Loop Until lRow2 = lRow - 1
    lblRowColumn = "Row " & lRow + 1 & ", Column " & lCol2 - lCol
  Else
    lblRowColumn = "Row " & lRow + 1 & ", Column " & lCol + 1
  End If

End Sub
Private Sub rtViewFile_SelChange()
  If Not m_bInLoad Then Call UpdateRowCol
End Sub

Private Sub txtFindWhat_GotFocus()
  txtFindWhat.SelStart = 0
  txtFindWhat.SelLength = Len(txtFindWhat.Text)
End Sub

Private Sub txtFindWhat_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    DoEvents
    Call FindString
  End If
End Sub
