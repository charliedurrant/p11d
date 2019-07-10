VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Frm_DLim 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Delimit Source File"
   ClientHeight    =   6372
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   7632
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6372
   ScaleWidth      =   7632
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Fra_FieldNames 
      Height          =   612
      Left            =   2700
      TabIndex        =   22
      Top             =   1260
      Width           =   3312
      Begin MSComCtl2.UpDown UpD_FieldNames 
         Height          =   288
         Left            =   2944
         TabIndex        =   25
         Top             =   240
         Width           =   192
         _ExtentX        =   423
         _ExtentY        =   508
         _Version        =   393216
         BuddyControl    =   "Txt_FieldNames"
         BuddyDispid     =   196610
         OrigLeft        =   3120
         OrigTop         =   240
         OrigRight       =   3312
         OrigBottom      =   528
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   0   'False
      End
      Begin VB.TextBox Txt_FieldNames 
         Enabled         =   0   'False
         Height          =   288
         Left            =   2400
         TabIndex        =   24
         Text            =   "0"
         Top             =   240
         Width           =   543
      End
      Begin VB.CheckBox Chk_FieldNames 
         Height          =   372
         Left            =   180
         TabIndex        =   23
         Top             =   180
         Width           =   192
      End
      Begin VB.Label Lbl_FieldNames 
         Caption         =   "Take field names from line"
         Height          =   252
         Left            =   420
         TabIndex        =   26
         Top             =   240
         Width           =   1992
      End
   End
   Begin VB.Frame Fra_Omit 
      Caption         =   "Omit Rows"
      Height          =   612
      Left            =   360
      TabIndex        =   14
      Top             =   1920
      Width           =   5655
      Begin VB.TextBox Txt_Omit 
         Height          =   288
         Index           =   0
         Left            =   1320
         TabIndex        =   17
         Text            =   "0"
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox Txt_Omit 
         Height          =   288
         Index           =   1
         Left            =   3960
         TabIndex        =   16
         Text            =   "0"
         Top             =   240
         Width           =   735
      End
      Begin MSComCtl2.UpDown UpD_Omit 
         Height          =   288
         Index           =   0
         Left            =   2056
         TabIndex        =   15
         Top             =   240
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   508
         _Version        =   393216
         BuddyControl    =   "Txt_Omit(0)"
         BuddyDispid     =   196614
         BuddyIndex      =   0
         OrigLeft        =   2280
         OrigTop         =   240
         OrigRight       =   2520
         OrigBottom      =   525
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown UpD_Omit 
         Height          =   288
         Index           =   1
         Left            =   4696
         TabIndex        =   18
         Top             =   240
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   508
         _Version        =   393216
         BuddyControl    =   "Txt_Omit(1)"
         BuddyDispid     =   196614
         BuddyIndex      =   1
         OrigLeft        =   4920
         OrigTop         =   240
         OrigRight       =   5160
         OrigBottom      =   525
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.Label Lbl_Omit 
         Caption         =   "Header"
         Height          =   255
         Index           =   0
         Left            =   720
         TabIndex        =   20
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Lbl_Omit 
         Caption         =   "Footer"
         Height          =   255
         Index           =   1
         Left            =   3360
         TabIndex        =   19
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.ComboBox Cbo_TextQual 
      Height          =   315
      ItemData        =   "Frm_dlim.frx":0000
      Left            =   1440
      List            =   "Frm_dlim.frx":000D
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   1440
      Width           =   975
   End
   Begin VB.Frame Fra_Delimiters 
      Caption         =   "Choose the delimiter(s) that separate(s) your fields:"
      Height          =   1095
      Left            =   360
      TabIndex        =   4
      Top             =   120
      Width           =   5655
      Begin VB.TextBox Txt_Delim 
         Height          =   285
         Left            =   4920
         TabIndex        =   13
         Top             =   660
         Width           =   492
      End
      Begin VB.CheckBox Chk_Delims 
         Caption         =   "Other:"
         Height          =   192
         Index           =   5
         Left            =   4080
         TabIndex        =   12
         Top             =   720
         Width           =   732
      End
      Begin VB.CheckBox Chk_Delims 
         Caption         =   "Space"
         Height          =   192
         Index           =   4
         Left            =   2040
         TabIndex        =   11
         Top             =   720
         Width           =   855
      End
      Begin VB.CheckBox Chk_Delims 
         Caption         =   "Tab ()"
         Height          =   192
         Index           =   3
         Left            =   4080
         TabIndex        =   10
         Top             =   360
         Width           =   855
      End
      Begin VB.CheckBox Chk_Delims 
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   192
         Index           =   2
         Left            =   2040
         TabIndex        =   9
         Top             =   360
         Width           =   372
      End
      Begin VB.CheckBox Chk_Delims 
         Caption         =   ";"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   192
         Index           =   1
         Left            =   360
         TabIndex        =   8
         Top             =   720
         Width           =   372
      End
      Begin VB.CheckBox Chk_Delims 
         Caption         =   ","
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   192
         Index           =   0
         Left            =   360
         TabIndex        =   7
         Top             =   360
         Value           =   1  'Checked
         Width           =   372
      End
   End
   Begin VB.CommandButton Cmd_Next 
      Caption         =   "&Next >"
      Height          =   375
      Left            =   4680
      TabIndex        =   2
      Top             =   5760
      Width           =   1215
   End
   Begin VB.CommandButton Cmd_Back 
      Caption         =   "< &Back"
      Height          =   375
      Left            =   3480
      TabIndex        =   1
      Top             =   5760
      Width           =   1215
   End
   Begin VB.CommandButton Cmd_Cancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   6120
      TabIndex        =   0
      Top             =   5760
      Width           =   1215
   End
   Begin MSFlexGridLib.MSFlexGrid FlG_Source 
      DragIcon        =   "Frm_dlim.frx":001F
      Height          =   2187
      Left            =   240
      TabIndex        =   21
      Top             =   3228
      Width           =   7095
      _ExtentX        =   12510
      _ExtentY        =   3852
      _Version        =   393216
      Cols            =   1
      AllowBigSelection=   -1  'True
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   0
      FillStyle       =   1
      SelectionMode   =   2
      AllowUserResizing=   1
   End
   Begin VB.Label Lbl_TextQual 
      Caption         =   "Text Qualifier"
      Height          =   255
      Left            =   360
      TabIndex        =   6
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label Lbl_DelimitedFile 
      Caption         =   "Delimited contents of source file:"
      Height          =   252
      Left            =   360
      TabIndex        =   3
      Top             =   2940
      Width           =   2652
   End
End
Attribute VB_Name = "Frm_DLim"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_ImpWiz As ImportWizard
Implements IImportForm

Private Sub Chk_FieldNames_Click()
  If Chk_FieldNames.Value = vbChecked Then
    Txt_FieldNames.Enabled = True
    UpD_FieldNames.Enabled = True
    Txt_FieldNames.Text = "1"
    If CLng(Txt_FieldNames.Text) > CLng(Txt_Omit(0).Text) Then Txt_Omit(0).Text = Txt_FieldNames.Text
  Else
    Txt_FieldNames.Enabled = False
    UpD_FieldNames.Enabled = False
    Txt_FieldNames.Text = "0"
  End If
  Call m_ImpWiz.ReCalc_DLim(False)
End Sub

Private Sub Command1_Click()
  Dim impvar As Variant, s As String, i As Long
  
  Call m_ImpWiz.ImportFieldNames(impvar)
  s = ""
  For i = 1 To 4
    s = s & impvar(1, i) & vbCrLf
  Next i
  MsgBox s
End Sub

Private Property Get IImportForm_FormType() As IMPORT_GOTOFORM
  IImportForm_FormType = TCSIMP_DLIM
End Property

Private Property Set IImportForm_ImpWiz(RHS As ImportWizard)
  Set m_ImpWiz = RHS
End Property

Private Property Get IImportForm_ImpWiz() As ImportWizard
  Set IImportForm_ImpWiz = m_ImpWiz
End Property

Private Sub Cbo_TextQual_LostFocus()
  Call m_ImpWiz.ReCalc_DLim(False)
End Sub

Private Sub Chk_Delims_Click(Index As Integer)
  'PC GET Delimeter string here
  If m_ImpWiz.GotoForm = TCSIMP_DLIM Then Call m_ImpWiz.ReCalc_DLim(True)
End Sub

Private Sub Cmd_Back_Click()
  Call m_ImpWiz.ReCalc_Src(Nothing)
  Call SwitchForm(Me, TCSIMP_SOURCE)
End Sub

Private Sub Cmd_Cancel_Click()
  Call SwitchForm(Me, TCSIMP_CANCEL)
End Sub

Private Sub Cmd_Next_Click()
  Call m_ImpWiz.ReCalc_Misc(False)
  Call SwitchForm(Me, TCSIMP_MISC)
End Sub

Private Sub Txt_Delim_LostFocus()
  Call m_ImpWiz.ReCalc_DLim(False)
End Sub

Private Sub Txt_FieldNames_LostFocus()
  Txt_FieldNames.Text = CStr(Max(Min(CLngEx(Txt_FieldNames.Text, 0), UpD_FieldNames.Max), 1))
  If CLng(Txt_FieldNames.Text) > CLngEx(Txt_Omit(0).Text, 0) Then Txt_Omit(0).Text = Txt_FieldNames.Text
  Call m_ImpWiz.ReCalc_DLim(False)
End Sub

Private Sub Txt_Omit_LostFocus(Index As Integer)
  Call OmitLines(Me.Txt_Omit(0), Me.Txt_Omit(1), CLngEx(Me.Txt_FieldNames.Text, 0), m_ImpWiz.NumLines)
  Call m_ImpWiz.ReCalc_DLim(False)
End Sub

Private Sub UpD_FieldNames_Change()
  Call Txt_FieldNames_LostFocus
End Sub

Private Sub UpD_Omit_Change(Index As Integer)
  Call Txt_Omit_LostFocus(Index)
End Sub

Public Function GetDelimiters() As String
  Dim i As Long, s As String
  Static inGetDelimiter As Boolean
    
  Call xSet("GetDelimiter")
  On Error Resume Next
  If Not inGetDelimiter Then
    inGetDelimiter = True
    For i = 0 To Me.Chk_Delims.Count - 1
      If Me.Chk_Delims(i).Value = vbChecked Then
        Select Case i
          Case 0: s = s & ","
          Case 1: s = s & ";"
          Case 2: s = s & ":"
          Case 3: s = s & vbTab
          Case 4: s = s & " "
          Case 5
            If Len(Me.Txt_Delim.Text) = 0 Then
              Me.Txt_Delim.Text = DEFAULT_TEXT_DELIMITER
              Me.Chk_Delims(5).Value = vbUnchecked
            Else
              s = s & Left$(Me.Txt_Delim.Text, 1)
            End If
        End Select
      End If
    Next i
    GetDelimiters = s
    inGetDelimiter = False
  End If
End Function

Public Sub SetDelimsEscChar(ByVal DLims As String, EChar As String)
  Dim i As Long
  Static inSetDelimiter As Boolean
    
  Call xSet("SetDelimiter")
  On Error Resume Next
  If Not inSetDelimiter Then
    inSetDelimiter = True
    
    For i = 0 To 5
      Me.Chk_Delims(i).Value = vbUnchecked
    Next i
    
    Do
      i = InStr(DLims, ",")
      If i = 0 Then Exit Do
      Me.Chk_Delims(0).Value = vbChecked
      DLims = Left$(DLims, i - 1) & Right$(DLims, Len(DLims) - i)
    Loop
    
    Do
      i = InStr(DLims, ";")
      If i = 0 Then Exit Do
      Me.Chk_Delims(1).Value = vbChecked
      DLims = Left$(DLims, i - 1) & Right$(DLims, Len(DLims) - i)
    Loop
    
    Do
      i = InStr(DLims, ":")
      If i = 0 Then Exit Do
      Me.Chk_Delims(2).Value = vbChecked
      DLims = Left$(DLims, i - 1) & Right$(DLims, Len(DLims) - i)
    Loop
    
    Do
      i = InStr(DLims, vbTab)
      If i = 0 Then Exit Do
      Me.Chk_Delims(3).Value = vbChecked
      DLims = Left$(DLims, i - 1) & Right$(DLims, Len(DLims) - i)
    Loop
    
    Do
      i = InStr(DLims, " ")
      If i = 0 Then Exit Do
      Me.Chk_Delims(4).Value = vbChecked
      DLims = Left$(DLims, i - 1) & Right$(DLims, Len(DLims) - i)
    Loop
    
    If Len(DLims) > 0 Then
      Me.Chk_Delims(5).Value = vbChecked
      Me.Txt_Delim.Text = Left$(DLims, 1)
    Else
      Me.Txt_Delim.Text = ""
    End If
     
    Select Case EChar
      Case "'"
        Me.Cbo_TextQual.Text = "'"
      Case Chr$(34)
        Me.Cbo_TextQual.Text = Chr$(34)
      Case Else
        Me.Cbo_TextQual.Text = "{None}"
    End Select
 
    inSetDelimiter = False
  End If
End Sub

Public Function GetTextQualifier() As String
  If Len(Me.Cbo_TextQual.Text) = 0 Then
    Me.Cbo_TextQual.Text = Me.Cbo_TextQual.List(0)
  End If
  If StrComp(Me.Cbo_TextQual.Text, "{None}", vbTextCompare) = 0 Then
    GetTextQualifier = ""
  Else
    GetTextQualifier = Me.Cbo_TextQual.Text
  End If
End Function

