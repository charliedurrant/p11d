VERSION 5.00
Object = "{5725D6E7-E490-4440-A4B0-2B60A8B971AA}#1.0#0"; "atc2Qbe.ocx"
Begin VB.Form frmFilter 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Filter Wizard"
   ClientHeight    =   2760
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   7125
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2760
   ScaleWidth      =   7125
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin atc2QBE.QBEGrid QBEFilter 
      Height          =   1905
      Left            =   450
      TabIndex        =   3
      Top             =   150
      Width           =   6405
      _ExtentX        =   11298
      _ExtentY        =   3360
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Height          =   330
      Left            =   4470
      TabIndex        =   2
      Top             =   2220
      Width           =   1250
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   330
      Left            =   2730
      TabIndex        =   1
      Top             =   2220
      Width           =   1250
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   330
      Left            =   990
      TabIndex        =   0
      Top             =   2220
      Width           =   1250
   End
   Begin VB.Menu mnuFilter 
      Caption         =   "&Filter"
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open Filter"
         Index           =   0
         Begin VB.Menu mnuNew 
            Caption         =   "&New Filter"
            Index           =   0
         End
      End
      Begin VB.Menu mnuSave 
         Caption         =   "&Save Filter"
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "&Delete Filter"
      End
   End
End
Attribute VB_Name = "frmFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Enum FILTER_BUTTON
  CANCEL_BUTTON = 0
  APPLY_BUTTON
  OK_BUTTON
End Enum
Public ExitMode As FILTER_BUTTON
Private iSelect As Integer
Private m_AGrid As AutoGrid

Private Sub cmdApply_Click()
  ExitMode = APPLY_BUTTON
  frmFilter.Hide
End Sub

Private Sub cmdCancel_Click()
  ExitMode = CANCEL_BUTTON
  frmFilter.Hide
End Sub

Private Sub cmdOK_Click()
  ExitMode = OK_BUTTON
  frmFilter.Hide
End Sub

Private Sub Form_Load()
  ExitMode = CANCEL_BUTTON
  iSelect = 0
End Sub

Private Sub mnuDelete_Click()
  Dim sFilter As String
  
  If iSelect > 0 Then
    sFilter = mnuNew(iSelect).Caption
    'Delete ini entries
    'Renumber remaining
    'change number of entries
  End If
End Sub

Private Function FillMenu()
  Dim i As Integer
  Dim iNoFilters As Integer
  Dim rs As Recordset
  
  Set rs = QBEFilter.rsFilter
  For i = 1 To mnuNew.Count - 1
    Unload mnuNew(i)
  Next i
  iNoFilters = CInt(GetIniEntry(rs.Name, "Filters", "0"))
  For i = 1 To iNoFilters
    Load mnuNew(i)
    mnuNew(i).Caption = GetIniEntry(QBEFilter.rsFilter.Name, "Filter" & CStr(i))
  Next i
End Function

Private Sub mnuNew_Click(Index As Integer)
  'Clear form
  QBEFilter.Reset
  If Index = 0 Then
    iSelect = 0
    frmFilter.Caption = "Filter Wizard"
  Else
    iSelect = Index
    QBEFilter.sSort = GetIniEntry(QBEFilter.rsFilter.Name, mnuNew(iSelect).Caption & "Sort")
    QBEFilter.sFilter = GetIniEntry(QBEFilter.rsFilter.Name, mnuNew(iSelect).Caption & "Filter")
    frmFilter.Caption = "Filter Wizard: " & mnuNew(iSelect).Caption
  End If
End Sub

Private Sub mnuFilter_Click()
  Call FillMenu
End Sub

Private Sub mnuSave_Click()
  Dim sName As String
  
  If iSelect > 0 Then
    sName = mnuNew(iSelect).Caption
  End If
  sName = InputBox("Please enter name of Filter", "Filter Save", sName)
  If sName = "" Then Exit Sub
  If iSelect = 0 Then
    iSelect = CInt(GetIniEntry(QBEFilter.rsFilter.Name, "Filters", "0") + 1)
    Call WriteIniEntry(QBEFilter.rsFilter.Name, "Filters", CStr(iSelect))
    Call WriteIniEntry(QBEFilter.rsFilter.Name, "Filter" & CStr(iSelect), sName)
  End If
  Call WriteIniEntry(QBEFilter.rsFilter.Name, sName & "Sort", QBEFilter.sSort)
  Call WriteIniEntry(QBEFilter.rsFilter.Name, sName & "Filter", QBEFilter.sFilter)
  sName = ""
End Sub

Public Function FillRCMenu(AGrid As AutoGrid)
  Dim i As Long, NoFilters As Long
  
  Set m_AGrid = AGrid
  For i = 1 To frmPopupMenu.mnuSavedFilter.UBound
    Unload frmPopupMenu.mnuSavedFilter(i)
  Next i
  NoFilters = CInt(GetIniEntry(AGrid.ParentAC.AutoName, "Filters", "0"))
  For i = 1 To NoFilters
    Load frmPopupMenu.mnuSavedFilter(i)
    frmPopupMenu.mnuSavedFilter(i).Caption = GetIniEntry(AGrid.ParentAC.AutoName, "Filter" & CStr(i))
    frmPopupMenu.mnuSavedFilter(i).Visible = True
  Next i
End Function

