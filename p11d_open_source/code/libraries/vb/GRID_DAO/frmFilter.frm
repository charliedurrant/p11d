VERSION 5.00
Object = "{EC655C08-40A0-4D95-AF62-863C56B762FB}#1.0#0"; "atcQbe.ocx"
Begin VB.Form frmFilter 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Filter Wizard"
   ClientHeight    =   2760
   ClientLeft      =   48
   ClientTop       =   612
   ClientWidth     =   7128
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2760
   ScaleWidth      =   7128
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin atcQBE.QBEGrid QBEFilter 
      Height          =   1905
      Left            =   450
      TabIndex        =   3
      Top             =   150
      Width           =   6405
      _ExtentX        =   11282
      _ExtentY        =   3344
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
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
Public mode As Integer
Private iSelect As Integer
Private myag As AutoGrid

Private Sub cmdApply_Click()
  mode = 1
  frmFilter.Hide
End Sub

Private Sub cmdCancel_Click()
  mode = 0
  frmFilter.Hide
End Sub

Private Sub cmdOK_Click()
  mode = 2
  frmFilter.Hide
End Sub

Private Sub Form_Load()
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
  Dim RS As Recordset
  
  Set RS = QBEFilter.rsFilter
  For i = 1 To mnuNew.Count - 1
    Unload mnuNew(i)
  Next i
  iNoFilters = CInt(GetIniEntry(RS.Name, "Filters", "0"))
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
  FillMenu
End Sub

Private Sub mnuSave_Click()
  Dim sName As String
  
  If iSelect > 0 Then
    sName = mnuNew(iSelect).Caption
  End If
  Do While sName = ""
    sName = InputBox("Please enter name of Filter", "Filter Save", sDef)
  Loop
  If iSelect = 0 Then
    iSelect = CInt(GetIniEntry(QBEFilter.rsFilter.Name, "Filters", "0") + 1)
    Call WriteIniEntry(QBEFilter.rsFilter.Name, "Filters", CStr(iSelect))
    Call WriteIniEntry(QBEFilter.rsFilter.Name, "Filter" & CStr(iSelect), sName)
  End If
  Call WriteIniEntry(QBEFilter.rsFilter.Name, sName & "Sort", QBEFilter.sSort)
  Call WriteIniEntry(QBEFilter.rsFilter.Name, sName & "Filter", QBEFilter.sFilter)
  sName = ""
End Sub

Public Function FillRCMenu(ag As AutoGrid)
  Dim i As Integer
  Dim iNoFilters As Integer
  
  Set myag = ag
  For i = 1 To frmSaveOptions.mnuSavedFilter.Count - 1
    Unload frmSaveOptions.mnuSavedFilter(i)
  Next i
  iNoFilters = CInt(GetIniEntry(myag.pseCols.RS.Name, "Filters", "0"))
  For i = 1 To iNoFilters
    Load frmSaveOptions.mnuSavedFilter(i)
    frmSaveOptions.mnuSavedFilter(i).Caption = GetIniEntry(myag.pseCols.RS.Name, "Filter" & CStr(i))
    frmSaveOptions.mnuSavedFilter(i).Visible = True
  Next i
End Function

Public Sub SavedFilterClick(Index As Integer)
    Call myag.pseCols.Sort(GetIniEntry(myag.pseCols.RS.Name, frmSaveOptions.mnuSavedFilter(Index).Caption & "Sort"))
    Call myag.pseCols.Filter(GetIniEntry(myag.pseCols.RS.Name, frmSaveOptions.mnuSavedFilter(Index).Caption & "Filter"))
End Sub
