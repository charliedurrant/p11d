VERSION 5.00
Begin VB.Form frmErr 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " "
   ClientHeight    =   2370
   ClientLeft      =   945
   ClientTop       =   1785
   ClientWidth     =   7620
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2370
   ScaleWidth      =   7620
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraApp 
      Height          =   585
      Left            =   90
      TabIndex        =   8
      Top             =   1740
      Width           =   4620
      Begin VB.Label lblApplication 
         Height          =   180
         Left            =   120
         TabIndex        =   9
         Top             =   210
         Width           =   4395
      End
   End
   Begin VB.Frame fraDetails 
      Caption         =   "Details"
      Height          =   2310
      Left            =   75
      TabIndex        =   6
      Top             =   2415
      Width           =   7440
      Begin VB.TextBox txtPath 
         BackColor       =   &H80000004&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   90
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   675
         Width           =   5100
      End
      Begin VB.ListBox lstStack 
         Height          =   1620
         Left            =   5265
         TabIndex        =   12
         Top             =   450
         Width           =   2055
      End
      Begin VB.Label lblErrType 
         Height          =   285
         Left            =   90
         TabIndex        =   14
         Top             =   315
         Width           =   5100
      End
      Begin VB.Label lblStack 
         Caption         =   "Function call stack:"
         Height          =   255
         Left            =   5325
         TabIndex        =   13
         Top             =   225
         Width           =   1965
      End
      Begin VB.Label lblCoreVersion 
         Height          =   285
         Left            =   90
         TabIndex        =   11
         Top             =   1755
         Width           =   5100
      End
      Begin VB.Label lblFunction 
         Height          =   285
         Left            =   90
         TabIndex        =   10
         Top             =   1395
         Width           =   5100
      End
      Begin VB.Label lblExeName 
         Height          =   285
         Left            =   90
         TabIndex        =   7
         Top             =   1035
         Width           =   5100
      End
   End
   Begin VB.CommandButton cmdDetails 
      Caption         =   "&Details >>"
      Height          =   375
      Left            =   4875
      TabIndex        =   4
      Top             =   1950
      Width           =   1260
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   6225
      TabIndex        =   3
      Top             =   1950
      Width           =   1260
   End
   Begin VB.Frame fraContact 
      Height          =   525
      Left            =   75
      TabIndex        =   2
      Top             =   1125
      Width           =   7440
      Begin VB.Label lblHelp 
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   7200
      End
   End
   Begin VB.Frame fraErr 
      Height          =   1080
      Left            =   75
      TabIndex        =   0
      Top             =   0
      Width           =   7440
      Begin VB.Label lblErrMsg 
         AutoSize        =   -1  'True
         Height          =   675
         Left            =   150
         TabIndex        =   1
         Top             =   240
         Width           =   7200
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Menu mnuDebug 
      Caption         =   "&Debug"
      Visible         =   0   'False
      Begin VB.Menu mnuBreak 
         Caption         =   "&Break"
      End
      Begin VB.Menu mnuDebugItem 
         Caption         =   "&Environment"
         Index           =   0
         Begin VB.Menu mnuEnvirItem 
            Caption         =   "&System"
            Index           =   0
         End
         Begin VB.Menu mnuEnvirItem 
            Caption         =   "&Application"
            Index           =   1
         End
         Begin VB.Menu mnuEnvirItem 
            Caption         =   "&Files"
            Enabled         =   0   'False
            Index           =   2
         End
      End
      Begin VB.Menu mnuDebugItem 
         Caption         =   "&Database"
         Index           =   1
         Begin VB.Menu mnuDatabaseItem 
            Caption         =   "&Repair and Compact"
            Enabled         =   0   'False
            Index           =   0
         End
         Begin VB.Menu mnuDatabaseItem 
            Caption         =   "&SQL"
            Enabled         =   0   'False
            Index           =   1
         End
      End
      Begin VB.Menu mnuUserDebugItem 
         Caption         =   "&Application"
         Enabled         =   0   'False
         Begin VB.Menu mnuOtherItem 
            Caption         =   "AppItem"
            Enabled         =   0   'False
            Index           =   0
         End
      End
   End
   Begin VB.Menu mnuGrid 
      Caption         =   "Grid"
      Visible         =   0   'False
      Begin VB.Menu mnuFilterSelection 
         Caption         =   "Filter by Selection"
      End
      Begin VB.Menu mnuFilterEx 
         Caption         =   "Filter excluding Selection"
      End
      Begin VB.Menu mnuFilters 
         Caption         =   "Filters"
         Begin VB.Menu mnuWizard 
            Caption         =   "New Filter"
         End
      End
      Begin VB.Menu mnuRemoveFilter 
         Caption         =   "Remove Filter/Sort"
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSortAsc 
         Caption         =   "Sort Ascending"
      End
      Begin VB.Menu mnuSortDesc 
         Caption         =   "Sort Descending"
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCopyCol 
         Caption         =   "Copy Columns"
      End
      Begin VB.Menu mnuInsCol 
         Caption         =   "Insert Columns"
      End
      Begin VB.Menu mnusep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCutRow 
         Caption         =   "Cut"
      End
      Begin VB.Menu mnuCopyRow 
         Caption         =   "Copy"
      End
      Begin VB.Menu mnuPasteRow 
         Caption         =   "Paste"
      End
      Begin VB.Menu mnusep4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFind 
         Caption         =   "Find..."
      End
      Begin VB.Menu mnuOther 
         Caption         =   "Other"
         Enabled         =   0   'False
         Index           =   0
         Visible         =   0   'False
      End
   End
End
Attribute VB_Name = "frmErr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const ERRFRM_INCREASE = 2395

Private Sub cmdDetails_Click()
Static bDetails As Boolean

  If bDetails Then
    Me.height = Me.height - ERRFRM_INCREASE
    Me.cmdDetails.Caption = "&Details >>"
    bDetails = False
  Else
    Me.height = Me.height + ERRFRM_INCREASE
    Me.cmdDetails.Caption = "&Details <<"
    bDetails = True
  End If

End Sub

Private Sub cmdOK_Click()
  Unload Me
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF12 And Shift = False Then
    Call gSystem.ShowDebugPopup
  End If
End Sub

Private Sub Form_Load()
  Call FillDebugMenu
End Sub

Private Sub mnuBreak_Click()
  Call debugmenu(0, MNU_BREAK, "BREAK")
End Sub

Private Sub mnuDatabaseItem_Click(Index As Integer)
  Call debugmenu(CLng(Index), MNU_DATABASE, mnuDatabaseItem(Index).Tag)
End Sub

Private Sub mnuEnvirItem_Click(Index As Integer)
  Call debugmenu(CLng(Index), MNU_ENVIRONMENT, mnuEnvirItem(Index).Tag)
End Sub

Private Sub mnuOtherItem_Click(Index As Integer)
  Call debugmenu(CLng(Index), MNU_APPLICATION, mnuOtherItem(Index).Tag)
End Sub
Private Sub FillDebugMenu()
Dim m As MenuClass
  Dim l As Long
  If gMenusCollection.Count > 0 Then
    mnuUserDebugItem.Enabled = True
    For Each m In gMenusCollection
      Select Case m.Parent
        Case MNU_DATABASE
          l = Me.mnuDatabaseItem.Count
          Load mnuDatabaseItem(l)
          mnuDatabaseItem(l).Caption = m.Caption
          mnuDatabaseItem(l).Enabled = m.Enabled
          mnuDatabaseItem(l).Tag = m.Name
        Case MNU_ENVIRONMENT
          l = Me.mnuEnvirItem.Count
          Load mnuEnvirItem(l)
          mnuEnvirItem(l).Caption = m.Caption
          mnuEnvirItem(l).Enabled = m.Enabled
          mnuEnvirItem(l).Tag = m.Name
        Case MNU_APPLICATION
          l = Me.mnuOtherItem.Count
          Load mnuOtherItem(l)
          mnuOtherItem(l).Caption = m.Caption
          mnuOtherItem(l).Enabled = m.Enabled
          mnuOtherItem(l).Tag = m.Name
        Case Else
          gCore.ECASE "CREATE_MENUS - INVALID PARENT"
      End Select
    Next m
    'special case - no default items in menu
    If mnuOtherItem.Count > 1 Then mnuOtherItem(0).Visible = False
  End If
End Sub

