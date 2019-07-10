VERSION 5.00
Begin VB.Form frmPopupMenu 
   Caption         =   "Form1"
   ClientHeight    =   2355
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   4485
   LinkTopic       =   "Form1"
   ScaleHeight     =   2355
   ScaleWidth      =   4485
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu mnuGrid 
      Caption         =   "Grid"
      Visible         =   0   'False
      Begin VB.Menu mnuFilterSelectionInc 
         Caption         =   "Filter by Selection"
      End
      Begin VB.Menu mnuFilterSelectionEx 
         Caption         =   "Filter excluding Selection"
      End
      Begin VB.Menu mnuFilterS 
         Caption         =   "Filters"
         Begin VB.Menu mnuWizard 
            Caption         =   "Filter Wizard"
         End
         Begin VB.Menu sep1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuSavedFilter 
            Caption         =   "SavedFilter"
            Index           =   0
            Visible         =   0   'False
         End
      End
      Begin VB.Menu mnuRemoveFilter 
         Caption         =   "Remove Filter/Sort"
      End
      Begin VB.Menu sep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbortEdit 
         Caption         =   "Abort Edit"
      End
      Begin VB.Menu mnuCommitEdit 
         Caption         =   "Commit Edit"
      End
      Begin VB.Menu sep6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSortAsc 
         Caption         =   "Sort Ascending"
      End
      Begin VB.Menu mnuSortDesc 
         Caption         =   "Sort Descending"
      End
      Begin VB.Menu sep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCopyCol 
         Caption         =   "Move Columns"
      End
      Begin VB.Menu mnuInsCol 
         Caption         =   "Insert Columns"
      End
      Begin VB.Menu mnuAutoSize 
         Caption         =   "Resize Column"
      End
      Begin VB.Menu mnuAutoSizeAll 
         Caption         =   "Resize All Columns"
      End
      Begin VB.Menu sep4 
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
      Begin VB.Menu sep5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFind 
         Caption         =   "Find..."
      End
      Begin VB.Menu mnuDebugSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDebugMenu 
         Caption         =   "Show Format"
      End
      Begin VB.Menu UserSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuUserMenus 
         Caption         =   "UserMenus"
         Index           =   0
         Visible         =   0   'False
      End
   End
End
Attribute VB_Name = "frmPopupMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public AGrid As AutoGrid

Private Sub mnuAbortEdit_Click()
  AGrid.TDBGrid.DataChanged = False
  Call AGrid.TDBGrid.Refresh
End Sub

Private Sub mnuCommitEdit_Click()
  Call AGrid.PopupMenuAction(COMMIT_EDIT)
End Sub

Private Sub mnuAutoSize_Click()
  Call AGrid.PopupMenuAction(AUTOSIZE_COLUMN)
End Sub

Private Sub mnuAutoSizeAll_Click()
  Call AGrid.PopupMenuAction(AUTOSIZEALL_COLUMNS)
End Sub

Private Sub mnuCopyCol_Click()
  Call AGrid.PopupMenuAction(MOVE_COLUMN)
End Sub

Private Sub mnuCopyRow_Click()
  Call AGrid.PopupMenuAction(MENU_COPY)
End Sub

Private Sub mnuCutRow_Click()
  Call AGrid.PopupMenuAction(MENU_CUT)
End Sub

Private Sub mnuDebugMenu_Click()
  Call AGrid.PopupMenuAction(MENU_DEBUG)
End Sub

Private Sub mnuFilterSelectionEx_Click()
  Call AGrid.PopupMenuAction(FILTER_SELECTION_EX)
End Sub

Private Sub mnuFilterSelectionInc_Click()
  Call AGrid.PopupMenuAction(FILTER_SELECTION_INC)
End Sub

Private Sub mnuFind_Click()
  Call AGrid.PopupMenuAction(MENU_FIND_INCOL)
End Sub

Private Sub mnuInsCol_Click()
  Call AGrid.PopupMenuAction(INSERT_COLUMN)
End Sub

Private Sub mnuPasteRow_Click()
  Call AGrid.PopupMenuAction(MENU_PASTE)
End Sub

Private Sub mnuRemoveFilter_Click()
  Call AGrid.PopupMenuAction(FILTER_REMOVE)
End Sub

Private Sub mnuSavedFilter_Click(Index As Integer)
  AGrid.ParentAC.SortFilterRefresh = True
  Call AGrid.ParentAC.SetFilterSort(GetIniEntry(AGrid.ParentAC.AutoName, Me.mnuSavedFilter(Index).Caption & "Sort"), GetIniEntry(AGrid.ParentAC.AutoName, Me.mnuSavedFilter(Index).Caption & "Filter"))
End Sub

Private Sub mnuSortAsc_Click()
  Call AGrid.PopupMenuAction(SORT_ASC)
End Sub

Private Sub mnuSortDesc_Click()
  Call AGrid.PopupMenuAction(SORT_DESC)
End Sub

Private Sub mnuUserMenus_Click(Index As Integer)
  Dim MenuHandler As IAutoPopupHandler, Caption As String
  Dim um As UserPopupMenu, vbmk As Variant
  
  Set MenuHandler = AGrid.ParentAC.UserPopupMenuHandler
  If Not MenuHandler Is Nothing Then
    Caption = Me.mnuUserMenus(Index).Caption
    Set um = AGrid.ParentAC.UserPopupMenus(Caption)
    If Not AGrid.DAORecordset Is Nothing Then
      If Not (AGrid.DAORecordset.BOF Or AGrid.DAORecordset.EOF) Then vbmk = AGrid.DAORecordset.Bookmark
    End If
    If Not AGrid.RDOResultset Is Nothing Then
      If Not (AGrid.RDOResultset.BOF Or AGrid.RDOResultset.EOF) Then vbmk = AGrid.RDOResultset.Bookmark
    End If
    Call MenuHandler.MenuAction(um.Caption, um.Tag, AGrid.ParentAC, vbmk)
  End If
End Sub

Private Sub mnuWizard_Click()
  Call AGrid.PopupMenuAction(FILTER_WIZARD)
End Sub
