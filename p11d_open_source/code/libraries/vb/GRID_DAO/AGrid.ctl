VERSION 5.00
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "tdbg6.ocx"
Object = "{89056D22-ECDA-4A64-B90B-25EBB3AE8DB8}#1.0#0"; "ATC2HOOK.OCX"
Begin VB.UserControl AutoGridCtrl_DAO 
   ClientHeight    =   4245
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10350
   ScaleHeight     =   4245
   ScaleWidth      =   10350
   ToolboxBitmap   =   "AGrid.ctx":0000
   Begin TrueDBGrid60.TDBGrid grid_i 
      Bindings        =   "AGrid.ctx":0312
      Height          =   2580
      Left            =   45
      OleObjectBlob   =   "AGrid.ctx":032A
      TabIndex        =   4
      Top             =   45
      Width           =   10230
   End
   Begin VB.PictureBox picTest 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   8730
      ScaleHeight     =   735
      ScaleWidth      =   1005
      TabIndex        =   3
      Top             =   2925
      Visible         =   0   'False
      Width           =   1005
   End
   Begin atc2hook.HOOK KeyHook 
      Left            =   2280
      Top             =   3120
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin VB.Data Data_16 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Enabled         =   0   'False
      Exclusive       =   0   'False
      Height          =   345
      Left            =   5280
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10680
      Visible         =   0   'False
      Width           =   1545
   End
   Begin VB.Data Data_15 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Enabled         =   0   'False
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3600
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10680
      Visible         =   0   'False
      Width           =   1545
   End
   Begin VB.Data Data_13 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Enabled         =   0   'False
      Exclusive       =   0   'False
      Height          =   345
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10680
      Visible         =   0   'False
      Width           =   1545
   End
   Begin VB.Data Data_14 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Enabled         =   0   'False
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1800
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   10680
      Visible         =   0   'False
      Width           =   1545
   End
   Begin VB.Data Data_12 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Enabled         =   0   'False
      Exclusive       =   0   'False
      Height          =   345
      Left            =   8640
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   9240
      Visible         =   0   'False
      Width           =   1545
   End
   Begin VB.Data Data_11 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Enabled         =   0   'False
      Exclusive       =   0   'False
      Height          =   345
      Left            =   7080
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   9240
      Visible         =   0   'False
      Width           =   1545
   End
   Begin VB.Data Data_10 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Enabled         =   0   'False
      Exclusive       =   0   'False
      Height          =   345
      Left            =   5400
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   9240
      Visible         =   0   'False
      Width           =   1545
   End
   Begin VB.Data Data_9 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Enabled         =   0   'False
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3720
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   9240
      Visible         =   0   'False
      Width           =   1545
   End
   Begin VB.Data Data_8 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Enabled         =   0   'False
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1920
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   9240
      Visible         =   0   'False
      Width           =   1545
   End
   Begin VB.Data Data_6 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Enabled         =   0   'False
      Exclusive       =   0   'False
      Height          =   345
      Left            =   8640
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   7920
      Visible         =   0   'False
      Width           =   1545
   End
   Begin VB.Data Data_7 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Enabled         =   0   'False
      Exclusive       =   0   'False
      Height          =   345
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   9240
      Visible         =   0   'False
      Width           =   1545
   End
   Begin VB.Data Data_5 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Enabled         =   0   'False
      Exclusive       =   0   'False
      Height          =   345
      Left            =   6960
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   7920
      Visible         =   0   'False
      Width           =   1545
   End
   Begin VB.Data Data_4 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Enabled         =   0   'False
      Exclusive       =   0   'False
      Height          =   345
      Left            =   5160
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   7920
      Visible         =   0   'False
      Width           =   1545
   End
   Begin VB.Data Data_3 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Enabled         =   0   'False
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   7920
      Visible         =   0   'False
      Width           =   1545
   End
   Begin VB.Data Data_1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Enabled         =   0   'False
      Exclusive       =   0   'False
      Height          =   345
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   7920
      Visible         =   0   'False
      Width           =   1545
   End
   Begin VB.Data Data_2 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Enabled         =   0   'False
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1800
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   7920
      Visible         =   0   'False
      Width           =   1545
   End
   Begin VB.Data datGrid_i 
      Caption         =   "dat"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3120
      Width           =   1965
   End
   Begin TrueDBGrid60.TDBDropDown Combo_1 
      Bindings        =   "AGrid.ctx":2D22
      Height          =   930
      Left            =   45
      OleObjectBlob   =   "AGrid.ctx":2D37
      TabIndex        =   5
      Top             =   6795
      Visible         =   0   'False
      Width           =   1560
   End
   Begin TrueDBGrid60.TDBDropDown Combo_2 
      Bindings        =   "AGrid.ctx":4F58
      Height          =   930
      Left            =   1845
      OleObjectBlob   =   "AGrid.ctx":4F6D
      TabIndex        =   6
      Top             =   6915
      Visible         =   0   'False
      Width           =   1560
   End
   Begin TrueDBGrid60.TDBDropDown Combo_3 
      Bindings        =   "AGrid.ctx":718E
      Height          =   930
      Left            =   3525
      OleObjectBlob   =   "AGrid.ctx":71A3
      TabIndex        =   7
      Top             =   6915
      Visible         =   0   'False
      Width           =   1560
   End
   Begin TrueDBGrid60.TDBDropDown Combo_4 
      Bindings        =   "AGrid.ctx":93C4
      Height          =   930
      Left            =   5205
      OleObjectBlob   =   "AGrid.ctx":93D9
      TabIndex        =   8
      Top             =   6915
      Visible         =   0   'False
      Width           =   1560
   End
   Begin TrueDBGrid60.TDBDropDown Combo_5 
      Bindings        =   "AGrid.ctx":B5FA
      Height          =   930
      Left            =   7005
      OleObjectBlob   =   "AGrid.ctx":B60F
      TabIndex        =   9
      Top             =   6915
      Visible         =   0   'False
      Width           =   1560
   End
   Begin TrueDBGrid60.TDBDropDown Combo_6 
      Bindings        =   "AGrid.ctx":D830
      Height          =   930
      Left            =   8685
      OleObjectBlob   =   "AGrid.ctx":D845
      TabIndex        =   10
      Top             =   6915
      Visible         =   0   'False
      Width           =   1560
   End
   Begin TrueDBGrid60.TDBDropDown Combo_7 
      Bindings        =   "AGrid.ctx":FA66
      Height          =   930
      Left            =   45
      OleObjectBlob   =   "AGrid.ctx":FA7B
      TabIndex        =   11
      Top             =   8235
      Visible         =   0   'False
      Width           =   1560
   End
   Begin TrueDBGrid60.TDBDropDown Combo_8 
      Bindings        =   "AGrid.ctx":11C9C
      Height          =   930
      Left            =   1845
      OleObjectBlob   =   "AGrid.ctx":11CB1
      TabIndex        =   12
      Top             =   8235
      Visible         =   0   'False
      Width           =   1560
   End
   Begin TrueDBGrid60.TDBDropDown Combo_9 
      Bindings        =   "AGrid.ctx":13ED2
      Height          =   930
      Left            =   3645
      OleObjectBlob   =   "AGrid.ctx":13EE7
      TabIndex        =   13
      Top             =   8235
      Visible         =   0   'False
      Width           =   1560
   End
   Begin TrueDBGrid60.TDBDropDown Combo_10 
      Bindings        =   "AGrid.ctx":16108
      Height          =   930
      Left            =   5325
      OleObjectBlob   =   "AGrid.ctx":1611E
      TabIndex        =   14
      Top             =   8235
      Visible         =   0   'False
      Width           =   1560
   End
   Begin TrueDBGrid60.TDBDropDown Combo_11 
      Bindings        =   "AGrid.ctx":18340
      Height          =   930
      Left            =   7005
      OleObjectBlob   =   "AGrid.ctx":18356
      TabIndex        =   15
      Top             =   8235
      Visible         =   0   'False
      Width           =   1560
   End
   Begin TrueDBGrid60.TDBDropDown Combo_12 
      Bindings        =   "AGrid.ctx":1A578
      Height          =   930
      Left            =   8685
      OleObjectBlob   =   "AGrid.ctx":1A58E
      TabIndex        =   16
      Top             =   8235
      Visible         =   0   'False
      Width           =   1560
   End
   Begin TrueDBGrid60.TDBDropDown Combo_13 
      Bindings        =   "AGrid.ctx":1C7B0
      Height          =   930
      Left            =   165
      OleObjectBlob   =   "AGrid.ctx":1C7C6
      TabIndex        =   17
      Top             =   9675
      Visible         =   0   'False
      Width           =   1560
   End
   Begin TrueDBGrid60.TDBDropDown Combo_14 
      Bindings        =   "AGrid.ctx":1E9E8
      Height          =   930
      Left            =   1845
      OleObjectBlob   =   "AGrid.ctx":1E9FE
      TabIndex        =   18
      Top             =   9675
      Visible         =   0   'False
      Width           =   1560
   End
   Begin TrueDBGrid60.TDBDropDown Combo_15 
      Bindings        =   "AGrid.ctx":20C20
      Height          =   930
      Left            =   3645
      OleObjectBlob   =   "AGrid.ctx":20C36
      TabIndex        =   19
      Top             =   9675
      Visible         =   0   'False
      Width           =   1560
   End
   Begin TrueDBGrid60.TDBDropDown Combo_16 
      Bindings        =   "AGrid.ctx":22E58
      Height          =   930
      Left            =   5325
      OleObjectBlob   =   "AGrid.ctx":22E6E
      TabIndex        =   20
      Top             =   9675
      Visible         =   0   'False
      Width           =   1560
   End
   Begin VB.Label lblFastKey_i 
      Caption         =   "lblFastKey"
      Height          =   360
      Left            =   4365
      TabIndex        =   2
      Top             =   3015
      Width           =   3855
   End
   Begin VB.Label lblsort_i 
      Caption         =   "lblSort"
      Height          =   255
      Left            =   45
      TabIndex        =   1
      Top             =   3720
      Width           =   2535
   End
   Begin VB.Label lblFilter_i 
      Caption         =   "lblFilter"
      Height          =   255
      Left            =   3480
      TabIndex        =   0
      Top             =   3840
      Width           =   6135
   End
End
Attribute VB_Name = "AutoGridCtrl_DAO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Option Compare Text

Implements ILibraryVersion

Private mDataComboCount As Long
Public Event Resize()
Public Event Show()
Public Event BeginDrag()
Public Event FetchRowStyle(ByVal Bookmark As Variant, ByVal RowStyle As TrueDBGrid60.StyleDisp)

' Events raised from TDBGrid
Public Event AfterColUpdate(ByVal ColIndex As Integer)
Public Event BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
Public Event ButtonClick(ByVal ColIndex As Integer)
Public Event KeyDown(KeyCode As Integer, Shift As Integer)

Private mRecords As Collection
Private mAllowDrag As Boolean
Private mShown As Boolean
Private mAllowSetFocus As Boolean
Private mDCVisible As Boolean

Public Property Get grid() As Object
  Set grid = UserControl.grid_i
End Property

Public Property Get GridDataControl() As Object
  Set GridDataControl = UserControl.datGrid_i
End Property

Public Property Get FastKeyLabel() As Object
  Set FastKeyLabel = UserControl.lblFastKey_i
End Property

Public Property Get SortLabel() As Object
  Set SortLabel = UserControl.lblsort_i
End Property

Public Property Get FilterLabel() As Object
  Set FilterLabel = UserControl.lblFilter_i
End Property

Public Property Get ContainerForm() As Object
  Set ContainerForm = UserControl.Parent
End Property

Public Property Get picTest() As Object
  Set picTest = UserControl.picTest
End Property

Public Property Get Enabled() As Boolean
  Enabled = UserControl.grid_i.Enabled
End Property

Public Property Let Enabled(ByVal NewValue As Boolean)
  UserControl.grid_i.Enabled = NewValue
  PropertyChanged "Enabled"
End Property

Public Property Get AllowAddNew() As Boolean
  AllowAddNew = UserControl.grid_i.AllowAddNew
End Property

Public Property Let AllowAddNew(ByVal NewValue As Boolean)
  UserControl.grid_i.AllowAddNew = NewValue
  PropertyChanged "AllowAddNew"
End Property

Public Property Get AllowDelete() As Boolean
  AllowDelete = UserControl.grid_i.AllowDelete
End Property

Public Property Let AllowDelete(ByVal NewValue As Boolean)
  UserControl.grid_i.AllowDelete = NewValue
  PropertyChanged "AllowDelete"
End Property

Public Property Get AllowUpdate() As Boolean
  AllowUpdate = UserControl.grid_i.AllowUpdate
End Property

Public Property Let AllowUpdate(ByVal NewValue As Boolean)
  UserControl.grid_i.AllowUpdate = NewValue
  PropertyChanged "AllowUpdate"
End Property

Public Function AddDataCombo() As Long
  mDataComboCount = mDataComboCount + 1
  If mDataComboCount > DataComboCount Then Err.Raise 380, "AddDataCombo", "Unable to add bound Dropdown. Maximum number of bound dropdowns is " & CStr(DataComboCount)
  AddDataCombo = mDataComboCount
End Function

Public Property Get LabelSortVisible() As Boolean
  LabelSortVisible = UserControl.lblsort_i.visible
End Property

Public Property Let LabelSortVisible(ByVal NewValue As Boolean)
  UserControl.lblsort_i.visible = NewValue
  PropertyChanged "LabelSortVisible"
  Call ResizeGridControl(UserControl.Width, UserControl.Height, UserControl.grid_i, UserControl.datGrid_i, UserControl.lblsort_i, UserControl.lblFilter_i, mDCVisible, lblFastKey_i)
End Property

Public Property Get LabelFilterVisible() As Boolean
  LabelFilterVisible = UserControl.lblFilter_i.visible
End Property

Public Property Let LabelFilterVisible(ByVal NewValue As Boolean)
  UserControl.lblFilter_i.visible = NewValue
  PropertyChanged "LabelFilterVisible"
  Call ResizeGridControl(UserControl.Width, UserControl.Height, UserControl.grid_i, UserControl.datGrid_i, UserControl.lblsort_i, UserControl.lblFilter_i, mDCVisible, lblFastKey_i)
End Property

Public Property Get LabelFastKeyVisible() As Boolean
  LabelFastKeyVisible = UserControl.lblFastKey_i.visible
End Property

Public Property Let LabelFastKeyVisible(ByVal NewValue As Boolean)
  UserControl.lblFastKey_i.visible = NewValue
  PropertyChanged "LabelFastKeyVisible"
  Call ResizeGridControl(UserControl.Width, UserControl.Height, UserControl.grid_i, UserControl.datGrid_i, UserControl.lblsort_i, UserControl.lblFilter_i, mDCVisible, lblFastKey_i)
End Property

Public Property Get RecordNavigatorVisible() As Boolean
  RecordNavigatorVisible = mDCVisible
End Property

Public Property Let RecordNavigatorVisible(ByVal NewValue As Boolean)
  mDCVisible = NewValue
  PropertyChanged "RecordNavigatorVisible"
  Call SetDCProp(UserControl.datGrid_i, mDCVisible)
  Call ResizeGridControl(UserControl.Width, UserControl.Height, UserControl.grid_i, UserControl.datGrid_i, UserControl.lblsort_i, UserControl.lblFilter_i, mDCVisible, lblFastKey_i)
End Property

Public Sub Kill()
  Dim i As Long, d As Data
  
  On Error Resume Next
  Set mRecords = Nothing
  Call UserControl.grid_i.Close
  For i = 1 To mDataComboCount
    Set d = UserControl.Controls("Data_" & CStr(i))
    Call d.Recordset.Close
    Set d.Recordset = Nothing
    d.Enabled = False
  Next i
  mDataComboCount = 0
  Call UserControl.datGrid_i.Recordset.Close
  Set UserControl.datGrid_i.Recordset = Nothing
  Debug.Print "Kill ACTRL"
End Sub

Public Property Get Combo(ByVal Index As Long) As Object
  If (Index < 1) Or (Index > DataComboCount) Then Err.Raise 380
  Set Combo = UserControl.Controls("Combo_" & CStr(Index))
End Property

Public Property Get DataCombo(ByVal Index As Long) As Object
  If (Index < 1) Or (Index > DataComboCount) Then Err.Raise 380
  Set DataCombo = UserControl.Controls("Data_" & CStr(Index))
End Property

Private Sub Grid_i_FetchRowStyle(ByVal Split As Integer, Bookmark As Variant, ByVal RowStyle As TrueDBGrid60.StyleDisp)
  RaiseEvent FetchRowStyle(Bookmark, RowStyle)
End Sub

Private Sub grid_i_GotFocus()
  Debug.Print "Grid gf"
  mAllowSetFocus = False
End Sub

Private Sub grid_i_LostFocus()
  Debug.Print "Grid lf"
  mAllowSetFocus = True
End Sub

Private Sub Grid_i_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim hDepth  As Single
  
  If CheckMouseButton(VK_LBUTTON) Then
    hDepth = 250 * grid_i.HeadLines
    If (grid_i.ColContaining(X) = -1) And (X < 250) And (grid_i.RowContaining(Y) = -1) And (Y < hDepth) Then
      Call SelectAll
    End If
  End If
  Call trySetGridFocus
End Sub

Private Sub trySetGridFocus()
  If Not mAllowSetFocus Then Exit Sub
  On Error Resume Next
  grid_i.SetFocus
End Sub

Private Sub UserControl_Initialize()
  mShown = False
  mDCVisible = True
  mAllowSetFocus = True
  Debug.Print "Init ACTRL"
End Sub

Private Sub UserControl_Terminate()
  Debug.Print "Term ACTRL"
End Sub

Private Sub UserControl_Resize()
  If mShown Then Call ResizeGridControl(UserControl.Width, UserControl.Height, UserControl.grid_i, UserControl.datGrid_i, UserControl.lblsort_i, UserControl.lblFilter_i, mDCVisible, UserControl.lblFastKey_i)
  mAllowSetFocus = True
End Sub

Private Sub UserControl_Show()
  mShown = True
  mAllowSetFocus = True
  Call SetDCProp(UserControl.datGrid_i, mDCVisible)
  Call ResizeGridControl(UserControl.Width, UserControl.Height, UserControl.grid_i, UserControl.datGrid_i, UserControl.lblsort_i, UserControl.lblFilter_i, mDCVisible, lblFastKey_i)
  UserControl.KeyHook.hwnd = UserControl.grid_i.hwnd
  UserControl.KeyHook.Messages(WM_CHAR) = True
  RaiseEvent Show
End Sub

Private Sub KeyHook_WndProc(Discard As Boolean, MsgReturn As Long, ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long)
  If (wParam = CTRL_KEY_C) Or (wParam = CTRL_KEY_V) Or (wParam = CTRL_KEY_X) Then
    Discard = True
    MsgReturn = 0
  End If
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
  Call WriteProperties(PropBag, UserControl.grid_i, mDCVisible, UserControl.lblsort_i, UserControl.lblFilter_i, UserControl.lblFastKey_i)
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
  Call ReadProperties(PropBag, UserControl.grid_i, mDCVisible, UserControl.lblsort_i, UserControl.lblFilter_i, UserControl.lblFastKey_i)
End Sub

Private Sub Grid_i_DragCell(ByVal SplitIndex As Integer, RowBookmark As Variant, ByVal ColIndex As Integer)
  Set mRecords = GridDragCell(UserControl.grid_i, Nothing, UserControl.datGrid_i, RowBookmark, ColIndex, mAllowDrag)
  If Not mRecords Is Nothing Then RaiseEvent BeginDrag
End Sub

Public Function TrackDragDrop(ByVal X As Single, ByVal Y As Single) As Variant
  Dim rIndex As Integer, cIndex As Integer
  Dim vbmk As Variant, vbmkfirst As Variant
  Dim ColIndex As Long
  Dim vScrollHeight As Long
  Static InTrackDragDrop As Boolean
  
  On Error Resume Next
  If InTrackDragDrop Then Exit Function
  InTrackDragDrop = True
  TrackDragDrop = ""
  cIndex = grid_i.ColContaining(X)
  rIndex = grid_i.RowContaining(Y)
  If rIndex = -1 Then
    Call ClearSelRows(grid_i)
  Else
    vbmkfirst = grid_i.FirstRow
    vbmk = grid_i.RowBookmark(rIndex)
    If Len(vbmk) > 0 Then grid_i.Bookmark = vbmk
    If cIndex <> -1 Then grid_i.Col = cIndex
    If rIndex = 0 Then
      Call grid_i.Scroll(0, -1)
      vScrollHeight = 1
    End If
    If rIndex = (grid_i.VisibleRows - 1) Then
      Call grid_i.Scroll(0, 1)
      vScrollHeight = -1
    End If
    If vbmkfirst = grid_i.FirstRow Then vScrollHeight = 0
    If vScrollHeight <> 0 Then Call MoveMouseCursor(0, vScrollHeight * 2)
    If (cIndex > 0) And (cIndex < (grid_i.Columns.Count - 1)) Then
      ColIndex = cIndex - grid_i.LeftCol
      If (ColIndex = (grid_i.VisibleCols - 1)) Then Call grid_i.Scroll(1, 0)
      If cIndex = grid_i.LeftCol Then Call grid_i.Scroll(-1, 0)
    End If
    grid_i.CurrentCellVisible = True
    TrackDragDrop = vbmk
  End If
  InTrackDragDrop = False
End Function

Public Property Get DraggedRecords() As Collection
  Set DraggedRecords = mRecords
End Property

Public Property Get AllowDrag() As Boolean
  AllowDrag = mAllowDrag
End Property

Public Property Let AllowDrag(ByVal NewVal As Boolean)
  mAllowDrag = NewVal
End Property

' no errors if select fails
Public Sub SelectAll()
  Dim vbmk As Variant
  Dim iRow As Long
  
  On Error GoTo SelectAll_err
  Call SetCursor
  If Not ConfirmSelectAll(Nothing, datGrid_i.Recordset, UserControl.Parent) Then GoTo SelectAll_end
  Do While grid_i.SelBookmarks.Count > 0
    Call grid_i.SelBookmarks.Remove(0)
  Loop
  
  ' retrieve First Grid bookmark
  iRow = 0
  Do
    iRow = iRow - 1
    vbmk = grid_i.GetBookmark(iRow)
  Loop Until IsNull(vbmk)
  iRow = iRow + 1
  
  ' go through all rows adding to SelBookmarks collection
  vbmk = grid_i.GetBookmark(iRow)
  Do While Not IsNull(vbmk)
    Call grid_i.SelBookmarks.Add(vbmk)
    iRow = iRow + 1
    vbmk = grid_i.GetBookmark(iRow)
  Loop
  
SelectAll_end:
  Call ClearCursor
  Exit Sub
  
SelectAll_err:
  'Err.Raise ERR_SELECTALL, "SelectAll", "Error selecting all grid rows"
  Resume SelectAll_end
End Sub

Private Property Get ILibraryVersion_Name() As String
  ILibraryVersion_Name = "DAO Auto Control"
End Property

Private Property Get ILibraryVersion_Version() As String
  ILibraryVersion_Version = App.Major & "." & App.Minor & "." & App.Revision
End Property


Private Sub Grid_i_AfterColUpdate(ByVal ColIndex As Integer)
  RaiseEvent AfterColUpdate(ColIndex)
End Sub

Private Sub Grid_i_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
  RaiseEvent BeforeColUpdate(ColIndex, OldValue, Cancel)
End Sub

Private Sub Grid_i_ButtonClick(ByVal ColIndex As Integer)
  RaiseEvent ButtonClick(ColIndex)
End Sub

Private Sub Grid_i_KeyDown(KeyCode As Integer, Shift As Integer)
  RaiseEvent KeyDown(KeyCode, Shift)
End Sub


