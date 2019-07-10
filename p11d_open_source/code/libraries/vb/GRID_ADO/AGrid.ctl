VERSION 5.00
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Object = "{89056D22-ECDA-4A64-B90B-25EBB3AE8DB8}#1.0#0"; "ATC2HOOK.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Begin VB.UserControl AutoGridCtrl_ADO 
   ClientHeight    =   4335
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10350
   ScaleHeight     =   4335
   ScaleWidth      =   10350
   ToolboxBitmap   =   "AGrid.ctx":0000
   Begin MSAdodcLib.Adodc datGrid_i 
      Height          =   375
      Left            =   120
      Top             =   2640
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "dat"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
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
      TabIndex        =   4
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
   Begin TrueOleDBGrid60.TDBGrid Grid_i 
      Bindings        =   "AGrid.ctx":0312
      Height          =   2385
      Left            =   0
      OleObjectBlob   =   "AGrid.ctx":032A
      TabIndex        =   0
      Top             =   120
      Width           =   10215
   End
   Begin VB.Label lblFastKey_i 
      Caption         =   "lblFastKey"
      Height          =   360
      Left            =   4365
      TabIndex        =   3
      Top             =   3030
      Width           =   3855
   End
   Begin VB.Label lblsort_i 
      Caption         =   "lblSort"
      Height          =   255
      Left            =   45
      TabIndex        =   2
      Top             =   3720
      Width           =   2535
   End
   Begin VB.Label lblFilter_i 
      Caption         =   "lblFilter"
      Height          =   255
      Left            =   3480
      TabIndex        =   1
      Top             =   3840
      Width           =   6135
   End
End
Attribute VB_Name = "AutoGridCtrl_ADO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Option Compare Text
Private mDataComboCount As Long
Public Event Resize()
Public Event Show()
Public Event BeginDrag()
Public Event FetchRowStyle(ByVal Bookmark As Variant, ByVal RowStyle As TrueOleDBGrid60.StyleDisp)
Public Event MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)

' Events raised from TDBGrid
Public Event AfterColUpdate(ByVal ColIndex As Integer)
Public Event BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
Public Event ButtonClick(ByVal ColIndex As Integer)
Public Event KeyDown(KeyCode As Integer, Shift As Integer)

Private mDragRS As Recordset
Private mAllowDrag As Boolean
Private mShown As Boolean
Private mAllowSetFocus  As Boolean
Private mDCVisible As Boolean
Implements ILibraryVersion

Public Property Get grid() As Object
  Set grid = UserControl.Grid_i
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
  Enabled = UserControl.Grid_i.Enabled
End Property

Public Property Let Enabled(ByVal NewValue As Boolean)
  UserControl.Grid_i.Enabled = NewValue
  PropertyChanged "Enabled"
End Property

Public Property Get AllowAddNew() As Boolean
  AllowAddNew = UserControl.Grid_i.AllowAddNew
End Property

Public Property Let AllowAddNew(ByVal NewValue As Boolean)
  UserControl.Grid_i.AllowAddNew = NewValue
  PropertyChanged "AllowAddNew"
End Property

Public Property Get AllowDelete() As Boolean
  AllowDelete = UserControl.Grid_i.AllowDelete
End Property

Public Property Let AllowDelete(ByVal NewValue As Boolean)
  UserControl.Grid_i.AllowDelete = NewValue
  PropertyChanged "AllowDelete"
End Property

Public Property Get AllowUpdate() As Boolean
  AllowUpdate = UserControl.Grid_i.AllowUpdate
End Property

Public Property Let AllowUpdate(ByVal NewValue As Boolean)
  UserControl.Grid_i.AllowUpdate = NewValue
  PropertyChanged "AllowUpdate"
End Property

Public Property Get LabelSortVisible() As Boolean
  LabelSortVisible = UserControl.lblsort_i.visible
End Property

Public Property Let LabelSortVisible(ByVal NewValue As Boolean)
  UserControl.lblsort_i.visible = NewValue
  PropertyChanged "LabelSortVisible"
  Call ResizeGridControl(UserControl.Width, UserControl.Height, UserControl.Grid_i, UserControl.datGrid_i, UserControl.lblsort_i, UserControl.lblFilter_i, mDCVisible, lblFastKey_i)
End Property

Public Property Get LabelFilterVisible() As Boolean
  LabelFilterVisible = UserControl.lblFilter_i.visible
End Property

Public Property Let LabelFilterVisible(ByVal NewValue As Boolean)
  UserControl.lblFilter_i.visible = NewValue
  PropertyChanged "LabelFilterVisible"
  Call ResizeGridControl(UserControl.Width, UserControl.Height, UserControl.Grid_i, UserControl.datGrid_i, UserControl.lblsort_i, UserControl.lblFilter_i, mDCVisible, lblFastKey_i)
End Property

Public Property Get LabelFastKeyVisible() As Boolean
  LabelFastKeyVisible = UserControl.lblFastKey_i.visible
End Property

Public Property Let LabelFastKeyVisible(ByVal NewValue As Boolean)
  UserControl.lblFastKey_i.visible = NewValue
  PropertyChanged "LabelFastKeyVisible"
  Call ResizeGridControl(UserControl.Width, UserControl.Height, UserControl.Grid_i, UserControl.datGrid_i, UserControl.lblsort_i, UserControl.lblFilter_i, mDCVisible, lblFastKey_i)
End Property

Public Property Get RecordNavigatorVisible() As Boolean
  RecordNavigatorVisible = mDCVisible
End Property

Public Property Let RecordNavigatorVisible(ByVal NewValue As Boolean)
  mDCVisible = NewValue
  PropertyChanged "RecordNavigatorVisible"
  Call SetDCProp(UserControl.datGrid_i, mDCVisible)
  Call ResizeGridControl(UserControl.Width, UserControl.Height, UserControl.Grid_i, UserControl.datGrid_i, UserControl.lblsort_i, UserControl.lblFilter_i, mDCVisible, lblFastKey_i)
End Property

Public Sub Kill()
  Dim i As Long, tdd As TDBDropDown
  
  On Error Resume Next
  Set mDragRS = Nothing
  Call UserControl.Grid_i.Close
  For i = 1 To mDataComboCount
    Set tdd = UserControl.Controls.Item("Combo_" & CStr(i))
    Set tdd.DataSource = Nothing
    Call UserControl.Controls.Remove("Combo_" & CStr(i))
  Next i
  mDataComboCount = 0
  'Set UserControl.datGrid_i.RecordSource = Nothing
End Sub

Public Function AddDataCombo() As Long
  Dim tdd As TDBDropDown
  mDataComboCount = mDataComboCount + 1
  Set tdd = UserControl.Controls.Add("TrueOleDBGrid60.TDBDropDown", "Combo_" & CStr(mDataComboCount))
  AddDataCombo = mDataComboCount
End Function

Public Property Get Combo(ByVal Index As Long) As Object
  If (Index < 1) Or (Index > mDataComboCount) Then Err.Raise 380, "Combo", "Invalid Auto Control Combo index " & Index
  Set Combo = UserControl.Controls("Combo_" & CStr(Index))
End Property

Private Sub Grid_i_FetchRowStyle(ByVal Split As Integer, Bookmark As Variant, ByVal RowStyle As TrueOleDBGrid60.StyleDisp)
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

Private Sub Grid_i_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
  Dim hDepth  As Single
  
  If CheckMouseButton(VK_LBUTTON) Then
    hDepth = 250 * Grid_i.HeadLines
    If (Grid_i.ColContaining(x) = -1) And (x < 250) And (Grid_i.RowContaining(Y) = -1) And (Y < hDepth) Then
      Call SelectAll
    End If
  End If
  Call trySetGridFocus
  RaiseEvent MouseDown(Button, Shift, x, Y)
End Sub

Private Sub trySetGridFocus()
  If Not mAllowSetFocus Then Exit Sub
  On Error Resume Next
  Grid_i.SetFocus
End Sub

Private Sub UserControl_Initialize()
  mShown = False
  mDCVisible = True
  mDataComboCount = 0
  mAllowSetFocus = True
End Sub

Private Sub UserControl_Terminate()
  Debug.Print "Term AGRID"
End Sub

Private Sub UserControl_Resize()
  If mShown Then Call ResizeGridControl(UserControl.Width, UserControl.Height, UserControl.Grid_i, UserControl.datGrid_i, UserControl.lblsort_i, UserControl.lblFilter_i, mDCVisible, UserControl.lblFastKey_i)
  mAllowSetFocus = True
End Sub

Private Sub UserControl_Show()
  mShown = True
  mAllowSetFocus = True
  Call SetDCProp(UserControl.datGrid_i, mDCVisible)
  Call ResizeGridControl(UserControl.Width, UserControl.Height, UserControl.Grid_i, UserControl.datGrid_i, UserControl.lblsort_i, UserControl.lblFilter_i, mDCVisible, lblFastKey_i)
  UserControl.KeyHook.hwnd = UserControl.Grid_i.hwnd
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
  Call WriteProperties(PropBag, UserControl.Grid_i, mDCVisible, UserControl.lblsort_i, UserControl.lblFilter_i, UserControl.lblFastKey_i)
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
  Call ReadProperties(PropBag, UserControl.Grid_i, mDCVisible, UserControl.lblsort_i, UserControl.lblFilter_i, UserControl.lblFastKey_i)
End Sub

Private Sub Grid_i_DragCell(ByVal SplitIndex As Integer, RowBookmark As Variant, ByVal ColIndex As Integer)
  If mAllowDrag Then
    Set mDragRS = GridDragCell(UserControl.Grid_i, UserControl.datGrid_i, RowBookmark)
    If Not mDragRS Is Nothing Then RaiseEvent BeginDrag
  End If
End Sub

Public Property Get DraggedRecordset() As Recordset
  Set DraggedRecordset = mDragRS
End Property

Public Function TrackDragDrop(ByVal x As Single, ByVal Y As Single) As Variant
  Dim rIndex As Integer, cIndex As Integer
  Dim vbmk As Variant, vbmkfirst As Variant
  Dim ColIndex As Long
  Dim vScrollHeight As Long
  Static InTrackDragDrop As Boolean
  
  On Error Resume Next
  If InTrackDragDrop Then Exit Function
  InTrackDragDrop = True
  TrackDragDrop = Null
  cIndex = Grid_i.ColContaining(x)
  rIndex = Grid_i.RowContaining(Y)
  If rIndex = -1 Then
    Call ClearSelRows(Grid_i)
  Else
    vbmkfirst = Grid_i.FirstRow
    vbmk = Grid_i.RowBookmark(rIndex)
    If IsNull(vbmk) Then Grid_i.Bookmark = vbmk
    If cIndex <> -1 Then Grid_i.Col = cIndex
    If rIndex = 0 Then
      Call Grid_i.Scroll(0, -1)
      vScrollHeight = 1
    End If
    If rIndex = (Grid_i.VisibleRows - 1) Then
      Call Grid_i.Scroll(0, 1)
      vScrollHeight = -1
    End If
    If vbmkfirst = Grid_i.FirstRow Then vScrollHeight = 0
    If vScrollHeight <> 0 Then Call MoveMouseCursor(0, vScrollHeight * 2)
    If (cIndex > 0) And (cIndex < (Grid_i.Columns.Count - 1)) Then
      ColIndex = cIndex - Grid_i.LeftCol
      If (ColIndex = (Grid_i.VisibleCols - 1)) Then Call Grid_i.Scroll(1, 0)
      If cIndex = Grid_i.LeftCol Then Call Grid_i.Scroll(-1, 0)
    End If
    Grid_i.CurrentCellVisible = True
    TrackDragDrop = vbmk
  End If
  InTrackDragDrop = False
End Function

Public Property Get AllowDrag() As Boolean
  AllowDrag = mAllowDrag
End Property

Public Property Let AllowDrag(ByVal NewVal As Boolean)
  mAllowDrag = NewVal
End Property

' no errors if select fails
Public Sub SelectAll()
  Dim rsClone As Recordset, AllBookmarks As SelBookmarks
  
  On Error GoTo SelectAll_err
  Call SetCursor
  If Not ConfirmSelectAll(datGrid_i.Recordset, UserControl.Parent) Then GoTo SelectAll_end
  Set rsClone = datGrid_i.Recordset.Clone
  Set AllBookmarks = Grid_i.SelBookmarks
  Do While AllBookmarks.Count > 0
    Call AllBookmarks.Remove(0)
  Loop
  
  ' retrieve First Grid bookmark
  If Not (rsClone.EOF And rsClone.BOF) Then
    rsClone.MoveFirst
    Do
      Call AllBookmarks.Add(rsClone.Bookmark)
      rsClone.MoveNext
    Loop Until rsClone.EOF
  End If
  
SelectAll_end:
  Call ClearCursor
  Exit Sub
  
SelectAll_err:
  'Err.Raise ERR_SELECTALL, "SelectAll", "Error selecting all grid rows"
  Resume SelectAll_end
End Sub

Private Property Get ILibraryVersion_Name() As String
  ILibraryVersion_Name = "ADO Auto Control"
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

