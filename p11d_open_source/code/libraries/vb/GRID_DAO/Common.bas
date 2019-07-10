Attribute VB_Name = "Common"
Option Explicit
Public Const SELECT_ALL_LIMIT_RDO As Long = 100
Public Const SELECT_ALL_LIMIT_DAO As Long = 1000
Private Const STATUS_HEIGHT As Single = 365
Private Const OFFSET As Single = 10
Public Const DataComboCount As Long = 16


Public Const CTRL_KEY_C As Integer = &H3
Public Const CTRL_KEY_V As Integer = &H16
Public Const CTRL_KEY_X As Integer = &H18
Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Public Const VK_CONTROL As Long = &H11
Public Const WM_CHAR As Long = &H102

Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
Public Type POINTAPI
        X As Long
        Y As Long
End Type

Public Function ConfirmSelectAll(ByVal rsRDO As rdoResultset, ByVal rs As Recordset, ByVal frm As Form) As Boolean
  Dim rCount As Long
  Dim rLimit As Long
  On Error GoTo ConfirmSelectAll_err
  
  ConfirmSelectAll = True
  If Not rsRDO Is Nothing Then
    rCount = rsRDO.RowCount
    rLimit = SELECT_ALL_LIMIT_RDO
  End If
  If Not rs Is Nothing Then
    rCount = rs.RecordCount
    rLimit = SELECT_ALL_LIMIT_DAO
  End If
  If rCount > rLimit Then
    ConfirmSelectAll = DisplayMessage(frm, "There are " & rCount & " rows in the current grid." & vbCrLf & "Selecting all rows will take a long time" & vbCrLf & "Do you wish to continue?", "Select all rows", "Yes", "No")
  End If
  
ConfirmSelectAll_end:
  Exit Function
  
ConfirmSelectAll_err:
  ConfirmSelectAll = False
  Resume ConfirmSelectAll_end
End Function
  

Public Sub ResizeGridControl(ByVal Width As Single, ByVal Height As Single, ByVal grd As TDBGrid, ByVal dc As Control, ByVal lblsort As Label, ByVal lblfilter As Label, ByVal dcVisible As Boolean, ByVal lblFastKey As Label)
  Dim GridHeight As Single, dcWidth As Single
  Dim xGridWidth As Single
  Dim tmp As Single
  Const DC_WIDTH As Single = 2500
  Const FK_WIDTH As Single = 1500
  
  On Error Resume Next
  xGridWidth = Width - 2 * OFFSET
  If dcVisible Then xGridWidth = xGridWidth - (DC_WIDTH + OFFSET)
  If lblFastKey.visible Then xGridWidth = xGridWidth - (FK_WIDTH + OFFSET)
  If lblfilter.visible And lblsort.visible Then xGridWidth = xGridWidth / 2
  
  If dcVisible Or lblfilter.visible Or lblsort.visible Or lblFastKey.visible Then
    GridHeight = Height - STATUS_HEIGHT - (OFFSET * 2)
  Else
    GridHeight = Height
  End If
  If (xGridWidth <= 0) Or (GridHeight <= 0) Then Exit Sub
    
  'GRID
  Call grd.Move(0, 0, Width, GridHeight)
  
  'DATACONTROL
  If dcVisible Then
    Call dc.Move(0, GridHeight + OFFSET, DC_WIDTH, STATUS_HEIGHT)
  End If
  
  If lblFastKey.visible Then
    Call lblFastKey.Move(grd.Width - FK_WIDTH - OFFSET, GridHeight + OFFSET, FK_WIDTH, STATUS_HEIGHT)
  End If
  
  'LABELS
  If lblsort.visible Then
    If dcVisible Then
      tmp = dc.Left + dc.Width + OFFSET
    Else
      tmp = 0
    End If
    Call lblsort.Move(tmp, GridHeight + OFFSET, xGridWidth, STATUS_HEIGHT)
  End If
  
  If lblfilter.visible Then
    If lblsort.visible Then
      tmp = lblsort.Left + lblsort.Width + OFFSET
    ElseIf dcVisible Then
      tmp = dc.Left + dc.Width + OFFSET
    Else
      tmp = 0
    End If
    Call lblfilter.Move(tmp, GridHeight + OFFSET, xGridWidth, STATUS_HEIGHT)
  End If
End Sub

Public Function IsBookmarkSelected(ByVal grd As TDBGrid, ByVal vbmk As Variant) As Boolean
  Dim i As Long
  
  IsBookmarkSelected = False
  For i = 0 To grd.SelBookmarks.Count - 1
    If grd.SelBookmarks(i) = vbmk Then
      IsBookmarkSelected = True
      Exit For
    End If
  Next i
End Function

Public Sub ClearSelRows(ByVal grd As TDBGrid)
  Do While grd.SelBookmarks.Count > 0
    Call grd.SelBookmarks.Remove(0)
  Loop
End Sub

Public Function GridDragCell(ByVal grd As TDBGrid, ByVal rdoDC As Object, ByVal daoDC As Object, ByVal RowBookmark As Variant, ByVal ColIndex As Integer, ByVal AllowDrag As Boolean) As Collection
  Dim iFld As FieldDetails, SelectedField As String
  Dim fld As Field, col As rdoColumn
  Dim i As Long, RowBookmarkFound As Boolean, bMax As Integer
  Dim Fields As Collection, Records As Collection
  
  On Error GoTo GridDragCell_err
  If Not AllowDrag Then GoTo GridDragCell_end
  If grd.DataChanged Then Beep: GoTo GridDragCell_end
  
  Set Records = New Collection
  grd.ClearSelCols
  RowBookmarkFound = IsBookmarkSelected(grd, RowBookmark)
  If Not RowBookmarkFound Then
    Call ClearSelRows(grd)
    Call grd.SelBookmarks.Add(RowBookmark)
  End If

  SelectedField = grd.Columns(ColIndex).DataField
  bMax = grd.SelBookmarks.Count - 1
  For i = 0 To bMax
    Set Fields = New Collection
    If Not daoDC Is Nothing Then
      daoDC.Recordset.Bookmark = grd.SelBookmarks(i)
      For Each fld In daoDC.Recordset.Fields
        Set iFld = New FieldDetails
        iFld.Name = fld.Name
        iFld.Value = fld.Value
        iFld.DataType = DAOtoDatatype(fld.Type)
        iFld.Tag = (StrComp(SelectedField, iFld.Name, vbTextCompare) = 0) And (grd.SelBookmarks(i) = RowBookmark)
        Call Fields.Add(iFld, iFld.Name)
      Next fld
    End If
    If Not rdoDC Is Nothing Then
      rdoDC.Resultset.Bookmark = grd.SelBookmarks(i)
      For Each col In rdoDC.Resultset.rdoColumns
        Set iFld = New FieldDetails
        iFld.Name = col.Name
        iFld.Value = col.Value
        iFld.DataType = RDOtoDatatype(col.Type)
        iFld.Tag = (StrComp(SelectedField, iFld.Name, vbTextCompare) = 0) And (grd.SelBookmarks(i) = RowBookmark)
        Call Fields.Add(iFld, iFld.Name)
      Next col
    End If
    Call Records.Add(Fields)
  Next
  
  Set GridDragCell = Records
GridDragCell_end:
  Exit Function
  
GridDragCell_err:
  Resume GridDragCell_end
End Function

Public Sub SetDCProp(ByVal dc As Control, ByVal visible As Boolean)
  dc.visible = visible
End Sub

Public Sub WriteProperties(ByVal PropBag As PropertyBag, ByVal grd As TDBGrid, ByVal dcVisible As Boolean, ByVal lblsort As Label, ByVal lblfilter As Label, ByVal lblfk As Label)
  On Error Resume Next
  Call PropBag.WriteProperty("Enabled", grd.Enabled, True)
  Call PropBag.WriteProperty("AllowAddNew", grd.AllowAddNew, False)
  Call PropBag.WriteProperty("AllowDelete", grd.AllowDelete, False)
  Call PropBag.WriteProperty("AllowUpdate", grd.AllowUpdate, False)
  Call PropBag.WriteProperty("LabelSortVisible", lblsort.visible, True)
  Call PropBag.WriteProperty("LabelFilterVisible", lblfilter.visible, True)
  Call PropBag.WriteProperty("LabelFastKeyVisible", lblfk.visible, True)
  Call PropBag.WriteProperty("RecordNavigatorVisible", dcVisible, True)
End Sub

Public Sub ReadProperties(ByVal PropBag As PropertyBag, ByVal grd As TDBGrid, ByVal dcVisible As Boolean, ByVal lblsort As Label, ByVal lblfilter As Label, ByVal lblfk As Label)
  On Error Resume Next
  grd.Enabled = PropBag.ReadProperty("Enabled", True)
  grd.AllowAddNew = PropBag.ReadProperty("AllowAddNew", False)
  grd.AllowDelete = PropBag.ReadProperty("AllowDelete", False)
  grd.AllowUpdate = PropBag.ReadProperty("AllowUpdate", False)
  lblsort.visible = PropBag.ReadProperty("LabelSortVisible", True)
  lblfilter.visible = PropBag.ReadProperty("LabelFilterVisible", True)
  lblfk.visible = PropBag.ReadProperty("LabelFastKeyVisible", True)
  dcVisible = PropBag.ReadProperty("RecordNavigatorVisible", True)
End Sub

Public Function ToUpper(ByVal KeyAscii As Long) As Integer
  If (KeyAscii >= 97) And (KeyAscii <= 122) Then
    KeyAscii = KeyAscii - 97 + 65
  End If
  ToUpper = KeyAscii
End Function

Public Sub MoveMouseCursor(ByVal XOffset As Long, ByVal YOffset As Long)
  Dim pt As POINTAPI
  If GetCursorPos(pt) Then Call SetCursorPos(pt.X + XOffset, pt.Y + YOffset)
End Sub


