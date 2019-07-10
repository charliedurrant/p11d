Attribute VB_Name = "Common"
Option Explicit
Public Const SELECT_ALL_LIMIT_ADO As Long = 1000
Private Const STATUS_HEIGHT As Single = 365
Private Const OFFSET As Single = 10

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

Public Function ConfirmSelectAll(ByVal rs As Recordset, ByVal frm As Form) As Boolean
  Dim rCount As Long
  On Error GoTo ConfirmSelectAll_err
  
  ConfirmSelectAll = True
  rCount = rs.RecordCount
  If rCount > SELECT_ALL_LIMIT_ADO Then
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

Public Sub ClearSelRows(ByVal grd As TDBGrid)
  Do While grd.SelBookmarks.Count > 0
    Call grd.SelBookmarks.Remove(0)
  Loop
End Sub

Public Function GridDragCell(ByVal grd As TDBGrid, ByVal aDC As Adodc, ByVal RowBookmark As Variant) As Recordset
  Dim vRecs() As Variant, bMax As Long, i As Long
  Dim rs As Recordset
  
  On Error GoTo GridDragCell_err
  Set rs = Nothing
  If grd.DataChanged Then Beep: GoTo GridDragCell_end
  Call grd.ClearSelCols
  If grd.IsSelected(RowBookmark) = -1 Then
    Call ClearSelRows(grd)
    Call grd.SelBookmarks.Add(RowBookmark)
  End If
  
  bMax = grd.SelBookmarks.Count - 1
  If bMax >= 0 Then
    ReDim vRecs(0 To bMax)
    For i = 0 To bMax
      vRecs(i) = grd.SelBookmarks(i)
    Next i
    Set rs = aDC.Recordset.Clone
    rs.Filter = vRecs
  End If
    
GridDragCell_end:
  Set GridDragCell = rs
  Exit Function
  
GridDragCell_err:
  Set rs = Nothing
  Resume GridDragCell_end
End Function

Public Sub SetDCProp(dc As Control, ByVal visible As Boolean)
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

