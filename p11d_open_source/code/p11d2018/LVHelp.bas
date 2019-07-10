Attribute VB_Name = "ListviewHelp"
Option Explicit

Public Sub AllocateListview(lv As ListView, ByVal nItems As Long, Optional ByVal ClearKeys As Boolean = False)
  Dim v As Variant
  Dim i As Long, j As Long, lCount As Long
  
On Error GoTo AllocateListview_ERR
  
Call xSet("AllocateListview")
  lCount = lv.listitems.Count
  If nItems > lCount Then
    For i = lCount To (nItems - 1)
      Call AllocateListviewItem(lv)
    Next i
  ElseIf nItems < lCount Then
    j = lCount
    For i = lCount To (nItems + 1) Step -1
      Call lv.listitems.Remove(i)
      lCount = lCount - 1
    Next i
  End If
  For i = 1 To nItems
    lv.listitems(i).Key = vbNullString
  Next i
  
AllocateListview_END:
  Call xReturn("AllocateListview")
  Exit Sub
AllocateListview_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "AllocateListview", "Allocate Lis tview", "Error allocating a list view.")
  Resume AllocateListview_END
End Sub

Public Sub AllocateListviewItem(lv As ListView)
  Dim i As Long, s As String
  Dim nSubItems As Long, li As ListItem
  
  s = String$(64, " ")
  nSubItems = (lv.ColumnHeaders.Count - 1)
  Set li = lv.listitems.Add(, , s)
  For i = 1 To nSubItems
    li.SubItems(i) = s
  Next i
End Sub
Public Function CopyListView(lvDst As ListView, lvSrc As ListView, Optional ByVal CopyColumnSizes As Boolean = False, Optional ByVal SelectSourceSelected As Boolean = True) As Boolean
  Dim m As Long, ch As ColumnHeader, chDst As ColumnHeader, lst As ListItem
  Dim nCols As Long
  Dim lst2 As ListItem
  
  On Error GoTo CopyListView_ERR
  Call xSet("CopyListView")
  Call SetCursor
  lvDst.Sorted = False
    
  nCols = lvSrc.ColumnHeaders.Count
  If lvSrc.ColumnHeaders.Count <> lvDst.ColumnHeaders.Count Then
    For Each ch In lvSrc.ColumnHeaders
      If CopyColumnSizes Then
        Set chDst = lvDst.ColumnHeaders.Add(, ch.Key, ch.Text, ch.Width)
      Else
        Set chDst = lvDst.ColumnHeaders.Add(, ch.Key, ch.Text)
      End If
      chDst = ch.Tag
    Next ch
  End If
  
  Call AllocateListview(lvDst, lvSrc.listitems.Count)
  
  For Each lst In lvSrc.listitems
    Set lst2 = lvDst.listitems.Item(lst.Index)
    If StrComp(lst.Key, lst2.Key, vbBinaryCompare) <> 0 Then
      lst2.Key = lst.Key
    End If
    If StrComp(lst.Text, lst2.Text, vbBinaryCompare) <> 0 Then
      lst2.Text = lst.Text
    End If
    If StrComp(lst.Tag, lst2.Tag, vbBinaryCompare) <> 0 Then
      lst2.Tag = lst.Tag
    End If
    For m = 1 To (nCols - 1)  'use column header count to determine subitems
      If StrComp(lst2.SubItems(m), lst.SubItems(m), vbBinaryCompare) <> 0 Then
        lst2.SubItems(m) = lst.SubItems(m)
      End If
    Next
  Next lst
  If SelectSourceSelected Then
    Set lvDst.SelectedItem = lvDst.listitems(lvSrc.SelectedItem.Index)
  End If
  lvDst.Sorted = True
  CopyListView = True
  
  
  For m = 1 To lvSrc.ColumnHeaders.Count
    Set ch = lvSrc.ColumnHeaders(m)
    Set chDst = lvDst.ColumnHeaders(m)
    chDst.Text = ch.Text
  Next
  
CopyListView_END:
  Set lst2 = Nothing
  Set lst = Nothing
  Set ch = Nothing
  Call ClearCursor
  xReturn ("CopyListView")
  Exit Function
  
CopyListView_ERR:
  CopyListView = False
  Call ErrorMessage(ERR_ERROR, Err, "CopyListView", "Copy List View", "Error copying a list view.")
  Resume CopyListView_ERR
End Function

