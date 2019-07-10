Attribute VB_Name = "System"
Option Explicit
Public Function ClearCollection(C As Collection)
  Dim i As Long
  For i = 1 To C.count
    Call C.Remove(1)
  Next i
End Function
Public Sub ShowMaximized(frm As Form)
  frm.WindowState = 2
  frm.Show
  frm.ZOrder
  If frm.Enabled Then frm.SetFocus
End Sub



Public Function GetObjectIndex(ObjList As ObjectList, Item As Object) As Long
  Dim i As Long
  
  For i = 1 To ObjList.count
    If ObjList.Item(i) Is Item Then
      GetObjectIndex = i
      Exit Function
    End If
  Next i
  GetObjectIndex = -1
End Function

  
