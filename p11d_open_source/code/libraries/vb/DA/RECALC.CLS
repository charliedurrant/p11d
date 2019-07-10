VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Recalc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public name As String
Public recalctype As viewtypes
Public modes As String
Public parent As DADatabase

Dim strquerystring As String
Dim strparenttables() As String
Dim strchildtables() As String
Dim parents As New Collection
Dim children As New Collection
Dim booinlink As Boolean
Dim boolinked As Boolean
Dim boocalced As Boolean

Implements calc

Public Property Let querystring(text As String)
  Dim rs As Recordset
  Dim sttext As String

  On Error GoTo err_querystring

  strquerystring = text
  
  sttext = "Select * from sys_tablesfrom where viewname='" + name + "'"
  Set rs = parent.database.openrecordset(sttext)
  Do While Not rs.EOF
    addlist CStr(rs!tablename), strparenttables
    rs.MoveNext
  Loop
  rs.Close
  
  sttext = "Select * from sys_tablesto where viewname='" + name + "'"
  Set rs = parent.database.openrecordset(sttext)
  Do While Not rs.EOF
    addlist CStr(rs!tablename), strchildtables
    rs.MoveNext
  Loop
  rs.Close
  
  Exit Property
err_querystring:
  logerr "Error in recalc querystring " + Err.Description
  
  
End Property

Private Sub Calc_addchild(c As calc)
  Dim i As Long
  Dim pc As calc
  
  On Error GoTo err_addchild

  children.Add c, c.name
  
  If booinlink Then
    logerr "Circular calc in recalc addchild " + name
  ElseIf Not boolinked Then
    booinlink = True
    For i = 1 To UBound(strparenttables)
      Set pc = parent.getcalc(strparenttables(i))
      pc.addchild Me
      parents.Add pc, pc.name
    Next i
    booinlink = False
    boolinked = True
  End If
  
  Exit Sub
  
err_addchild:
  logerr "Error in recalc addchild " + Err.Description
End Sub

Private Sub Calc_addparent(c As calc)
' should not happen; only inserts/updates/deletes have explicit parents
' and these do not add to any query, only tables
  logerr "Invalid addparent in recalc " + name
End Sub

Private Property Let Calc_calced(RHS As Boolean)
  Dim c As calc
  
  If RHS = False And boocalced = True Then
    For Each c In children
      c.calced = False
    Next c
  End If
  boocalced = RHS
End Property

Private Sub calc_calc()
  Dim c As calc
  
  On Error GoTo err_calc
  
  If Not boocalced Then
    For Each c In parents
      c.calc
    Next c
    parent.database.Execute strquerystring
    boocalced = True
  End If
  
  Exit Sub
  
err_calc:
  logerr "Error in recalc calc " + Err.Description
End Sub

Private Property Get Calc_name() As String
  Calc_name = name
End Property

Private Sub Calc_removechild(c As calc)
  On Error GoTo err_removechild
  
  children.Remove c.name
  
  Exit Sub
  
err_removechild:
  logerr "Error in recalc removechild - not found " + c.name
End Sub

Private Sub Calc_removeparent(c As calc)
  On Error GoTo err_removeparent
  
  parents.Remove c.name
  
  Exit Sub

err_removeparent:
  logerr "Error in recalc removeparent - not found " + c.name
End Sub

Private Sub Class_Initialize()
  ReDim strparenttables(0)
  ReDim strchildtables(0)
End Sub

Public Sub linkchildren()
  Dim i As Long
  Dim c As calc
  
  On Error GoTo err_linkchildren
  
  For i = 1 To UBound(strchildtables)
    Set c = parent.getcalc(strchildtables(i))
    c.addparent Me
  Next i
    
  Exit Sub
err_linkchildren:
  logerr "Error in linkchildren " + Err.Description
End Sub

Public Sub delete()
  Dim c As calc
  Dim i As Long
  
  On Error GoTo err_delete
  
  For i = parents.Count To 1 Step -1
    Set c = parents(i)
    c.removechild Me
    parents.Remove (i)
  Next i
  
  For i = children.Count To 1 Step -1
    Set c = children(i)
    c.removeparent Me
    children.Remove i
  Next i
  
  Exit Sub
err_delete:
  logerr "Error in delete recalc " + Err.Description
End Sub
