VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "view"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public name As String
Public viewtype As viewtypes
Public parent As DADatabase

'Dim strquery As String
Dim strparenttables() As String
Dim strquerystring As String
Dim parents As New Collection
Dim children As New Collection
Dim booinlink As Boolean
Dim boolinked As Boolean
Dim boocalced As Boolean

Implements calc

Public Property Let querystring(text As String)
  Dim rs As recordset
  strquerystring = text
  
  On Error GoTo err_querystring
  
  text = "Select * from sys_tablesfrom where viewname='" + name + "'"
  Set rs = parent.database.openrecordset(text)
  Do While Not rs.EOF
    addlist CStr(rs!tablename), strparenttables
    rs.MoveNext
  Loop
  rs.Close

  Exit Property

err_querystring:
  logerr "Error in view querystring " + Err.Description
  
End Property

Private Sub Calc_addparent(c As calc)
  parents.Add c, c.name
End Sub

Private Sub Calc_addchild(c As calc)

  children.Add c, c.name
  If booinlink Then
    logerr "Circular calc in " + name
  ElseIf Not boolinked Then
    linkparents
  End If
End Sub

Friend Sub linkparents()
  Dim i As Long
  Dim c As calc
  
  On Error GoTo err_linkparents
 
  If Not boolinked Then
    booinlink = True
    
    For Each c In parents
      c.addchild Me
    Next c
    
    For i = 1 To UBound(strparenttables)
      Set c = parent.getcalc(strparenttables(i))
      c.addchild Me
      parents.Add c, c.name
    Next i
    
    booinlink = False
    boolinked = True
  End If
  
  Exit Sub
  
err_linkparents:
  logerr "Error in view linkparents " + Err.Description
End Sub

Private Property Get Calc_name() As String
  Calc_name = name
End Property

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
    boocalced = True
  End If
    
  Exit Sub
  
err_calc:
  logerr "Error in recalc calc " + Err.Description
End Sub

Private Sub Calc_removechild(c As calc)
  On Error GoTo err_removechild
  
  children.Remove c.name
  
  Exit Sub
  
err_removechild:
  logerr "Error in view removechild - not found " + c.name
End Sub

Private Sub Calc_removeparent(c As calc)
  On Error GoTo err_removeparent
  
  parents.Remove c.name
  
  Exit Sub

err_removeparent:
  logerr "Error in view removeparent - not found " + c.name
End Sub

Private Sub Class_Initialize()
  ReDim strparenttables(0)
End Sub

Public Sub delete()
  Dim i As Long
  Dim c As calc
  
  On Error GoTo err_delete
  
  For i = parents.Count To 1 Step -1
    Set c = parents(i)
    c.removechild Me
    parents.Remove i
  Next i
  
  For i = children.Count To 1 Step -1
    Set c = children(i)
    c.removeparent Me
    children.Remove i
  Next i
  
  Exit Sub
err_delete:
  logerr "Error in view delete " + Err.Description
End Sub

Public Function createrecordset() As DARecordset
  Dim dars As New DARecordset
  Dim c As calc
  
  On Error GoTo err_createrecordset
  
  calc_calc
  dars.name = name + nextid
  If viewtype = tabletype Then
    Set dars.recordset = parent.database.openrecordset(name)
  ElseIf viewtype = querytype Then
    Set dars.recordset = parent.database.openrecordset(strquerystring)
  Else
    logerr "Query type not yet supported"
  End If
  Set c = dars
  c.addparent Me
  children.Add c, c.name
  c.calced = True
  Set createrecordset = dars
  
  Exit Function
err_createrecordset:
  logerr "Error in create recordset " + Err.Description
End Function
