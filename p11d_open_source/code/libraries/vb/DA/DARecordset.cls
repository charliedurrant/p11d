VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DARecordset"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim strname As String
Dim objparent As Query
Dim boocalced As Boolean
Dim objrs As recordset
Dim boonotify As Boolean

Public parentdadb As DADatabase

Public Event DARSRefreshRequired()
Public Event darsRefreshReminder()

Implements calc

Public Function Kill()
  On Error Resume Next
  If Not objparent Is Nothing Then
    objparent.KillKid Me.name
  End If
  Set objparent = Nothing
  If Not parentdadb Is Nothing Then
    If InCollection(parentdadb.Recordsets, Me.name) Then parentdadb.Recordsets.Remove Me.name
  End If
  Set parentdadb = Nothing
  Set objrs = Nothing
End Function

Public Property Set recordset(rs As recordset)
  Set objrs = rs
End Property

Public Property Get recordset() As recordset
  Set recordset = objrs
End Property

Public Sub refresh()
  On Error GoTo err_refresh
  
  If Not boocalced Then
    Calc_Calc
'    objrs.Close
'    Set objrs = objparent.createrecordset
    objrs.Requery
    Calc_calced(Nothing) = True
  End If
  Exit Sub
err_refresh:
  logerr "Error in refresh " + Err.Description

End Sub

Public Function eof() As Boolean
  eof = objrs.eof
End Function

Public Function bof() As Boolean
  bof = objrs.bof
End Function

Public Sub movefirst()
  objrs.movefirst
End Sub

Public Sub movenext()
  objrs.movenext
End Sub

Public Sub movelast()
  objrs.movelast
End Sub

Public Sub moveprevious()
  objrs.moveprevious
End Sub

Public Property Get bookmark() As Variant
  bookmark = objrs.bookmark
End Property

Public Property Let bookmark(val As Variant)
  objrs.bookmark = val
End Property

Public Property Get fields(index As Variant) As Variant
  fields = objrs.fields(index).Value
End Property

Public Property Let fields(index As Variant, val As Variant)
  objrs.fields(index).Value = val
End Property

Public Sub update()
  Dim pc As calc
  Dim c As calc
  
  objrs.update
  Set pc = objparent
  pc.touchparents
  
End Sub

Public Sub addnew()
  objrs.addnew
End Sub

Public Sub edit()
  objrs.edit
End Sub

Public Sub closedars()
  Dim q As Query
  Dim c As calc
  
  On Error Resume Next
  
  If Not objrs Is Nothing Then objrs.Close
  Set objrs = Nothing
  Set q = objparent
  If Not q Is Nothing Then
    Set c = q
    c.removechild Me
    If q.Transient Then q.Kill
  End If
  If Not parentdadb Is Nothing Then parentdadb.Recordsets.Remove strname
  Set q = Nothing
  Set c = Nothing
End Sub

Friend Property Let name(val As String)
  strname = val
End Property

Private Sub Calc_addassocparent(pname As String)
  ' not used
End Sub

Private Sub Calc_addassocparenttoCollection(pCalc As calc)
  ' not used

End Sub

Private Sub Calc_addchild(c As calc, assoc As Boolean)
  ' not used
End Sub

Private Sub Calc_addchildtoCollection(pCalc As calc)
  ' not used

End Sub

Private Sub calc_addparent(pname As String)
  ' not used
End Sub

Private Sub Calc_addparenttoCollection(pCalc As calc)
  ' not used

End Sub

Private Property Get Calc_boocalced() As Boolean
Calc_boocalced = boocalced
End Property

Private Sub Calc_Calc()
  Dim c As calc
   
  If gCalc Then
    Set c = objparent
    c.calc
  End If
  
End Sub

Private Property Let Calc_calced(parent As calc, RHS As Boolean)
  Dim oldstate As Boolean
  
  #If DEBUGVER Then
    logmessage "Marking " + strname + " as " + CStr(RHS)
  #End If
  
  oldstate = boocalced
  If RHS = True Then
    boocalced = True
  Else
    boocalced = False
    If oldstate Then boonotify = True
  End If
End Property

Private Property Let Calc_CircCount(RHS As Integer)

End Property

Private Property Get Calc_CircCount() As Integer

End Property

Private Property Let Calc_HasParents(RHS As Boolean)
' not used
End Property

Private Property Get Calc_HasParents() As Boolean
  Calc_HasParents = False
End Property

Private Sub Calc_LogRelations(obj As Object, i As REL_TYPE, Optional nn As Node = Nothing)

Dim lv As ListView
Dim tli As ListItem
Dim tv As TreeView

Dim n As Node
Dim n1 As Node

Dim c As calc

  On Error Resume Next
  Set c = objparent
  If i = LIST_REL Then
    If Not TypeOf obj Is ListView Then Exit Sub
    Set lv = obj
    Set tli = lv.ListItems.Add
    tli.text = strname
    Call tli.ListSubItems.Add(, , "Bid RS")
    Call tli.ListSubItems.Add(, , objparent.name)
    Call tli.ListSubItems.Add(, , CStr(c.boocalced))
  Else
    If Not TypeOf obj Is TreeView Then Exit Sub
    Set tv = obj
    Set n = tv.Nodes.Add
    n.text = strname
    Set c = objparent
    c.LogRelations tv, TREE_BID_REL
  End If
  Set n = Nothing
  Set n1 = Nothing
  Set tv = Nothing
  Set tli = Nothing
  Set lv = Nothing
  Set c = Nothing

End Sub

Private Property Get Calc_name() As String
  Calc_name = strname
End Property

Private Sub Calc_removechild(c As calc)
  ' not used
End Sub

Private Sub Calc_removeparent(c As calc)
  Call Kill
  'Set objparent = Nothing
End Sub

Friend Sub addparent(parent As calc)
  Set objparent = parent
End Sub

Public Property Get name() As String
  name = strname
End Property

Friend Sub notify()
  If boonotify Then
    boonotify = False
    RaiseEvent DARSRefreshRequired
  End If
End Sub

Private Sub Calc_touchparents()
  ' not used
End Sub

Public Sub touch()
  Dim pc As calc
  
  Set pc = objparent
  #If DEBUGVER Then
    logmessage "Touching " & strname & " in DARecordet.Touch"
  #End If
  pc.touchparents
End Sub

Private Sub Class_Initialize()
  Debug.Print "Create DARS"
End Sub

Private Sub Class_Terminate()
  Debug.Print "Terminate DARS - " & strname
  Call Kill
  
End Sub
