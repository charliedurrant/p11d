VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Query"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mParentDb As DADatabase

Private mName As String
Private mTransient As Boolean
Private mType As Long
Private mDrillQuery As String
Private mHasBids As Boolean  ' whether the query has parents
                             ' which should be calced before it
                             ' false for batched sql

Private Enum ENTRY_TYPE
  HEADER_ENTRY
  CHILD_ENTRY
  PARENT_ENTRY
  SOURCEFIELD_ENTRY
  DESTFIELD_ENTRY
  ASSOCPARENT_ENTRY
  ASSOCCHILD_ENTRY
End Enum
Dim assocquery As String
Dim strsql As String

' Parents will be calced before calcing self
' self will be calced before children

' Query type:
' SELECT has parents the contents of the FROM clause
' SELECT has no children in SQL, but can have children if it is the parent for other queries (nested select)

' APPEND has parents the contents of the FROM clause
' APPEND has children the contents of the INSERT INTO clause upto the (...)
' APPEND may have no parents if it has keyword VALUES

' UPDATE has parents the contents of the UPDATE clause upto the SET clause
' UPDATE has children the destination fields of the SET clause

' DELETE has parents the contents of the FROM clause
' DELETE has children the tables in the DELETE clause
' zzzz DELETE is not complete yet

Dim strparents() As String        'parent objects by name
Dim strchildren() As String       'child objects by name
Dim strassocparents() As String   'breaks circularity for append/update where it updatres itself
Dim strsourcefields() As String   'fields/values in query which provide updates
Dim strdestfields() As String     'fields in query which are updated
Dim strassocchildren() As String  'breaks circularity for append/update where it updatres itself
Dim strparamnames() As String
Dim varparams() As Variant
Dim strmodes() As String          'flat/tree
Dim parents As New Collection     'parent as Object
Dim children As New Collection    'child onjects
Dim assocparents As New Collection 'circular parent objects
Dim booinlink As Boolean          'is it being linked - to detect circularity
Dim boolinked As Boolean          'is it linked
Dim boocalced As Boolean          'not dirty
Dim intcalcno As Long

Implements Calc
Private Function KillDatabase() As Boolean
  On Error Resume Next
  If Not mParentDb Is Nothing Then
    If mTransient And Not mParentDb.database Is Nothing Then
      If InCollection(mParentDb.database.QueryDefs, mName) And InStr(mName, "temp_", vbTextCompare) = 1 Then mParentDb.database.QueryDefs.Delete mName
    End If
    mParentDb.queries.Remove mName
  End If
  Set mParentDb = Nothing

End Function
Private Property Get SdbQType() As String
Select Case mType
Case dbQAppend
  SdbQType = "APPEND"
Case dbQUpdate
  SdbQType = "UPDATE"
Case dbQSelect
  SdbQType = "SELECT"
Case dbQDelete
  SdbQType = "DELETE"
Case Else
  SdbQType = "UNSUPPORTED"
End Select

End Property
Public Function Kill() As Boolean
  On Error Resume Next
  KillRelationships
  KillDatabase

End Function

Private Function KillRelationships() As Boolean
  Dim c As Calc
  Dim i As Long
  
  On Error Resume Next
  
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
  
  For i = assocparents.Count To 1 Step -1
    Set c = assocparents(i)
    c.removechild Me
    assocparents.Remove i
  Next i

End Function

Friend Sub setupquery(ParentDb As DADatabase, Optional qd As QueryDef = Nothing, Optional SQL As String, Optional HasParents As Boolean = True) ', Optional ParseQD As QueryDef)
Dim b As Boolean
' zzzz errorhandling and should be a function to return b
  If ParentDb Is Nothing Then Exit Sub
  Set mParentDb = ParentDb
  mHasBids = HasParents
  If qd Is Nothing Then
    If Len(SQL) = 0 Then Exit Sub
    If Not LoadHeaderSQL(SQL) Then Exit Sub
  Else
    If Not LoadHeader(qd) Then Exit Sub
  End If
  
  If Not mTransient Then Call ClearQueryinDB
  
  Select Case mType
  Case dbQAppend
    b = SetUpQuerySQLappend(strsql)
  Case dbQUpdate
    b = SetUpQuerySQLupdate(strsql)
  Case dbQSelect, dbQDelete
    If mHasBids Then
      b = SetUpQuerySQLselect(strsql)
    Else
      b = True
    End If
  Case Else
    Err.Raise ERR_SQL_PARSE, "tcsda.SetUpQueryEx", "Unknown query type " & CStr(mType) & " in " & mName & ", " & vbCr & vbCr & qd.SQL
  End Select

  If Not b Then
    Exit Sub  ' zzzz need some sort of error return to indicate failure
  End If
  If InCollection(mParentDb.queries, mName) Then mParentDb.queries.Remove mName
  mParentDb.queries.Add Me, mName
  If Not mTransient Then
    readmodes
    Call AddQueryEntryinDb(QUERY_HEADER_RS, HEADER_ENTRY, strsql & "|ASSOC|" & assocquery)
  End If
  
  
End Sub
Friend Sub setuptable(td As TableDef, parent As DADatabase)
  On Error GoTo err_setuptable
  Dim f As Field
  
  If parent Is Nothing Then Exit Sub
  Set mParentDb = parent
  mName = td.name
  mType = dbQTable
  Set f = td.fields("source")
  parent.queries.Add Me, mName
  readmodes
  boocalced = mParentDb.flag(mName, intcalcno)
  
  Exit Sub
  
err_setuptable:
  Select Case Err.Number
  Case 1
    logerr "Table without source column " & mName
  Case Else
    logerr "Error in setuptable for table " & mName
  End Select
  Resume Next
  Resume
End Sub


Friend Function createrecordset() As recordset
  Dim rs As recordset
  Dim qd As QueryDef
  Dim i As Long
  
  On Error GoTo err_createrecordset
  If mParentDb Is Nothing Then
    Set createrecordset = Nothing
    Exit Function
  End If
  
  Call Calc_Calc
  If mType = dbQSelect Then
    If InCollection(mParentDb.database.QueryDefs, mName) Then
      Set qd = mParentDb.database.QueryDefs(mName)
      For i = 1 To UBound(strparamnames)
        qd.Parameters(strparamnames(i)) = varparams(i)
      Next i
    End If
  End If
  Select Case mType
  Case dbQSelect
    If Not qd Is Nothing Then
      Set createrecordset = qd.OpenRecordset
      qd.Close
    Else
      Set createrecordset = mParentDb.database.OpenRecordset(mName)
    End If
  Case dbQTable
    Set createrecordset = mParentDb.database.OpenRecordset(mName)
  Case Else
    logerr "Can not create a recordset for query type " & CStr(mType) & " for query " & mName
    Set createrecordset = Nothing
  End Select
  
  Exit Function
err_createrecordset:
  logerr "Error in create recordset " & mName
 'Resume
End Function

Public Function createdars() As DARecordset
  Dim dars As New DARecordset
  Dim c As Calc
  
  On Error GoTo err_createdars
    
  dars.name = mName & NextID
  Set dars.recordset = Me.createrecordset
  If Not dars.recordset Is Nothing Then
    dars.addparent Me
    Set c = dars
    c.calced(Nothing) = True
    If Not mTransient Then
      ' zzzz may want to change this is due course
      children.Add dars, dars.name
    End If
    Set dars.parentdadb = mParentDb
    mParentDb.Recordsets.Add dars, dars.name
  
    Set createdars = dars
  Else
    Set createdars = Nothing
  End If
end_createdars:
  Exit Function
err_createdars:
  logerr "Error in createdars " & mName
  Resume end_createdars
End Function


Friend Sub linkchildren()
  Dim i As Long
  Dim c As Calc
  
  On Error GoTo err_linkchildren
  
  For i = 1 To UBound(strchildren)
    Set c = mParentDb.queries(strchildren(i))
    c.addparent mName
  Next i
  
  For i = 1 To UBound(strassocchildren)
    Set c = mParentDb.queries(strassocchildren(i))
    c.addassocparent mName
  Next i
  
  Exit Sub
err_linkchildren:
  logerr "Error in linkchildren for " & mName
  Resume Next
  Resume
End Sub

Private Function CreateMissingQuery(sName As String) As Boolean
Dim q As Query
Dim qd As QueryDef
Dim b As Boolean

  On Error GoTo CreateMissingQuery_err
  b = False
  If StrComp(Left(sName, 4), "sys_", vbTextCompare) = 0 Then
    logmessage "Link relations: The system query '" & sName & "' was not linked as an relation to " & mName
  Else
    ' check if in querydefs and is so then setup
    ' handles nested select queries which must be setup
    If InCollection(mParentDb.database.QueryDefs, sName) Then
      Set qd = mParentDb.database.QueryDefs(sName)
      Set q = mParentDb.CreateQuery(qd, "", mTransient)
      b = (Not q Is Nothing)
'  redundant as createquery adds to collection
'      If b Then
'        If Not InCollection(mParentDb.queries, q.name) Then mParentDb.queries.Add q, q.name
'      End If
    Else
      logerr "The expected relation, " & sName & ", of " & mName & " could not be found in the database."
    End If
  End If
  CreateMissingQuery = b
  
CreateMissingQuery_end:
  Set qd = Nothing
  Exit Function
CreateMissingQuery_err:
  logerr "The expected relation, " & sName & ", of " & mName & " has not been linked."
  Resume CreateMissingQuery_end
  Resume
End Function
Friend Sub linkinrelations()
' adds in relations after relations have been set up, used by createquery
Dim cRelation As Calc
Dim c As Calc
Dim i As Long

  On Error Resume Next
  Set c = Me
  For i = 1 To UBound(strassocchildren)
    If InCollection(mParentDb.queries, strassocchildren(i)) Then
      Set cRelation = mParentDb.queries(strassocchildren(i))
      cRelation.addassocparenttoCollection c
      cRelation.addchildtoCollection c
      c.addassocparenttoCollection cRelation
      c.addchildtoCollection cRelation
    End If
  Next i
  For i = 1 To UBound(strassocparents)
    If InCollection(mParentDb.queries, strassocparents(i)) Then
      Set cRelation = mParentDb.queries(strassocparents(i))
      cRelation.addassocparenttoCollection c
      cRelation.addchildtoCollection c
      c.addassocparenttoCollection cRelation
      c.addchildtoCollection cRelation
    End If
  Next i

  For i = 1 To UBound(strchildren)
    If Not (InList(strchildren(i), strassocchildren) Or InList(strchildren(i), strassocparents)) Then
      If InCollection(mParentDb.queries, strchildren(i)) Then
        Set cRelation = mParentDb.queries(strchildren(i))
        cRelation.addparenttoCollection Me
        c.addchildtoCollection cRelation
      End If
    End If
  Next i

  For i = 1 To UBound(strparents)
    If Not (InList(strparents(i), strassocchildren) Or InList(strparents(i), strassocparents)) Then
      If InCollection(mParentDb.queries, strparents(i)) Then
        Set cRelation = mParentDb.queries(strparents(i))
        cRelation.addchildtoCollection Me
        c.addparenttoCollection cRelation
      End If
    End If
  Next i

End Sub
Friend Sub linkparents()
  Dim i As Long
  Dim c As Calc
  Dim errDesc As String
  Dim b As Boolean
  
  On Error GoTo err_linkparents
 
  If Not boolinked Then
    booinlink = True
    
    For i = 1 To UBound(strassocparents)
      b = (InCollection(mParentDb.queries, strassocparents(i)))
      If Not b Then b = CreateMissingQuery(strassocparents(i))

      If b Then
        Set c = mParentDb.queries(strassocparents(i))
        assocparents.Add c, c.name
        c.addchild Me, True
      End If
    Next i
    
    For i = 1 To UBound(strparents)
      b = (InCollection(mParentDb.queries, strparents(i)))
      If Not b Then b = CreateMissingQuery(strparents(i))
      If b Then
        Set c = mParentDb.queries(strparents(i))
        parents.Add c, c.name
        c.addchild Me, False
      End If
    Next i
    
    booinlink = False
    boolinked = True
  End If
  
  Exit Sub
  
err_linkparents:
  Select Case Err.Number
  Case 457
    '  already in collection
    Resume Next
  Case Else
    logerr "Error in query linkparents for " & mName
    Resume Next
    
  End Select
  Resume
End Sub

Private Sub Calc_addassocparent(pname As String)
  On Error GoTo err_addparents
  
  AddList pname, strassocparents
  
  Exit Sub
err_addparents:
  logerr "Error in addassocparents "

End Sub

Private Sub Calc_addassocparenttoCollection(pCalc As Calc)
  On Error Resume Next
  AddList pCalc.name, strassocparents
  assocparents.Add pCalc, pCalc.name

End Sub

Private Sub Calc_addchildtoCollection(pCalc As Calc)
  On Error Resume Next
  AddList pCalc.name, strchildren
  children.Add pCalc, pCalc.name

End Sub

Private Sub Calc_addparenttoCollection(pCalc As Calc)
  On Error Resume Next
  AddList pCalc.name, strparents
  parents.Add pCalc, pCalc.name

End Sub

Private Sub Calc_addchild(c As Calc, assoc As Boolean)
  Dim i As Long
  Dim pc As Calc
  Dim allow As Boolean
  Dim b As Boolean
  
  On Error GoTo err_addchild

  children.Add c, c.name
  
  If booinlink Then
    If assoc Then
      allow = True
    Else
      For Each pc In assocparents
        If pc Is c Then allow = True
      Next pc
    End If
    If Not allow Then logerr "Circular calc in recalc addchild " & mName & " adding " & c.name
  ElseIf Not boolinked Then
    booinlink = True
    For i = 1 To UBound(strassocparents)
      b = (InCollection(mParentDb.queries, strassocparents(i)))
      If Not b Then b = CreateMissingQuery(strassocparents(i))
        
      If b Then
        Set pc = mParentDb.queries(strassocparents(i))
        If Not pc Is Nothing Then
          assocparents.Add pc, pc.name
          pc.addchild Me, True
        End If
      End If
    Next i
    For i = 1 To UBound(strparents)
      b = (InCollection(mParentDb.queries, strparents(i)))
      If Not b Then b = CreateMissingQuery(strparents(i))
        
      If b Then
        Set pc = mParentDb.queries(strparents(i))
        If Not pc Is Nothing Then
          parents.Add pc, pc.name
          pc.addchild Me, False
        End If
      End If
    Next i
    booinlink = False
    boolinked = True
  End If
  
  Exit Sub
  
err_addchild:
  Select Case Err.Number
  Case 457
    ' already in collection
    Resume Next
  Case Else
    logerr "Error in query add child, " & c.name & " to " & mName
    Resume Next
  End Select
  Resume
End Sub

Private Sub calc_addparent(pname As String)
  On Error GoTo err_addparents
  
  AddList pname, strparents
  ' why not add to the collection
  Exit Sub
err_addparents:
  logerr "Error in addparents " + Err.Description
End Sub
Private Property Get CalcedStatus() As Boolean
  On Error GoTo CalcedStatus_ERR
  
  If Not mTransient Then
    boocalced = mParentDb.flag(mName, intcalcno)
  End If
  CalcedStatus = boocalced
  
CalcedStatus_END:
Exit Property

CalcedStatus_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "GetCalcedStatus", "Getting Status", "The program has failed to establish the calculation status of " & mName & vbCr & vbCr & "It will be set to false.")
  CalcedStatus = False
  Resume CalcedStatus_END
End Property
Public Property Let CalcedStatus(RHS As Boolean)
  On Error GoTo CalcedStatusLET_ERR
  
  If Not mTransient Then
    mParentDb.flag(mName, intcalcno) = RHS
  End If
  boocalced = RHS

CalcedStatusLET_End:
Exit Property

CalcedStatusLET_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "LetCalcedStatus", "Setting Status", "The program has failed to set the calculation status of " & mName & " to " & CStr(RHS) & vbCr & vbCr & "This may lead to a failure to recalculate.")
  Resume CalcedStatusLET_End
End Property


Public Property Get Calc_boocalced() As Boolean
  
  Calc_boocalced = CalcedStatus
  
End Property
Public Property Let Calc_calced(parent As Calc, RHS As Boolean)
  Dim c As Calc
  
  On Error GoTo err_letcalced
  
  If Not Me.inmode Then Exit Property
  
  logmessage "Marking " + mName + " as " + CStr(RHS)
  
  If RHS = False Then
    For Each c In children
      If Not c Is parent Then
        ' gs change to reduce amount of touching
        'If StrComp(c.name, "PLCREATE", vbTextCompare) = 0 Then
        '   MsgBox ("PLCREATE")
        'End If
        If c.boocalced Then
          c.calced(Me) = False
        End If

        'c.calced(Me) = False
      End If
    Next c
  End If
  
  CalcedStatus = RHS
  
  Exit Property
  
err_letcalced:
  logerr "error in let calced " + Err.Description
  Resume Next
End Property

Private Sub Calc_Calc()
  Dim c As Calc
  Dim qd As QueryDef
  Dim rs As recordset
  Dim iter As Long
  Dim i As Long
  
  On Error GoTo err_calc
  
  If Not Me.inmode Then Exit Sub
  
  If Not CalcedStatus Then   'boocalced Then
    If mHasBids Then
      For Each c In parents
        c.Calc
      Next c
    End If
    Calc_calced(Nothing) = True
    If mHasBids Then
      For Each c In assocparents
        c.Calc
      Next c
    End If
    
    Select Case mType
    Case dbQUpdate
      logmessage "Calcing " + mName
      If mTransient Then
        mParentDb.database.execute mName
      Else
        Set rs = mParentDb.database.OpenRecordset(assocquery)
        Do While Not rs.eof And iter < 10 ' mak this a flag
          rs.Close
          iter = iter + 1
          Set qd = mParentDb.database.QueryDefs(mName)
          For i = 1 To UBound(strparamnames)
            logerr "Paramaters not yet supported in associated queries"
            qd.Parameters(strparamnames(i)).Value = varparams(i)
          Next i
          qd.execute
          Set rs = mParentDb.database.OpenRecordset(assocquery)
        Loop
        rs.Close
      End If
      logmessage "Complete on loop " + CStr(iter)
    Case dbQAppend
      logmessage "Calcing " + mName
      If mTransient Then
        ' ignores associated children of transients and associated queries for deletes
        ' ZZZZ CHANGE SETUP FOR TRANSIENTS implies that for transients do not need to construct associated query
        mParentDb.database.execute mName
      Else
        mParentDb.database.execute assocquery
        Set qd = mParentDb.database.QueryDefs(mName)
        For i = 1 To UBound(strparamnames)
          qd.Parameters(strparamnames(i)).Value = varparams(i)
        Next i
        If UBound(strassocchildren) > 0 Then
          Do
            qd.execute
            iter = iter + 1
          Loop While qd.RecordsAffected > 0 And iter < 10
        Else
          qd.execute
        End If
      End If
      logmessage "Complete on loop " + CStr(iter - 1)
    ' CASE DBQDELETE IS OUTSTANDING
    End Select
    
  End If
  
  Exit Sub
  
err_calc:
  logerr "Error in recalc calc of " & mName
  Resume Next
  Resume
End Sub

Private Property Let Calc_HasParents(RHS As Boolean)
  mHasBids = RHS
End Property

Private Property Get Calc_HasParents() As Boolean
  Calc_HasParents = mHasBids
End Property

Private Sub Calc_LogRelations()
Dim c As Calc


  For Each c In children
    logmessage mName & " Kid " & c.name
  Next c
  For Each c In parents
    logmessage mName & " Bid " & c.name
  Next c
  For Each c In assocparents
    logmessage mName & " KidBid " & c.name
  Next c


End Sub

Private Property Get Calc_name() As String
  Calc_name = mName
End Property

Private Sub Calc_removechild(c As Calc)
  On Error GoTo err_removechild
  
  children.Remove c.name
  
  Exit Sub
  
err_removechild:
  logmessage "Error in recalc removechild from " & mName & " child not found " & c.name
  'Resume
End Sub

Private Sub Calc_removeparent(c As Calc)
  Dim bparent As Boolean
  
  On Error Resume Next
  If c Is Nothing Then Exit Sub
  bparent = False
  If InCollection(parents, c.name) Then
    parents.Remove c.name
    bparent = True
  End If
  If InCollection(assocparents, c.name) Then
    assocparents.Remove c.name
    bparent = True
  End If
  If Not bparent Then
    logmessage "Error in recalc removeparent from " & mName & " parent not found " & c.name
  End If

End Sub



Private Sub Calc_touchparents()
  Dim c As Calc
  Dim dars As DARecordset
  
  For Each c In parents
    If c.boocalced Then
      c.calced(Me) = False
    End If
  Next c
  For Each dars In mParentDb.Recordsets
    dars.notify
  Next dars
End Sub

Private Sub Class_Initialize()
  ReDim strparents(0)
  ReDim strchildren(0)
  ReDim strsourcefields(0)
  ReDim strdestfields(0)
  ReDim strassocparents(0)
  ReDim strassocchildren(0)
  ReDim strmodes(0)
  ReDim strparamnames(0)
  ReDim strparams(0)
  mHasBids = True  ' zzzz OK for now, but should initialise in setup
End Sub

Private Sub readmodes()
  Dim rs As recordset
  
  On Error GoTo err_modes
  
  Set rs = mParentDb.database.OpenRecordset("SELECT * FROM sys_queries WHERE query='" + mName + "'")
  Do While Not rs.eof
    AddList UCase$(rs!Mode), strmodes
    rs.movenext
  Loop
  rs.Close
  
  Exit Sub
err_modes:
  logerr "Error in readmodes " + Err.Description
End Sub

Public Function inmode() As Boolean
  If UBound(strmodes) = 0 Then
    inmode = True
    Exit Function
  End If
  inmode = InList(mParentDb.Mode, strmodes)
End Function

Public Property Let parameter(pname As String, RHS As Variant)
  Dim i As Long
  Dim dars As DARecordset
  
  For i = 1 To UBound(strparamnames)
    If StrComp(pname, strparamnames(i), vbTextCompare) = 0 Then
      varparams(i) = RHS
      Calc_calced(Nothing) = False
' zzzz why notify every recordset when this query is changing
' surely should 1 compare varparams(i) with rhs and if same then no action
'               2 only notify those recordsets with this q as parent
      If Not mParentDb.Recordsets Is Nothing Then
        For Each dars In mParentDb.Recordsets
          dars.notify
        Next dars
      End If
      Exit Property
    End If
  Next i
  
  Err.Raise ERR_NOT_IN_LIST, "Parameter", "Unknown parameter " & pname & " set for query " & mName
  
End Property

Public Property Get parameter(pname As String) As Variant
  Dim i As Long
  
  For i = 1 To UBound(strparamnames)
    If StrComp(pname, strparamnames(i), vbTextCompare) = 0 Then
      parameter = varparams(i)
      Exit Property
    End If
  Next i
  
  Err.Raise ERR_NOT_IN_LIST, "Parameter", "Unknown parameter " & pname & " get for query " & mName

End Property

Private Sub Class_Terminate()
On Error Resume Next
  If mTransient And Not mParentDb Is Nothing Then
    If Not mParentDb.database Is Nothing Then
      If InCollection(mParentDb.database.QueryDefs, mName) Then mParentDb.database.QueryDefs.Delete mName
    End If
  End If
End Sub

Public Sub execute()
  Calc_calced(Nothing) = False
End Sub

Public Property Get name() As String
  name = mName
End Property
Public Property Let name(NewVal As String)
  mName = NewVal
End Property

Public Property Get Transient() As Boolean
  Transient = mTransient
End Property
Public Property Let Transient(NewVal As Boolean)
  mTransient = NewVal
End Property

Public Function LoadQuery(ParentDb As DADatabase, qd As QueryDef) As Boolean ', Optional ParseQD As QueryDef) As Boolean

Dim rsdetail As recordset
Dim rsrelations As recordset
Dim rsquery As recordset
Dim qname As String

  On Error GoTo loadquery_err
  xSet ("loadquery")
  
  If ParentDb Is Nothing Then Exit Function
  Set mParentDb = ParentDb
  qname = qd.name
  
  If Not LoadHeader(qd) Then Exit Function
  
  xSet ("loadquery getheader")
  Set rsquery = mParentDb.RS_ofType(QUERY_HEADER_RS)
  rsquery.FindFirst "Name='" & qname & "'"
  
  If rsquery.NoMatch Then Err.Raise ERR_INVALID_QUERY_LOAD
  strsql = "" & rsquery.fields("strsql")
  assocquery = "" & rsquery.fields("assoc_strsql")
  Set rsquery = Nothing
  xReturn ("loadquery getheader")
  
  xSet ("loadquery getdetail")
  Set rsdetail = mParentDb.RS_ofType(QUERY_DETAIL_RS)
  rsdetail.FindFirst "Name='" & qname & "'"
  While Not rsdetail.NoMatch
    Select Case rsdetail.fields("Detailtype")
      Case "SOURCEFIELD"
        Call AddList(rsdetail.fields("Detailname"), strsourcefields)
      Case "DESTFIELD"
        Call AddList(rsdetail.fields("Detailname"), strdestfields)
      Case Else
        ECASE "INVALID detailtype for " & mName & " - " & rsdetail.fields("Detailtype")
    End Select
    rsdetail.FindNext "Name='" & qname & "'"
  Wend
  Set rsdetail = Nothing
  xReturn ("loadquery getdetail")
  
  xSet ("loadquery getrelations")
  Set rsrelations = mParentDb.RS_ofType(QUERY_RELATIONS_RS)
  rsrelations.FindFirst "Name='" & qname & "'"
  While Not rsrelations.NoMatch
    Select Case rsrelations.fields("relationship")
      Case "CHILD"
        If rsrelations.fields("ASSOCIATED") Then
          Call AddList(rsrelations.fields("relationname"), strassocchildren)
        Else
          Call AddList(rsrelations.fields("relationname"), strchildren)
        End If
      Case "PARENT"
        If rsrelations.fields("ASSOCIATED") Then
          Call AddList(rsrelations.fields("relationname"), strassocparents)
        Else
          Call AddList(rsrelations.fields("relationname"), strparents)
        End If
      Case Else
        ECASE "INVALID RELATIONSHIP for " & mName & " - " & rsdetail.fields("RELATIONSHIP")
    End Select
    rsrelations.FindNext "Name='" & qname & "'"
  Wend
  Set rsrelations = Nothing
  xReturn ("loadquery getrelations")
  
  xSet ("loadquery getFlag")
  If Not InCollection(mParentDb.queries, mName) Then mParentDb.queries.Add Me, mName
  
  If Not mTransient Then
    readmodes
'    boocalced = mParentDb.flag(mName, intcalcno)
'  Else
'    boocalced = False
  End If
  LoadQuery = True
  xReturn ("loadquery getFlag")

loadquery_end:
  Set rsdetail = Nothing
  Set rsrelations = Nothing
  Set rsquery = Nothing
  xReturn ("loadquery")
  Exit Function
loadquery_err:
'zzzz
  
  Err.Raise Err.Number, "tcsda.loadquery", Err.Description
  Resume Next
Resume
End Function

Private Function LoadHeader(qd As QueryDef) As Boolean  ', Optional ParseQD As QueryDef) As Boolean
Dim i As Long

  On Error GoTo LoadHeader_Err
  Call xSet("LoadHeader")

  mName = qd.name
  mType = qd.Type
  strsql = qd.SQL
  
  If mType = 1 Or mType = 3 Or mType = 65 Then
    ' access internal type
    Exit Function
  End If
    
  If qd.Parameters.Count > 0 Then
    ReDim strparamnames(1 To qd.Parameters.Count)
    ReDim varparams(1 To qd.Parameters.Count)
    For i = 1 To qd.Parameters.Count
      strparamnames(i) = UCase$(qd.Parameters(i - 1).name)
    Next i
  End If
  LoadHeader = True

LoadHeader_End:
  Call xReturn("LoadHeader")
  Exit Function

LoadHeader_Err:
  Call ErrorMessage(ERR_ERROR, Err, "LoadHeader", "Error in LoadHeader", "Undefined error.")
  Resume LoadHeader_End
  Resume
End Function
Private Function LoadHeaderSQL(SQL As String) As Boolean  ', Optional ParseQD As QueryDef) As Boolean
  Dim bsel As Long
  Dim bapp As Long
  Dim bupd As Long
  Dim bdel As Long

  On Error GoTo LoadHeaderSQL_Err
  Call xSet("LoadHeaderSQL")
    
  mName = SQL
  strsql = SQL

  ' establish query type
  bsel = InstrSQL(1, SQL, "SELECT", vbTextCompare)
  bapp = InstrSQL(1, SQL, "INSERT", vbTextCompare)
  bupd = InstrSQL(1, SQL, "UPDATE", vbTextCompare)
  bdel = InstrSQL(1, SQL, "DELETE", vbTextCompare)
  
  If bsel = 0 And bapp = 0 And bupd = 0 And bdel = 0 Then
    Err.Raise ERR_SQL_PARSE, "tcsda.LoadHeaderSQL", "Unknown query type for," & vbCr & vbCr & SQL
  ElseIf (bsel < bapp Or bapp = 0) And (bsel < bupd Or bupd = 0) And (bsel < bdel Or bdel = 0) And bsel <> 0 Then
    mType = dbQSelect
  ElseIf (bapp < bsel Or bsel = 0) And (bapp < bupd Or bupd = 0) And (bapp < bdel Or bdel = 0) And bapp <> 0 Then
    mType = dbQAppend
  ElseIf (bupd < bapp Or bapp = 0) And (bupd < bsel Or bsel = 0) And (bupd < bdel Or bdel = 0) And bupd <> 0 Then
    mType = dbQUpdate
  Else  'If (bDEL < bapp Or bapp = 0) And (bDEL< bsel Or bsel = 0) And (bDEL < bUPD Or bUPD = 0) And bDEL <> 0 Then
    mType = dbQDelete
  End If
    
  LoadHeaderSQL = True

LoadHeaderSQL_End:
  Call xReturn("LoadHeaderSQL")
  Exit Function

LoadHeaderSQL_Err:
  Call ErrorMessage(ERR_ERROR, Err, "LoadHeaderSQL", "Error in LoadHeaderSQL", "Undefined error.")
  Resume LoadHeaderSQL_End
  Resume
End Function


Private Function AddQueryEntryinDb(rs_t As RS_TYPE, Entry_t As ENTRY_TYPE, s As String) As Boolean
Dim rs As recordset

  On Error GoTo AddQueryEntryinDb_Err
  Call xSet("AddQueryEntryinDb")
  Set rs = mParentDb.RS_ofType(rs_t)
  s = Trim(s)
  rs.addnew
  rs.fields("NAME") = mName
  Select Case Entry_t
  Case HEADER_ENTRY
    rs.fields("TYPE") = SdbQType
    rs.fields("DATEUPDATE") = mParentDb.database.QueryDefs(mName).LastUpdated
    rs.fields("STRSQL") = Left(s, InStr(s, "|ASSOC|") - 1)
    rs.fields("ASSOC_STRSQL") = RightPart(s, InStr(s, "|ASSOC|") + 7)
  Case ASSOCCHILD_ENTRY
    rs.fields("RELATIONSHIP") = "CHILD"
    rs.fields("ASSOCIATED") = True
    rs.fields("RELATIONNAME") = s
  Case ASSOCPARENT_ENTRY
    rs.fields("RELATIONSHIP") = "PARENT"
    rs.fields("ASSOCIATED") = True
    rs.fields("RELATIONNAME") = s
  Case CHILD_ENTRY
    rs.fields("RELATIONSHIP") = "CHILD"
    rs.fields("ASSOCIATED") = False
    rs.fields("RELATIONNAME") = s
  Case PARENT_ENTRY
    rs.fields("RELATIONSHIP") = "PARENT"
    rs.fields("ASSOCIATED") = False
    rs.fields("RELATIONNAME") = s
  Case DESTFIELD_ENTRY
    rs.fields("DETAILNAME") = s
    rs.fields("DETAILTYPE") = "DESTFIELD"
  Case SOURCEFIELD_ENTRY
    rs.fields("DETAILNAME") = s
    rs.fields("DETAILTYPE") = "SOURCEFIELD"
  Case Else
    ECASE "Invalid Entry type in " & mName & " - " & CStr(Entry_t)
  End Select
  rs.update

AddQueryEntryinDb_End:
  Call xReturn("AddQueryEntryinDb")
  Exit Function

AddQueryEntryinDb_Err:
If Err.Number = 3022 Then Resume Next ' zzzz could improve, merely avoids duplicate keys
  Call ErrorMessage(ERR_ERROR, Err, "AddQueryEntryinDb", "Recording Query Relationships", "An error has occurred with the query " & mName)
  Resume AddQueryEntryinDb_End
  Resume
End Function


Private Function ClearQueryinDB() As Boolean
Dim SQL As String

  On Error GoTo ClearQueryinDB_Err
  Call xSet("ClearQueryinDB")
  SQL = "delete * from sys_queries_Header where Name='" & mName & "'"
  mParentDb.database.execute (SQL)
  SQL = "delete * from sys_queries_detail where Name='" & mName & "'"
  mParentDb.database.execute (SQL)
  SQL = "delete * from sys_queries_relations where Name='" & mName & "'"
  mParentDb.database.execute (SQL)
' could use batchexecute
ClearQueryinDB_End:
  Call xReturn("ClearQueryinDB")
  Exit Function

ClearQueryinDB_Err:
  Call ErrorMessage(ERR_ERROR, Err, "ClearQueryinDB", "Error in ClearQueryinDB", "Undefined error.")
  Resume ClearQueryinDB_End
  Resume
End Function


Private Function SetUpQuerySQLappend(SQL As String) As Boolean
  Dim pos As Long
  Dim b As Boolean
  Dim s As String

  On Error GoTo SetUpQuerySQLappend_Err
  Call xSet("SetUpQuerySQLappend")
  
  pos = InStr(SQL, "(")
  If pos < 13 Then Err.Raise ERR_SQL_PARSE, "tcsda.setupquery", "Invalid append sql in " & mName & ", " & vbCr & vbCr & SQL
  s = Trim(Mid$(SQL, 13, pos - 13))
  If Not (InCollection(mParentDb.database.TableDefs, s) Or InCollection(mParentDb.database.QueryDefs, s)) Then Err.Raise ERR_SQL_PARSE, "tcsda.setupquery", "Invalid append sql in " & mName & ", " & vbCr & vbCr & SQL & vbCr & vbCr & " as non existent child, " & vbCr & s
  AddList s, strchildren
  
  pos = InstrSQL(1, SQL, "SELECT", vbTextCompare)
  If pos <> 0 Then
    If mHasBids Then
      b = SetUpQuerySQLselect(RightPart(SQL, pos))
    Else
      b = True
    End If
  Else
    pos = InstrSQL(1, SQL, "VALUES", vbTextCompare)
    If pos = 0 Then Err.Raise ERR_SQL_PARSE, "tcsda.setupquery", "No select statement found in append query, " & mName & ", " & vbCr & vbCr & SQL
  End If
  assocquery = "DELETE FROM " + strchildren(1) + " WHERE source='" + mName + "'"
  
  If InList(strchildren(1), strparents) Then
    AddList strchildren(1), strassocchildren
    If Not mTransient Then Call AddQueryEntryinDb(QUERY_RELATIONS_RS, ASSOCCHILD_ENTRY, strassocchildren(1))
    ReDim strchildren(0)
  Else
    If Not mTransient Then Call AddQueryEntryinDb(QUERY_RELATIONS_RS, CHILD_ENTRY, strchildren(1))
  End If

  SetUpQuerySQLappend = True

SetUpQuerySQLappend_End:
  Call xReturn("SetUpQuerySQLappend")
  Exit Function

SetUpQuerySQLappend_Err:
  Err.Raise Err.Number, Err.Description
'  Call ErrorMessage(ERR_ERROR, Err, "SetUpQuerySQLappend", "Error in SetUpQuerySQLappend", "Undefined error.")
'  Resume SetUpQuerySQLappend_End
End Function


Private Function SetUpQuerySQLupdate(SQL As String) As Boolean
  Dim tmpSQL As String
  Dim text As String
  Dim pos As Long
  Dim f As Field
  Dim sf As String
  Dim df As String
  Dim i As Long
  Dim desttable As String
  Dim b As Boolean

  On Error GoTo SetUpQuerySQLupdate_Err
  Call xSet("SetUpQuerySQLupdate")
  pos = InStr(SQL, "SET")
  If pos = 0 Then Err.Raise ERR_SQL_PARSE, "tcsda.setupquery", "Invalid update sql in " & mName & ", " & vbCr & vbCr & SQL
  assocquery = "SELECT * FROM " + Mid$(SQL, 8, pos - 9)
  If mHasBids Then
    b = SetUpQuerySQLselect(assocquery)
  Else
    b = True
  End If
  pos = pos + 3
  Do
    pos = SplitNextEqualsPair(SQL, pos, df, sf)
    If Len(sf) Then
      AddList sf, strsourcefields
      If Not mTransient Then Call AddQueryEntryinDb(QUERY_DETAIL_RS, SOURCEFIELD_ENTRY, sf)
      AddList df, strdestfields
      If Not mTransient Then Call AddQueryEntryinDb(QUERY_DETAIL_RS, DESTFIELD_ENTRY, df)
      desttable = TableName(df)
      If Len(desttable) Then
        If Not (InCollection(mParentDb.database.TableDefs, desttable) Or InCollection(mParentDb.database.QueryDefs, desttable)) Then Err.Raise ERR_SQL_PARSE, "tcsda.setupquery", "Invalid update sql in " & mName & ", " & vbCr & vbCr & SQL & vbCr & vbCr & " as non existent child, " & vbCr & desttable
        ' zzzz should this be if inlist(desttable, strparents) then add to strchildren else add to assocchildren
        If Not InList(desttable, strassocchildren) Then
          AddList desttable, strassocchildren
          If Not mTransient Then Call AddQueryEntryinDb(QUERY_RELATIONS_RS, ASSOCCHILD_ENTRY, desttable)
        End If
      Else
        logmessage "tcsda.setupquery: can not handle parsing update sql as need to get table from source tables in " & mName & ", " & vbCr & vbCr & SQL
      End If
    End If
  Loop While pos > 0
  
  assocquery = assocquery + " WHERE ("
  For i = 1 To UBound(strsourcefields)
    If i > 1 Then assocquery = assocquery + " OR "
    assocquery = assocquery + strsourcefields(i) + "<>" + strdestfields(i)
  Next i
  For i = 1 To UBound(strdestfields)
    assocquery = assocquery + " OR " + strdestfields(i) + " IS NULL "
  Next i
  assocquery = assocquery + ")"
  pos = InStr(SQL, "WHERE")
  If pos > 0 Then
    assocquery = assocquery + "and " + RightPart(SQL, pos + 5)
  End If

  SetUpQuerySQLupdate = True
  
SetUpQuerySQLupdate_End:
  Call xReturn("SetUpQuerySQLupdate")
  Exit Function

SetUpQuerySQLupdate_Err:
  
  Err.Raise Err.Number, Err.Description
'  Call ErrorMessage(ERR_ERROR, Err, "SetUpQuerySQLupdate", "Error in SetUpQuerySQLupdate", "Undefined error.")
'  Resume SetUpQuerySQLupdate_End
  Resume
End Function


Private Function SetUpQuerySQLselect(SQL As String) As Boolean

  Dim pos As Long
  Dim Startpos As Long
  Dim Endpos As Long
  Dim TempEndpos As Long
  Dim sparent As String
  Dim Newsql As String
  Dim blongEND As Boolean
  
  Dim sepstr(1 To 6) As String
  Dim endstr(1 To 5) As String
  Dim sepItem   As Long
  
  Dim i As Long
  
  Dim b As Boolean

  On Error GoTo SetUpQuerySQLselect_Err
  Call xSet("SetUpQuerySQLselect")
  
  sepstr(1) = ","
  sepstr(2) = " INNER JOIN "
  sepstr(3) = " LEFT JOIN "
  sepstr(4) = " RIGHT JOIN "
  sepstr(5) = " ON "
  sepstr(6) = " AS "
  
  endstr(1) = " GROUP BY "
  endstr(2) = " WHERE "
  endstr(3) = " ORDER BY "
  endstr(4) = " HAVING "
  endstr(5) = ";"
  
  Newsql = SQL
  ' replace vbcr and vblf with " "
  Newsql = Replace(Newsql, vbCr, " ", 1, -1, vbTextCompare)
  Newsql = Replace(Newsql, vbLf, " ", 1, -1, vbTextCompare)
    
  ' identify start
  Startpos = InstrSQL(1, Newsql, " FROM ", vbTextCompare)
  If Startpos = 0 Then Err.Raise ERR_SQL_PARSE, "tcsda.setupquery", "Invalid select sql in " & mName & ", " & vbCr & vbCr & SQL
  Startpos = Startpos + 6
  
  ' identify end
  blongEND = True
  Endpos = Len(Newsql)
  For i = LBound(endstr) To UBound(endstr)
    If Len(endstr(i)) Then
      TempEndpos = InstrSQL(Startpos, Newsql, endstr(i), vbTextCompare)
      If TempEndpos > 0 Then
        Endpos = Min(Endpos, TempEndpos)
        blongEND = (Len(endstr(i)) > 1)
      End If
    End If
  Next i
  Newsql = Mid$(Newsql, Startpos, Endpos - Startpos + IIf(blongEND, 1, 0))
  
  ' separate
  TempEndpos = 0
  Endpos = Len(Newsql)
  pos = 1
  Do
    pos = GetNextItemSeparatedByString(Newsql, pos, sepstr, sparent, sepItem)
    If Len(sparent) And Not (InStr(1, sparent, "sys_", vbTextCompare) = 1) Then
      If Not (InCollection(mParentDb.database.TableDefs, sparent) Or InCollection(mParentDb.database.QueryDefs, sparent)) Then Err.Raise ERR_SQL_PARSE, "tcsda.setupquery", "Invalid append sql in " & mName & ", " & vbCr & vbCr & SQL & vbCr & vbCr & " as non existent parent, " & vbCr & sparent
      AddList sparent, strparents
      If Not mTransient Then Call AddQueryEntryinDb(QUERY_RELATIONS_RS, PARENT_ENTRY, sparent)
    End If
    If sepItem = 0 Then Exit Do
    If sepstr(sepItem) = " ON " Or sepstr(sepItem) = " AS " Then
      ' find start of next sepstr
      Endpos = Len(Newsql)
      For i = LBound(sepstr) To UBound(sepstr)
        If Len(sepstr(i)) And (StrComp(sepstr(i), " ON ", vbTextCompare) <> 0) And (StrComp(sepstr(i), " AS ", vbTextCompare) <> 0) Then
          TempEndpos = InStr(pos, Newsql, sepstr(i), vbTextCompare)
          If TempEndpos > 0 Then Endpos = Min(Endpos, TempEndpos + Len(sepstr(i)))
        End If
      Next i
      pos = Endpos
    End If
  Loop While pos > 0 And pos < Len(Newsql)

  SetUpQuerySQLselect = True
SetUpQuerySQLselect_End:
  Call xReturn("SetUpQuerySQLselect")
  Exit Function

SetUpQuerySQLselect_Err:
  Err.Raise Err.Number, Err.Description
'  Call ErrorMessage(ERR_ERROR, Err, "SetUpQuerySQLselect", "Error in SetUpQuerySQLselect", "Undefined error.")
'  Resume SetUpQuerySQLselect_End
Resume
End Function

