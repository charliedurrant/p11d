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

Implements calc

Private mParentDb As DADatabase
Private mName As String
Private mTransient As Boolean
Private mType As Long
Private mDrillQuery As String
Private mHasBids As Boolean  ' whether the query has parents
                             ' which should be calced before it
                             ' false for batched sql
Private mCircCount As Integer ' count of circular recalcs which calc is involved in
Private mCached As Boolean

Private Enum ENTRY_TYPE
  HEADER_ENTRY
  CHILD_ENTRY
  PARENT_ENTRY
  SOURCEFIELD_ENTRY
  DESTFIELD_ENTRY
  ASSOCPARENT_ENTRY
  ASSOCCHILD_ENTRY
End Enum

Dim strAssocquery As String
Dim strSql As String
Dim strParameters As String
Dim lngParamcount As Long
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

Public Property Get Modes() As String
Dim i As Long
Dim b As Boolean
Dim s As String
  For i = LBound(strmodes) To UBound(strmodes)
    s = s & IIf(b, ", ", "") & strmodes(i)
    b = True
  Next i
  Modes = s
End Property

Private Property Get SdbQType() As Long
  
  Select Case mType
    Case dbQSelect
      SdbQType = 0 '"SELECT"
    Case dbQCrosstab
      SdbQType = 16 '"CROSSTAB"
    Case dbQDelete
      SdbQType = 32  '"DELETE"
    Case dbQUpdate
      SdbQType = 48 '"UPDATE"
    Case dbQAppend
      SdbQType = 64 '"APPEND"
    Case dbQSetOperation
      SdbQType = 128 '"UNION"
    Case Else
      SdbQType = 0 ' Set as default "SELECT"
  End Select

End Property

Public Function Kill() As Boolean
  On Error Resume Next
  KillRelationships
  KillDatabase
End Function

Private Function KillRelationships() As Boolean
  Dim c As calc
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
    If InCollection(children, c.name) Then children.Remove i
  Next i
  
  For i = assocparents.Count To 1 Step -1
    Set c = assocparents(i)
    c.removechild Me
    assocparents.Remove i
  Next i
' set all to nothing
End Function

Private Function KillDatabase() As Boolean
  On Error Resume Next
  
  If (Not mParentDb Is Nothing) Then
    If mTransient And InStr(1, mName, "temp_", vbTextCompare) = 1 And Not mParentDb.database Is Nothing Then
      If InCollection(mParentDb.database.QueryDefs, mName) Then mParentDb.database.QueryDefs.Delete mName
    End If
    'zzzz pc removed because caused vb to gpf
    If InCollection(mParentDb.queries, Me.name) Then Call mParentDb.queries.Remove(Me.name)
  End If
  Set mParentDb = Nothing

End Function

Friend Sub setupquery(ParentDB As DADatabase, Optional qd As QueryDef = Nothing, Optional sql As String, Optional HasParents As Boolean = True) ', Optional ParseQD As QueryDef)
  Dim b As Boolean
  
  ' zzzz errorhandling and should be a function to return b
    
  If ParentDB Is Nothing Then Exit Sub
  Set mParentDb = ParentDB
  mHasBids = HasParents
  If qd Is Nothing Then
    If Len(sql) = 0 Then Exit Sub
    If Not LoadHeaderSQL(sql) Then Exit Sub
  Else
    If Not LoadHeader(qd) Then Exit Sub
  End If
  
  If Not mTransient Then Call ClearQueryinDB
  
  Select Case mType
  Case dbQAppend
    b = SetUpQuerySQLappend(strSql)
  Case dbQUpdate
    b = SetUpQuerySQLupdate(strSql)
  Case dbQSelect, dbQDelete, dbQCrosstab
    If mHasBids Then
      b = SetUpQuerySQLselect(strSql)
    Else
      b = True
    End If
  ' zzzz PC added 15/11/00
  Case dbQSetOperation
    b = SetUpQuerySQLUnion(strSql)
  Case Else
    Err.Raise ERR_SQL_PARSE, "tcsda.SetUpQueryEx", "Unknown query type " & CStr(mType) & " in " & mName & ", " & vbCr & vbCr & qd.sql
  End Select

  If Not b Then
    Exit Sub  ' zzzz need some sort of error return to indicate failure
  End If
  
  Call AddtoParent
  
  If Not mTransient Then
    readmodes
    Call AddQueryEntryinDb(QUERY_HEADER_RS, HEADER_ENTRY, strSql & "|ASSOC|" & strAssocquery)
  End If
  
End Sub

Friend Sub setuptable(ParentDB As DADatabase, tdName As String)
  On Error GoTo err_setuptable
  Dim f As Field

  ' zzzz use tablename and drop check for source field (will cause problems if any exist)
  If ParentDB Is Nothing Then Exit Sub
  Set mParentDb = ParentDB
  mName = tdName
  mType = dbQTable
  'If Not InCollection(td.fields, "Source") Then logerr "There is a data Table without a source column, " & mName & vbCr & vbCr & "This should be reported."
  'Set f = td.fields("source")

  Call AddtoParent
  readmodes
  boocalced = mParentDb.flag(mName, intcalcno)
  Exit Sub
  
err_setuptable:
'  Select Case Err.Number
'  Case 1
'    logerr "Table without source column " & mName
'  Case Else
  logerr "Error in setuptable for table " & mName
'  End Select
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
    ' zzzz this has to be an error
    Exit Function
  End If
  
  Call Calc_Calc
  If mType = dbQSelect Or mType = dbQCrosstab Then
    If InCollection(mParentDb.database.QueryDefs, mName) Then
      Set qd = mParentDb.database.QueryDefs(mName)
      For i = 1 To UBound(strparamnames)
        qd.Parameters(strparamnames(i)) = varparams(i)
      Next i
    End If
  End If
  Select Case mType
  Case dbQSelect, dbQCrosstab
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

Friend Function KillKid(kname As String)
  
  If InCollection(children, kname) Then children.Remove kname

End Function
Public Function createdars() As DARecordset
  Dim dars As New DARecordset
  Dim c As calc
  
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
  Dim c As calc
  
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
    #If DEBUGVER Then
      logmessage "Link relations: The system query '" & sName & "' was not linked as an relation to " & mName
    #End If
  Else
    ' check if in querydefs and is so then setup
    ' handles nested select queries which must be setup
    If InCollection(mParentDb.database.QueryDefs, sName) Then
      Set qd = mParentDb.database.QueryDefs(sName)
      Set q = mParentDb.CreateQuery(qd, "", mTransient)
      b = (Not q Is Nothing)
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
Dim cRelation As calc
Dim c As calc
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
  Dim c As calc
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

Private Sub Calc_addassocparenttoCollection(pCalc As calc)
  On Error Resume Next
  AddList pCalc.name, strassocparents
  assocparents.Add pCalc, pCalc.name

End Sub

Private Sub Calc_addchildtoCollection(pCalc As calc)
  On Error Resume Next
  AddList pCalc.name, strchildren
  children.Add pCalc, pCalc.name

End Sub

Private Sub Calc_addparenttoCollection(pCalc As calc)
  On Error Resume Next
  AddList pCalc.name, strparents
  parents.Add pCalc, pCalc.name

End Sub

Private Sub Calc_addchild(c As calc, assoc As Boolean)
  Dim i As Long
  Dim pc As calc
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
    If Not allow Then
      #If DEBUGVER Then
        logmessage "A circular relationship has arisen so that a query is its own ancestor" & vbCr & vbCr & "In adding the child " & c.name & " to the circular parent " & mName
      #End If
    End If
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
  Resume
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

Public Property Let Calc_calced(parent As calc, RHS As Boolean)
  Dim c As calc
  Static inLetCalced As Boolean
  
  On Error GoTo err_letcalced
  
  
  If Not Me.inmode Then Exit Property
  If inLetCalced Then
    #If DEBUGVER Then
      logmessage "Circular marking " & mName & " as " & CStr(RHS) & " and exiting"
    #End If
    CalcedStatus = RHS
    Exit Property
  End If
  inLetCalced = True
  #If DEBUGVER Then
    logmessage "Marking " + mName + " as " + CStr(RHS)
  #End If
  
  If RHS = False Then
    For Each c In children
      If Not c Is parent Then
        If c.boocalced Then
          c.calced(Me) = False
        End If
      End If
    Next c
  End If
  
  CalcedStatus = RHS
  inLetCalced = False
  Exit Property
  
err_letcalced:
  logerr "error in let calced " + Err.Description
  Resume Next
End Property

Private Sub Calc_Calc()
  Dim c As calc
  Dim qd As QueryDef
  Dim rs As recordset
  Dim iter As Long
  Dim i As Long
  Static inCalc As Integer
  Dim errflag As Boolean
  
  On Error GoTo err_calc
  
  errflag = False
  
  If gCalc Then
    #If DEBUGVER Then
      logmessage "Query.Calc_Calc where gCalc = True and name is " & mName & " and " & IIf(CalcedStatus, "", "no ") & "calc required"
    #End If
    If Not Me.inmode Then Exit Sub
    If inCalc = 1 Then
      #If DEBUGVER Then
        logmessage ("Circular calc in " & mName & " starting circuit number " & CStr(inCalc))
      #End If
    End If
    If inCalc > 1 Then
      #If DEBUGVER Then
        logmessage ("Circular calc in " & mName & " stopping at end of circuit number " & CStr(inCalc))
      #End If
      'Calc_calced(Nothing) = True
      inCalc = 0
      Exit Sub
    End If
    
    If Not CalcedStatus Then   'boocalced Then
      inCalc = inCalc + 1
      If mHasBids Then
        For Each c In parents
          c.calc
        Next c
      End If
      Calc_calced(Nothing) = True
      If mHasBids Then
        For Each c In assocparents
          c.calc
        Next c
      End If
      
      Select Case mType
      Case dbQUpdate
        #If DEBUGVER Then
          logmessage "Executing " + mName
        #End If
        iter = 0
        If mTransient Then
          mParentDb.database.execute mName
        Else
          Set rs = mParentDb.database.OpenRecordset(strAssocquery)
          Do While Not rs.eof And iter < 10 ' mak this a flag
            rs.Close
            iter = iter + 1
            Set qd = mParentDb.database.QueryDefs(mName)
            For i = 1 To UBound(strparamnames)
              logerr "Paramaters not yet supported in associated queries"
              qd.Parameters(strparamnames(i)).Value = varparams(i)
            Next i
            qd.execute
            Set rs = mParentDb.database.OpenRecordset(strAssocquery)
          Loop
          rs.Close
        End If
        #If DEBUGVER Then
          logmessage "Executing " + mName & " - Complete on loop " + CStr(iter)
        #End If
      Case dbQAppend
        #If DEBUGVER Then
          logmessage "Executing " + mName
        #End If
        iter = 0
        If mTransient Then
          ' ignores associated children of transients and associated queries for deletes
          ' ZZZZ CHANGE SETUP FOR TRANSIENTS implies that for transients do not need to construct associated query
          mParentDb.database.execute mName
        Else
          mParentDb.database.execute strAssocquery
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
        #If DEBUGVER Then
          logmessage "Executing " + mName & " - Complete on loop " + CStr(iter)
        #End If
      Case dbQDelete
        #If DEBUGVER Then
          logmessage "Executing " + mName
        #End If
        If mTransient Then
          mParentDb.database.execute mName
        Else
          qd.execute
        End If
      End Select
      inCalc = 0
    End If
  Else
    #If DEBUGVER Then
      logmessage "Query.Calc_Calc where gCalc = False and name is " & mName
    #End If
  End If
  
err_errflag:
  Exit Sub

err_calc:
  If errflag Then
    logerr "Error in recalc calc of " & mName & " and the recalc will roll-up"
    Resume err_errflag
  End If
  errflag = True
  logerr "Error in recalc calc of " & mName
  Resume Next
  Resume
End Sub

Private Property Let Calc_CircCount(RHS As Integer)
  mCircCount = RHS

End Property

Private Property Get Calc_CircCount() As Integer
  Calc_CircCount = mCircCount
End Property

Private Property Let Calc_HasParents(RHS As Boolean)
  mHasBids = RHS
End Property

Private Property Get Calc_HasParents() As Boolean
  Calc_HasParents = mHasBids
End Property

Private Sub Calc_LogRelations(obj As Object, i As REL_TYPE, Optional nn As Node = Nothing)
Dim lv As ListView
Dim tli As ListItem
' zzzz could improve a lot by having a selector on form for choosing object for which relations should be displayed

Dim c As calc

Dim tv As TreeView
Dim n As Node
Dim n1 As Node
Dim n2 As Node
Static inLog As Boolean

  On Error GoTo Calc_LogRelations_err
  If i = LIST_REL Then
    If Not TypeOf obj Is ListView Then Exit Sub
    Set lv = obj
    For Each c In children
      Set tli = lv.ListItems.Add
      tli.text = mName
      Call tli.ListSubItems.Add(, , "Kid")
      Call tli.ListSubItems.Add(, , c.name)
      Call tli.ListSubItems.Add(, , CStr(c.boocalced))
    Next c
    For Each c In parents
      Set tli = lv.ListItems.Add
      tli.text = mName
      Call tli.ListSubItems.Add(, , "Bid")
      Call tli.ListSubItems.Add(, , c.name)
      Call tli.ListSubItems.Add(, , CStr(c.boocalced))
    Next c
    For Each c In assocparents
      Set tli = lv.ListItems.Add
      tli.text = mName
      Call tli.ListSubItems.Add(, , "Kid/Bid")
      Call tli.ListSubItems.Add(, , c.name)
      Call tli.ListSubItems.Add(, , CStr(c.boocalced))
    Next c
  Else
    If Not TypeOf obj Is TreeView Then Exit Sub
    Set tv = obj
    If inLog Then
      If Not nn Is Nothing Then
        Set n = tv.Nodes.Add(nn, tvwChild, mName & CStr(Rnd), mName & "   Circular")
      End If
    Else
      inLog = True
      Select Case i
      Case TREE_TOP_REL
        If Not mType = 1024 Then Exit Sub
        Call ShowProgress("Loading " & mName & "'s relations", "Loading relations")
        Set n = tv.Nodes.Add(, , mName & CStr(Rnd))
        n.text = mName
        ' load properties
        Set c = Me
        'Set n1 = tv.Nodes.Add(n.Key, tvwChild, mName & CStr(Rnd), "Properties")
        'Set n2 = tv.Nodes.Add(n1.Key, tvwChild, mName & CStr(Rnd), "Type:         " & mType)
        'Set n2 = tv.Nodes.Add(n1.Key, tvwChild, mName & CStr(Rnd), "Transient:    " & CStr(mTransient))
        'Set n2 = tv.Nodes.Add(n1.Key, tvwChild, mName & CStr(Rnd), "SQL:          " & Left$(strsql, 50))
        'Set n2 = tv.Nodes.Add(n1.Key, tvwChild, mName & CStr(Rnd), "AssocSQL:     " & Left$(strassocquery, 50))
        Set n2 = tv.Nodes.Add(n.Key, tvwChild, mName & CStr(Rnd), "Calced:       " & CStr(c.boocalced))
        If children.Count Then
          Set n1 = tv.Nodes.Add(n.Key, tvwChild, mName & CStr(Rnd), "Descendants")
          For Each c In children
            c.LogRelations tv, TREE_KID_REL, n1
          Next c
        End If
        If parents.Count Then
          Set n1 = tv.Nodes.Add(n.Key, tvwChild, mName & CStr(Rnd), "Ancestors")
          For Each c In parents
            c.LogRelations tv, TREE_BID_REL, n1
          Next c
        End If
        If assocparents.Count Then
          Set n1 = tv.Nodes.Add(n.Key, tvwChild, mName & CStr(Rnd), "Incestors")
          For Each c In assocparents
            c.LogRelations tv, TREE_KIDBID_REL, n1
          Next c
        End If
      Case TREE_KID_REL
        If Not nn Is Nothing Then
          Set n = tv.Nodes.Add(nn, tvwChild, mName & CStr(Rnd), mName & IIf(Calc_boocalced, "      Calc", "") & IIf(mTransient, "      Transient", ""))
          If children.Count Then
            For Each c In children
              If Not InCollection(assocparents, c.name) And Not InCollection(parents, c.name) Then
                c.LogRelations tv, TREE_KID_REL, n
              End If
            Next c
          End If
        End If
      Case TREE_BID_REL
        If Not nn Is Nothing Then
          Set n = tv.Nodes.Add(nn, tvwChild, mName & CStr(Rnd), mName & IIf(Calc_boocalced, "      Calc", "") & IIf(mTransient, "      Transient", ""))
          If parents.Count Then
            For Each c In parents
              If Not InCollection(assocparents, c.name) And Not InCollection(children, c.name) Then
                c.LogRelations tv, TREE_BID_REL, n
              End If
            Next c
          End If
        End If
      Case TREE_KIDBID_REL
        If Not nn Is Nothing Then
          Set n = tv.Nodes.Add(nn, tvwChild, mName & CStr(Rnd), mName & IIf(Calc_boocalced, "      Calc", "") & IIf(mTransient, "      Transient", ""))
          If children.Count Then
            Set n1 = tv.Nodes.Add(n, tvwChild, mName & CStr(Rnd), "Descendants")
            For Each c In children
              If StrComp(c.name, n.parent.parent.text) = 0 Then
                Set n2 = tv.Nodes.Add(n1, tvwChild, mName & CStr(Rnd), "Self")
              Else
                c.LogRelations tv, TREE_KID_REL, n1
              End If
            Next c
          End If
          If parents.Count Then
            Set n1 = tv.Nodes.Add(n, tvwChild, mName & CStr(Rnd), "Ancestors")
            For Each c In parents
              If StrComp(c.name, n.parent.parent.text) = 0 Then
                Set n2 = tv.Nodes.Add(n1, tvwChild, mName & CStr(Rnd), "Self")
              Else
                c.LogRelations tv, TREE_BID_REL, n1
              End If
            Next c
          End If
          If assocparents.Count Then
            Set n1 = tv.Nodes.Add(n, tvwChild, mName & CStr(Rnd), "Incestors")
            For Each c In assocparents
              If StrComp(c.name, n.parent.parent.text) = 0 Then
                Set n2 = tv.Nodes.Add(n1, tvwChild, mName & CStr(Rnd), "Self")
              Else
                c.LogRelations tv, TREE_KIDBID_REL, n1
              End If
            Next c
          End If
        End If
      Case Else
        ECASE "Invalid rel_type in LogRelations for Query"
      End Select
    End If
  End If
Calc_LogRelations_end:
  inLog = False
  Set c = Nothing
  Set n = Nothing
  Set n1 = Nothing
  Set n2 = Nothing
  Set tli = Nothing
  Set lv = Nothing
  Set tv = Nothing
  Exit Sub
Calc_LogRelations_err:
  Call ErrorMessage(ERR_ERROR, Err, "Calc_LogRelations", "Logging Relations", "")
  Resume Calc_LogRelations_end
  Resume
End Sub

Private Property Get Calc_name() As String
  Calc_name = mName
End Property

Private Sub Calc_removechild(c As calc)
Dim ccalc As calc
  On Error GoTo err_removechild
'  For Each ccalc In children
'    If ccalc Is c Then
'      children.Remove ccalc.name
'      Exit For
'    End If
'  Next ccalc
  children.Remove c.name
 
  Exit Sub
  
err_removechild:
  #If DEBUGVER Then
    logmessage "Error in recalc removechild from " & mName & " child not found " & c.name
  #End If
'  Resume
End Sub

Private Sub Calc_removeparent(c As calc)
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
    #If DEBUGVER Then
      logmessage "Error in recalc removeparent from " & mName & " parent not found " & c.name
    #End If
  End If

End Sub

Private Sub Calc_touchparents()
  Dim c As calc
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

Private Sub readmodes()
  Dim rs As recordset
  
  On Error GoTo err_modes
  Set rs = mParentDb.database.OpenRecordset("SELECT * FROM sys_queries WHERE query='" + mName + "'", dbOpenSnapshot)
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

Public Sub execute()
  Calc_calced(Nothing) = False
End Sub

Public Property Get name() As String
  name = mName
End Property

Public Property Let name(newval As String)
  mName = newval
End Property

Public Property Get Transient() As Boolean
  Transient = mTransient
End Property

Public Property Let Transient(newval As Boolean)
  mTransient = newval
End Property

Public Function LoadQuery(ByVal ParentDB As DADatabase, ByVal qd As QueryDef) As Boolean
  Dim rsdetail As recordset
  Dim rsrelations As recordset
  Dim rsQuery As recordset, rsParsedQry As recordset
  Dim qname As String

  On Error GoTo loadquery_err
  Call xSet("loadquery")
  
  If ParentDB Is Nothing Then Exit Function
  Set mParentDb = ParentDB
  qname = qd.name
  'mTransient = False
  
  If Not LoadHeader(qd) Then Exit Function
  Set rsQuery = mParentDb.ParsedQueries.GetParsedQueryRS(mParentDb.database, QUERY_HEADER_RS)
  rsQuery.FindFirst "Name='" & qname & "'"
  If rsQuery.NoMatch Then Err.Raise ERR_INVALID_QUERY_LOAD
  
  strSql = "" & rsQuery.fields("strsql")
  strAssocquery = "" & rsQuery.fields("assoc_strsql")
  Set rsQuery = Nothing
  
'  Set rsdetail = mParentDb.ParsedQueries.GetParsedQueryRS(mParentDb.database, QUERY_DETAIL_RS)
'  rsdetail.FindFirst "Name='" & qname & "'"
'  While Not rsdetail.NoMatch
'    Select Case rsdetail.fields("Detailtype")
'      Case "SOURCEFIELD"
'        Call AddList(rsdetail.fields("Detailname"), strsourcefields)
'      Case "DESTFIELD"
'        Call AddList(rsdetail.fields("Detailname"), strdestfields)
'      Case Else
'        ECASE "INVALID detailtype for " & mName & " - " & rsdetail.fields("Detailtype")
'    End Select
'    rsdetail.FindNext "Name='" & qname & "'"
'  Wend
'  Set rsdetail = Nothing
  
  Set rsrelations = mParentDb.ParsedQueries.GetParsedQueryRS(mParentDb.database, QUERY_RELATIONS_RS)
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
  
  Call AddtoParent
  
  If Not mTransient Then
    readmodes
'    boocalced = mParentDb.flag(mName, intcalcno)
'  Else
'    boocalced = False
  End If
  LoadQuery = True

loadquery_end:
  Set rsdetail = Nothing
  Set rsrelations = Nothing
  Set rsQuery = Nothing
  Call xReturn("loadquery")
  Exit Function
  
loadquery_err:
  Err.Raise ERR_LOAD_QUERY, "tcsda.loadquery", Err.Description
  Resume Next
Resume
End Function

Public Function LoadQueryName(ByVal ParentDB As DADatabase, ByVal qdname As String) As Boolean
  Dim rsdetail As recordset
  Dim rsrelations As recordset
  Dim rsQuery As recordset
  Dim rsParsedQry As recordset
  Dim i As Long
  Dim s As String
  Dim sCacheParams() As String

  On Error GoTo LoadQueryName_err
  Call xSet("LoadQueryName")
  
  If ParentDB Is Nothing Then Exit Function
  Set mParentDb = ParentDB
  
  ' next line redundant once incoproated below
  'If Not LoadHeader(qd) Then Exit Function
  
  ' load header information
  mTransient = False
  Set rsQuery = mParentDb.ParsedQueries.GetParsedQueryRS(mParentDb.database, QUERY_HEADER_RS)
  rsQuery.FindFirst "Name='" & qdname & "'"
  If rsQuery.NoMatch Then Err.Raise ERR_INVALID_QUERY_LOAD
  
  ' also set type of query and parameters as per ?... from change to sys-query-header
  ' makes loadheadername redundant
  strSql = "" & rsQuery.fields("strsql")
  strAssocquery = "" & rsQuery.fields("assoc_strsql")
  strParameters = "" & rsQuery.fields("parameters")
  lngParamcount = IsNullEx(rsQuery.fields("paramcount"), 0)
  mName = "" & rsQuery.fields("Name")
  mType = IsNullEx(rsQuery.fields("Type"), 0)
  mCached = True
  
  If mType = 1 Or mType = 3 Or mType = 65 Then  'Access internal type
    Exit Function
  End If
    
  If Len(IsNullEx(strParameters, "")) > 0 Then
    sCacheParams = Split(IsNullEx(strParameters, ""), ";")
    ReDim strparamnames(1 To UBound(sCacheParams) + 1)
    ReDim varparams(1 To UBound(sCacheParams) + 1)
    For i = 1 To UBound(sCacheParams) + 1
      s = sCacheParams(i - 1)
      If StrComp(Left$(s, 1), "[") = 0 And StrComp(Right$(s, 1), "]") = 0 Then s = Mid$(s, 2, Len(s) - 2)
      strparamnames(i) = UCase$(s)
    Next i
  End If
  
  Set rsQuery = Nothing
  
  
  ' load relationships
  Set rsrelations = mParentDb.ParsedQueries.GetParsedQueryRS(mParentDb.database, QUERY_RELATIONS_RS)
  rsrelations.FindFirst "Name='" & qdname & "'"
  While Not rsrelations.NoMatch
    Select Case rsrelations.fields("relationship")
      Case "CHILD"
        If rsrelations.fields("ASSOCIATED") Then
        ' leave as string arrays for now as quite quick
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
    rsrelations.FindNext "Name='" & qdname & "'"
  Wend
  Set rsrelations = Nothing
  
  Call AddtoParent
  
  If Not mTransient Then
    readmodes
  End If
  LoadQueryName = True

LoadQueryName_end:
  Set rsdetail = Nothing
  Set rsrelations = Nothing
  Set rsQuery = Nothing
  Call xReturn("LoadQueryName")
  Exit Function
  
LoadQueryName_err:
  Err.Raise ERR_LOAD_QUERY_NAME, "tcsda.LoadQueryName", Err.Description
  Resume Next
Resume
End Function

Private Function LoadHeader(ByVal qd As QueryDef) As Boolean  ', Optional ParseQD As QueryDef) As Boolean
  Dim i As Long
  Dim s As String
  Dim rsCache As recordset, bCached As Boolean, sCacheParams() As String
  
  On Error GoTo LoadHeader_Err
  Call xSet("LoadHeader")
  
  Set rsCache = mParentDb.ParsedQueries.GetParsedQueryRS(mParentDb.database, QUERY_HEADER_RS, True)
  
  If Not rsCache Is Nothing Then
    rsCache.FindFirst ("Name=" & ATC2CORE.strSql(qd.name))
    If Not rsCache.NoMatch Then
      bCached = (rsCache!LastUpdated = qd.LastUpdated)
    End If
  Else
    bCached = False
  End If

  If Not bCached Then
   ' apf this is the slow bit
    mName = qd.name
    mType = qd.Type
    strSql = qd.sql
  Else
    mName = rsCache!name
    mType = rsCache!Type
    strSql = rsCache!strSql
    mCached = True
  End If
  If mType = 1 Or mType = 3 Or mType = 65 Then  'Access internal type
    Exit Function
  End If
    
  If mCached Then
    If Len(IsNullEx(rsCache!Parameters, "")) > 0 Then
      sCacheParams = Split(IsNullEx(rsCache!Parameters, ""), ";")
      ReDim strparamnames(1 To UBound(sCacheParams) + 1)
      ReDim varparams(1 To UBound(sCacheParams) + 1)
      For i = 1 To UBound(sCacheParams) + 1
        s = sCacheParams(i - 1)
        If StrComp(Left$(s, 1), "[") = 0 And StrComp(Right$(s, 1), "]") = 0 Then s = Mid$(s, 2, Len(s) - 2)
        strparamnames(i) = UCase$(s)
      Next i
    End If
  Else
    If qd.Parameters.Count > 0 Then
      ReDim strparamnames(1 To qd.Parameters.Count)
      ReDim varparams(1 To qd.Parameters.Count)
      For i = 1 To qd.Parameters.Count
        s = qd.Parameters(i - 1).name
        If StrComp(Left$(s, 1), "[") = 0 And StrComp(Right$(s, 1), "]") = 0 Then s = Mid$(s, 2, Len(s) - 2)
        strparamnames(i) = UCase$(s)
      Next i
    End If
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

'Private Function LoadHeaderName(ByVal qdname As String) As Boolean  ', Optional ParseQD As QueryDef) As Boolean
'  Dim i As Long
'  Dim s As String
'  Dim rsCache As recordset, bCached As Boolean, sCacheParams() As String
'  On Error GoTo LoadHeaderName_Err
'  Call xSet("LoadHeaderName")
'
'  ' move functionality to LoadQueryName
'  Set rsCache = mParentDb.ParsedQueries.GetParsedQueryRS(mParentDb.database, QUERY_CACHE, True)
'  If Not rsCache Is Nothing Then
'    rsCache.FindFirst ("Name=" & TCSCORE.strsql(qdname))
'    If Not rsCache.NoMatch Then
'      bCached = (rsCache!LastUpdated = qd.LastUpdated)
'    End If
'  Else
'    bCached = False
'  End If
'
'  If Not bCached Then
'   ' apf this is the slow bit
'    mName = qd.name
'    mType = qd.Type
'    strsql = qd.sql
'  Else
'    mName = rsCache!name
'    mType = rsCache!Type
'    strsql = rsCache!sql
'    mCached = True
'  End If
'  If mType = 1 Or mType = 3 Or mType = 65 Then  'Access internal type
'    Exit Function
'  End If
'
'  If mCached Then
'
'    If Len(IsNullEx(rsCache!Parameters, "")) > 0 Then
'      sCacheParams = Split(IsNullEx(rsCache!Parameters, ""), ";")
'      ReDim strparamnames(1 To UBound(sCacheParams) + 1)
'      ReDim varparams(1 To UBound(sCacheParams) + 1)
'      For i = 1 To UBound(sCacheParams) + 1
'        s = sCacheParams(i - 1)
'        If StrComp(Left$(s, 1), "[") = 0 And StrComp(Right$(s, 1), "]") = 0 Then s = Mid$(s, 2, Len(s) - 2)
'        strparamnames(i) = UCase$(s)
'      Next i
'    End If
'  Else
'    If qd.Parameters.Count > 0 Then
'      ReDim strparamnames(1 To qd.Parameters.Count)
'      ReDim varparams(1 To qd.Parameters.Count)
'      For i = 1 To qd.Parameters.Count
'        s = qd.Parameters(i - 1).name
'        If StrComp(Left$(s, 1), "[") = 0 And StrComp(Right$(s, 1), "]") = 0 Then s = Mid$(s, 2, Len(s) - 2)
'        strparamnames(i) = UCase$(s)
'      Next i
'    End If
'  End If
'
'  LoadHeaderName = True
'
'LoadHeaderName_End:
'  Call xReturn("LoadHeaderName")
'  Exit Function
'
'LoadHeaderName_Err:
'  Call ErrorMessage(ERR_ERROR, Err, "LoadHeaderName", "Error in LoadHeaderName", "Undefined error.")
'  Resume LoadHeaderName_End
'  Resume
'End Function

'Friend Sub CacheQuery()
'  Dim rsCache As recordset
'  Dim qd As QueryDef
'  Dim i As Long, sParam As String
'
'  On Error GoTo CacheQuery_err
'  Set rsCache = mParentDb.ParsedQueries.GetParsedQueryRS(mParentDb.database, QUERY_HEADER_RS, False)
'  If rsCache Is Nothing Then Exit Sub
'  If Not InCollection(mParentDb.database.QueryDefs, mName) Then
'    Call ErrorMessage(ERR_ERRORSILENT, Err, "CacheQuery", "Query not found", "Query not found in database - " & mName)
'    GoTo CacheQuery_end
'  End If
'  Set qd = mParentDb.database.QueryDefs(mName)
'  rsCache.addnew
'    rsCache!name = mName
'    rsCache!strSql = strSql
'    rsCache!Type = mType
'    rsCache!LastUpdated = qd.LastUpdated
'    If UBound(strparamnames) > 0 Then
'      For i = 1 To UBound(strparamnames)
'        sParam = sParam & strparamnames(i) & ";"
'      Next i
'      sParam = Left(sParam, Len(sParam) - 1)
'    End If
'    rsCache!Parameters = sParam
'    rsCache!paramcount = UBound(strparamnames)
'  rsCache.update
'
'CacheQuery_end:
'  Exit Sub
'
'CacheQuery_err:
'  If Not rsCache Is Nothing Then
'    If rsCache.EditMode <> dbEditNone Then rsCache.CancelUpdate
'  End If
'  Call ErrorMessage(ERR_ERROR, Err, "CacheQuery", "CacheQuery", "An error occurred caching a query to make it faster to load.")
'  Resume CacheQuery_end
'  Resume
'End Sub

Private Function LoadHeaderSQL(sql As String) As Boolean  ', Optional ParseQD As QueryDef) As Boolean
  Dim bsel As Long
  Dim bapp As Long
  Dim bupd As Long
  Dim bdel As Long

  On Error GoTo LoadHeaderSQL_Err
  Call xSet("LoadHeaderSQL")
    
  mName = sql
  strSql = sql

  ' establish query type
  bsel = InstrSQL(1, sql, "SELECT", vbTextCompare)
  bapp = InstrSQL(1, sql, "INSERT", vbTextCompare)
  bupd = InstrSQL(1, sql, "UPDATE", vbTextCompare)
  bdel = InstrSQL(1, sql, "DELETE", vbTextCompare)
  
  If bsel = 0 And bapp = 0 And bupd = 0 And bdel = 0 Then
    Err.Raise ERR_SQL_PARSE, "tcsda.LoadHeaderSQL", "Unknown query type for," & vbCr & vbCr & sql
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

Private Function AddQueryEntryinDb(ByVal rs_t As RS_TYPE, ByVal Entry_t As ENTRY_TYPE, ByVal s As String) As Boolean
Dim rs As recordset
Dim i As Long
Dim sParam As String

  On Error GoTo AddQueryEntryinDb_Err
  Call xSet("AddQueryEntryinDb")
  Set rs = mParentDb.ParsedQueries.GetParsedQueryRS(mParentDb.database, rs_t, False)
  s = Trim(s)
  rs.addnew
  rs.fields("NAME") = mName
  Select Case Entry_t
  Case HEADER_ENTRY
    rs.fields("TYPE") = SdbQType
    rs.fields("LASTUPDATED") = mParentDb.database.QueryDefs(mName).LastUpdated
    rs.fields("STRSQL") = Left(s, InStr(s, "|ASSOC|") - 1)
    rs.fields("ASSOC_STRSQL") = RightPart(s, InStr(s, "|ASSOC|") + 7)
    If UBound(strparamnames) > 0 Then
      For i = 1 To UBound(strparamnames)
        sParam = sParam & strparamnames(i) & ";"
      Next i
      sParam = Left(sParam, Len(sParam) - 1)
    End If
    rs.fields("PARAMETERS") = sParam
    rs.fields("PARAMCOUNT") = UBound(strparamnames)
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
Dim sql As String

  On Error GoTo ClearQueryinDB_Err
  Call xSet("ClearQueryinDB")
  sql = "delete * from sys_queries_Header where Name='" & mName & "'"
  mParentDb.database.execute (sql)
'  sql = "delete * from sys_queries_detail where Name='" & mName & "'"
'  mParentDb.database.execute (sql)
  sql = "delete * from sys_queries_relations where Name='" & mName & "'"
  mParentDb.database.execute (sql)
' could use batchexecute
ClearQueryinDB_End:
  Call xReturn("ClearQueryinDB")
  Exit Function

ClearQueryinDB_Err:
  Call ErrorMessage(ERR_ERROR, Err, "ClearQueryinDB", "Error in ClearQueryinDB", "Undefined error.")
  Resume ClearQueryinDB_End
  Resume
End Function

Private Function SetUpQuerySQLappend(sql As String) As Boolean
  Dim pos As Long
  Dim b As Boolean
  Dim s As String

  On Error GoTo SetUpQuerySQLappend_Err
  Call xSet("SetUpQuerySQLappend")
  
  pos = InStr(sql, "(")
  If pos < 13 Then Err.Raise ERR_SQL_PARSE, "tcsda.setupquery", "Invalid append sql in " & mName & ", " & vbCr & vbCr & sql
  s = Trim(Mid$(sql, 13, pos - 13))
  If Not (InCollection(mParentDb.database.TableDefs, s) Or InCollection(mParentDb.database.QueryDefs, s)) Then Err.Raise ERR_SQL_PARSE, "tcsda.setupquery", "Invalid append sql in " & mName & ", " & vbCr & vbCr & sql & vbCr & vbCr & " as non existent child, " & vbCr & s
  AddList s, strchildren
  
  If mHasBids Then
    pos = InstrSQL(1, sql, "SELECT", vbTextCompare)
    If pos <> 0 Then
      b = SetUpQuerySQLselect(RightPart(sql, pos))
    Else
      pos = InstrSQL(1, sql, "VALUES", vbTextCompare)
      If pos = 0 Then Err.Raise ERR_SQL_PARSE, "tcsda.setupquery", "No select statement found in append query, " & mName & ", " & vbCr & vbCr & sql
    End If
    strAssocquery = "DELETE FROM " + strchildren(1) + " WHERE source='" + mName + "'"
    
    If InList(strchildren(1), strparents) Then
      AddList strchildren(1), strassocchildren
      If Not mTransient Then Call AddQueryEntryinDb(QUERY_RELATIONS_RS, ASSOCCHILD_ENTRY, strassocchildren(1))
      ReDim strchildren(0)
    Else
      If Not mTransient Then Call AddQueryEntryinDb(QUERY_RELATIONS_RS, CHILD_ENTRY, strchildren(1))
    End If
  Else
    b = True
  End If

  SetUpQuerySQLappend = True

SetUpQuerySQLappend_End:
  Call xReturn("SetUpQuerySQLappend")
  Exit Function

SetUpQuerySQLappend_Err:
  Err.Raise ERR_SETUP_SQL_APPEND, Err.Description
'  Call ErrorMessage(ERR_ERROR, Err, "SetUpQuerySQLappend", "Error in SetUpQuerySQLappend", "Undefined error.")
'  Resume SetUpQuerySQLappend_End
Resume
End Function

Private Function SetUpQuerySQLUnion(sql As String) As Boolean
  Dim pos As Long
  Dim posUnion As Long
  Dim bFirst As Boolean
  Dim bSecond As Boolean
  Dim s As String
  Dim Endpos As Long

  On Error GoTo SetUpQuerySQLUnion_Err
  Call xSet("SetUpQuerySQLUnion")
  ' zzzz Needs to check
  posUnion = InStr(sql, "UNION")
  
  If mHasBids Then
    pos = InstrSQL(1, sql, "SELECT", vbTextCompare)
    If pos <> 0 Then
      If StrComp(Mid$(sql, posUnion - 2, 1), ")", vbTextCompare) = 0 Then
        bFirst = SetUpQuerySQLselect(Mid$(sql, pos, posUnion - 4))
      Else
        bFirst = SetUpQuerySQLselect(Mid$(sql, pos, posUnion))
      End If
      pos = InstrSQL(posUnion, sql, "SELECT", vbTextCompare)
      Endpos = InStr(sql, ";")
      If StrComp(Mid$(sql, posUnion + 6, 1), "(", vbTextCompare) = 0 Then
        bSecond = SetUpQuerySQLselect(Mid$(sql, pos, Endpos - pos - 1))
      Else
        bSecond = SetUpQuerySQLselect(Mid$(sql, pos, Endpos - pos))
      End If
    Else
      Err.Raise ERR_SQL_PARSE, "tcsda.setupquery", "No select statement found in union query, " & mName & ", " & vbCr & vbCr & sql
    End If
  Else
    bFirst = True
    bSecond = True
  End If

  SetUpQuerySQLUnion = True

SetUpQuerySQLUnion_End:
  Call xReturn("SetUpQuerySQLUnion")
  Exit Function

SetUpQuerySQLUnion_Err:
  Err.Raise ERR_SETUP_SQL_UNION, Err.Description
  Resume
End Function

Private Function SetUpQuerySQLupdate(sql As String) As Boolean
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
  pos = InStr(sql, "SET")
  If pos = 0 Then Err.Raise ERR_SQL_PARSE, "tcsda.setupquery", "Invalid update sql in " & mName & ", " & vbCr & vbCr & sql

  If mHasBids Then
    strAssocquery = "SELECT * FROM " + Mid$(sql, 8, pos - 9)
    b = SetUpQuerySQLselect(strAssocquery)
  Else
    b = True
  End If
  pos = pos + 3
  Do
    pos = SplitNextEqualsPair(sql, pos, df, sf)
    If Len(sf) Then
      AddList sf, strsourcefields
'      If Not mTransient Then Call AddQueryEntryinDb(QUERY_DETAIL_RS, SOURCEFIELD_ENTRY, sf)
      AddList df, strdestfields
'      If Not mTransient Then Call AddQueryEntryinDb(QUERY_DETAIL_RS, DESTFIELD_ENTRY, df)
      desttable = TableName(df)
      If Len(desttable) Then
        If Not (InCollection(mParentDb.database.TableDefs, desttable) Or InCollection(mParentDb.database.QueryDefs, desttable)) Then Err.Raise ERR_SQL_PARSE, "tcsda.setupquery", "Invalid update sql in " & mName & ", " & vbCr & vbCr & sql & vbCr & vbCr & " as non existent child, " & vbCr & desttable
        ' zzzz should this be if inlist(desttable, strparents) then add to strchildren else add to assocchildren
        If Not InList(desttable, strassocchildren) Then
          AddList desttable, strassocchildren
          If Not mTransient Then Call AddQueryEntryinDb(QUERY_RELATIONS_RS, ASSOCCHILD_ENTRY, desttable)
        End If
      Else
      ' if only one dest table tyhen if none supplied it must be that one

        #If DEBUGVER Then
          logmessage "tcsda.setupquery: can not handle parsing update sql as need to get table from source tables in " & mName & ", " & vbCr & vbCr & sql
        #End If
      End If
    End If
  Loop While pos > 0
  If mHasBids Then
    strAssocquery = strAssocquery + " WHERE ("
    For i = 1 To UBound(strsourcefields)
      If i > 1 Then strAssocquery = strAssocquery + " OR "
      strAssocquery = strAssocquery + strsourcefields(i) + "<>" + strdestfields(i)
    Next i
    For i = 1 To UBound(strdestfields)
      strAssocquery = strAssocquery + " OR " + strdestfields(i) + " IS NULL "
    Next i
    strAssocquery = strAssocquery + ")"
    pos = InStr(sql, "WHERE")
    If pos > 0 Then
      strAssocquery = strAssocquery + "and " + RightPart(sql, pos + 5)
    End If
  End If
  SetUpQuerySQLupdate = True

SetUpQuerySQLupdate_End:
  Call xReturn("SetUpQuerySQLupdate")
  Exit Function

SetUpQuerySQLupdate_Err:

  Err.Raise ERR_SETUP_SQL_UPDATE, Err.Description
'  Call ErrorMessage(ERR_ERROR, Err, "SetUpQuerySQLupdate", "Error in SetUpQuerySQLupdate", "Undefined error.")
'  Resume SetUpQuerySQLupdate_End
  Resume
End Function

Private Function SetUpQuerySQLselect(sql As String) As Boolean

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
  
  Newsql = sql
  ' replace vbcr and vblf with " "
  Newsql = Replace(Newsql, vbCr, " ", 1, -1, vbTextCompare)
  Newsql = Replace(Newsql, vbLf, " ", 1, -1, vbTextCompare)
    
  ' identify start
  Startpos = InstrSQL(1, Newsql, " FROM ", vbTextCompare)
  If Startpos = 0 Then Err.Raise ERR_SQL_PARSE, "tcsda.setupquery", "Invalid select sql in " & mName & ", " & vbCr & vbCr & sql
  Startpos = Startpos + 6
  
  ' identify end
  blongEND = True
  Endpos = Len(Newsql)
  For i = LBound(endstr) To UBound(endstr)
    If Len(endstr(i)) Then
      TempEndpos = InstrSQL(Startpos, Newsql, endstr(i), vbTextCompare)
      If TempEndpos > 0 Then
        Endpos = min(Endpos, TempEndpos)
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
      If Not (InCollection(mParentDb.database.TableDefs, sparent) Or InCollection(mParentDb.database.QueryDefs, sparent)) Then Err.Raise ERR_SQL_PARSE, "tcsda.setupquery", "Invalid append sql in " & mName & ", " & vbCr & vbCr & sql & vbCr & vbCr & " as non existent parent, " & vbCr & sparent
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
          If TempEndpos > 0 Then Endpos = min(Endpos, TempEndpos + Len(sepstr(i)))
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
  Err.Raise ERR_SETUP_SQL_SELECT, Err.Description
'  Call ErrorMessage(ERR_ERROR, Err, "SetUpQuerySQLselect", "Error in SetUpQuerySQLselect", "Undefined error.")
'  Resume SetUpQuerySQLselect_End
Resume
End Function

Public Property Get Cached() As Boolean
  Cached = mCached
End Property

Public Property Get QueryType() As Long
  QueryType = mType
End Property

Private Sub Class_Terminate()
  On Error Resume Next
  Kill
End Sub

Private Sub AddtoParent()
  Dim qtemp As Query

  Set qtemp = Me
  If Not InCollection(mParentDb.queries, mName) Then mParentDb.queries.Add qtemp, mName

End Sub
