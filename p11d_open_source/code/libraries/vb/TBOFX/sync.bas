Attribute VB_Name = "sync"
Option Explicit

Public Function syncStructure(dbSrc As Database, dbDest As Database, mStructureFilters As Collection, sLog As QString, dbChange As dbChangeDetails, sCheckFldFlags() As String, sCheckIdxFlags() As String) As Boolean
  Dim flt As syncFilter
  Dim bRemoveFields As Boolean
  Dim i As Long, j As Long
  Dim k As Long, Max As Long
  Dim sFilter As String
  
  On Error GoTo syncStructure_Err
  Call xSet("syncStructure")
  
  Call SetCursor
  Call sLog.Append("<STRUCTURESYNC>")
  
  
  For i = 1 To mStructureFilters.Count
    Set flt = mStructureFilters(i)
    If flt.FilterType = EXCLUDE_FILTER Then
      sFilter = " (NOT " & flt.FilterString & ")"
    Else
      sFilter = " (" & flt.FilterString & ")"
    End If
    
    Call sLog.Append("<FILTER><FILTERSTRING>" & flt.FilterString & "</FILTERSTRING>")
    
    k = 0
    If (flt.FilterActions And OVERWRITE_TABLE_STRUCTURE) = OVERWRITE_TABLE_STRUCTURE Then
      bRemoveFields = Not ((flt.FilterActions And NO_REMOVE_FIELDS) = NO_REMOVE_FIELDS)
      Call sLog.Append("<TABLESTRUCT>" & vbCrLf)
      For j = 1 To dbChange.ChangedTables.Count
        Call ProvideFeedback(j, dbChange.ChangedTables.Count, "Synchronising table structures" & sFilter)
        If flt.FilterApplies(dbChange.ChangedTables(j).TableName) Then
          Call sLog.Append("<UPDATETABLE Start=""" & Now & """>" & dbChange.ChangedTables(j).TableName & "</UPDATETABLE>" & vbCrLf)
          Call ProvideFeedback(j, dbChange.ChangedTables.Count, "Synchronising table structures" & sFilter & " - " & dbChange.ChangedTables(j).TableName)
          Call syncTable(dbChange.ChangedTables(j), dbSrc, dbDest, bRemoveFields, sCheckIdxFlags, sCheckFldFlags, sLog)
          k = k + 1
        End If
      Next
      For j = 1 To dbChange.NewTables.Count
        Call ProvideFeedback(j, dbChange.NewTables.Count, "Adding new tables" & sFilter)
        If flt.FilterApplies(dbChange.NewTables.Item(j)) Then
          Call sLog.Append("<UPDATETABLE Start=""" & Now & """>" & dbChange.NewTables.Item(j) & "</UPDATETABLE>" & vbCrLf)
          Call ProvideFeedback(j, dbChange.NewTables.Count, "Adding new tables" & sFilter & " - " & dbChange.NewTables.Item(j))
          Call CopyEntireTable(dbSrc.TableDefs(dbChange.NewTables.Item(j)), dbDest, sCheckFldFlags, sCheckIdxFlags)
          k = k + 1
        End If
      Next
      Call sLog.Append("<COUNT>" & k & "</COUNT></TABLESTRUCT>" & vbCrLf)
    End If
    
    k = 0
    If (flt.FilterActions And OVERWRITE_QUERY_STRUCTURE) = OVERWRITE_QUERY_STRUCTURE Then
      Call sLog.Append("<QUERYSTRUCT>" & vbCrLf)
      For j = 1 To dbChange.ChangedQueries.Count
        Call ProvideFeedback(j, dbChange.ChangedQueries.Count, "Updating queries" & sFilter)
        If flt.FilterApplies(dbChange.ChangedQueries.Item(j)) Then
          Call syncQuery(dbChange.ChangedQueries.Item(j), dbSrc, dbDest)
          Call sLog.Append("<UPDATEQUERY>" & dbChange.ChangedQueries.Item(j) & "</UPDATEQUERY>" & vbCrLf)
          k = k + 1
        End If
      Next
      Call sLog.Append("<COUNT>" & k & "</COUNT></QUERYSTRUCT>" & vbCrLf)
      DoDBEvents (FREE_LOCKS + REFRESH_CACHE)
      dbDest.QueryDefs.Refresh
    End If
    
    k = 0
    If (flt.FilterActions And DELETE_TABLES) = DELETE_TABLES Then
      Call sLog.Append("<TABLEDEL>" & vbCrLf)
      For j = 1 To dbChange.OldTables.Count
        Call ProvideFeedback(j, dbChange.OldTables.Count, "Deleting tables" & sFilter)
        If flt.FilterApplies(dbChange.OldTables.Item(j)) Then
          Call dbDest.TableDefs.Delete(dbChange.OldTables.Item(j))
          Call sLog.Append("<DELETETABLE>" & dbChange.OldTables.Item(j) & "</DELETETABLE>" & vbCrLf)
          k = k + 1
        End If
      Next
      Call sLog.Append("<COUNT>" & k & "</COUNT></TABLEDEL>" & vbCrLf)
'      Call ProvideFeedback(j, dbChange.OldTables.Count + 2, "Refreshing tables" & sFilter)
      DoDBEvents (FREE_LOCKS + REFRESH_CACHE)
'      dbDest.TableDefs.Refresh
'      Call ProvideFeedback(j + 1, dbChange.OldTables.Count + 2, "Refreshing tables" & sFilter)
    End If
    
    k = 0
    If (flt.FilterActions And DELETE_QUERIES) = DELETE_QUERIES Then
      Call sLog.Append("<QUERYDEL>" & vbCrLf)
      For j = 1 To dbChange.OldQueries.Count
        Call ProvideFeedback(j, dbChange.OldQueries.Count, "Deleting queries(" & sFilter)
        If flt.FilterApplies(dbChange.OldQueries.Item(j)) Then
          Call dbDest.QueryDefs.Delete(dbChange.OldQueries.Item(j))
          Call sLog.Append("<DELETEQUERY>" & dbChange.OldQueries.Item(j) & "</DELETEQUERY>" & vbCrLf)
          k = k + 1
        End If
      Next
      Call sLog.Append("<COUNT>" & k & "</COUNT></QUERYDEL>" & vbCrLf)
'      Call ProvideFeedback(j, dbChange.OldQueries.Count + 2, "Refreshing queries(" & sFilter)
      DoDBEvents (FREE_LOCKS + REFRESH_CACHE)
'      dbDest.QueryDefs.Refresh
'      Call ProvideFeedback(j + 1, dbChange.OldQueries.Count + 2, "Refreshing queries(" & sFilter)
    End If
    
    Call sLog.Append("</FILTER>")
  Next i
  
  Call ProvideFeedback(0, 4, "Refreshing database")
  dbDest.TableDefs.Refresh
  Call ProvideFeedback(1, 4, "Refreshing database")
  dbSrc.TableDefs.Refresh
  Call ProvideFeedback(2, 4, "Refreshing database")
  dbDest.QueryDefs.Refresh
  Call ProvideFeedback(3, 4, "Refreshing database")
  dbSrc.QueryDefs.Refresh
  Call ProvideFeedback(4, 4, "Refreshing database")
  
syncStructure_End:
  Call sLog.Append("</STRUCTURESYNC>")
  Call ClearCursor
  Call xReturn("syncStructure")
  Exit Function

syncStructure_Err:
  Call ErrorMessage(ERR_ERROR, Err, "syncStructure", "Error in syncStructure", "Undefined error.")
  Resume syncStructure_End
  Resume
End Function

Private Function syncQuery(sQuery As String, dbSrc As Database, dbDest As Database) As Boolean
  Dim qd As QueryDef
  On Error GoTo syncQuery_Err
  Call xSet("syncQuery")

  If Not InCollection(dbDest.QueryDefs, sQuery) Then
    Set qd = New QueryDef
    qd.sql = dbSrc.QueryDefs(sQuery).sql
    qd.Name = sQuery
    Call dbDest.QueryDefs.Append(qd)
  Else
    Set qd = dbDest.QueryDefs(sQuery)
    qd.sql = dbSrc.QueryDefs(sQuery).sql
  End If
  Call DoDBEvents(FREE_LOCKS + REFRESH_CACHE)
  'Call dbDest.QueryDefs.Refresh
  
syncQuery_End:
  Call xReturn("syncQuery")
  Exit Function

syncQuery_Err:
  Call ErrorMessage(ERR_ERROR, Err, "syncQuery", "Error in syncQuery", "Undefined error.")
  Resume syncQuery_End
  Resume
End Function

Private Function syncTable(tblChange As TblChangeDetails, dbSrc As Database, dbDest As Database, bRemoveFields As Boolean, sCheckIdxProps() As String, sCheckFldProps() As String, sLog As QString) As Boolean
  Dim tdSrc As TableDef, tdDest As TableDef
  Dim fld As DAO.Field, fldDest As DAO.Field
  Dim idx As DAO.Index, idxDest As DAO.Index
  Dim fldChange As fldChangeDetails
  Dim i As Long, j As Long
  
  On Error GoTo syncTable_Err
  Call xSet("sycTable")

  Set tdSrc = dbSrc.TableDefs(tblChange.TableName)
  Set tdDest = dbDest.TableDefs(tblChange.TableName)
  
  'Remove fields not in src
  If bRemoveFields Then
    For i = 1 To tblChange.OldFields.Count
      If InCollection(tdDest.Fields, tblChange.OldFields.Item(i)) Then
        Call tdDest.Fields.Delete(tblChange.OldFields.Item(i))
      End If
    Next
  End If
  
  'Must remove all indexes as it is not possible to change their properties when appended to the indexes collection
  For i = tdDest.Indexes.Count - 1 To 0 Step -1
    Call tdDest.Indexes.Delete(tdDest.Indexes(i).Name)
  Next
   
  'Add in fields not in dest
  For i = 1 To tblChange.NewFields.Count
     Set fld = tdSrc.Fields(tblChange.NewFields.Item(i))
    Set fldDest = tdDest.CreateField(fld.Name, fld.Type)
    For j = LBound(sCheckFldProps) To UBound(sCheckFldProps)
      If Not (fldDest.Properties(sCheckFldProps(j)).Value = fld.Properties(sCheckFldProps(j)).Value) Then
        fldDest.Properties(sCheckFldProps(j)).Value = fld.Properties(sCheckFldProps(j)).Value
      End If
    Next
    Call tdDest.Fields.Append(fldDest)
  Next
  
  For i = 1 To tblChange.ChangedFields.Count
    Set fldChange = tblChange.ChangedFields(i)
    If fldChange.RequireNewField Then
      Set fld = tdDest.Fields(fldChange.Name)
      Set fldDest = tdDest.CreateField(fld.Name & "____OLD", fld.Type)
      For j = LBound(sCheckFldProps) To UBound(sCheckFldProps)
        If Not (fldDest.Properties(sCheckFldProps(j)).Value = fld.Properties(sCheckFldProps(j)).Value) Then
          fldDest.Properties(sCheckFldProps(j)).Value = fld.Properties(sCheckFldProps(j)).Value
        End If
      Next
      Call tdDest.Fields.Append(fldDest)
      Call dbDest.Execute("UPDATE " & tdDest.Name & " SET [" & fld.Name & "____OLD]=[" & fld.Name & "]", dbFailOnError)
      Set fld = Nothing
      Call tdDest.Fields.Delete(fldChange.Name)
      Set fld = tdSrc.Fields(fldChange.Name)
      Set fldDest = tdDest.CreateField(fld.Name, fld.Type)
      For j = LBound(sCheckFldProps) To UBound(sCheckFldProps)
        If Not (fldDest.Properties(sCheckFldProps(j)).Value = fld.Properties(sCheckFldProps(j)).Value) Then
          fldDest.Properties(sCheckFldProps(j)).Value = fld.Properties(sCheckFldProps(j)).Value
        End If
      Next
      Call tdDest.Fields.Append(fldDest)
      Call dbDest.Execute("UPDATE " & tdDest.Name & " SET [" & fld.Name & "]=[" & fld.Name & "____OLD]", dbFailOnError)
      Call tdDest.Fields.Delete(fldChange.Name & "____OLD")
    Else
      Set fld = tdSrc.Fields(fldChange.Name)
      Set fldDest = tdDest.Fields(fldChange.Name)
      For j = 1 To fldChange.ChangedProperties.Count
        fldDest.Properties(fldChange.ChangedProperties.Item(j)).Value = fld.Properties(fldChange.ChangedProperties.Item(j)).Value
      Next
    End If
  Next
  
  Call AddIndexes(tdSrc, tdDest, sCheckIdxProps)
  Call DoDBEvents(FREE_LOCKS + REFRESH_CACHE)
  'Call dbDest.TableDefs.Refresh
   
syncTable_End:
  Call xReturn("syncTable")
  Exit Function

syncTable_Err:
  Call ErrorMessage(ERR_ERROR, Err, "syncTable", "Error in syncTable", "An error occurred synchronising tables.")
  Resume syncTable_End
  Resume
End Function

Public Function CopyEntireTable(tdSrc As TableDef, dbDest As Database, sCheckFldProps() As String, sCheckIdxProps() As String) As Boolean
  Dim tdDest As TableDef
  Dim fld As DAO.Field, fldDest As DAO.Field
  Dim idx As DAO.Index, idxDest As DAO.Index
  Dim i As Long
  
  On Error GoTo CopyEntireTable_Err
  Call xSet("CopyEntireTable")
  
  Set tdDest = dbDest.CreateTableDef(tdSrc.Name)
  
  For Each fld In tdSrc.Fields
    Set fldDest = tdDest.CreateField(fld.Name, fld.Type)
    For i = LBound(sCheckFldProps) To UBound(sCheckFldProps)
      If Not (fldDest.Properties(sCheckFldProps(i)).Value = fld.Properties(sCheckFldProps(i)).Value) Then
        fldDest.Properties(sCheckFldProps(i)).Value = fld.Properties(sCheckFldProps(i)).Value
      End If
    Next
    Call tdDest.Fields.Append(fldDest)
  Next
  
  Call AddIndexes(tdSrc, tdDest, sCheckIdxProps)
  Call dbDest.TableDefs.Append(tdDest)
  Call DoDBEvents(FREE_LOCKS + REFRESH_CACHE)
  'Call dbDest.TableDefs.Refresh
  
  CopyEntireTable = True

CopyEntireTable_End:
  Call xReturn("CopyEntireTable")
  Exit Function

CopyEntireTable_Err:
  CopyEntireTable = False
  Call ErrorMessage(ERR_ERROR, Err, "CopyEntireTable", "Error Copying Table", "An error occurred copying the table '" & tdSrc.Name & "'.")
  Resume CopyEntireTable_End
  Resume
End Function

Private Function AddIndexes(tdSrc As TableDef, tdDest As TableDef, sCheckIdxProps() As String) As Boolean
  Dim idxDest As DAO.Index
  Dim fld As DAO.Field
  Dim idx As DAO.Index
  Dim i As Long
  
  On Error GoTo AddIndexes_Err
  Call xSet("AddIndexes")
  
  For Each idx In tdSrc.Indexes
    Set idxDest = tdDest.CreateIndex(idx.Name)
    For Each fld In idx.Fields
      Call idxDest.Fields.Append(idxDest.CreateField(fld.Name))
    Next
    For i = LBound(sCheckIdxProps) To UBound(sCheckIdxProps)
      If Not (idx.Properties(sCheckIdxProps(i)).Value = idxDest.Properties(sCheckIdxProps(i)).Value) Then
        idxDest.Properties(sCheckIdxProps(i)).Value = idx.Properties(sCheckIdxProps(i)).Value
      End If
    Next
    Call tdDest.Indexes.Append(idxDest)
  Next

AddIndexes_End:
  Call xReturn("AddIndexes")
  Exit Function

AddIndexes_Err:
  Call ErrorMessage(ERR_ERROR, Err, "AddIndexes", "Error in AddIndexes", "Undefined error.")
  Resume AddIndexes_End
End Function

Public Function syncData(mDataFilters As Collection, dbSrc As Database, dbDest As Database) As Boolean
  Dim flt As syncFilter
  Dim sql As String
  Dim i As Long, j As Long
  On Error GoTo syncData_Err
  Call xSet("syncData")
  
  Call ProvideFeedback(0, 0, "Synchronising data")
  For i = 0 To dbSrc.TableDefs.Count - 1
    For j = 1 To mDataFilters.Count
      If mDataFilters(j).DataFilter Then
        If mDataFilters(j).FilterApplies(dbSrc.TableDefs(i).Name) And ((dbSrc.TableDefs(i).Attributes And dbAttachedTable) = 0) Then
          Call dbDest.Execute("DELETE * FROM " & dbSrc.TableDefs(i).Name, dbFailOnError)
          sql = "INSERT INTO " & dbSrc.TableDefs(i).Name & " IN " & StrSQL(dbDest.Name) & " SELECT * FROM " & dbSrc.TableDefs(i).Name
          Call dbSrc.Execute(sql, dbFailOnError)
        End If
      End If
    Next
    Call ProvideFeedback(i + 1, dbSrc.TableDefs.Count, "Synchronising data")
  Next i
  

syncData_End:
  Call xReturn("syncData")
  Exit Function

syncData_Err:
  Call ErrorMessage(ERR_ERROR, Err, "syncData", "Error in syncData", "Undefined error.")
  Resume syncData_End
  Resume
End Function

