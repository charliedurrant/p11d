Attribute VB_Name = "Check"
Option Explicit

Public Sub CheckTables(ByVal dbSrc As Database, ByVal dbDest As Database, sCheckFldProps() As String, sCheckIdxProps() As String, sLog As QString, dbChange As dbChangeDetails)
  Dim tblChange As TblChangeDetails
  Dim td As TableDef
  Dim tdDest As TableDef
  Dim i As Long, j As Long, k As Long, sReason As String
  
  On Error GoTo CheckTables_Err
  
  Call SetCursor
  
  Call sLog.Append("<CHECKTABLES><TIME> " & Now & "</TIME><FAILURES>")
  i = 1
  For Each td In dbSrc.TableDefs
    Call ProvideFeedback(i, dbSrc.TableDefs.Count, "Comparing table structure - " & td.Name)
    If Not IsSysTable(td) Then
      If Not InCollection(dbDest.TableDefs, td.Name) Then
        j = j + 1
        Call sLog.Append("<FAIL Type=""DBTable""><OBJNAME>" & td.Name & "</OBJNAME><REASON><MAINREASON>Table not found in destination</MAINREASON></REASON></FAIL>")
        Call dbChange.NewTables.Add(td.Name)
        GoTo Next_Table:
      End If
      Set tdDest = dbDest.TableDefs(td.Name)
      If tblChange Is Nothing Then Set tblChange = New TblChangeDetails
      tblChange.Setup (td.Name)
      If Not IsTableSame(td, tdDest, sLog, sCheckFldProps, sCheckIdxProps, tblChange) Then
        j = j + 1
        Call sLog.Append("<FAIL Type=""DBTable""><OBJNAME>" & td.Name & "</OBJNAME>" & sReason & "</FAIL>")
        Call dbChange.ChangedTables.Add(tblChange, tblChange.TableName)
        Set tblChange = Nothing
      End If
Next_Table:
      sReason = ""
    End If
    i = i + 1
  Next
  k = 1
  For Each td In dbDest.TableDefs
    Call ProvideFeedback(k, dbDest.TableDefs.Count, "Checking for deleted tables")
    If Not IsSysTable(td) Then
      If Not InCollection(dbSrc.TableDefs, td.Name) Then
        Call sLog.Append("<FAIL Type=""DBTable""><OBJNAME>" & td.Name & "</OBJNAME><REASON><MAINREASON>Table not found in source</MAINREASON></REASON></FAIL>")
        Call dbChange.OldTables.Add(td.Name)
      End If
    End If
    k = k + 1
  Next
  Call sLog.Append("</FAILURES><TOTALCOUNT>" & i - 1 & "</TOTALCOUNT><FAILCOUNT>" & j & "</FAILCOUNT><TIME>" & Now & "</TIME></CHECKTABLES>")
  
  
CheckTables_End:
  Set tblChange = Nothing
  Call ClearCursor
  Exit Sub

CheckTables_Err:
  Call ErrorMessage(ERR_ERROR, Err, "CheckTables", "Error in CheckTables", "Error checking the table designs to identify differences.")
  Resume CheckTables_End
  Resume
End Sub

Private Function IsTableSame(ByVal tdSrc As TableDef, ByVal tdDest As TableDef, sLog As QString, sCheckFldProps() As String, sCheckIdxProps() As String, tblChange As TblChangeDetails) As Boolean
  Dim fldDest As DAO.Field
  Dim fldSrc As DAO.Field
  Dim idxSrc As Index
  Dim idxDest As Index
  Dim propSrc As DAO.Property, propDest As DAO.Property
  Dim fldChange As fldChangeDetails, bFieldDiffers As Boolean
  Dim i As Long
  
  On Error GoTo IsTableSame_Err
  
  IsTableSame = True
  If tdSrc.LastUpdated = tdDest.LastUpdated Then GoTo IsTableSame_End
  If Not (tdSrc.Indexes.Count = tdDest.Indexes.Count) Then
    Call LogReason(sLog, "Index", "Different Count", tdSrc.Indexes.Count, tdDest.Indexes.Count)
    IsTableSame = False
    GoTo Check_Fields
  End If

'Indexes - if they differ set TableSame to false, they will always be recreated if TableSame is false
  For Each idxSrc In tdSrc.Indexes
    If Not InCollection(tdDest.Indexes, idxSrc.Name) Then
      Call LogReason(sLog, "Index", "Not in destination", idxSrc.Name)
      IsTableSame = False
      GoTo Check_Fields
    End If
    Set idxDest = tdDest.Indexes(idxSrc.Name)
    For i = LBound(sCheckIdxProps) To UBound(sCheckIdxProps)
      If Not (idxSrc.Properties(sCheckIdxProps(i)) = idxDest.Properties(sCheckIdxProps(i))) Then
        Call LogReason(sLog, "Index", sCheckIdxProps(i), idxSrc.Properties(sCheckIdxProps(i)).Value, idxSrc.Properties(sCheckIdxProps(i)).Value)
        IsTableSame = False
        GoTo Check_Fields
      End If
    Next i
    If Not (idxSrc.Fields.Count = idxDest.Fields.Count) Then
      Call LogReason(sLog, "Index", "Field count different", idxSrc.Fields.Count, idxDest.Fields.Count)
      IsTableSame = False
      GoTo Check_Fields
    End If
    For Each fldSrc In idxSrc.Fields
      If Not InCollection(idxDest.Fields, fldSrc.Name) Then
        Call LogReason(sLog, "Index", "Field not in destination", fldSrc.Name)
        IsTableSame = False
        GoTo Check_Fields
      End If
    Next
  Next
  
Check_Fields:
  For Each fldSrc In tdSrc.Fields
    bFieldDiffers = False
    If fldChange Is Nothing Then Set fldChange = New fldChangeDetails
    fldChange.Setup (fldSrc.Name)
    If Not InCollection(tdDest.Fields, fldSrc.Name) Then
      Call LogReason(sLog, "Field - " & fldSrc.Name, "Not in destination", fldSrc.Name)
      Call tblChange.AddNewField(fldSrc.Name)
      IsTableSame = False
      GoTo next_field
    End If
    Set fldDest = tdDest.Fields(fldSrc.Name)
    For i = LBound(sCheckFldProps) To UBound(sCheckFldProps)
      If Not (fldSrc.Properties(sCheckFldProps(i)) = fldDest.Properties(sCheckFldProps(i))) Then
        Call LogReason(sLog, "Field - " & fldSrc.Name, sCheckFldProps(i), fldSrc.Properties(sCheckFldProps(i)).Value, fldDest.Properties(sCheckFldProps(i)).Value)
        Call fldChange.AddChangedProperty(sCheckFldProps(i))
        IsTableSame = False
        bFieldDiffers = True
      End If
    Next i
next_field:
    If bFieldDiffers Then
      Call tblChange.ChangedFields.Add(fldChange, fldChange.Name)
      Set fldChange = Nothing
    End If
  Next
  For Each fldDest In tdDest.Fields
    If Not InCollection(tdSrc.Fields, fldDest.Name) Then
      Call LogReason(sLog, "Field - " & fldDest.Name, "Not in source", fldDest.Name)
      Call tblChange.AddOldField(fldDest.Name)
      IsTableSame = False
    End If
  Next

IsTableSame_End:
  Exit Function

IsTableSame_Err:
  Call ErrorMessage(ERR_ERROR, Err, "IsTableSame", "Error in IsTableSame", "Error checking if a table is the same.")
  Resume IsTableSame_End
  Resume
End Function

Private Function IsQuerySame(ByVal qdSrc As QueryDef, ByVal qdDest As QueryDef, sLog As QString) As Boolean

  On Error GoTo IsQuerySame_Err
  
  If qdSrc.LastUpdated = qdDest.LastUpdated Then
    IsQuerySame = True
    GoTo IsQuerySame_End
  End If
  If StrComp(qdSrc.SQL, qdDest.SQL, vbTextCompare) = 0 Then
    IsQuerySame = True
    GoTo IsQuerySame_End
  End If
  Call LogReason(sLog, "Query different", , qdSrc.SQL, qdDest.SQL)

IsQuerySame_End:
  Exit Function

IsQuerySame_Err:
  Call ErrorMessage(ERR_ERROR, Err, "IsQuerySame", "Error in IsQuerySame", "Undefined error.")
  Resume IsQuerySame_End
  Resume
End Function

Public Function CheckAllQueries(mdbSrc As Database, mdbDest As Database, dbChange As dbChangeDetails, sLog As QString) As Boolean
  Dim qdSrc As QueryDef, qdDest As QueryDef
  Dim i As Long, j As Long, k As Long
  Dim sReason As String
  
  On Error GoTo CheckAllQueries_Err
  
  sLog.Append ("<CHECKQUERIES><STARTTIME> " & Now & "</STARTTIME><FAILURES>")
  Call SetCursor
  i = 0
  For Each qdSrc In mdbSrc.QueryDefs
    Call ProvideFeedback(i, mdbSrc.QueryDefs.Count, "Comparing query structures")
    If Not InCollection(mdbDest.QueryDefs, qdSrc.Name) Then
      j = j + 1
      Call sLog.Append("<FAIL Type=""DbQuery""><OBJNAME>" & qdSrc.Name & "</OBJNAME><REASON><MAINREASON>QUERY not found in destination</MAINREASON></REASON></FAIL>")
      Call dbChange.ChangedQueries.Add(qdSrc.Name)
      GoTo next_query
    End If
    Set qdDest = mdbDest.QueryDefs(qdSrc.Name)
    If Not IsQuerySame(qdSrc, qdDest, sLog) Then
      j = j + 1
      Call sLog.Append("<FAIL Type=""DBQuery""><OBJNAME>" & qdSrc.Name & "</OBJNAME>" & sReason & "</FAIL>")
      Call dbChange.ChangedQueries.Add(qdSrc.Name)
    End If
next_query:
    sReason = ""
    i = i + 1
  Next
  k = 1
  For Each qdDest In mdbDest.QueryDefs
    Call ProvideFeedback(k, mdbDest.QueryDefs.Count, "Checking for deleted queries")
    If Not InCollection(mdbSrc.QueryDefs, qdDest.Name) Then
      Call sLog.Append("<FAIL Type=""DbQuery""><OBJNAME>" & qdDest.Name & "</OBJNAME><REASON><MAINREASON>QUERY not found in source</MAINREASON></REASON></FAIL>")
      Call dbChange.OldQueries.Add(qdDest.Name)
    End If
    k = k + 1
  Next
  Call sLog.Append("</FAILURES><TOTALCOUNT>" & i - 1 & "</TOTALCOUNT><FAILCOUNT>" & j & "</FAILCOUNT><ENDTIME>" & Now & "</ENDTIME></CHECKQUERIES>")

CheckAllQueries_End:
  ClearCursor
  Exit Function

CheckAllQueries_Err:
  Call ErrorMessage(ERR_ERROR, Err, "CheckAllQueries", "Error in CheckAllQueries", "Undefined error.")
  Resume CheckAllQueries_End
  Resume
End Function

Private Sub LogReason(sLog As QString, ByVal sMainReason As String, Optional ByVal sSubReason As String = "", Optional ByVal sSrcVal As Variant = "", Optional ByVal sDestVal As Variant = "")
  Dim s As String

  Call sLog.Append("<REASON>" & vbCrLf)
  Call sLog.Append("  <MAINREASON>" & sMainReason & "</MAINREASON>" & vbCrLf)
  Call sLog.Append("  <SUBREASON>" & sSubReason & "</SUBREASON>" & vbCrLf)
  Call sLog.Append("  <SRCVAL><![CDATA[" & sSrcVal & "]]></SRCVAL>" & vbCrLf)
  Call sLog.Append("  <DESTVAL><![CDATA[" & sDestVal & "]]></DESTVAL>" & vbCrLf)
  Call sLog.Append("</REASON>" & vbCrLf)
End Sub

