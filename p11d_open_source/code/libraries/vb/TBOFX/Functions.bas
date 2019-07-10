Attribute VB_Name = "Functions"
Option Explicit

Public Sub AddTextField(tdDest As TableDef, sName As String, iLen As Long, Optional bRequired As Boolean = False, Optional bAllowZeroLength As Boolean = True)
  Dim fld As Field

  On Error GoTo AddTextField_end
  
  If InCollection(tdDest.Fields, sName) Then Exit Sub
  
  Set fld = tdDest.CreateField(sName, dbText, iLen)
  fld.AllowZeroLength = bAllowZeroLength
  fld.Required = bRequired
  Call tdDest.Fields.Append(fld)
  Call tdDest.Fields.Refresh
   
AddTextField_end:
  Set fld = Nothing
  Exit Sub
  
AddTextField_err:
  Err.Raise Err.Number, ErrorSource(Err, "AddTextField"), "An error occurred applying a database fix."
  Resume
End Sub

Public Sub AddDoubleField(tdDest As TableDef, sName As String)
  Dim fld As Field

  On Error GoTo AddDoubleField_err
  
  If InCollection(tdDest.Fields, sName) Then Exit Sub
  
  Set fld = tdDest.CreateField(sName, dbDouble)
  fld.Required = False
  fld.DefaultValue = 0
  
  Call tdDest.Fields.Append(fld)
  tdDest.Fields.Refresh
   
AddDoubleField_end:
  
  Set fld = Nothing
  Exit Sub
  
AddDoubleField_err:
  Err.Raise Err.Number, ErrorSource(Err, "AddDoubleField"), "An error occurred applying a database fix."
End Sub

Public Sub AddBooleanField(tdDest As TableDef, sName As String)
  Dim fld As Field

  On Error GoTo AddBooleanField_err
  
  If InCollection(tdDest.Fields, sName) Then Exit Sub
  
  Set fld = tdDest.CreateField(sName, dbBoolean)
  fld.Required = False
  fld.DefaultValue = False
  
  Call tdDest.Fields.Append(fld)
  tdDest.Fields.Refresh
   
AddBooleanField_end:
  
  Set fld = Nothing
  Exit Sub
  
AddBooleanField_err:
  Err.Raise Err.Number, ErrorSource(Err, "AddBooleanField"), "An error occurred applying a database fix."
End Sub

Public Sub AddLongField(tdDest As TableDef, sName As String)
  Dim fld As Field

  On Error GoTo AddTextField_end
  
  If InCollection(tdDest.Fields, sName) Then Exit Sub
  
  Set fld = tdDest.CreateField(sName, dbText)
  fld.Required = False
  Call tdDest.Fields.Append(fld)
  Call tdDest.Fields.Refresh
   
AddTextField_end:
  Set fld = Nothing
  Exit Sub
  
AddTextField_err:
  Err.Raise Err.Number, ErrorSource(Err, "AddTextField"), "An error occurred applying a database fix."
  Resume
End Sub

Public Sub SyncScheduleSetupAndMenu(db As Database, dbTemplate As Database)
  Static bAlreadyRun As Boolean
  Dim sql As String
  
  If bAlreadyRun Then Exit Sub
  
  bAlreadyRun = False
  If InCollection(dbTemplate.TableDefs("Menu").Fields, "Expanded") Then
    If Not InCollection(db.TableDefs("Menu").Fields, "Expanded") Then
      Call AddBooleanField(db.TableDefs("Menu"), "Expanded")
    End If
  End If
  
  On Error GoTo SyncScheduleSetup_err
  Call ConnectAllFilterPrefix(db, dbTemplate, False, "ScheduleSetup", "SYNCDB")
  Call ConnectAllFilterPrefix(db, dbTemplate, False, "sys_ScheduleSetupStatics", "SYNCDB")
  Call ConnectAllFilterPrefix(db, dbTemplate, False, "Menu", "SYNCDB")
  Call ConnectAllFilterPrefix(db, dbTemplate, False, "sys_MenuStatic", "SYNCDB")
  
  'Updates default options field
  Call db.Execute("UPDATE SYNCDBScheduleSetup INNER JOIN ScheduleSetup ON (SYNCDBScheduleSetup.ColumnField = ScheduleSetup.ColumnField) AND (SYNCDBScheduleSetup.ObjectName = ScheduleSetup.ObjectName) SET ScheduleSetup.DefaultOptions = [SYNCDBScheduleSetup].[DefaultOptions];", dbFailOnError)
  'Adds in new lines not already there
  sql = "INSERT INTO ScheduleSetup ( ObjectName, ColumnField, DefaultOptions, Displayed ) "
  sql = sql & "SELECT SYNCDBsys_ScheduleSetupstatics.ObjectName, SYNCDBsys_ScheduleSetupstatics.ColumnField, 0 AS Expr1, False AS Expr2 "
  sql = sql & "FROM SYNCDBsys_ScheduleSetupstatics LEFT JOIN ScheduleSetup ON (SYNCDBsys_ScheduleSetupstatics.ColumnField = ScheduleSetup.ColumnField) AND (SYNCDBsys_ScheduleSetupstatics.ObjectName = ScheduleSetup.ObjectName) "
  sql = sql & "WHERE (((ScheduleSetup.ObjectName) Is Null));"
  Call db.Execute(sql, dbFailOnError)
  'Updates default options field in Menu
  sql = "UPDATE SYNCDBMenu INNER JOIN Menu ON SYNCDBMenu.ObjectName = Menu.ObjectName SET Menu.DefaultOptions = [SYNCDBMenu].[DefaultOptions], Menu.DataEntryTaxReview = [SYNCDBMenu].[DataEntryTaxReview];"
  Call db.Execute(sql, dbFailOnError)
  'Adds new lines into Menu
  sql = "INSERT INTO Menu "
  sql = sql & "SELECT SYNCDBMenu.* "
  sql = sql & "FROM SYNCDBMenu LEFT JOIN Menu ON SYNCDBMenu.ObjectName = Menu.ObjectName "
  sql = sql & "WHERE (Menu.ObjectName Is Null);"
  Call db.Execute(sql, dbFailOnError)
  
SyncScheduleSetup_end:
  bAlreadyRun = True
  Call RemoveLinkedTables(db)
  Exit Sub
  
SyncScheduleSetup_err:
  Call RemoveLinkedTables(db)
  Err.Raise Err.Number, ErrorSource(Err, "SyncScheduleSetupAndMenu"), "An error occurred updating the Schedule Setup and Menu tables." & Err.Description
  Resume
End Sub

Public Sub SyncTaxQuestions(db As Database, dbTemplate As Database)
  Dim sql As String
  Static bAlreadyRun As Boolean
  
  If bAlreadyRun Then Exit Sub
  
  On Error GoTo SyncTaxQuestions_Err
  
  bAlreadyRun = False
  Call ConnectAllFilterPrefix(db, dbTemplate, False, "TaxQuestions", "SYNCDB")
  sql = "UPDATE SYNCDBTaxQuestions INNER JOIN TaxQuestions ON SYNCDBTaxQuestions.QuestionCode = TaxQuestions.QuestionCode SET "
  sql = sql & "TaxQuestions.QuestionOrder = [SYNCDBTaxQuestions].[QuestionOrder], "
  sql = sql & "TaxQuestions.QuestionGroup = [SYNCDBTaxQuestions].[QuestionGroup], "
  sql = sql & "TaxQuestions.QuestionSch = [SYNCDBTaxQuestions].[QuestionSch], "
  sql = sql & "TaxQuestions.Help = [SYNCDBTaxQuestions].[Help], "
  sql = sql & "TaxQuestions.Question = [SYNCDBTaxQuestions].[Question], "
  sql = sql & "TaxQuestions.Persist = [SYNCDBTaxQuestions].[Persist], "
  sql = sql & "TaxQuestions.DivisionalType = [SYNCDBTaxQuestions].[DivisionalType], "
  sql = sql & "TaxQuestions.QuestionType = [SYNCDBTaxQuestions].[QuestionType];"
  Call db.Execute(sql, dbFailOnError)

SyncTaxQuestions_End:
  bAlreadyRun = True
  Call RemoveLinkedTables(db)
  Exit Sub

SyncTaxQuestions_Err:
  Call RemoveLinkedTables(db)
  Err.Raise Err.Number, ErrorSource(Err, "SyncTaxQuestions"), "An error occurred updating the TaxQuestions table." & Err.Description
  Resume
End Sub

Public Sub SyncTaxTypes(db As Database, dbTemplate As Database)
  Dim sql As String
  Dim td As TableDef
  Static bAlreadyRun As Boolean
    
  On Error GoTo syncTaxTypes_Err
  
  If bAlreadyRun Then Exit Sub
  
  bAlreadyRun = False
  Set td = db.TableDefs("TaxTypes")
  If Not FieldPresent(td.Fields, "RowSupport") Then
    Call AddTextField(td, "RowSupport", 50)
  End If
  If Not FieldPresent(td.Fields, "") Then
    Call AddTextField(td, "Packname", 255)
  End If
  Call ConnectAllFilterPrefix(db, dbTemplate, False, "TaxTypes", "SYNCDB")
  Call db.Execute("DELETE * FROM SYNCDBTaxTypes WHERE UserDefined=false", dbFailOnError)
  
  sql = "INSERT INTO TaxTypes ( [Table], Type, [Disallow], Dlink, [Dlink text], Source, UserDefined, TableDescrption, IncludeinPandL, RowSupport, Packname )"
  sql = sql & " SELECT SYNCDBTaxTypes.Table, SYNCDBTaxTypes.Type, SYNCDBTaxTypes.Disallow, SYNCDBTaxTypes.Dlink, SYNCDBTaxTypes.[Dlink text], SYNCDBTaxTypes.Source, SYNCDBTaxTypes.UserDefined, SYNCDBTaxTypes.TableDescrption, SYNCDBTaxTypes.IncludeinPandL, SYNCDBTaxTypes.RowSupport, SYNCDBTaxTypes.Packname"
  sql = sql & " FROM SYNCDBTaxTypes LEFT JOIN TaxTypes ON (SYNCDBTaxTypes.Table = TaxTypes.Table) AND (SYNCDBTaxTypes.Type = TaxTypes.Type)"
  sql = sql & " WHERE (((TaxTypes.Table) Is Null));"
  Call db.Execute(sql, dbFailOnError)

syncTaxTypes_End:
  bAlreadyRun = True
  Call RemoveLinkedTables(db)
  Exit Sub

syncTaxTypes_Err:
  Call RemoveLinkedTables(db)
  Err.Raise Err.Number, ErrorSource(Err, "SyncTaxTypes"), "An error occurred updating the TaxTypes table." & Err.Description
  Resume
End Sub

Public Sub SyncValidation(db As Database, dbTemplate As Database)
  Dim sql As String
  Static bAlreadyRun As Boolean
  
  On Error GoTo SyncValidation_err
  
  If bAlreadyRun Then Exit Sub
  
  bAlreadyRun = False
  Call ConnectAllFilterPrefix(db, dbTemplate, False, "Validation", "SYNCDB")
  Call ConnectAllFilterPrefix(db, dbTemplate, False, "sys_ValidationStatic", "SYNCDB")
  
  'Adds new lines into Validation
  If InCollection(db.TableDefs("Validation").Indexes, "PrimaryKey") Then
    Call db.TableDefs("Validation").Indexes.Delete("PrimaryKey")
  End If
  sql = "INSERT INTO Validation "
  sql = sql & "SELECT SYNCDBValidation.* "
  sql = sql & "FROM SYNCDBValidation LEFT JOIN Validation ON SYNCDBValidation.SupportObjectName = Validation.SupportObjectName "
  sql = sql & "WHERE (Validation.SupportObjectName Is Null);"
  Call db.Execute(sql, dbFailOnError)
  
SyncValidation_end:
  bAlreadyRun = True
  Call RemoveLinkedTables(db)
  Exit Sub
  
SyncValidation_err:
  Call RemoveLinkedTables(db)
  Err.Raise Err.Number, ErrorSource(Err, "SyncValidation"), "An error occurred updating the Validation table." & Err.Description
  Resume
End Sub

Public Sub SyncCache(db As Database, dbTemplate As Database)
  Dim sql As String
  
  On Error GoTo SyncCache_Err

  Call ConnectAllFilterPrefix(db, dbTemplate, False, "sys_Queries_Header", "SYNCDB")
  Call ConnectAllFilterPrefix(db, dbTemplate, False, "sys_Queries_Relations", "SYNCDB")

  Call CopyTable(gwsMain, db, "sys_Queries_Header", "SYNCDBsys_Queries_Header", True)
  Call CopyTable(gwsMain, db, "sys_Queries_Relations", "SYNCDBsys_Queries_Relations", True)
  
SyncCache_End:
  Call RemoveLinkedTables(db)
  Exit Sub
  
SyncCache_Err:
  Call RemoveLinkedTables(db)
  Err.Raise Err.Number, ErrorSource(Err, "SyncCache"), "An error occurred synchronising the cache." & Err.Description
  Resume
End Sub

Public Sub SyncExpCar(db As Database, dbTemplate As Database)
  Dim sql As String
  
  On Error GoTo SyncExpCar_Err
  Call ConnectAllFilterPrefix(db, dbTemplate, False, "TaxTypes", "SYNCDB")

  sql = "INSERT INTO TaxTypes "
  sql = sql & "SELECT SYNCDBTaxTypes.* "
  sql = sql & "FROM SYNCDBTaxTypes LEFT JOIN TaxTypes ON SYNCDBTaxTypes.Type = TaxTypes.Type"
  sql = sql & " WHERE (TaxTypes.Type Is Null);"
  Call db.Execute(sql, dbFailOnError)
  
  sql = "DELETE FROM TaxTypes WHERE Type='Car'"
  Call db.Execute(sql, dbFailOnError)
    
  sql = "UPDATE FAAdds SET TaxType = 'Expensive car' WHERE (TaxType='Car' AND Cost>=12000)"
  Call db.Execute(sql, dbFailOnError)
  
  sql = "UPDATE FAAdds SET TaxType = 'Cheap car' WHERE (TaxType='Car' AND Cost<12000)"
  Call db.Execute(sql, dbFailOnError)
  
  sql = "UPDATE FADisps SET TaxType = 'Expensive car' WHERE (TaxType='Car' AND Cost>=12000)"
  Call db.Execute(sql, dbFailOnError)
  
  sql = "UPDATE FADisps SET TaxType = 'Cheap car' WHERE (TaxType='Car' AND Cost<12000)"
  Call db.Execute(sql, dbFailOnError)
  
SyncExpCar_End:
  Call RemoveLinkedTables(db)
  Exit Sub
  
SyncExpCar_Err:
  Call RemoveLinkedTables(db)
  Err.Raise Err.Number, ErrorSource(Err, "SyncExpCar"), "An error occurred synchronising expensive and cheap car information." & Err.Description
  Resume
End Sub

Public Sub SyncProcessHelp(db As Database, dbTemplate As Database)
  Dim rs As Recordset
  
  On Error GoTo SyncProcessHelp_Err
  
  If InCollection(db.TableDefs, "ProcessHelp") Then
    Set rs = db.OpenRecordset("ProcessHelp", dbOpenDynaset, dbFailOnError)
    If rs.BOF And rs.EOF Then
      With rs
        .AddNew
          !ProcessPage = "INDEX"
          !ContextID = 1006
        .Update
        .AddNew
          !ProcessPage = "PAGE1"
          !ContextID = 1007
        .Update
        .AddNew
          !ProcessPage = "PAGE2"
          !ContextID = 1008
        .Update
        .AddNew
          !ProcessPage = "PAGE3"
          !ContextID = 1009
        .Update
      End With
    End If
  End If
  
SyncProcessHelp_End:
  Set rs = Nothing
  Exit Sub
  
SyncProcessHelp_Err:
  Err.Raise Err.Number, ErrorSource(Err, "SyncProcessHelp"), "An error occurred synchronising the process help table." & Err.Description
  Resume
End Sub

Public Sub SyncMultiSchedules(db As Database, dbTemplate As Database)
  
  On Error GoTo SyncMultiSchedules_Err

  Call SyncReserveTree(db)
  Call SyncCarTree(db)
  Call SyncIBATree(db)
  Call SyncOtherAnalysisTree(db)
  
SyncMultiSchedules_End:
  Exit Sub
  
SyncMultiSchedules_Err:
  Err.Raise Err.Number, ErrorSource(Err, "SyncMultiSchedules"), "An error occurred synchronising multi schedules." & Err.Description
  Resume
End Sub

Private Sub SyncReserveTree(db As Database)
  Dim sqlTree As String
  Dim sqlSched As String
  Dim sqlNonTree As String
  Dim rsTree As Recordset
  Dim rsSched As Recordset
 
  On Error GoTo SyncReserveTree_Err
  
  ' Set ChildSchedNo in ReserveTree to correct value from Schedules
  sqlTree = "SELECT * FROM ReserveTree WHERE ChildSchedNo is null"
  Set rsTree = db.OpenRecordset(sqlTree, dbOpenDynaset, dbFailOnError)
  If Not (rsTree.BOF And rsTree.EOF) Then
    ' Update ReserveTree
    sqlSched = "UPDATE ReserveTree INNER JOIN SCHEDULES ON (ReserveTree.Description = Schedules.Description) "
    sqlSched = sqlSched & "SET ReserveTree.ChildSchedNo = Schedules.InternalSchedNo "
    sqlSched = sqlSched & "WHERE Schedules.Packname='C5'"
    Call db.Execute(sqlSched, dbFailOnError)
    Set rsTree = Nothing
    sqlTree = "SELECT * FROM ReserveTree WHERE ChildSchedNo is null"
    Set rsTree = db.OpenRecordset(sqlTree, dbOpenDynaset, dbFailOnError)
    If Not (rsTree.BOF And rsTree.EOF) Then
      ' Still records in ReserveTree with ChildSchedNo null so count those in Schedules
      sqlSched = "SELECT * FROM Schedules WHERE Packname = 'C5'"
      Set rsSched = db.OpenRecordset(sqlSched, dbOpenDynaset, dbFailOnError)
      If rsSched.RecordCount = 1 Then
        ' Only one in Schedules, so set all to unique InternalSchedNo
        sqlNonTree = "UPDATE ReserveTree SET ReserveTree.ChildSchedNo = " & rsSched!InternalSchedNo
        Call db.Execute(sqlNonTree, dbFailOnError)
      Else
        While Not rsTree.EOF
          ' More than one in Schedules so try looking in Title of schedule
          sqlSched = "SELECT * FROM Schedules WHERE (InStr(1,[Title],'" & rsSched!Description & "',1)<>0) AND "
          sqlSched = sqlSched & " Packname = 'C5'"
          Set rsSched = db.OpenRecordset(sqlSched, dbOpenDynaset, dbFailOnError)
          If rsSched.RecordCount = 1 Then
            'One match, so update row in ReserveTree
            rsTree.Edit
              rsTree!ChildSchedNo = rsSched!InternalSchedNo
            rsTree.Update
          ElseIf rsSched.RecordCount = 0 Then
            'No match, so raise error
            Call MultiDialog("SyncReserveTree", "Unable to update ChildSchedNo in ReserveTree - no match in Schedules table", "OK")
          Else
            'More than one match so raise error
            Call MultiDialog("SyncReserveTree", "Unable to update ChildSchedNo in ReserveTree - multiple matches in Schedules table", "OK")
          End If
          rsTree.MoveNext
        Wend
      End If
    End If
    ' Fix Reserves table with new ChildSchedNo in ReserveTree
    sqlNonTree = "UPDATE ReserveTree INNER JOIN Reserves ON (ReserveTree.Description = Reserves.MultiDescription) AND "
    sqlNonTree = sqlNonTree & "(ReserveTree.InternalSchedNo = Reserves.AbacusSchedNo) "
    sqlNonTree = sqlNonTree & " SET Reserves.ChildSchedNo  = ReserveTree.ChildSchedNo "
    Call db.Execute(sqlNonTree, dbFailOnError)
  End If

SyncReserveTree_End:
  Set rsTree = Nothing
  Set rsSched = Nothing
  Exit Sub
  
SyncReserveTree_Err:
  Err.Raise Err.Number, ErrorSource(Err, "SyncReserveTree"), "An error occurred synchronising reserve schedules." & Err.Description
End Sub

Private Sub SyncCarTree(db As Database)
  Dim sqlTree As String
  Dim sqlSched As String
  Dim sqlNonTree As String
  Dim rsTree As Recordset
  Dim rsSched As Recordset
 
  On Error GoTo SyncCarTree_Err
  
  ' Set ChildSchedNo in CarTree to correct value from Schedules
  sqlTree = "SELECT * FROM CarTree WHERE ChildSchedNo is null"
  Set rsTree = db.OpenRecordset(sqlTree, dbOpenDynaset, dbFailOnError)
  If Not (rsTree.BOF And rsTree.EOF) Then
    ' Update CarTree
    sqlSched = "UPDATE CarTree INNER JOIN SCHEDULES ON (CarTree.Description = Schedules.Description) "
    sqlSched = sqlSched & "SET CarTree.ChildSchedNo = Schedules.InternalSchedNo "
    sqlSched = sqlSched & "WHERE Schedules.Packname='B15'"
    Call db.Execute(sqlSched, dbFailOnError)
    Set rsTree = Nothing
    sqlTree = "SELECT * FROM CarTree WHERE ChildSchedNo is null"
    Set rsTree = db.OpenRecordset(sqlTree, dbOpenDynaset, dbFailOnError)
    If Not (rsTree.BOF And rsTree.EOF) Then
      ' Still records in CarTree with ChildSchedNo null so count those in Schedules
      sqlSched = "SELECT * FROM Schedules WHERE Packname = 'B15'"
      Set rsSched = db.OpenRecordset(sqlSched, dbOpenDynaset, dbFailOnError)
      If rsSched.RecordCount = 1 Then
        ' Only one in Schedules, so set all to unique InternalSchedNo
        sqlNonTree = "UPDATE CarTree SET CarTree.ChildSchedNo = " & rsSched!InternalSchedNo
        Call db.Execute(sqlNonTree, dbFailOnError)
      Else
        While Not rsTree.EOF
          ' More than one in Schedules so try looking in Title of schedule
          sqlSched = "SELECT * FROM Schedules WHERE (InStr(1,[Title],'" & rsSched!Description & "',1)<>0) AND "
          sqlSched = sqlSched & " Packname = 'B15'"
          Set rsSched = db.OpenRecordset(sqlSched, dbOpenDynaset, dbFailOnError)
          If rsSched.RecordCount = 1 Then
            'One match, so update row in ReserveTree
            rsTree.Edit
              rsTree!ChildSchedNo = rsSched!InternalSchedNo
            rsTree.Update
          ElseIf rsSched.RecordCount = 0 Then
            'No match, so raise error
            Call MultiDialog("SyncCarTree", "Unable to update ChildSchedNo in CarTree - no match in Schedules table", "OK")
          Else
            'More than one match so raise error
            Call MultiDialog("SyncCarTree", "Unable to update ChildSchedNo in CarTree - multiple matches in Schedules table", "OK")
          End If
          rsTree.MoveNext
        Wend
      End If
    End If
    ' Fix FAAdds and FADisps tables with new ChildSchedNo in CarTree
    sqlNonTree = "UPDATE CarTree INNER JOIN FAAdds ON (ReserveTree.Description = FAAdds.MultiDescription) AND "
    sqlNonTree = sqlNonTree & "(CarTree.InternalSchedNo = FAAdds.AbacusSchedNo) "
    sqlNonTree = sqlNonTree & " SET FAAdds.ChildSchedNo  = CarTree.ChildSchedNo"
    Call db.Execute(sqlNonTree, dbFailOnError)
    sqlNonTree = "UPDATE CarTree INNER JOIN FADisps ON (ReserveTree.Description = FADisps.MultiDescription) AND "
    sqlNonTree = sqlNonTree & "(CarTree.InternalSchedNo = FADisps.AbacusSchedNo) "
    sqlNonTree = sqlNonTree & " SET FADisps.ChildSchedNo  = CarTree.ChildSchedNo"
    Call db.Execute(sqlNonTree, dbFailOnError)
  End If

SyncCarTree_End:
  Set rsTree = Nothing
  Set rsSched = Nothing
  Exit Sub
  
SyncCarTree_Err:
  Err.Raise Err.Number, ErrorSource(Err, "SyncCarTree"), "An error occurred synchronising car schedules." & Err.Description
End Sub

Private Sub SyncIBATree(db As Database)
  Dim sqlTree As String
  Dim sqlSched As String
  Dim sqlNonTree As String
  Dim rsTree As Recordset
  Dim rsSched As Recordset
 
  On Error GoTo SyncIBATree_Err
  
  ' Set ChildSchedNo in IBATree to correct value from Schedules
  sqlTree = "SELECT * FROM IBATree WHERE ChildSchedNo is null"
  Set rsTree = db.OpenRecordset(sqlTree, dbOpenDynaset, dbFailOnError)
  If Not (rsTree.BOF And rsTree.EOF) Then
    ' Update IBATree
    sqlSched = "UPDATE IBATree INNER JOIN SCHEDULES ON (IBATree.Description = Schedules.Description) "
    sqlSched = sqlSched & "SET IBATree.ChildSchedNo = Schedules.InternalSchedNo "
    sqlSched = sqlSched & "WHERE Schedules.Packname='B21'"
    Call db.Execute(sqlSched, dbFailOnError)
    Set rsTree = Nothing
    sqlTree = "SELECT * FROM IBATree WHERE ChildSchedNo is null"
    Set rsTree = db.OpenRecordset(sqlTree, dbOpenDynaset, dbFailOnError)
    If Not (rsTree.BOF And rsTree.EOF) Then
      ' Still records in IBATree with ChildSchedNo null so count those in Schedules
      sqlSched = "SELECT * FROM Schedules WHERE Packname = 'B21'"
      Set rsSched = db.OpenRecordset(sqlSched, dbOpenDynaset, dbFailOnError)
      If rsSched.RecordCount = 1 Then
        ' Only one in Schedules, so set all to unique InternalSchedNo
        sqlNonTree = "UPDATE IBATree SET IBATree.ChildSchedNo = " & rsSched!InternalSchedNo
        Call db.Execute(sqlNonTree, dbFailOnError)
      Else
        While Not rsTree.EOF
          ' More than one in Schedules so try looking in Title of schedule
          sqlSched = "SELECT * FROM Schedules WHERE (InStr(1,[Title],'" & rsSched!Description & "',1)<>0) AND "
          sqlSched = sqlSched & " Packname = 'B21'"
          Set rsSched = db.OpenRecordset(sqlSched, dbOpenDynaset, dbFailOnError)
          If rsSched.RecordCount = 1 Then
            'One match, so update row in IBATree
            rsTree.Edit
              rsTree!ChildSchedNo = rsSched!InternalSchedNo
            rsTree.Update
          ElseIf rsSched.RecordCount = 0 Then
            'No match, so raise error
            Call MultiDialog("SyncIBATree", "Unable to update ChildSchedNo in IBATree - no match in Schedules table", "OK")
          Else
            'More than one match so raise error
            Call MultiDialog("SyncIBATree", "Unable to update ChildSchedNo in IBATree - multiple matches in Schedules table", "OK")
          End If
          rsTree.MoveNext
        Wend
      End If
    End If
    ' Fix FAAdds and FADisps tables with new ChildSchedNo in IBATree
    sqlNonTree = "UPDATE IBATree INNER JOIN FAAdds ON (IBATree.Description = FAAdds.MultiDescription) AND "
    sqlNonTree = sqlNonTree & "(IBATree.InternalSchedNo = FAAdds.AbacusSchedNo) "
    sqlNonTree = sqlNonTree & " SET FAAdds.ChildSchedNo  = IBATree.ChildSchedNo"
    Call db.Execute(sqlNonTree, dbFailOnError)
    sqlNonTree = "UPDATE IBATree INNER JOIN FADisps ON (IBATree.Description = FADisps.MultiDescription) AND "
    sqlNonTree = sqlNonTree & "(IBATree.InternalSchedNo = FADisps.AbacusSchedNo) "
    sqlNonTree = sqlNonTree & " SET FADisps.ChildSchedNo  = IBATree.ChildSchedNo"
    Call db.Execute(sqlNonTree, dbFailOnError)
  End If

SyncIBATree_End:
  Set rsTree = Nothing
  Set rsSched = Nothing
  Exit Sub
  
SyncIBATree_Err:
  Err.Raise Err.Number, ErrorSource(Err, "SyncIBATree"), "An error occurred synchronising IBA schedules." & Err.Description
End Sub

Private Sub SyncOtherAnalysisTree(db As Database)
  Dim sqlTree As String
  Dim sqlSched As String
  Dim sqlNonTree As String
  Dim rsTree As Recordset
 
  On Error GoTo SyncOtherAnalysisTree_Err
  
  ' Set ChildSchedNo in OtherAnalysisTree to correct value from Schedules
  sqlTree = "SELECT * FROM OtherAnalysisTree WHERE ChildSchedNo is null"
  Set rsTree = db.OpenRecordset(sqlTree, dbOpenDynaset, dbFailOnError)
  If Not (rsTree.BOF And rsTree.EOF) Then
    ' Update OtherAnalysisTree
    sqlSched = "UPDATE OtherAnalysisTree SET ChildSchedNo = 0"
    Call db.Execute(sqlSched, dbFailOnError)
    ' Fix OtherAnalysis table with new ChildSchedNo in OtherAnalysisTree
    sqlNonTree = "UPDATE OtherAnalysis SET ChildSchedNo = 0"
    Call db.Execute(sqlNonTree, dbFailOnError)
  End If

SyncOtherAnalysisTree_End:
  Set rsTree = Nothing
  Exit Sub
  
SyncOtherAnalysisTree_Err:
  Err.Raise Err.Number, ErrorSource(Err, "SyncOtherAnalysisTree"), "An error occurred synchronising other analysis schedules." & Err.Description
End Sub

Public Sub CopyQuery(sQuery As String, db As Database, dbTemplate As Database)
  Dim qd As QueryDef
  Dim sSysQuery As String
  
  On Error GoTo CopyQuery_Err

  If InCollection(db.QueryDefs, sQuery) Then
    Call db.QueryDefs.Delete(sQuery)
  End If
  sSysQuery = "sys_" & sQuery
  If Not InCollection(db.QueryDefs, sSysQuery) Then
    Set qd = New QueryDef
    qd.sql = dbTemplate.QueryDefs(sSysQuery).sql
    qd.Name = sSysQuery
    Call db.QueryDefs.Append(qd)
  End If
  Call DoDBEvents(FREE_LOCKS + REFRESH_CACHE)
  
CopyQuery_End:
  Set qd = Nothing
  Exit Sub
  
CopyQuery_Err:
  Err.Raise Err.Number, ErrorSource(Err, "CopyQuery"), "An error occurred copying queries from the template." & Err.Description
  Resume
End Sub

