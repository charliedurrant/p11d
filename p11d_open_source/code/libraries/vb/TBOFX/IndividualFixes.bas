Attribute VB_Name = "IndividualFixes"
Option Explicit

Public Sub ApplyIndividualFixes(db As Database, dbTemplate As Database, mlFixLevel As Long, mlPostFixLevel As Long, mlDbVersion As Long, mlDBSubVersion As Long)
  
  On Error GoTo ApplyIndividualFixes_err
  
  '*************************************************************************************
  '         VERY IMPORTANT
  '         Make sure that any fixes are already in the template db and the FixLevel is
  '         incremented accordingly
  '*************************************************************************************
  
  Call ProvideFeedback(-1, -1, "Applying Fixes")
  If mlFixLevel < 1 Then
    If Fix1(db, dbTemplate) Then mlFixLevel = setFixLevel(db, 2)
  End If
  If mlFixLevel < 3 Then
    If Fix2(db, dbTemplate) Then mlFixLevel = setFixLevel(db, 4)
  End If
  If mlFixLevel < 5 Then
    If Fix2(db, dbTemplate) Then mlFixLevel = setFixLevel(db, 5)
  End If
  If mlFixLevel < 6 Then
    If Fix5(db, dbTemplate) Then mlFixLevel = setFixLevel(db, 6)
  End If
  If mlFixLevel < 7 Then
    If Fix6(db, dbTemplate) Then mlFixLevel = setFixLevel(db, 7)
  End If
  If mlFixLevel < 8 Then
    If Fix7(db, dbTemplate) Then mlFixLevel = setFixLevel(db, 8)
  End If
  If mlFixLevel < 9 Then
    If Fix8(db, dbTemplate) Then mlFixLevel = setFixLevel(db, 9)
  End If
  If mlFixLevel < 10 Then
    If Fix9(db, dbTemplate) Then mlFixLevel = setFixLevel(db, 10)
  End If
  If mlFixLevel < 11 Then
    If Fix10(db, dbTemplate) Then mlFixLevel = setFixLevel(db, 11)
  End If
  If mlFixLevel < 12 Then
    If Fix11(db, dbTemplate) Then mlFixLevel = setFixLevel(db, 12)
  End If
  If mlFixLevel < 13 Then
    If Fix12(db, dbTemplate) Then mlFixLevel = setFixLevel(db, 13)
  End If
  If mlFixLevel < 14 Then
    If Fix13(db, dbTemplate) Then mlFixLevel = setFixLevel(db, 14)
  End If
  If mlFixLevel < 15 Then
    If Fix14(db, dbTemplate) Then mlFixLevel = setFixLevel(db, 15)
  End If
  If mlFixLevel < 16 Then
    If Fix15(db, dbTemplate) Then mlFixLevel = setFixLevel(db, 16)
  End If
  If mlFixLevel < 17 Then
    If Fix16(db, dbTemplate) Then mlFixLevel = setFixLevel(db, 17)
  End If

  'Next fix
   
ApplyIndividualFixes_end:
  Exit Sub
  
ApplyIndividualFixes_err:
  Err.Raise Err.Number, ErrorSource(Err, "ApplyIndividualFixes"), "An error occurred applying a database fix." & vbCrLf & vbCrLf & Err.Description
  Resume
End Sub

'Public Function FixExample(db As Database) As Boolean
'  On Error GoTo Fixexample_Err:
'
'  Fix code goes here
'
'  End of fix code
'
'  FixExample = True
'
'FixExample_end:
'  Exit Function
'
'FixExample_Err:
'  Err.Raise Err.Number, ErrorSource(Err, "FixExample"), "An error occurred running FixExample" & vbcrlf & vbcrlf & err.Description
'End Function

Public Function Fix1(db As Database, dbTemplate As Database) As Boolean
  On Error GoTo Fix1_Err:
  
  Call DisplayMessagePopup(Nothing, "Backing up Questions", "Updating TBO file")
  Call CopyTable(gwsMain, db, "TaxQuestions_Backup", "TaxQuestions", False)
  Call CopyTable(gwsMain, db, "TaxQuestionsDivisions_Backup", "TaxQuestionsDivisions", False)
  
  Call Fix1_DivisionalType(db)
  Call Fix1_TaxQuestionsAnswer(db, "TaxQuestions")
  Call Fix1_TaxQuestionsAnswer(db, "TaxQuestionsDivisions")
  Call Fix1_TaxQuestions(db, dbTemplate)
  Call Fix1_Schedules(db)
  Call Fix1_BFWD_Fields(db)
  Call Fix1_USAnalysis_Fields(db)
  Call Fix1_ScheduleSetup(db)
  Call Fix1_sysImports(db)

  Fix1 = True

Fix1_end:
  Exit Function

Fix1_Err:
  Err.Raise Err.Number, ErrorSource(Err, "Fix1"), "An error occurred running Fix 1" & vbCrLf & vbCrLf & Err.Description
  Resume
End Function
    
Private Sub Fix1_DivisionalType(db As Database)
  Dim td As TableDef
  Dim fld As Field
  
  On Error GoTo Fix1_DivisionalType_err
  
  Set td = db.TableDefs("TaxQuestions")
  
  Set fld = td.Fields("DivisionalType")
  If fld.Type = vbString Then
    Call AddLongField(td, "DTypeNum")
    Call db.Execute("UPDATE TaxQuestions SET DTypeNum=" & NumSQL(1) & " WHERE DivisionalType=" & StrSQL("MASTER"))
    Call db.Execute("UPDATE TaxQuestions SET DTypeNum=" & NumSQL(2) & " WHERE DivisionalType=" & StrSQL("DIVISIONAL"))
    Call db.Execute("UPDATE TaxQuestions SET DTypeNum=" & NumSQL(3) & " WHERE DivisionalType=" & StrSQL("BOTH"))
    If Not KillField(td, "DivisionalType") Then Err.Raise ERR_NO_KILL_FIELD, "Fix1_DivisionalType", "Error removing the field DivisionalType"
    Call AddLongField(td, "DivisionalType")
    Call db.Execute("UPDATE TaxQuestions SET DivisionalType=[DTypeNum]")
    Call KillField(td, "DTypeNum")
  End If
  
Fix1_DivisionalType_end:
  Set fld = Nothing
  Set td = Nothing
  Exit Sub

Fix1_DivisionalType_err:
  Err.Raise Err.Number, ErrorSource(Err, "Fix1_DivisionalType"), "An error applying a database fix."
  Resume
End Sub

Private Sub Fix1_TaxQuestionsAnswer(db As Database, sTableName As String)
  Dim td As TableDef
  
  On Error GoTo Fix1_TaxQuestionsAnswer_err:
  
  Set td = db.TableDefs(sTableName)
  Call AddTextField(td, "Answertext", 255)
  Call db.Execute("UPDATE " & sTableName & " SET AnswerText=Left$([Answer],255)")
  If Not KillField(td, "Answer") Then Err.Raise ERR_NO_KILL_FIELD, "Fix1_TaxQuestionaAnswer", "Error removing the field Answer"
  Call AddTextField(td, "Answer", 255)
  Call db.Execute("UPDATE " & sTableName & " SET Answer=[Answertext]")
  Call KillField(td, "Answertext")
  
Fix1_TaxQuestionsAnswer_end:
  Exit Sub
  
Fix1_TaxQuestionsAnswer_err:
  Err.Raise Err.Number, ErrorSource(Err, "Fix1_TaxQuestionsAnswer"), "An error applying a database fix."
  Resume
End Sub
     
Private Sub Fix1_TaxQuestions(db As Database, dbTemplate As Database)
  Dim td As TableDef
  Dim rs As Recordset
  Dim s As String
  
  On Error GoTo Fix1_TaxQuestions_err
  
  ' update taxquestions
  Call DisplayMessagePopup(Nothing, "Updating Questions", "Updating TBO file")
  s = "Select * from TaxQuestions"
  Set rs = db.OpenRecordset(s)
  Call Fix1_TaxQuestions_sub(db, dbTemplate, rs, dbTemplate.Name)
  Call DisplayMessagePopup(Nothing, "Updating Divisional Questions", "Updating TBO file")
  s = "Select * from TaxQuestionsDivisions"
  Set rs = db.OpenRecordset(s)
  Call Fix1_TaxQuestions_sub(db, dbTemplate, rs, dbTemplate.Name)
    
Fix1_TaxQuestions_end:
  Set rs = Nothing
  Exit Sub
  
Fix1_TaxQuestions_err:
  Err.Raise Err.Number, ErrorSource(Err, "Fix1_TaxQuestions"), "An error occurred applying a database fix." & vbCrLf & vbCrLf & Err.Description
  Resume
End Sub

Private Sub Fix1_TaxQuestions_sub(db As Database, dbTemplate As Database, rs As Recordset, sTemplateFile As String)
  Dim s1 As String
  Dim s2 As String
  
  On Error GoTo Fix1_TaxQuestions_sub_err
  
  While Not rs.EOF
    ' update question codes
    rs.Edit
    Select Case rs.Fields("QuestionCode")
    Case "NBVBF"
      rs.Fields("QuestionCode") = "BFNBV"
      rs.Update
    Case "NBVCF"
      rs.Fields("QuestionCode") = "CFNBV"
      rs.Update
    Case "REVBS"
      rs.Fields("QuestionCode") = "DATAREVP"
      rs.Update
    Case "DEPBS"
      rs.Fields("QuestionCode") = "DATADEPC"
      rs.Update
    Case "NBVDISP"
      rs.Fields("QuestionCode") = "DATAFADISP"
      rs.Update
    Case "EXTI"
      rs.Fields("QuestionCode") = "EXTRAORD"
      rs.Update
    Case "PENBF"
      rs.Fields("QuestionCode") = "BFPEN"
      rs.Update
    Case "PENCF"
      rs.Fields("QuestionCode") = "CFPEN"
      rs.Update
    Case "LESBF"
      rs.Fields("QuestionCode") = "BALBF"
      rs.Update
    Case "LESADDS"
      rs.Fields("QuestionCode") = "DATADRADDS"
      rs.Update
    Case "LESPL"
      rs.Fields("QuestionCode") = "FLDCPL"
      rs.Update
    Case "LESPRO"
      rs.Fields("QuestionCode") = "FLDP"
      rs.Update
    Case "LESCF"
      rs.Fields("QuestionCode") = "BALCF"
      rs.Update
    Case "FINLEADEPPOST91"
      rs.Fields("QuestionCode") = "NLDEP"
      rs.Update
    Case "FINLEABFPOST91"
      rs.Fields("QuestionCode") = "BFLEASEPOST"
      rs.Update
    Case "FINLEAADDSPOST91"
      rs.Fields("QuestionCode") = "NEWLAA"
      rs.Update
    Case "FINLEAPLPOST91"
      rs.Fields("QuestionCode") = "NLIP"
      rs.Update
    Case "FINLEAPAYPOST91"
      rs.Fields("QuestionCode") = "NLAP"
      rs.Update
    Case "FINLEACFPOST91"
      rs.Fields("QuestionCode") = "CFLEASEPOST"
      rs.Update
    Case "FINLEABFPRE91"
      rs.Fields("QuestionCode") = "BFLEASE"
      rs.Update
    Case "FINLEAPLPRE91"
      rs.Fields("QuestionCode") = "OLIP"
      rs.Update
    Case "FINLEACFPRE91"
      rs.Fields("QuestionCode") = "CFLEASE"
    Case "CURRYTC"
      rs.Fields("QuestionCode") = "CURTCT"
      rs.Update
    Case "PRIORYTC"
      rs.Fields("QuestionCode") = "PYPLCT"
      rs.Update
    Case "PAID"
      rs.Fields("QuestionCode") = "CTPAYPY"
      rs.Update
    Case Else
      rs.CancelUpdate
    End Select
    rs.MoveNext
  Wend
  If InCollection(rs.Fields, "QuestionGroup") Then
    ' replace existing foreign branch questions and abacus options questions
    ' there is the possibility of losing foreign tax answers, but this is minor
    s1 = "DELETE * FROM TAXQUESTIONS"
    s2 = " WHERE QuestionGroup=" & StrSQL("Foreign activities") & " or QuestionSch=" & StrSQL("qAbacusOptions")
    db.Execute (s1 & s2)
    s1 = "INSERT INTO TAXQUESTIONS (QuestionOrder, QuestionGroup, QuestionSch, Help, QuestionCode, Question, Answer, Persist, Source, DivisionalType, QuestionType) SELECT QuestionOrder, QuestionGroup, QuestionSch, Help, QuestionCode, Question, Answer, Persist, Source, DivisionalType, QuestionType from TAXQUESTIONS in " & StrSQL(sTemplateFile)
    Call db.Execute(s1 & s2)
    Call CopyQuery("AddDivisionalQuestionsSub", db, dbTemplate)
    Call CopyQuery("AddDivisionalQuestions", db, dbTemplate)
    Call db.Execute("sys_AddDivisionalQuestions")
  End If
  
Fix1_TaxQuestions_sub_end:
  Exit Sub
  
Fix1_TaxQuestions_sub_err:
  Err.Raise Err.Number, ErrorSource(Err, "Fix1_TaxQuestions_sub"), "An error occurred applying a database fix." & vbCrLf & vbCrLf & Err.Description
  Resume
End Sub

Private Sub Fix1_Schedules(db As Database)
  Dim td As TableDef
  
  On Error GoTo Fix1_Schedules_err
  ' add new field Description
  Call DisplayMessagePopup(Nothing, "Updating Schedules", "Updating TBO file")
  Set td = db.TableDefs("Schedules")
  Call AddTextField(td, "Description", 50)
  
Fix1_Schedules_end:
  Exit Sub
  
Fix1_Schedules_err:
  Err.Raise Err.Number, ErrorSource(Err, "Fix1_Schedules"), "An error occurred applying a database fix." & vbCrLf & vbCrLf & Err.Description
End Sub

Private Sub Fix1_BFWD_Fields(db As Database)
  Dim td As TableDef

  On Error GoTo Fix1_BFWD_Fields_err
  ' add new field AccrualbfAbacus
  Call DisplayMessagePopup(Nothing, "Adding brought forwards per computation", "Updating TBO file")
  Set td = db.TableDefs("Charges")
  Call AddDoubleField(td, "AccrualbfAbacus")
  Set td = db.TableDefs("DividendsReceived")
  Call AddDoubleField(td, "AccrualbfAbacus")
  Set td = db.TableDefs("InvestmentIncome")
  Call AddDoubleField(td, "AccrualbfAbacus")
  Set td = db.TableDefs("Repairs")
  Call AddDoubleField(td, "AccrualbfAbacus")
  ' add new field AmountbfAbacus
  Set td = db.TableDefs("Reserves")
  Call AddDoubleField(td, "AmountbfAbacus")
  
Fix1_BFWD_Fields_end:
  Set td = Nothing
  Exit Sub
  
Fix1_BFWD_Fields_err:
  Err.Raise Err.Number, ErrorSource(Err, "Fix1_BFWD_Fields"), "An error occurred applying a database fix." & vbCrLf & vbCrLf & Err.Description
  Resume
End Sub

Private Sub Fix1_USAnalysis_Fields(db As Database)
  Dim td As TableDef

  On Error GoTo Fix1_USAnalysis_Fields_err
  ' add new field AccrualbfAbacus
  Call DisplayMessagePopup(Nothing, "Adding US/UK options", "Updating TBO file")
  Set td = db.TableDefs("ProfitAndLoss")
  Call AddTextField(td, "USAnalysis", 255)
  Set td = db.TableDefs("DividendsReceived")
  Call AddTextField(td, "USAnalysis", 255)
  Set td = db.TableDefs("RentalIncome")
  Call AddTextField(td, "USAnalysis", 255)
  Set td = db.TableDefs("InvestmentIncome")
  Call AddTextField(td, "USAnalysis", 255)
  
Fix1_USAnalysis_Fields_end:
  Set td = Nothing
  Exit Sub
  
Fix1_USAnalysis_Fields_err:
  Err.Raise Err.Number, ErrorSource(Err, "Fix1_USAnalysis_Fields"), "An error occurred applying a database fix." & vbCrLf & vbCrLf & Err.Description
  Resume
End Sub

Private Sub Fix1_ScheduleSetup(db As Database)

  On Error GoTo Fix1_ScheduleSetup_err
  
  Call db.Execute("UPDATE ScheduleSetup SET ColumnField=" & StrSQL("Enterprize Zone") & " WHERE ColumnField=" & StrSQL("Enterprise Zone"))

Fix1_ScheduleSetup_end:
  Exit Sub

Fix1_ScheduleSetup_err:
  Err.Raise Err.Number, ErrorSource(Err, "Fix1_ScheduleSetup"), "An error occurred applying a database fix." & vbCrLf & vbCrLf & Err.Description
  Resume
End Sub

Private Sub Fix1_sysImports(db As Database)
  Dim td As TableDef
  
  On Error GoTo Fix1_sysImports_err
  
  Set td = db.TableDefs("sys_Imports")
  Call KillField(td, "UpdateType")
  Call AddLongField(td, "UpdateType")
  
Fix1_sysImports_end:
  Set td = Nothing
  Exit Sub

Fix1_sysImports_err:
  Err.Raise Err.Number, ErrorSource(Err, "Fix1_sysImports"), "An error occurred applying a database fix." & vbCrLf & vbCrLf & Err.Description
  Resume
End Sub

Public Function Fix2(db As Database, dbTemplate As Database) As Boolean
  On Error GoTo Fix2_Err:
  
  Call Fix2_UpdateMenu(db)
  Fix2 = True

Fix2_end:
  Exit Function

Fix2_Err:
  Err.Raise Err.Number, ErrorSource(Err, "Fix2"), "An error occurred running Fix 2" & vbCrLf & vbCrLf & Err.Description
  Resume
End Function

Private Sub Fix2_UpdateMenu(db As Database)
  Dim rs As Recordset
  
  On Error GoTo Fix2_UpdateMenu_err
  
  Set rs = db.OpenRecordset("SELECT * FROM MENU", dbOpenDynaset, dbFailOnError)
  rs.FindFirst ("ObjectName=" & StrSQL("qAbacusOptions"))
  If rs.NoMatch Then
    rs.AddNew
    rs!ObjectName = "qAbacusOptions"
    rs!DataEntryTaxReview = True
    rs!TaxPackSchedule = True
    rs!DefaultOptions = 0
    rs!Source = "Added by Fix 2"
    rs.Update
  End If
  rs.FindFirst ("ObjectName=" & StrSQL("pOtherAnalysisMulti"))
  If rs.NoMatch Then
    rs.AddNew
    rs!ObjectName = "pOtherAnalysisMulti"
    rs!DataEntryTaxReview = True
    rs!TaxPackSchedule = True
    rs!DefaultOptions = 0
    rs!Source = "Added by Fix 2"
    rs.Update
  End If
  
Fix2_UpdateMenu_end:
  Exit Sub
  
Fix2_UpdateMenu_err:
  Err.Raise Err.Number, ErrorSource(Err, "Fix2 - UpdateManu"), "Fix 2 - Update Menu" & vbCrLf & vbCrLf & Err.Description
  Resume
End Sub

Public Function Fix5(db As Database, dbTemplate As Database) As Boolean
  On Error GoTo Fix5_Err:

  Call SyncScheduleSetupAndMenu(db, dbTemplate)
  Fix5 = True

Fix5_end:
  Exit Function

Fix5_Err:
  Err.Raise Err.Number, ErrorSource(Err, "Fix5"), "An error occurred running Fix 5" & vbCrLf & vbCrLf & Err.Description
End Function

Public Function Fix6(db As Database, dbTemplate As Database) As Boolean

  On Error GoTo Fix6_Err:
  
  Call SyncScheduleSetupAndMenu(db, dbTemplate)
  If TablePresent(db.TableDefs, "OtherInformation") Then
    If Not FieldPresent(db.TableDefs("OtherInformation").Fields, "InfoType") Then
      Call AddTextField(db.TableDefs("OtherInformation"), "InfoType", 255, False, False)
    End If
    Call db.Execute("UPDATE OtherInformation SET InfoType=" & StrSQL("None"), dbFailOnError)
  End If
  Call SyncTaxQuestions(db, dbTemplate)
  Call SyncTaxTypes(db, dbTemplate)
  Call Fix6_AddPLOnDisposal(db, dbTemplate)
  Call db.TableDefs.Refresh
  
  Fix6 = True

Fix6_end:
  Exit Function

Fix6_Err:
  Err.Raise Err.Number, ErrorSource(Err, "Fix6"), "An error occurred running Fix 6" & vbCrLf & vbCrLf & Err.Description
  Resume
End Function

Private Sub Fix6_AddPLOnDisposal(db As Database, dbTemplate As Database)
  Dim s1() As String, s2() As String
  Dim td As TableDef
  Dim sql As String
    
  On Error GoTo Fix6_AddPLOnDisposal_err
  s1 = Split("AllowZeroLength;Attributes;DefaultValue;Required;Size;Type;ValidationRule;ValidationText;OrdinalPosition", ";")
  s2 = Split("Clustered;Foreign;IgnoreNulls;Primary;Required;Unique", ";")
  If Not TablePresent(dbTemplate.TableDefs, "PLonDisposal") Then Err.Raise ERR_APPLY_FIXES, "Fix 6", "The PLonDisposal table cannot be found in the template database"
  If Not TablePresent(db.TableDefs, "PLonDisposal") Then
    Set td = dbTemplate.TableDefs("PLonDisposal")
    If Not CopyEntireTable(td, db, s1, s2) Then Err.Raise ERR_APPLY_FIXES, "Fix6", "An error occurred copying the PLonDisposal from the template into your file."
  End If
  If Not FieldPresent(db.TableDefs("Depreciation").Fields, "PLOnDisplosal") Then Exit Sub
  If FieldPresent(db.TableDefs("Depreciation").Fields, "QueryID") Then
    sql = "INSERT INTO PLonDisposal ( Schedule, Row, Source, AccountCode, Division, TransactionID, Description, Amount, TaxType, TaxTypeDefault, TaxTypeRule, RuleName, OverrideDescription, Disclose, QueryID ) "
    sql = sql & "SELECT Depreciation.Schedule, Depreciation.Row, Depreciation.Source, Depreciation.AccountCode, Depreciation.Division, Depreciation.TransactionID, Depreciation.Description, Depreciation.PLOnDisposal, Depreciation.TaxType, Depreciation.TaxTypeDefault, Depreciation.TaxTypeRule, Depreciation.RuleName, Depreciation.OverrideDescription, Depreciation.Disclose, Depreciation.QueryID "
    sql = sql & "FROM Depreciation WHERE TaxType=" & StrSQL("Profit or loss on disposal")
  Else
    sql = "INSERT INTO PLonDisposal ( Schedule, Row, Source, AccountCode, Division, TransactionID, Description, Amount, TaxType, TaxTypeDefault, TaxTypeRule, RuleName, OverrideDescription, Disclose) "
    sql = sql & "SELECT Depreciation.Schedule, Depreciation.Row, Depreciation.Source, Depreciation.AccountCode, Depreciation.Division, Depreciation.TransactionID, Depreciation.Description, Depreciation.PLOnDisposal, Depreciation.TaxType, Depreciation.TaxTypeDefault, Depreciation.TaxTypeRule, Depreciation.RuleName, Depreciation.OverrideDescription, Depreciation.Disclose "
    sql = sql & "FROM Depreciation WHERE TaxType=" & StrSQL("Profit or loss on disposal")
  End If
  
  Call db.Execute(sql, dbFailOnError)
  Call db.Execute("DELETE * FROM Depreciation WHERE TaxType=" & StrSQL("Profit or loss on disposal"), dbFailOnError)
  
Fix6_AddPLOnDisposal_end:
  Exit Sub
  
Fix6_AddPLOnDisposal_err:
  Err.Raise Err.Number, ErrorSource(Err, "Fix6_AddPLOnDisposal"), "An error occurred adding in a Profit or Loss on disposal table." & vbCrLf & vbCrLf & Err.Description
  Resume
End Sub

Public Function Fix7(db As Database, dbTemplate As Database) As Boolean

  On Error GoTo Fix7_Err:
  
  Call SyncScheduleSetupAndMenu(db, dbTemplate)
  Call SyncTaxQuestions(db, dbTemplate)
  Call SyncTaxTypes(db, dbTemplate)
  Fix7 = True

Fix7_end:
  Exit Function

Fix7_Err:
  Err.Raise Err.Number, ErrorSource(Err, "Fix7"), "An error occurred running Fix 7" & vbCrLf & vbCrLf & Err.Description
  Resume
End Function

Public Function Fix8(db As Database, dbTemplate As Database) As Boolean

  On Error GoTo Fix8_Err:
  
  Call SyncScheduleSetupAndMenu(db, dbTemplate)
  Call SyncTaxQuestions(db, dbTemplate)
  Call SyncTaxTypes(db, dbTemplate)
  Fix8 = True

Fix8_end:
  Exit Function

Fix8_Err:
  Err.Raise Err.Number, ErrorSource(Err, "Fix8"), "An error occurred running Fix 8" & vbCrLf & vbCrLf & Err.Description
  Resume
End Function

Public Function Fix9(db As Database, dbTemplate As Database) As Boolean

  On Error GoTo Fix9_Err:
  
  Call SyncScheduleSetupAndMenu(db, dbTemplate)
  Call SyncTaxQuestions(db, dbTemplate)
  Call SyncTaxTypes(db, dbTemplate)
  Fix9 = True

Fix9_end:
  Exit Function

Fix9_Err:
  Err.Raise Err.Number, ErrorSource(Err, "Fix9"), "An error occurred running Fix 9" & vbCrLf & vbCrLf & Err.Description
  Resume
End Function

Public Function Fix10(db As Database, dbTemplate As Database) As Boolean
  Dim td As TableDef
  Dim i As Long
  On Error GoTo Fix10_Err:
  
  If InCollection(db.TableDefs, "InvestmentTransactions") Then
    Set td = db.TableDefs("InvestmentTransactions")
    If Not FieldPresent(td.Fields, "ShareNumber") Then
      Call AddDoubleField(td, "ShareNumber")
      If FieldPresent(td.Fields, "Number") Then
        Call db.Execute("UPDATE InvestmentTransactions SET InvestmentTransactions.ShareNumber = [Number]", dbFailOnError)
      End If
      If InCollection(td.Indexes, "Number") Then
        Call td.Indexes.Delete("Number")
      End If
      Call KillField(td, "Number")
    End If
  End If
  Fix10 = True

Fix10_end:
  Exit Function

Fix10_Err:
  Err.Raise Err.Number, ErrorSource(Err, "Fix10"), "An error occurred running Fix 10" & vbCrLf & vbCrLf & Err.Description
  Resume
End Function

Public Function Fix11(db As Database, dbTemplate As Database) As Boolean

  On Error GoTo Fix11_Err:
  
  Call SyncScheduleSetupAndMenu(db, dbTemplate)
  
  Fix11 = True

Fix11_end:
  Exit Function

Fix11_Err:
  Err.Raise Err.Number, ErrorSource(Err, "Fix11"), "An error occurred running Fix 11" & vbCrLf & vbCrLf & Err.Description
  Resume
End Function

Public Function Fix12(db As Database, dbTemplate As Database) As Boolean

  On Error GoTo Fix12_Err:
  
  Call SyncScheduleSetupAndMenu(db, dbTemplate)
  Call SyncValidation(db, dbTemplate)
  
  Fix12 = True

Fix12_end:
  Exit Function

Fix12_Err:
  Err.Raise Err.Number, ErrorSource(Err, "Fix12"), "An error occurred running Fix 12" & vbCrLf & vbCrLf & Err.Description
  Resume
End Function

Public Function Fix13(db As Database, dbTemplate As Database) As Boolean

  On Error GoTo Fix13_Err:
  
  Call SyncScheduleSetupAndMenu(db, dbTemplate)
  Call SyncValidation(db, dbTemplate)
  Call SyncCache(db, dbTemplate)
  Call SyncExpCar(db, dbTemplate)
  
  Fix13 = True

Fix13_end:
  Exit Function

Fix13_Err:
  Err.Raise Err.Number, ErrorSource(Err, "Fix13"), "An error occurred running Fix 13" & vbCrLf & vbCrLf & Err.Description
  Resume
End Function

Public Function Fix14(db As Database, dbTemplate As Database) As Boolean
  Dim td As TableDef
  Dim sql As String
  
  On Error GoTo Fix14_Err:
  
  If Not CopyTable2(gwsMain, db, "sys_CarryForward", dbTemplate, "sys_CarryForward", TBL_COPY_OVERWRITE, PROP_ALL, "SELECT * FROM sys_CarryForward", True, False) Then Err.Raise ERR_COPY_TEMPLATE_FILE, "Fix14", "An error occurred updating the system table sys_CarryForward."
  If Not CopyTable2(gwsMain, db, "sys_CarryForward_Questions", dbTemplate, "sys_CarryForward_Questions", TBL_COPY_OVERWRITE, PROP_ALL, "SELECT * FROM sys_CarryForward_Questions", True, False) Then Err.Raise ERR_COPY_TEMPLATE_FILE, "Fix14", "An error occurred updating the system table sys_CarryForward_Questions."
  Call CopyQuery("AddDivisionalQuestionsSub", db, dbTemplate)
  Call CopyQuery("AddDivisionalQuestions", db, dbTemplate)
  Call CopyQuery("AddToPandLFlat", db, dbTemplate)
  Call CopyQuery("AddToPandLDownloadFlat", db, dbTemplate)
  If InCollection(db.TableDefs, "Output_D3") Then
    Set td = db.TableDefs("Output_D3")
    td.Name = "Output_D3_Old"
    If Not CopyTable2(gwsMain, db, "Output_D3", dbTemplate, "Output_D3", TBL_COPY, PROP_ALL, "SELECT * FROM Output_D3", True, False) Then Err.Raise ERR_COPY_TEMPLATE_FILE, "Fix14", "An error occurred copying the table Output_D3."
    sql = "INSERT INTO Output_D3 SELECT * FROM Output_D3_Old"
    Call db.Execute(sql, dbFailOnError)
    Call db.TableDefs.Delete("Output_D3_Old")
  End If
  
  Fix14 = True

Fix14_end:
  Set td = Nothing
  Exit Function

Fix14_Err:
  Err.Raise Err.Number, ErrorSource(Err, "Fix14"), "An error occurred running Fix 14" & vbCrLf & vbCrLf & Err.Description
  Resume
End Function

Public Function Fix15(db As Database, dbTemplate As Database) As Boolean
  Dim rs As Recordset
  Dim rsControl As Recordset
  Dim bExternal As Boolean
  Dim bInternal As Boolean
  
  On Error GoTo Fix15_Err:
 
  bExternal = False
  If TablePresent(db.TableDefs, "sys_Control") Then
    Set rsControl = db.OpenRecordset("sys_Control", dbOpenDynaset, dbFailOnError)
    If Not (rsControl.BOF And rsControl.EOF) Then
      rsControl.MoveFirst
      rsControl.FindFirst ("Flag=" & StrSQL("TemplateFile"))
      If rsControl.NoMatch Then
        GoTo Fix15_Err
      Else
        If StrComp(Mid$(rsControl!Text, 4, 1), "x", vbTextCompare) = 0 Then
          bExternal = True
        Else
          bInternal = True
        End If
      End If
      rsControl.FindFirst ("Flag=" & StrSQL("FileType"))
      If rsControl.NoMatch Then
        GoTo Fix15_Err
      Else
        If StrComp(rsControl!Text, "Investment Company", vbTextCompare) = 0 Then
          Call EstablishTemplateFile(db, , bExternal, bInternal)
          Set rs = db.OpenRecordset("SELECT * from TaxTypes Where Table='ManagementExpenses' AND Dlink='A10'", dbOpenDynaset)
          If Not (rs.BOF And rs.EOF) Then
            rs.Edit
              rs!Dlink = Null
              rs.Fields("[Dlink text]").Value = Null
            rs.Update
            rs.Close
          End If
        End If
      End If
    End If
  End If
        
  Fix15 = True

Fix15_end:
  Set rs = Nothing
  Set rsControl = Nothing
  Exit Function

Fix15_Err:
  Err.Raise Err.Number, ErrorSource(Err, "Fix15"), "An error occurred running Fix 15" & vbCrLf & vbCrLf & Err.Description
  Resume
End Function

Public Function Fix16(db As Database, dbTemplate As Database) As Boolean
  Dim sql As String
  
  On Error GoTo Fix16_Err:
 
  sql = "UPDATE FAAdds SET BroadCategory = 'Motor Vehicles' WHERE BroadCategory = 'Cars'"
  Call db.Execute(sql, dbFailOnError)
  
  sql = "UPDATE FADisps SET BroadCategory = 'Motor Vehicles' WHERE BroadCategory = 'Cars'"
  Call db.Execute(sql, dbFailOnError)
        
  Fix16 = True

Fix16_end:
  Exit Function

Fix16_Err:
  Err.Raise Err.Number, ErrorSource(Err, "Fix16"), "An error occurred running Fix 16" & vbCrLf & vbCrLf & Err.Description
  Resume
End Function



