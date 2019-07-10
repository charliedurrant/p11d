VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DataAccess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private mDSN As String

'apf check for close connection / recordset
Friend Property Let DSN(ByVal NewValue As String)
  mDSN = NewValue
  If IsJetDSN(mDSN) Then gDBHelper.DatabaseTarget = DB_TARGET_JET
  'DEFAULT IS SQLSERVER
End Property

Friend Function NewSubmission(ByVal EntityID As String, ByVal EntityName As String, ByVal PeriodEnd As Date, ByVal Body As String, ByVal Username As String, ByVal Source As String, ByVal EmailAddress As String) As String
  Dim sql As String, GUID As String
  Dim cn As Connection
  Dim rs As Recordset
  'returns SubmissionID
  
  On Error GoTo NewSubmission_err
  GUID = GenerateGUID
  
  ' Insert without submission body
  sql = "INSERT INTO Submissions (SubmissionID, EntityID, EntityName, PeriodEnd, SubmissionDate, SubmissionBody, Status, Username, EmailAddress, LastUpdated, Source) VALUES " & "(" & gDBHelper.StrSQL(GUID) & "," & gDBHelper.StrSQL(EntityID) & "," & gDBHelper.StrSQL(EntityName) & "," & gDBHelper.DateTimeSQL(PeriodEnd) & "," & gDBHelper.DateTimeSQL(Now) & "," & gDBHelper.StrSQL("Insert Body here") & ",'" & S_STATUS_NOT_SUBMITTED & "'," & gDBHelper.StrSQL(Username) & "," & gDBHelper.StrSQL(EmailAddress) & "," & gDBHelper.DateTimeSQL(Now) & "," & gDBHelper.StrSQL(Source) & ")" ' removed gDBHelper.StrSQL("not_submitted") and added constant
  Set cn = NewADOConnection(True)
  cn.Execute (sql)
  
  'create recordset with write access with that single line
  sql = "Select * from Submissions where SubmissionID =" & gDBHelper.StrSQL(GUID)
  Set rs = New Recordset
  With rs
    .Open sql, cn, adOpenKeyset, adLockOptimistic
    .Fields("SubmissionBody").AppendChunk (Body)
    .Update
  End With
  
  Call CloseRecordSet(rs)
  Call CloseConnection(cn)
  NewSubmission = GUID

NewSubmission_end:
  Set rs = Nothing
  Set cn = Nothing
  Exit Function
NewSubmission_err:
  Err.Raise Err.Number, ErrorSourceComponent(Err, "NewSubmission", App.EXEName), Err.Description
  Resume
End Function

Friend Function SubmissionRS(ByVal SubmissionID As String, Optional ByVal WriteRequired As Boolean = False) As Recordset
  Dim sql As String
  Dim cn As Connection, rs As Recordset
  
  On Error GoTo SubmissionRS_err
  sql = "SELECT * FROM Submissions WHERE SubmissionID = " & gDBHelper.StrSQL(SubmissionID)
  Set cn = NewADOConnection(WriteRequired)
  
  If WriteRequired Then
    Set rs = New Recordset
    rs.Open sql, cn, adOpenKeyset, adLockOptimistic
    Set SubmissionRS = rs
  Else
    Set SubmissionRS = cn.Execute(sql)
  End If
  
SubmissionRS_end:
  Set rs = Nothing
  Set cn = Nothing
  Exit Function
SubmissionRS_err:
  Err.Raise Err.Number, ErrorSourceComponent(Err, "SubmissionRS", App.EXEName), Err.Description
  Resume
End Function

Friend Sub UpdateStatus(ByVal SubmissionID As String, ByRef status As String, Optional ByVal ResponseEndPoint As String, Optional ByVal PollInterval As String, Optional ByVal Message As String = "", Optional ByVal CorrelationID As String, Optional ByVal AppendMessage As Boolean)
  Dim sql As String, oldStatus As String
  Dim cn As Connection, rs As Recordset
  Dim Position As Integer, oldPosition As Integer
  
  On Error GoTo UpdateStatus_err
  
'  ' SET STATUS POSITION
'  Select Case status
'    Case S_STATUS_NONE
'      Position = 0
'    Case S_STATUS_NOT_SUBMITTED 'new submission
'      Position = 1
'    Case S_STATUS_ERROR
'      Position = 2
'    Case S_STATUS_ERROR_REQUIRES_DELETE 'new submission
'      Position = 3
'    Case S_STATUS_SUBMISSION_ACKNOWLEDGEMENT
'      Position = 4
'    Case S_STATUS_SUBMISSION_RESPONSE
'      Position = 5
'    Case S_STATUS_DELETE_ACKNOWLEDGEMENT
'      Position = 6
'    Case S_STATUS_DELETE_RESPONSE
'      Position = 7
'    Case S_STATUS_DELETE_REQUEST_CLIENT
'      Position = 8
'  End Select
'
'  ' check status in table before updating
'  Set cn = NewADOConnection(True)
'  sql = "SELECT Status From Submissions"
'  Set rs = cn.Execute(sql)
'
'  oldStatus = rs.Fields("Status").Value
'
'  ' SET OLDSTATUS POSITION
'  Select Case oldStatus
'    Case S_STATUS_NONE
'      oldPosition = 0
'    Case S_STATUS_NOT_SUBMITTED 'new submission
'      oldPosition = 1
'    Case S_STATUS_ERROR
'      oldPosition = 2
'    Case S_STATUS_ERROR_REQUIRES_DELETE 'new submission
'      oldPosition = 3
'    Case S_STATUS_SUBMISSION_ACKNOWLEDGEMENT
'      oldPosition = 4
'    Case S_STATUS_SUBMISSION_RESPONSE
'      oldPosition = 5
'    Case S_STATUS_DELETE_ACKNOWLEDGEMENT
'      oldPosition = 6
'    Case S_STATUS_DELETE_RESPONSE
'      oldPosition = 7
'    Case S_STATUS_DELETE_REQUEST_CLIENT
'      oldPosition = 8
'  End Select
'
'  If oldPosition > Position Then
'    GoTo UpdateStatus_end
'  Else
  
  Set cn = NewADOConnection(True)

    If AppendMessage Then
      sql = "UPDATE Submissions SET Status = " & gDBHelper.StrSQL(status) & ", LastUpdated =  " & gDBHelper.DateTimeSQL(Now) & ", ResponseEndPoint =  " & gDBHelper.StrSQL(ResponseEndPoint) & ", PollInterval =  " & gDBHelper.NumSQL(PollInterval) & IIf(Len(Message) > 0, ", Message = Message & " & gDBHelper.StrSQL(Message), "") & ", CorrelationID =  " & gDBHelper.StrSQL(CorrelationID) & " WHERE SubmissionID = " & "'" & SubmissionID & "'"
    Else
      sql = "UPDATE Submissions SET Status = " & gDBHelper.StrSQL(status) & ", LastUpdated =  " & gDBHelper.DateTimeSQL(Now) & ", ResponseEndPoint =  " & gDBHelper.StrSQL(ResponseEndPoint) & ", PollInterval =  " & gDBHelper.NumSQL(PollInterval) & IIf(Len(Message) > 0, ", Message =  " & gDBHelper.StrSQL(Message), "") & ", CorrelationID =  " & gDBHelper.StrSQL(CorrelationID) & " WHERE SubmissionID = " & "'" & SubmissionID & "'"
    End If
    cn.Execute (sql)
'  End If
  
  Call CloseRecordSet(rs)
  Call CloseConnection(cn)
  
UpdateStatus_end:
  Set rs = Nothing
  Set cn = Nothing
  Exit Sub
UpdateStatus_err:
  Err.Raise Err.Number, ErrorSourceComponent(Err, "UpdateStatus", App.EXEName), Err.Description
  Resume
End Sub

Friend Function StalledSubmissions(Optional ByVal Filter As String = "", Optional ByVal OrderBy As String = "SubmissionID", Optional ByVal StatusNow As Boolean = False) As Recordset
  Dim sql As String, cn As Connection
  Dim lUpdated As String, PollInt As String, Stat As String
    
  On Error GoTo StalledSubmissions_err
  
  'depends on current status, last updated, poll interval
  'Status Condition - Exclude Status completed or errors
  Stat = "Status NOT IN (" & gDBHelper.StrSQL(S_STATUS_COMPLETED) & "," & gDBHelper.StrSQL(S_STATUS_ERROR) & "," & gDBHelper.StrSQL(S_STATUS_CLIENT_DELETED) & ")"
  
  ' lastUpdated condition - should be more than 30 seconds
  If Not StatusNow Then lUpdated = " AND LastUpdated < " & gDBHelper.DateTimeSQL(CDate(Now - CDate("0:00:30"))) 'pq
  
  sql = "SELECT * FROM Submissions WHERE " & Stat & lUpdated
  
  If Len(Filter) > 0 Then sql = sql & " AND " & Filter
  sql = sql & " ORDER BY " & OrderBy
  
  Set cn = NewADOConnection(False)
  Set StalledSubmissions = cn.Execute(sql)
  
StalledSubmissions_end:
  Set cn = Nothing
  Exit Function
  
StalledSubmissions_err:
  Err.Raise Err.Number, ErrorSourceComponent(Err, "StalledSubmissions", App.EXEName), Err.Description
  Resume
End Function

Private Function NewADOConnection(ByVal RequireWriteConnection As Boolean, Optional ByVal adCmdType As CommandTypeEnum = adCmdText) As Connection
  
  On Error GoTo NewADOConnection_err
  Set NewADOConnection = New Connection
  NewADOConnection.CommandTimeout = 30
  If RequireWriteConnection Then
    NewADOConnection.Mode = adModeReadWrite
  Else
    NewADOConnection.Mode = adModeRead
  End If
  NewADOConnection.Open mDSN
  Exit Function
  
NewADOConnection_err:
  Set NewADOConnection = Nothing
  Err.Raise Err.Number, ErrorSourceComponent(Err, "NewADOConnection", App.EXEName), Err.Description
End Function

Friend Sub ArchiveUpdate(ByVal ArchiveLevel As ARCHIVE_LEVEL, ByVal SubmissionID As String, ByVal Accounts As String, ByVal Computations As String)
  Dim Body As String
  Dim doc As DOMDocument, rs As New Recordset
  Dim currnode As IXMLDOMNode
  Dim cn As Connection, sql As String
        
  On Error GoTo ArchiveUpdate_Err
  
  'get recordset here as it is needed to ensure if the archive level is complete, we do not lose the submission body
      
  Select Case ArchiveLevel
  Case "0"
    ' keep all of it, i.e, do nothing
    GoTo ArchiveUpdate_End
  Case "1"
    ' read into dom doc and remove the attachment tag content only, not tags
    Set rs = SubmissionRS(SubmissionID, True)
    If rs.BOF And rs.EOF Then Err.Raise ERR_INVALIDID, "ArchiveUpdate", "Invalid SubmissionID [" & SubmissionID & "]"
    Body = rs.Fields("SubmissionBody").Value
    If Len(Accounts) > 0 Or Len(Computations) > 0 Then
      Set doc = New DOMDocument
      With doc
        .loadXML Body
        If Len(Accounts) > 0 Then
        Set currnode = .selectSingleNode("//CompanyTaxReturn/AttachedFiles/Attachment[@Type = ""accounts""]") '.selectSingleNode("//CompanyTaxReturn/AttachedFiles/Attachment")
        currnode.Text = "Not archived"
        End If
        If Len(Computations) > 0 Then
        Set currnode = .selectSingleNode("//CompanyTaxReturn/AttachedFiles/Attachment[@Type = ""computations""]") '.selectSingleNode("//CompanyTaxReturn/AttachedFiles/Attachment")
        currnode.Text = "Not archived"
        End If
        Body = .XML
      End With
    End If
    Call CloseRecordSet(rs)
  Case "2"
    ' remove the whole field value
    Body = "Not archived"
  End Select
  
  sql = "UPDATE Submissions SET SubmissionBody = " & gDBHelper.StrSQL(Body) & " WHERE SubmissionID = " & "'" & SubmissionID & "'"
  Set cn = NewADOConnection(True)
  cn.Execute (sql)
  cn.Close
  Call CloseConnection(cn)

ArchiveUpdate_End:
  Set rs = Nothing
  Set cn = Nothing
  Exit Sub
  
ArchiveUpdate_Err:
  Err.Raise Err.Number, ErrorSourceComponent(Err, "ArchiveUpdate", App.EXEName), Err.Description
  Resume ArchiveUpdate_End
End Sub

Friend Function GetFilteredRecordSet(Optional ByVal Filter As String = "") As Recordset
  Dim sql As String, cn As Connection
    
  On Error GoTo GetFilteredRecordSet_err
    
  sql = "SELECT SubmissionID, Status, Message, LastUpdated FROM Submissions "
  If Len(Filter) > 0 Then sql = sql & "WHERE " & Filter & " ORDER BY Status, EntityName,SubmissionDate"
  
  Set cn = NewADOConnection(False)
  Set GetFilteredRecordSet = cn.Execute(sql)
  'cn.Close ' check if works
  
GetFilteredRecordSet_end:
  Set cn = Nothing
  Exit Function
  
GetFilteredRecordSet_err:
  Err.Raise Err.Number, ErrorSourceComponent(Err, "GetFilteredRecordSet", App.EXEName), Err.Description
  Resume
End Function

'Private Sub AddSubmissionBody(ByVal Body As String, ByVal SubmissionID As String)
'  Dim cn As Connection, sql As String, rs As Recordset
'
'  On Error GoTo AddSubmissionBody_Err
'
''  sql = "UPDATE Submissions SET SubmissionBody = " & gDBHelper.StrSQL(Body) & " WHERE SubmissionID = " & "'" & SubmissionID & "'"
''  Set cn = NewADOConnection(True)
''  cn.open
''  cn.Execute (sql)
''  cn.Close
'  Set cn = NewADOConnection(True)
'  cn.open
'  sql = "Select * from Submissions where SubmissionID =" & gDBHelper.StrSQL(SubmissionID)
'  Set rs = New Recordset
'  With rs
'    .open sql, cn, adOpenKeyset, adLockOptimistic
'    .Fields("SubmissionBody").AppendChunk (Body)
'    .Update
'    .Close
'  End With
'  cn.Close
'
'AddSubmissionBody_End:
'  Set rs = Nothing
'  Set cn = Nothing
'  Exit Sub
'
'AddSubmissionBody_Err:
'  Err.Raise Err.Number, ErrorSourceComponent(Err, "AddSubmissionBody", App.EXEName), Err.Description
'  Resume AddSubmissionBody_End
'End Sub
  
Friend Function IsAmendedSubmission(ByVal EntityID As String) As Boolean
  Dim sql As String, cn As Connection, rs As Recordset
    
  On Error GoTo IsAmendedSubmission_err
    
  sql = "SELECT SubmissionID FROM Submissions WHERE EntityID = " & gDBHelper.StrSQL(EntityID) & " AND Status = " & gDBHelper.StrSQL(S_STATUS_COMPLETED)
  
  Set cn = NewADOConnection(False)
  Set rs = cn.Execute(sql)
  If Not (rs.BOF And rs.EOF) Then IsAmendedSubmission = True
  
IsAmendedSubmission_end:
  Set rs = Nothing
  Set cn = Nothing
  Exit Function
  
IsAmendedSubmission_err:
  Err.Raise Err.Number, ErrorSourceComponent(Err, "IsAmendedSubmission", App.EXEName), Err.Description
  Resume
End Function
