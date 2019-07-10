Attribute VB_Name = "DBFunctions"
Option Explicit

Public Function ADOSQLConnectString(ByVal xmlNode As IXMLDOMElement) As String
  Dim s As String
  ' DB Provider test goes here, assume default is SQL
  s = "PROVIDER=" & xmlNode.getAttribute("provider")
  s = s & ";Data Source=" & xmlNode.getAttribute("servername")
  s = s & ";Initial Catalog=" & xmlNode.getAttribute("database")
  s = s & ";User ID=" & xmlNode.getAttribute("userid")
  s = s & ";Password=" & xmlNode.getAttribute("password")
  ADOSQLConnectString = s
End Function
Public Sub LogAuditInfo(ByVal ispec As ImporterSpec, ByVal Cn As Connection, ByVal sGUID As String, ByVal sComments As String, ByVal enumSpecType As XML_SPEC_TYPE, AuditLogExists As Boolean)
  Dim sInsert As String
  
On Error GoTo LogAuditInfo_err

  If AuditLogExists Then
    sInsert = "UPDATE " & ispec.AuditLog & " SET user_name = " & gDBHelper.StrSQL(ispec.ImportUsername) & ", date_time_stamp = " & gDBHelper.StrSQL(Now()) & ",file_name = " & gDBHelper.StrSQL(ispec.ImportFileName) & ", spec_name = " & IIf(enumSpecType = XML_FILE, gDBHelper.StrSQL(ispec.ImportSpecName), "NULL") & ", comments = " & gDBHelper.StrSQL(sComments) & " WHERE guid=" & gDBHelper.StrSQL(sGUID)
  Else
    sInsert = "INSERT INTO " & ispec.AuditLog & " ( guid, user_name, date_time_stamp, file_name, spec_name, comments) VALUES (" & _
    gDBHelper.StrSQL(sGUID) & "," & _
    gDBHelper.StrSQL(ispec.ImportUsername) & "," & _
    gDBHelper.StrSQL(Now()) & "," & _
    gDBHelper.StrSQL(ispec.ImportFileName) & "," & _
    IIf(enumSpecType = XML_FILE, gDBHelper.StrSQL(ispec.ImportSpecName), "NULL") & "," & _
    gDBHelper.StrSQL(sComments) & ")"
  End If
  Call Cn.Execute(sInsert)
  
LogAuditInfo_end:
  Exit Sub

LogAuditInfo_err:
  Err.Raise ERR_IMPORT_LOG_ERROR, ErrorSource(Err, "DBFunctions.LogAuditInfo"), "Could not log import information. " & vbCrLf & Err.Description
  Resume Next
End Sub

Public Sub LogError(ByVal ispec As ImporterSpec, ByVal Cn As Connection, ByVal sGUID As String, ByVal lLineNumber As Long, lErrorNum As Long, ByVal sErrorString As String, Optional berror As Boolean = False)
  Dim sInsert As String
 
On Error GoTo LogError_err
  sInsert = "INSERT INTO " & ispec.ErrorLog & " (guid,line_number,b_error,error_number,error_description) VALUES (" & _
  gDBHelper.StrSQL(sGUID) & "," & _
  gDBHelper.NumSQL(lLineNumber) & "," & _
  gDBHelper.NumSQL(berror) & "," & _
  gDBHelper.NumSQL(lErrorNum) & "," & _
  gDBHelper.StrSQL(Left$(sErrorString, 1000)) & ")"
  Cn.Execute sInsert
  
LogError_end:
  Exit Sub

LogError_err:
  Err.Raise ERR_IMPORT_LOG_ERROR, ErrorSource(Err, "DBFunctions.LogError"), "Could not log errors information. " & vbCrLf & Err.Description
  Resume Next
End Sub
Public Function ImportRecordExist(ByVal Cn As Connection, ByVal sSQL As String) As Boolean
Dim lRecCount As Long

On Error GoTo ImportRecordExist_err
  
  ImportRecordExist = False
  Cn.Execute sSQL, lRecCount
  ImportRecordExist = CBool(lRecCount)

ImportRecordExist_end:
  Exit Function

ImportRecordExist_err:
  Err.Raise Err.Number, ErrorSource(Err, "DBFunctions.LogAuditInfo"), "Error checking Import log " & vbCrLf & Err.Description
  Resume Next
End Function

Public Function FormatImportDate(ByVal dateVal As String, ByVal ispec As ImporterSpec) As String
Dim sTemp As String
Dim dRetVal As String


On Error GoTo FormatImportDate_err

' test if the given value is a valid date
If IsDate(dateVal) Then ' Do formating
    dRetVal = Format(dateVal, ispec.DateTo)
Else
  Select Case ispec.DateFrom
  
  Case "ddmmyyyy"
    sTemp = Format(dateVal, "##/##/####")
  Case "yyyymmdd"
    sTemp = Format(dateVal, "####/##/##")
  Case "ddmmyy"
    sTemp = Format(dateVal, "##/##/##")
  Case "yymmdd"
    sTemp = Format(dateVal, "##/##/##")
  Case Else
  '??
  End Select
  dRetVal = Format(sTemp, ispec.DateTo)
End If

If IsDate(dRetVal) Then
  FormatImportDate = dRetVal
Else
  FormatImportDate = dateVal
End If

FormatImportDate_end:
  Exit Function
  
FormatImportDate_err:
  Err.Raise ERR_CONVERTING_DATES, ErrorSource(Err, "DBFunctions.FormatImportDate"), "Error formating date " & dateVal & " to " & "SSD" & "." & vbCrLf & Err.Description
  Resume
End Function
