VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Import"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private m_Import_Objects As Long
Private Cn As ADODB.Connection
Private sGUID As String
Private mUseSQL As Boolean
Private mAuditLogExists As Boolean
Private col_FileColumns As Collection
Private col_TableColumns As Collection
Private col_HeaderColumns As Collection
Private col_FooterColumns As Collection
Private lImportedLines As Long
Private lLinesWithWarningAndErrors As Long

Public Enum XML_SPEC_TYPE
  XML_FILE
  XML_STRING
End Enum

Public Enum IMPORTER_DB_TYPE
  ORACLE_DB
  SQL_DB
End Enum

Public Function ImportFile(sSpecFile As String, enumSpecType As XML_SPEC_TYPE, sImportFile As String, ByVal sUserName As String, Optional WErrors As WebErrors, Optional impSchemaPath As String, Optional UsrCnObj As ADODB.Connection) As String
  ImportFile = ImportFileEx(sSpecFile, enumSpecType, sImportFile, sUserName, WErrors, impSchemaPath, UsrCnObj)
End Function
Public Function ImportFile2(sSpecFile As String, enumSpecType As XML_SPEC_TYPE, sImportFile As String, ByVal sUserName As String, Optional WErrors As WebErrors, Optional impSchemaPath As String, Optional UsrCnObj As ADODB.Connection, Optional IncludeTrailing As Boolean = False) As String
  ImportFile2 = ImportFileEx(sSpecFile, enumSpecType, sImportFile, sUserName, WErrors, impSchemaPath, UsrCnObj, IncludeTrailing)
End Function

Private Sub ReadImportSpec(sSpecificationName As String, ByVal xmlType As XML_SPEC_TYPE, ispec As ImporterSpec, ByVal impSchemaPath As String)
Dim xmlDoc As DOMDocument30
Dim sline As String
Dim sXmlString As String
Dim sTemp As String
Dim node As IXMLDOMElement
Dim xmlNodeList As IXMLDOMNodeList


On Error GoTo ReadImportSpec_Err
  Set xmlDoc = New DOMDocument30
  Call ValidateSpecWithSchema(sSpecificationName, xmlType, impSchemaPath & XML_SCHEMA_FILE)
  If xmlType = XML_FILE Then
    sXmlString = GetXmlFromFile(sSpecificationName)
  Else
    sXmlString = sSpecificationName
  End If
  If Len(sXmlString) = 0 Then Err.Raise ERR_INVALID_SPECIFICATION_FILE, ErrorSource(Err, "Import.ReadImportSpec"), "Specification file is empty."
  If xmlDoc.loadXML(sXmlString) Then
   
    Set node = xmlDoc.selectSingleNode(XML_DSN_NODE)
    ispec.DBConnectionString = node.Text
    If Cn Is Nothing Then Set Cn = gADOHelper.ADOConnect(ispec.DBConnectionString)
    If Cn.State = adStateClosed Then Err.Raise ERR_INVALID_CONNECTION, ErrorSource(Err, "Import.ReadImportSpec"), "Invalid Connection object, no active conection."
    
    'read column collection
    Set xmlNodeList = xmlDoc.selectNodes(XML_COLUMN_NODE)
    Set col_FileColumns = ispec.SetFileColumnProperties(xmlNodeList)
    
    'read delimiter
    Set node = xmlDoc.selectSingleNode(XML_DELIMITER_NODE)
    If node Is Nothing Then
      ispec.Delimiter = DEFAULT_DELIMITER
    Else
      ispec.Delimiter = IsNullEx(node.Text, DEFAULT_DELIMITER)
    End If
    
    'read qualifier
    Set node = xmlDoc.selectSingleNode(XML_QUALIFIER_NODE)
    If node Is Nothing Then
      ispec.Qualifier = DEFAULT_QUALIFIER
    Else
      ispec.Qualifier = IsNullEx(node.Text, DEFAULT_QUALIFIER)
    End If
    
    'Read date formats
    Set node = xmlDoc.selectSingleNode(XML_DATE_FROM)
    If node Is Nothing Then
      ispec.DateFrom = "dd/mm/yyyy"
    Else
      ispec.DateFrom = IsNullEx(node.Text, "dd/mm/yyyy")
    End If
    Set node = xmlDoc.selectSingleNode(XML_DATE_TO)
    If node Is Nothing Then
      ispec.DateTo = "dd/mm/yyyy"
    Else
       ispec.DateTo = IsNullEx(node.Text, "dd/mm/yyyy")
    End If
    
    'read columns in dest table
    Set node = xmlDoc.selectSingleNode(XML_TARGETTABLE_NODE)
    ispec.ImportDest = IsNullEx(node.Text, "")
    If Len(ispec.ImportDest) = 0 Then Err.Raise ERR_MISSING_DEST_NAME, ErrorSource(Err, "Import.ReadImportSpec"), "No name given for destination table/view in the specification file. "
    Set xmlNodeList = xmlDoc.selectNodes(XML_TABLECOLUMN_NODE)
    Set col_TableColumns = ispec.SetTableColumnProperties(xmlNodeList, Cn)
    
    'Validate column names
    Call ValidateColumnNames(col_TableColumns, col_FileColumns)
    
    'error log and audit log
    Set node = xmlDoc.selectSingleNode(XML_ERRORLOG_NODE)
    ispec.ErrorLog = IsNullEx(node.Text, "")
    If Len(ispec.ErrorLog) = 0 Then Err.Raise ERR_MISSING_ERRORLOG_NAME, ErrorSource(Err, "Import.ReadImportSpec"), "No name given for error log table in the specification file. "
    If Not gADOHelper.IsTablePresentADO(Cn, ispec.ErrorLog) Then Err.Raise ERR_MISSING_ERRORLOG_NAME, ErrorSource(Err, "Import.ReadImportSpec"), Err.Description & vbCrLf & "Could not find error log table " & ispec.ErrorLog & " in the DB."

    Set node = xmlDoc.selectSingleNode(XML_AUDITLOG_NODE)
    ispec.AuditLog = IsNullEx(node.Text, "")
    If Len(ispec.AuditLog) = 0 Then Err.Raise ERR_MISSING_AUDITLOG_NAME, ErrorSource(Err, "Import.ReadImportSpec"), "No name given for error log table in the specification file. "
    If Not gADOHelper.IsTablePresentADO(Cn, ispec.AuditLog) Then Err.Raise ERR_MISSING_ERRORLOG_NAME, ErrorSource(Err, "Import.ReadImportSpec"), Err.Description & vbCrLf & "Could not find audit log table " & ispec.AuditLog & "in the DB."

    'read row information
    Set node = xmlDoc.selectSingleNode(XML_ROW_NODE)
    Call ispec.SetRowProperties(node)
    
    'read header and footer
    If ispec.FileHeader Then
      Set xmlNodeList = xmlDoc.selectNodes(XML_HEADER_NODE)
      Set col_HeaderColumns = ispec.SetHeaderAndFooter(xmlNodeList)
      'If col_HeaderColumns.Count = 0 Then Err.Raise ERR_MISSING_HEADER, ErrorSource(Err, "Import.ReadImportSpec"), "Missing header node in the specification file."
    End If
    If ispec.FileFooter Then
      Set xmlNodeList = xmlDoc.selectNodes(XML_FOOTER_NODE)
      Set col_FooterColumns = ispec.SetHeaderAndFooter(xmlNodeList, True)
      If col_FooterColumns.Count Then Err.Raise ERR_MISSING_HEADER, ErrorSource(Err, "Import.ReadImportSpec"), "Missing footer node in the specification file."
    End If
    
    'read optional information
    Set node = Nothing
    Set xmlNodeList = Nothing
  Else
    'error has been handled by schema validation.
  End If
  
ReadImportSpec_End:
  Set xmlDoc = Nothing
  Exit Sub

ReadImportSpec_Err:
  Err.Raise ERR_IMPORT_SPEC_READ_ERROR, ErrorSource(Err, "Import.ReadImportSpec"), "Error parsing import specification." & vbCrLf & Err.Description
  Resume
End Sub


Private Function ImportSingleLine(ByVal sSingleLine As String, ByVal ispec As ImporterSpec, ByVal lLineNum As Long, ByVal sGUID As String, ByVal Cn As Connection, cmd As ADODB.Command) As Boolean

Dim sBuffer() As String
Dim iLoop As Integer
Dim iNoOfParameterRequired As Integer
Dim iRecordsAffected As Integer
Dim iNoOfItemsInLine As Integer
Dim iLinesSkipped As Integer, i As Integer

Dim SQL As String

On Error GoTo ImportSingleLine_err:
  iNoOfItemsInLine = GetDelimitedValues(sBuffer, sSingleLine, , , ispec.Delimiter, ispec.Qualifier)
  iNoOfParameterRequired = cmd.Parameters.Count - 2
  For i = 1 To col_FileColumns.Count
    If col_FileColumns.Item(i).ColSkip Then
      iLinesSkipped = iLinesSkipped + 1
    End If
  Next i
  
  If mUseSQL Then
    SQL = GetSQLFromCmd(cmd, iNoOfItemsInLine, iNoOfParameterRequired, lLineNum, sGUID, sBuffer, ispec, iLinesSkipped)
    If Len(SQL) > 0 Then
      Call cmd.ActiveConnection.Execute(SQL, iRecordsAffected, adExecuteNoRecords)
    End If
  Else
    If AppendDataToCmd(cmd, iNoOfItemsInLine, iNoOfParameterRequired, lLineNum, sGUID, sBuffer, ispec, iLinesSkipped) Then
      cmd.Execute iRecordsAffected, , adCmdText + adExecuteNoRecords
    Else
      ' append failed
    End If
  End If
ImportSingleLine_end:
  ImportSingleLine = CBool(iRecordsAffected)
  Exit Function
  
ImportSingleLine_err:
  Err.Raise ERR_IMPORT_SPEC_READ_ERROR, ErrorSource(Err, "Import.ImportSingleLine"), "Error in importing single line at line " & lLineNum & ". " & vbCrLf & Err.Description
  Resume
End Function

Private Function CreateCmdFromSpec(ByVal tblColumns As Collection, ByVal filColumns As Collection, ByVal ispec As ImporterSpec, ByVal Cn As ADODB.Connection) As ADODB.Command
' Check for mapping information in the column collection

Dim sSQL As String
Dim sSqlAppend As String
Dim cmdObj As ADODB.Command
Dim iCounter As Integer

On Error GoTo CreateCmdFromSpec_err
  sSQL = "Insert into " & ispec.ImportDest & "("
  sSqlAppend = "("
  
   
  If tblColumns.Count < filColumns.Count Then  'Log warnings
    'there are extra columns in the specification file
    'raise error / log warning
    Err.Raise ERR_COLUMN_MISMATCH, ErrorSource(Err, "Import.CreateCmdFromSpec"), "There are " & filColumns.Count - 2 & " column(s) in the spec and " & tblColumns.Count - 2 & " column(s) in the table!"
  ElseIf tblColumns.Count > filColumns.Count Then
    ' fewer columns in the specification file, warn and continue with import
    ' Err.Raise ERR_COLUMN_MISMATCH, ErrorSource(Err, "Import.CreateCmdFromSpec"), "There are " & filColumns.Count - 2 & " column(s) in the spec and " & tblColumns.Count - 2 & " column(s) in the table!"
  End If
  
  
  'Call ValidateColumnDataTypes(tblColumns, filColumns)
  
  For iCounter = 1 To filColumns.Count - 1
    If Not filColumns.Item(iCounter).ColSkip Then
      If gDBHelper.DatabaseTarget = DB_TARGET_SQLSERVER Then
        sSQL = sSQL & "[" & filColumns.Item(iCounter).ColName & "], "
      Else
        sSQL = sSQL & filColumns.Item(iCounter).ColName & ", "
      End If
      sSqlAppend = sSqlAppend & "?,"
    End If
  Next iCounter
    If Not filColumns.Item(iCounter).ColSkip Then
      If gDBHelper.DatabaseTarget = DB_TARGET_SQLSERVER Then
        sSQL = sSQL & "[" & filColumns.Item(iCounter).ColName & "])"
      Else
        sSQL = sSQL & filColumns.Item(iCounter).ColName & ")"
      End If
      sSqlAppend = sSqlAppend & "?)"
    Else
      sSQL = Left(sSQL, Len(sSQL) - 2) & ")"
      sSqlAppend = Left(sSqlAppend, Len(sSqlAppend) - 1) & ")"
    End If
    
  'Create new command object
  Set cmdObj = New ADODB.Command
  Set cmdObj.ActiveConnection = Cn
  'Set the SQL Query
  cmdObj.CommandText = sSQL & " VALUES " & sSqlAppend
  cmdObj.Prepared = True
  
  'Configure Query Parameters
  For iCounter = 1 To filColumns.Count
    If Not filColumns.Item(iCounter).ColSkip Then
      cmdObj.Parameters.Append cmdObj.CreateParameter(filColumns.Item(iCounter).ColName, tblColumns.Item(filColumns.Item(iCounter).ColName).ColDataType, , tblColumns.Item(filColumns.Item(iCounter).ColName).ColWidth)
    End If
  Next iCounter
  Set CreateCmdFromSpec = cmdObj

CreateCmdFromSpec_end:
  Set cmdObj = Nothing
  Exit Function
  
CreateCmdFromSpec_err:
  Set CreateCmdFromSpec = Nothing
  Err.Raise ERR_BUILDING_SQL_STRING, ErrorSource(Err, "Import.CreateCmdFromSpec"), "Error in building Insert SQL." & vbCrLf & Err.Description
End Function

Private Sub Class_Initialize()
  m_Import_Objects = m_Import_Objects + 1
End Sub

Private Sub Class_Terminate()
  m_Import_Objects = m_Import_Objects - 1
End Sub
Private Sub ValidateSpecWithSchema(ByVal sSpecificationFile As String, ByVal xmlType As XML_SPEC_TYPE, ByVal sSchemaFile As String)
Dim sErrorString As String
Dim xmlschema As MSXML2.XMLSchemaCache40
Dim xmldom As MSXML2.DOMDocument40
Dim fs As FileSystemObject
   
On Error GoTo ValidateSpecWithSchema_err
  Set fs = New FileSystemObject
  
  If xmlType = XML_FILE Then
    If Not fs.FileExists(sSpecificationFile) Then
      Err.Raise ERR_NO_SPECIFICATION_FILE, ErrorSource(Err, "Import.ValidateSpecWithSchema"), "Specification file does not exist. " & sSpecificationFile & vbCrLf & Err.Description
    End If
  End If
  
  If Not fs.FileExists(sSchemaFile) Then
    Err.Raise ERR_NO_SCHEMA_FILE, ErrorSource(Err, "Import.ValidateSpecWithSchema"), "Schema file does not exist. " & sSchemaFile & vbCrLf & Err.Description
  End If
  
  Set xmlschema = New MSXML2.XMLSchemaCache40
  xmlschema.Add "xsd_importer_spec", sSchemaFile
  Set xmldom = New MSXML2.DOMDocument40
  Set xmldom.schemas = xmlschema
  
  xmldom.async = False
  xmldom.validateOnParse = True
  If xmlType = XML_FILE Then
    xmldom.Load sSpecificationFile
  Else
    xmldom.loadXML sSpecificationFile
  End If
  If xmldom.parseError.errorCode <> 0 Then
     sErrorString = "Error at line " & xmldom.parseError.Line & " in " & xmldom.parseError.srcText & vbCrLf
     sErrorString = sErrorString & xmldom.parseError.reason & "Error code " & xmldom.parseError.errorCode & vbCrLf
     sErrorString = sErrorString & xmldom.parseError.url
     Err.Raise ERR_VALIDATING_SPEC_WITH_SCHEMA, , sErrorString
  End If
  If xmldom.namespaces.Length = 0 Then Err.Raise ERR_VALIDATING_SPEC_WITH_SCHEMA, , "No namespace found in the xml specification file."
  If StrComp(xmldom.namespaces(0), xmlschema.namespaceURI(0), vbBinaryCompare) <> 0 Then
    sErrorString = "Namespace in the specification file (" & xmldom.namespaces(0) & ")does not match with schema namespace (" & xmlschema.namespaceURI(0) & ")"
    Err.Raise ERR_VALIDATING_SPEC_WITH_SCHEMA, , sErrorString
  End If

ValidateSpecWithSchema_end:
  Set fs = Nothing
  Set xmlschema = Nothing
  Set xmldom = Nothing
  Exit Sub
  
ValidateSpecWithSchema_err:
  Err.Raise ERR_VALIDATING_SPEC_WITH_SCHEMA, ErrorSource(Err, "Import.ValidateSpecWithSchema"), "Spec validation error." & vbCrLf & Err.Description
  Resume
End Sub
Private Function AppendDataToCmd(cmd As ADODB.Command, ByVal iNoOfItemsInLine As Integer, ByVal iNoOfParameterRequired As Integer, ByVal lLineNum As Long, ByVal sGUID As String, sBuffer() As String, ByVal ispec As ImporterSpec, ByVal iLinesToSkip As Integer) As Boolean
Dim iLoop, j As Integer
Dim ptr As Integer
Dim s As String

On Error GoTo AppendDataToCmd_err
  'Populate first two columns
  cmd(0) = sGUID
  cmd(1) = Str(lLineNum)
  If iNoOfItemsInLine < iNoOfParameterRequired + iLinesToSkip Then
        For iLoop = 1 To iNoOfItemsInLine
          For j = iLoop + ptr To iNoOfItemsInLine
            If Not col_FileColumns.Item(j + 2).ColSkip Then
              If cmd(iLoop + 1).Type = adDBTimeStamp Then
                  cmd(iLoop + 1) = FormatImportDate(sBuffer(j), ispec)
                Else
                  cmd(iLoop + 1) = sBuffer(j)
                End If
              Exit For
            End If
            ptr = ptr + 1
          Next j
        Next iLoop
        If iLinesToSkip > 0 Then
          For iLoop = iNoOfItemsInLine + 1 To iNoOfParameterRequired
            cmd(iLoop + 1) = Null
          Next iLoop
        Else
           For iLoop = iNoOfItemsInLine + 1 To iNoOfParameterRequired
            cmd(iLoop + 1) = Null
          Next iLoop
        End If
        Call LogError(ispec, Cn, sGUID, lLineNum, 2, "Missing " & iNoOfParameterRequired + iLinesToSkip - iNoOfItemsInLine & " parameter(s)/item(s) in this line.")
        lLinesWithWarningAndErrors = lLinesWithWarningAndErrors + 1
    
    ElseIf iNoOfItemsInLine > iNoOfParameterRequired + iLinesToSkip Then    ' Ignore extra columns and error log
        For iLoop = 1 To iNoOfParameterRequired
         For j = iLoop + ptr To iNoOfItemsInLine
            If Not col_FileColumns.Item(j + 2).ColSkip Then
              If cmd(iLoop + 1).Type = adDBTimeStamp Then
                  cmd(iLoop + 1) = FormatImportDate(sBuffer(j), ispec)
                Else
                  cmd(iLoop + 1) = sBuffer(j)
                End If
              Exit For
            End If
            ptr = ptr + 1
          Next j
        Next iLoop
        Call LogError(ispec, Cn, sGUID, lLineNum, 2, "There are " & iNoOfItemsInLine - iNoOfParameterRequired - iLinesToSkip & " extra parameter(s)/item(s) in this line.")
        lLinesWithWarningAndErrors = lLinesWithWarningAndErrors + 1
    Else
        For iLoop = 1 To iNoOfParameterRequired
          For j = iLoop + ptr To iNoOfItemsInLine
            If Not col_FileColumns.Item(j + 2).ColSkip Then
              If Len(sBuffer(j)) = 0 Then
                cmd(iLoop + 1) = Null
              Else
                If cmd(iLoop + 1).Type = adDBTimeStamp Then
                  cmd(iLoop + 1) = FormatImportDate(sBuffer(j), ispec)
                Else
                  cmd(iLoop + 1) = sBuffer(j)
                End If
              End If
              s = s + sBuffer(j) + ","
              Exit For
            End If
            ptr = ptr + 1
          Next j
        Next iLoop
  End If
  
  AppendDataToCmd = True

AppendDataToCmd_end:
  Exit Function

AppendDataToCmd_err:
  If Err.Number = adErrDataConversion Then
    AppendDataToCmd = False
    Call LogError(ispec, Cn, sGUID, lLineNum, 3, "Data type conversion error for field '" & cmd(iLoop + 1).Name & "'. Cannot convert value '" & sBuffer(j) & "' to type " & TypeName(cmd(iLoop + 1).Type) & " at line " & lLineNum & ".", True)
    lLinesWithWarningAndErrors = lLinesWithWarningAndErrors + 1
  Else
    Err.Raise ERR_APPEND_DATA, ErrorSource(Err, "Import.AppendDataToCmd"), "Error appending data to command object at line " & lLineNum & ". " & vbCrLf & Err.Description
  End If
  GoTo AppendDataToCmd_end
  Resume
End Function
Private Function ValidateColumnNames(ByVal table As Collection, ByVal file As Collection) As Boolean
Dim i, j As Integer
Dim errstring As String

On Error GoTo ValidateColumnDataTypes_err:

 For i = 1 To file.Count
    For j = 1 To table.Count
      If file.Item(i).ColSkip Then Exit For
      If StrComp(table.Item(j).ColName, file.Item(i).ColName, vbTextCompare) = 0 Then
        Exit For
      ElseIf j = table.Count Then
        Err.Raise ERR_VALIDATING_DATA_TYPE, ErrorSource(Err, "Import.ValidateColumnNames"), "Column name " & file.Item(i).ColName & " not found in the table. " & vbCrLf & Err.Description
      End If
    Next j
 Next i
  
ValidateColumnDataTypes_end:
  Exit Function
  
ValidateColumnDataTypes_err:
  Err.Raise ERR_VALIDATING_DATA_TYPE, ErrorSource(Err, "Import.ValidateColumnNames"), "Error validating column names." & vbCrLf & Err.Description
  Resume
End Function

Public Function UndoImport(ByVal sImportID As String, sUserName As String, ByVal sSpecFile As String, enumSpecType As XML_SPEC_TYPE, Optional bClearErrorLog As Boolean = False, Optional lLineNumber As Long = 0, Optional lUndoLineCount As Long, Optional WErrors As WebErrors, Optional impSchemaPath As String, Optional UsrCnObj As ADODB.Connection) As Boolean
Dim sSQL As String
Dim impSpec As ImporterSpec
Dim sSql2 As String


On Error GoTo UndoImport_err

  UndoImport = False
  Set impSpec = New ImporterSpec
  If Len(impSchemaPath) = 0 Then impSchemaPath = App.Path
  Call ReadImportSpec(sSpecFile, enumSpecType, impSpec, impSchemaPath)
  If lLineNumber > 0 Then  ' remove by line
    sSQL = "Delete from " & impSpec.ImportDest & " WHERE GUID = " & gDBHelper.StrSQL(sImportID) & "AND line_number=" & gDBHelper.NumSQL(lLineNumber)
  Else 'remove entire import
    sSQL = "Delete from " & impSpec.ImportDest & " WHERE GUID = " & gDBHelper.StrSQL(sImportID)
  End If
  sSql2 = "Select * from " & impSpec.AuditLog & " WHERE GUID = " & gDBHelper.StrSQL(sImportID)
  If ImportRecordExist(Cn, sSql2) Then
    Cn.BeginTrans
      Call Cn.Execute(sSQL, lUndoLineCount)
      sSql2 = "UPDATE " & impSpec.AuditLog & " SET user_name=" & gDBHelper.StrSQL(sUserName) & ", " & _
              "date_time_stamp=" & gDBHelper.StrSQL(Now()) & ", " & _
              "comments=" & gDBHelper.StrSQL("Import undone") & _
              " WHERE guid=" & gDBHelper.StrSQL(sImportID)
      Cn.Execute sSql2
      If bClearErrorLog Then sSql2 = "Delete from " & impSpec.ErrorLog & " WHERE GUID = " & gDBHelper.StrSQL(sImportID)
      Cn.Execute sSql2
      UndoImport = True
    Cn.CommitTrans
  Else
    UndoImport = False
    Err.Raise ERR_No_IMPORT_ID_FOUND, ErrorSource(Err, "Import.UndoImport"), Err.Description & vbCrLf & "No import records found for import ID - " & sImportID
  End If

UndoImport_end:
    Set Cn = Nothing
    Set impSpec = Nothing
  Exit Function

UndoImport_err:
  
  If Not WErrors Is Nothing Then
    Call WErrors.Add(Err, "abatec Web Import Errors")
  Else
    Err.Raise ERR_UNDO_IMPORT, ErrorSource(Err, "Import.ImportFile"), "Undo Import failed." & vbCrLf & Err.Description
  End If
  GoTo UndoImport_end
  Resume
End Function

Public Property Let GUID(new_guid As String)
  sGUID = new_guid
  mAuditLogExists = True
End Property

Public Property Get GUID() As String
  GUID = sGUID
End Property

Public Property Let ImporterDatabaseTarget(db_Type As DATABASE_TARGET)
  gDBHelper.DatabaseTarget = db_Type
End Property

Public Property Get ImporterDatabaseTarget() As DATABASE_TARGET
 ImporterDatabaseTarget = gDBHelper.DatabaseTarget
End Property

Public Property Let UseSQL(new_value As Boolean)
  mUseSQL = new_value
End Property

Public Property Get UseSQL() As Boolean
  UseSQL = mUseSQL
End Property


Private Function GetSQLFromCmd(cmd As ADODB.Command, ByVal iNoOfItemsInLine As Integer, ByVal iNoOfParameterRequired As Integer, ByVal lLineNum As Long, ByVal sGUID As String, sBuffer() As String, ByVal ispec As ImporterSpec, ByVal iLinesToSkip As Integer) As String
  Dim iLoop, j As Integer
  Dim ptr As Integer
  Dim s As String
  Dim SQL As String
  Dim Fields As String
  Dim Values As String
  Dim Value As String
  
  On Error GoTo GetSQLFromCmd_err
  
  SQL = "INSERT INTO " & ispec.ImportDest & " ("
  Fields = "GUID, LineNumber"
  Values = gDBHelper.StrSQL(sGUID) & ", " & gDBHelper.StrSQL(lLineNum)
  
  If iNoOfItemsInLine < iNoOfParameterRequired + iLinesToSkip Then
    For iLoop = 1 To iNoOfItemsInLine
      For j = iLoop + ptr To iNoOfItemsInLine
        If Not col_FileColumns.Item(j + 2).ColSkip Then
          If cmd(iLoop + 1).Type = adDBTimeStamp Then
            Value = gDBHelper.DateSQL(FormatImportDate(sBuffer(j), ispec))
          ElseIf cmd(iLoop + 1).Type = adBoolean Then
            Value = gDBHelper.BoolSQL(sBuffer(j))
          ElseIf cmd(iLoop + 1).Type = adVarChar Or cmd(iLoop + 1).Type = adVarWChar Then
            Value = gDBHelper.StrSQL(sBuffer(j))
          Else
            Value = sBuffer(j)
          End If
          Fields = Fields & ", " & cmd(iLoop + 1).Name
          Values = Values & ", " & Value
          Exit For
        End If
        ptr = ptr + 1
      Next j
    Next iLoop
    If iLinesToSkip > 0 Then
      For iLoop = iNoOfItemsInLine + 1 To iNoOfParameterRequired
        Fields = Fields & ", " & cmd(iLoop + 1).Name
        Values = Values & ", NULL"
      Next iLoop
    Else
      For iLoop = iNoOfItemsInLine + 1 To iNoOfParameterRequired
        Fields = Fields & ", " & cmd(iLoop + 1).Name
        Values = Values & ", NULL"
      Next iLoop
    End If
    Call LogError(ispec, Cn, sGUID, lLineNum, 2, "Missing " & iNoOfParameterRequired + iLinesToSkip - iNoOfItemsInLine & " parameter(s)/item(s) in this line.")
    lLinesWithWarningAndErrors = lLinesWithWarningAndErrors + 1
  ElseIf iNoOfItemsInLine > iNoOfParameterRequired + iLinesToSkip Then    ' Ignore extra columns and error log
    For iLoop = 1 To iNoOfParameterRequired
     For j = iLoop + ptr To iNoOfItemsInLine
        If Not col_FileColumns.Item(j + 2).ColSkip Then
          If cmd(iLoop + 1).Type = adDBTimeStamp Then
            Value = gDBHelper.DateSQL(FormatImportDate(sBuffer(j), ispec))
          ElseIf cmd(iLoop + 1).Type = adBoolean Then
            Value = gDBHelper.BoolSQL(sBuffer(j))
          ElseIf cmd(iLoop + 1).Type = adVarChar Or cmd(iLoop + 1).Type = adVarWChar Then
            Value = gDBHelper.StrSQL(sBuffer(j))
          Else
            Value = sBuffer(j)
          End If
          Fields = Fields & ", " & cmd(iLoop + 1).Name
          Values = Values & ", " & Value
        End If
        ptr = ptr + 1
      Next j
    Next iLoop
    Call LogError(ispec, Cn, sGUID, lLineNum, 2, "There are " & iNoOfItemsInLine - iNoOfParameterRequired - iLinesToSkip & " extra parameter(s)/item(s) in this line.")
    lLinesWithWarningAndErrors = lLinesWithWarningAndErrors + 1
  Else
    For iLoop = 1 To iNoOfParameterRequired
      For j = iLoop + ptr To iNoOfItemsInLine
        If Not col_FileColumns.Item(j + 2).ColSkip Then
          If Len(sBuffer(j)) = 0 Then
            Fields = Fields & ", " & cmd(iLoop + 1).Name
            Values = Values & ", NULL"
          Else
            If cmd(iLoop + 1).Type = adDBTimeStamp Then
              Value = gDBHelper.DateSQL(FormatImportDate(sBuffer(j), ispec))
            ElseIf cmd(iLoop + 1).Type = adBoolean Then
              Value = gDBHelper.BoolSQL(sBuffer(j))
            ElseIf cmd(iLoop + 1).Type = adVarChar Or cmd(iLoop + 1).Type = adVarWChar Then
              Value = gDBHelper.StrSQL(sBuffer(j))
            Else
              Value = sBuffer(j)
            End If
            Fields = Fields & ", " & cmd(iLoop + 1).Name
            Values = Values & ", " & Value
          End If
          s = s + sBuffer(j) + ","
          Exit For
        End If
        ptr = ptr + 1
      Next j
    Next iLoop
  End If
  
  SQL = SQL & Fields & ") VALUES (" & Values & ")"
  GetSQLFromCmd = SQL
GetSQLFromCmd_end:
  Exit Function

GetSQLFromCmd_err:
  If Err.Number = adErrDataConversion Then
    SQL = ""
    Call LogError(ispec, Cn, sGUID, lLineNum, 3, "Data type conversion error for field '" & cmd(iLoop + 1).Name & "'. Cannot convert value '" & sBuffer(j) & "' to type " & TypeName(cmd(iLoop + 1).Type) & ".", True)
    lLinesWithWarningAndErrors = lLinesWithWarningAndErrors + 1
  Else
    Err.Raise ERR_APPEND_DATA, ErrorSource(Err, "Import.GetSQLFromCmd"), "Error appending data to command object" & vbCrLf & Err.Description
  End If
  GoTo GetSQLFromCmd_end
  Resume
End Function

Private Function ImportSingleLine2(ByVal sSingleLine As String, ByVal ispec As ImporterSpec, ByVal lLineNum As Long, ByVal sGUID As String, ByVal Cn As Connection, cmd As ADODB.Command, Optional IncludeTrailing As Boolean = False) As Boolean

Dim sBuffer() As String
Dim iLoop As Integer
Dim iNoOfParameterRequired As Integer
Dim iRecordsAffected As Integer
Dim iNoOfItemsInLine As Integer
Dim iLinesSkipped As Integer, i As Integer

Dim SQL As String

On Error GoTo ImportSingleLine2_err:
  iNoOfItemsInLine = GetDelimitedValues2(sBuffer, sSingleLine, , , ispec.Delimiter, ispec.Qualifier, IncludeTrailing)
  iNoOfParameterRequired = cmd.Parameters.Count - 2
  For i = 1 To col_FileColumns.Count
    If col_FileColumns.Item(i).ColSkip Then
      iLinesSkipped = iLinesSkipped + 1
    End If
  Next i
  
  If mUseSQL Then
    SQL = GetSQLFromCmd(cmd, iNoOfItemsInLine, iNoOfParameterRequired, lLineNum, sGUID, sBuffer, ispec, iLinesSkipped)
    If Len(SQL) > 0 Then
      Call cmd.ActiveConnection.Execute(SQL, iRecordsAffected, adExecuteNoRecords)
    End If
  Else
    If AppendDataToCmd(cmd, iNoOfItemsInLine, iNoOfParameterRequired, lLineNum, sGUID, sBuffer, ispec, iLinesSkipped) Then
      cmd.Execute iRecordsAffected, , adCmdText + adExecuteNoRecords
    Else
      ' append failed
    End If
  End If
ImportSingleLine2_end:
  ImportSingleLine2 = CBool(iRecordsAffected)
  Exit Function
  
ImportSingleLine2_err:
  Err.Raise ERR_IMPORT_SPEC_READ_ERROR, ErrorSource(Err, "Import.ImportSingleLine"), "Error in importing single line at line " & lLineNum & ". " & vbCrLf & Err.Description
  Resume
End Function

Private Function ImportFileEx(sSpecFile As String, enumSpecType As XML_SPEC_TYPE, sImportFile As String, ByVal sUserName As String, Optional WErrors As WebErrors, Optional impSchemaPath As String, Optional UsrCnObj As ADODB.Connection, Optional IncludeTrailing As Boolean = False) As String
Dim impSpec As ImporterSpec
Dim frImportFile As FileSystemObject
Dim sline As String
Dim lRowCount As Long
Dim sSQL As String
Dim cmd As ADODB.Command
Dim trans As Transaction
Dim sComments As String
Dim tsImport As TextStream
Dim ptrEndOfFile As Long

On Error GoTo ImportFileEx_err
  'Set file
  Set impSpec = New ImporterSpec
  If Len(impSchemaPath) = 0 Then impSchemaPath = App.Path
  Set Cn = UsrCnObj
  Call ReadImportSpec(sSpecFile, enumSpecType, impSpec, impSchemaPath)
  impSpec.ImportUsername = sUserName
  impSpec.ImportFileName = sImportFile
  impSpec.ImportSpecName = sSpecFile
  
  Set frImportFile = New FileSystemObject
  Set cmd = CreateCmdFromSpec(col_TableColumns, col_FileColumns, impSpec, Cn)
  If cmd Is Nothing Then
    Err.Raise ERR_OBJECT_IS_NOTHING, ErrorSource(Err, "Import.ImportFile"), "Command object is nothing."
  End If
  'open and read import file
  Set tsImport = frImportFile.OpenTextFile(sImportFile)
  If Not tsImport Is Nothing Then
    ptrEndOfFile = PreProcessImportFile(sImportFile)
    lRowCount = 1
    If Len(sGUID) = 0 Then sGUID = GenerateGUID
    If impSpec.FileHeader Then
      sline = tsImport.ReadLine
      ' call headerinfo
      lRowCount = lRowCount + 1
    End If
    Set trans = New Transaction
    Call trans.Transaction(Cn, TRANS_BEGIN)
    Do While Not tsImport.AtEndOfStream
      sline = tsImport.ReadLine
      If ptrEndOfFile = lRowCount Then
        If impSpec.FileFooter Then Exit Do ' call readfooterinfo - not current version
      End If
      If Len(sline) > 0 Then
        If IncludeTrailing Then
          If ImportSingleLine2(sline, impSpec, lRowCount, sGUID, Cn, cmd, IncludeTrailing) Then ' send connection string and sql header
            lImportedLines = lImportedLines + 1
          End If
        Else
          If ImportSingleLine(sline, impSpec, lRowCount, sGUID, Cn, cmd) Then ' send connection string and sql header
            lImportedLines = lImportedLines + 1
          End If
        End If
      Else
        Call LogError(impSpec, Cn, sGUID, lRowCount, 1, "Empty row.")
        lLinesWithWarningAndErrors = lLinesWithWarningAndErrors + 1
      End If
      lRowCount = lRowCount + 1
    Loop
  Else
    Err.Raise ERR_IMPORT_FILE_MISSING, ErrorSource(Err, "Import.ImportFile"), "Import file missing. " & sImportFile
  End If
  
  'Log data
    If impSpec.FileHeader Or impSpec.FileFooter Then
      'Adjust for header/footer
      lRowCount = lRowCount + IIf(impSpec.FileHeader, -1, 0) + IIf(impSpec.FileFooter, -1, 0)
      sComments = "Imported lines = " & lImportedLines & "; error/warning = " & lLinesWithWarningAndErrors & "; Total lines in file (including header/footer)= " & lRowCount
    Else
      sComments = "Imported lines = " & lImportedLines & "; error/warning = " & lLinesWithWarningAndErrors & "; Total lines in file = " & lRowCount
    End If
  
  Call LogAuditInfo(impSpec, Cn, sGUID, sComments, enumSpecType, mAuditLogExists)
  Call trans.Transaction(Cn, TRANS_COMMIT)
  ImportFileEx = sGUID
ImportFileEx_end:
    Set Cn = Nothing
    Set frImportFile = Nothing
    Set impSpec = Nothing
    Set trans = Nothing
    Exit Function
  
ImportFileEx_err:
  Call App.LogEvent("ImportFile failed: " & Err.Description)
  
  Call gErrHelp.Push(Err)
  If Not Cn Is Nothing And Not trans Is Nothing Then Call trans.Transaction(Cn, TRANS_ROLLBACK)
  Call gErrHelp.Pop(Err)
  If Not WErrors Is Nothing Then
    Call WErrors.Add(Err, "abatec Web Import Errors")
  Else
    Err.Raise ERR_IMPORT, ErrorSource(Err, "Import.ImportFileEx"), "Import failed." & vbCrLf & Err.Description
  End If
  GoTo ImportFileEx_end
  Resume
End Function
