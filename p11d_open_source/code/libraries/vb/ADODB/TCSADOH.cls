VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TCSADOHelper"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Enum FILTER_TYPES
  INCLUDE_FILTER = 0
  EXCLUDE_FILTER
End Enum
Private Const XML_OPEN_COMMENT As String = "<!--"
Private Const XML_CLOSE_COMMENT As String = "-->"

Public Function GetADOError(ByVal ErrObj As ErrObject, ByVal ADOcn As Connection) As String
  Dim eObj As ADODB.Error
  Dim errstring As String, edesc As String
    
  If Not ADOcn Is Nothing Then
    If ADOcn.Errors.Count > 0 Then
      errstring = ""
      For Each eObj In ADOcn.Errors
        edesc = RTrimAny(eObj.Description, vbCrLf)
        errstring = errstring & "ADO Error (0x" & Left$(Hex$(eObj.Number) & "00000000", 8) & ")" & vbCrLf & _
                                IIf(Len(edesc) = 0, "No error description provided", edesc) & vbCrLf & _
                                "Source: " & IIf(Len(eObj.Source) = 0, "(not provided)", eObj.Source) & vbCrLf & _
                                "SQL State: " & eObj.SQLState & vbCrLf & _
                                "NativeError: " & eObj.NativeError & vbCrLf & vbCrLf
      Next eObj
      ADOcn.Errors.Clear
      GetADOError = errstring
      If Not ErrObj Is Nothing Then Err.Description = errstring
    End If
  End If
End Function

  ' allows us to generate a conflict on AddNew
Public Sub ChangeFirstRecord(ByVal ADOrs As Recordset, ByVal FieldName As String)
  Dim v As Variant
  
  On Error GoTo ChangeFirstRecord_err
  ADOrs.MoveFirst
  v = ADOrs.Fields(FieldName).Value
  ADOrs.Update FieldName, v
  'ADOrs.Update FieldName, v
  Exit Sub
  
ChangeFirstRecord_err:
  Err.Raise Err.Number, ErrorSource(Err, "ChangeFirstRecord"), Err.Description
End Sub

Public Function IsDSN(ByVal DSN As String) As Boolean
  IsDSN = (InStr(1, DSN, "PROVIDER=", vbTextCompare) > 0)
End Function

Public Function IsJetDSN(ByVal DSN As String) As Boolean
  IsJetDSN = (InStr(1, DSN, "PROVIDER=Microsoft.Jet", vbTextCompare) > 0)
End Function

Public Function IsTablePresentADO(ByVal ADOcn As Connection, ByVal TableName As String) As Boolean
  Dim adx As ADOX.Catalog
  Dim atbl As ADOX.Table
  
  On Error GoTo IsTablePresentADO_err
  Set adx = New ADOX.Catalog
  Set adx.ActiveConnection = ADOcn
  Set atbl = adx.Tables(TableName)
  
  IsTablePresentADO = True
IsTablePresentADO_end:
  Set adx = Nothing
  Exit Function
  
IsTablePresentADO_err:
  IsTablePresentADO = False
  Resume IsTablePresentADO_end
End Function

Public Function IsFieldPresentADO(ByVal cFields As Fields, ByVal FieldName As String) As Boolean
  Call xSet("IsFieldPresentADO")
  IsFieldPresentADO = InCollection(cFields, FieldName)
  Call xReturn("IsFieldPresentADO")
End Function

Public Function ADOtoDatatype(ByVal ADOType As ADODB.DataTypeEnum) As DATABASE_FIELD_TYPES
  Select Case ADOType
    Case adBoolean
      ADOtoDatatype = TYPE_BOOL
    Case adBinary, adInteger, adTinyInt, adSmallInt, adUnsignedInt, adUnsignedSmallInt, adUnsignedTinyInt
      ADOtoDatatype = TYPE_LONG
    Case adChar, adBSTR, adLongVarChar, adLongVarWChar, adVarWChar, adWChar, adVarChar
      ADOtoDatatype = TYPE_STR
    Case adCurrency, adDecimal, adDouble, adNumeric, adSingle, adVarNumeric
      ADOtoDatatype = TYPE_DOUBLE
    Case adDate, adDBTime, adDBTimeStamp
      ADOtoDatatype = TYPE_DATE
    Case adLongVarBinary, adVarBinary
      ADOtoDatatype = TYPE_BLOB
    Case Else ' adBigInt , adarray, adbyref, adDBDate,  adEmpty, adError, adGUID, adIDispatch, adIUnknown, adUnsignedBigInt, adUserDefined, adVariant, advector
      Call ECASE("Unrecognised data Type: " & CStr(ADOType))
  End Select
End Function

Public Function ADOConnect(ByVal ConnectString As String, Optional ByVal CursorLocation As ADODB.CursorLocationEnum = adUseServer) As ADODB.Connection
  Set ADOConnect = ADOConnectEx(ConnectString, CursorLocation, -1)
End Function

Public Function ADOConnectEx(ByVal ConnectString As String, Optional ByVal CursorLocation As ADODB.CursorLocationEnum = adUseServer, Optional ByVal timeOut As Long = -1) As ADODB.Connection
  On Error GoTo ADOConnectEx_err
  
  Set ADOConnectEx = New Connection
  ADOConnectEx.CursorLocation = CursorLocation
  If timeOut <> -1 Then ADOConnectEx.ConnectionTimeout = timeOut
  ADOConnectEx.Open ConnectString
  'apf no obvoius difference !!! SQL 7.0 ADOConnect.Properties("Enable Fastload").Value = True
  Exit Function
  
ADOConnectEx_err:
  Set ADOConnectEx = Nothing
  Err.Raise Err.Number, "ADOConnectEx", "Unable to open the connection" & vbCrLf & ConnectString & vbCrLf & Err.Description
End Function

Public Function ShapeDSN(ByVal ConnectString As String) As String
  ShapeDSN = "PROVIDER=MSDataShape; DATA " & ConnectString
End Function

Public Function ADOAccessConnectString(ByVal dbPath As String, Optional ByVal SystemDBPath As Variant, Optional ByVal UserID As Variant, Optional ByVal UserPassword As Variant, Optional ByVal DBPassword As Variant) As String
  Dim s As String
  
  s = "PROVIDER=Microsoft.Jet.OLEDB.3.51"
  s = s & ";Data Source=" & dbPath
  If Not IsMissing(SystemDBPath) Then s = s & ";Jet OLEDB:System Database=" & SystemDBPath
  If Not IsMissing(DBPassword) Then s = s & ";Jet OLEDB:Database Password=" & DBPassword
  If Not IsMissing(UserID) Then s = s & ";User ID=" & UserID
  If Not IsMissing(UserPassword) Then s = s & ";Password=" & UserPassword
  ADOAccessConnectString = s
End Function
Public Function ADOOracleConnectString(ByVal DataSource As String, Optional ByVal UserID As Variant, Optional ByVal Password As Variant) As String
  Dim s As String
  
  s = "Provider=MSDAORA.1;Persist Security Info=True"
  s = s & ";Data Source=" & DataSource
  If Not IsMissing(UserID) Then s = s & ";User ID=" & UserID
  If Not IsMissing(Password) Then s = s & ";Password=" & Password
  ADOOracleConnectString = s
  
End Function

Public Function ADOAccess4ConnectString(ByVal dbPath As String, Optional ByVal SystemDBPath As Variant, Optional ByVal UserID As Variant, Optional ByVal UserPassword As Variant, Optional ByVal DBPassword As Variant) As String
  Dim s As String
  
  s = "PROVIDER=Microsoft.Jet.OLEDB.4.0"
  s = s & ";Data Source=" & dbPath
  If Not IsMissing(SystemDBPath) Then s = s & ";Jet OLEDB:System Database=" & SystemDBPath
  If Not IsMissing(DBPassword) Then s = s & ";Jet OLEDB:Database Password=" & DBPassword
  If Not IsMissing(UserID) Then s = s & ";User Id=" & UserID
  If Not IsMissing(UserPassword) Then s = s & ";Password=" & UserPassword
  ADOAccess4ConnectString = s
End Function

Public Function ADOSQLConnectString(ByVal Server As String, Optional ByVal UserID As Variant, Optional ByVal Password As Variant, Optional ByVal DatabaseName As Variant) As String
  Dim s As String
  
  s = "PROVIDER=SQLOLEDB"
  s = s & ";Data Source=" & Server
  If Not IsMissing(DatabaseName) Then s = s & ";Initial Catalog=" & DatabaseName
  If Not IsMissing(UserID) Then s = s & ";User ID=" & UserID
  If Not IsMissing(Password) Then s = s & ";Password=" & Password
  ADOSQLConnectString = s
End Function

Public Function ADOSQLConnectStringEx(ByVal Server As String, Optional ByVal UserID As Variant, Optional ByVal Password As Variant, Optional ByVal DatabaseName As Variant, Optional ByVal NetLibDLLName As Variant) As String
  Dim s As String
  
  s = "PROVIDER=SQLOLEDB"
  s = s & ";Data Source=" & Server
  If Not IsMissing(DatabaseName) Then s = s & ";Initial Catalog=" & DatabaseName
  If Not IsMissing(UserID) Then s = s & ";User ID=" & UserID
  If Not IsMissing(Password) Then s = s & ";Password=" & Password
  If Not IsMissing(NetLibDLLName) Then s = s & ";Network Library=" & NetLibDLLName
  ADOSQLConnectStringEx = s
End Function

Public Sub RefreshDAO(ByVal ADOcn As Connection)
  Dim JetDB As JetEngine
  
  Set JetDB = New JetEngine
  Call JetDB.RefreshCache(ADOcn)
End Sub

Public Sub ADOCopyRS(ByVal rsDst As Recordset, ByVal rsSrc As Recordset, Optional ByVal DeleteDestTableName As String = "", Optional ByVal IgnoreUpdateErrors As Boolean = False)
  Dim cn As ADODB.Connection
  Dim fld As Field, sErrorField As String
  
  On Error GoTo ADOCopyRS_err
  If rsSrc Is Nothing Then Err.Raise ERR_RS_IS_NOTHING, "ADOCopyRS", "The source recordset is nothing"
  If rsDst Is Nothing Then Err.Raise ERR_RS_IS_NOTHING, "ADOCopyRS", "The destination recordset is nothing"
    
  If Len(DeleteDestTableName) > 0 Then
    Set cn = rsDst.ActiveConnection ' could be disconnected
    If cn Is Nothing Then Err.Raise ERR_CONNECTION_IS_NOTHING, "ADOCopyRS", "Attempt to delete destination table failed since ActiveConnection is nothing"
    Call cn.Execute("Delete from " & DeleteDestTableName)
  End If
  If Not (rsSrc.BOF And rsSrc.EOF) Then Call rsSrc.MoveFirst
  Do While Not rsSrc.EOF
    rsDst.AddNew
    For Each fld In rsSrc.Fields
      rsDst.Fields(fld.Name).Value = fld.Value
    Next fld
    rsDst.Update
ADOCopyRS_next:
    rsSrc.MoveNext
  Loop
  Exit Sub
  
ADOCopyRS_err:
  sErrorField = ""
  If rsDst.EditMode <> adEditNone Then
    Call rsDst.CancelUpdate
    If IgnoreUpdateErrors Then Resume ADOCopyRS_next
  End If
  If Not fld Is Nothing Then
    If Len(fld.Name) > 0 Then sErrorField = "Error updating field [" & fld.Name & "]" & vbCrLf
  End If
  Err.Raise Err.Number, ErrorSource(Err, "ADOCopyRS"), sErrorField & Err.Description
End Sub

Private Function IsEmptyStringArray(v() As String) As Boolean
  IsEmptyStringArray = (UBound(v) - LBound(v)) < 0
End Function

Public Sub SaveADOAsFile(ByVal FileName As String, ByVal cn As ADODB.Connection, ByVal sql As String, ByVal PersistFormat As PersistFormatEnum)
  Dim rs As ADODB.Recordset

  On Error GoTo SaveADOAsFile_err
  Set rs = New ADODB.Recordset
  rs.Open sql, cn, adOpenStatic, adLockReadOnly, adCmdText
  rs.Save FileName, PersistFormat
  rs.Close
  Exit Sub

SaveADOAsFile_err:
  Set rs = Nothing
  Err.Raise Err.Number, ErrorSource(Err, "SaveADOAsFile"), "Unable to save file" & vbCrLf & Err.Description
End Sub

Public Sub ExportDBtoXML(ByVal FileName As String, ByVal DSNName As String, Optional ByVal IncludeFilter As String = "*", Optional ByVal ExcludeFilter As String)
  Dim cn As Connection
  Dim rsTables As Recordset
  Dim aIncludeFilter() As String, aExcludeFilter() As String
  Dim sql As String, sTempFileName As String, sConnString As String
  Dim sTableName As String, fRead As TCSFileread, buffer As String
  Dim qs As QString
  Dim hFile As Integer

  On Error GoTo ExportDBtoXML_err
  Call SetCursor
  If Len(IncludeFilter) > 0 Then aIncludeFilter = Split(IncludeFilter, ";")
  If Len(ExcludeFilter) > 0 Then aExcludeFilter = Split(ExcludeFilter, ";")
  If IsEmptyStringArray(aIncludeFilter) Then Err.Raise ERR_DBTOXML, "ConvertDBtoXML", "An include filter must be specified"
  
  Set cn = ADOConnect(DSNName, adUseClient)
  Set qs = New QString
  Set fRead = New TCSFileread
  
  Set rsTables = cn.OpenSchema(adSchemaTables)
  qs.Append XML_OPEN_COMMENT & "DSN=" & DSNName & " USER=" & GetNetUser & " DATE_TIME=" & Format$(Now(), "DD/MM/YYYY hh:nn:ss") & XML_CLOSE_COMMENT & vbCrLf
  Do Until rsTables.EOF
    If StrComp(rsTables!TABLE_TYPE, "TABLE", vbTextCompare) = 0 And _
      StrComp(Left$(rsTables!TABLE_NAME, 11), "MSysModules", vbBinaryCompare) <> 0 Then
      sTableName = rsTables!TABLE_NAME
      If Not Filter(sTableName, EXCLUDE_FILTER, aExcludeFilter) Then
        If Filter(sTableName, INCLUDE_FILTER, aIncludeFilter) Then
          sql = "SELECT * FROM " & sTableName
          sTempFileName = GetTempFileName("", "TMP")
        
          Call SaveADOAsFile(sTempFileName, cn, sql, adPersistXML)
          If Not fRead.OpenFile(sTempFileName) Then Err.Raise ERR_CONVERT_DB_TO_XML, "File open", "Error on opening file - " & sTempFileName
          qs.Append XML_OPEN_COMMENT & sTableName & XML_CLOSE_COMMENT & vbCrLf
          Call fRead.GetFile(buffer)
          qs.Append buffer
          qs.Append vbCrLf & XML_OPEN_COMMENT & sTableName & XML_CLOSE_COMMENT & vbCrLf
          fRead.CloseFile
          Call xKill(sTempFileName)
        End If
       End If
     End If
     rsTables.MoveNext
  Loop
  
ExportDBtoXML_end:
  hFile = FreeFile
  Open FileName For Output As #hFile
  If qs.Length > 0 Then
    Print #hFile, qs
  End If
  Close #hFile
  Call ClearCursor
  Exit Sub
  
ExportDBtoXML_err:
  ErrorMessagePush Err
    Call ClearCursor
  ErrorMessagePopErr Err
  Err.Raise Err.Number, ErrorSource(Err, "ConvertDBtoXML"), "Error in ConvertDBtoXML" & vbCrLf & Err.Description & vbCrLf
End Sub

Private Function Filter(ByVal TableName As String, ByVal FilterType As FILTER_TYPES, Filters() As String) As Boolean
  Dim i As Long
  
  On Error GoTo Filter_err
  For i = LBound(Filters) To UBound(Filters)
    If TableName Like Filters(i) Then
      Filter = True
      Exit For
    End If
  Next i
  Exit Function
  
Filter_err:
  Filter = False
End Function

