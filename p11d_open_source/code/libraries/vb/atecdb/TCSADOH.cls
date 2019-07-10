VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ADOHelper"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Enum FILTER_TYPES
  INCLUDE_FILTER = 0
  EXCLUDE_FILTER
End Enum
Private Const SQLDB_DSN As String = ";Initial Catalog="
Private Const SQLSERVER_DSN As String = ";Data Source="
Private Const SQLUSERID_DSN As String = ";User ID="
Private Const SQLPASSWORD_DSN As String = ";Password="

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
  IsFieldPresentADO = InCollection(cFields, FieldName)
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
    Case Else ' adBigInt , adarray, adbyref, adDBDate,  adEmpty, adError, adGUID, adIDispatch, adIUnknown, adUnsignedBigInt, adUserDefined, adVariant, advector
      Err.Raise ERR_ADODB, "ADOtoDatatype", "Unrecognised data Type: [" & CStr(ADOType) & "]"
  End Select
End Function

Public Function ADOConnect(ByVal ConnectString As String, Optional ByVal CursorLocation As ADODB.CursorLocationEnum = adUseServer) As ADODB.Connection
  On Error GoTo ADOConnect_err
  
  Set ADOConnect = New Connection
  ADOConnect.CursorLocation = CursorLocation
  ADOConnect.Open ConnectString
  'apf no obvious difference !!! SQL 7.0
  ' ADOConnect.Properties("Enable Fastload").Value = True
  Exit Function
  
ADOConnect_err:
  Set ADOConnect = Nothing
  Err.Raise Err.Number, "ADOConnect", "Unable to open the connection" & vbCrLf & ConnectString & vbCrLf & Err.Description
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
  If Not IsMissing(DatabaseName) Then s = s & SQLDB_DSN & DatabaseName
  If Not IsMissing(UserID) Then s = s & SQLUSERID_DSN & UserID
  If Not IsMissing(Password) Then s = s & SQLPASSWORD_DSN & Password
  ADOSQLConnectString = s
End Function

Public Function ADOSQLConnectStringEx(ByVal Server As String, Optional ByVal UserID As Variant, Optional ByVal Password As Variant, Optional ByVal DatabaseName As Variant, Optional ByVal NetLibDLLName As Variant) As String
  Dim s As String
  
  s = "PROVIDER=SQLOLEDB"
  s = s & ";Data Source=" & Server
  If Not IsMissing(DatabaseName) Then s = s & SQLDB_DSN & DatabaseName
  If Not IsMissing(UserID) Then s = s & SQLUSERID_DSN & UserID
  If Not IsMissing(Password) Then s = s & SQLPASSWORD_DSN & Password
  If Not IsMissing(NetLibDLLName) Then s = s & ";Network Library=" & NetLibDLLName
  ADOSQLConnectStringEx = s
End Function

Public Function SQLDatabaseName(ByVal DataDSN As String) As String
  Dim p0 As Long, p1 As Long
   
  p0 = InStr(1, DataDSN, SQLDB_DSN, vbTextCompare)
  If p0 > 0 Then
    p0 = p0 + Len(SQLDB_DSN)
    p1 = InStr(p0, DataDSN, ";")
    If p1 = 0 Then p1 = Len(DataDSN) + 1
    SQLDatabaseName = Mid$(DataDSN, p0, p1 - p0)
  End If
End Function

Private Function SQLConnectReplace(ByVal DataDSN As String, ByVal DSNAttribute As String, ByVal NewValue As String, Optional ByVal ReplaceOnly As Boolean = False) As String
  Dim p0 As Long, p1 As Long
  Dim OldValue As String
  
  p0 = InStr(1, DataDSN, DSNAttribute, vbTextCompare)
  NewValue = DSNAttribute & NewValue
  If p0 > 0 Then
    p1 = InStr(p0 + 1, DataDSN, ";")
    If p1 = 0 Then p1 = Len(DataDSN) + 1
    OldValue = Mid$(DataDSN, p0, p1 - p0)
    SQLConnectReplace = Replace(DataDSN, OldValue, NewValue, , 1, vbTextCompare)
  Else
    If ReplaceOnly Then
      SQLConnectReplace = DataDSN
    Else
      SQLConnectReplace = DataDSN & NewValue
    End If
  End If
End Function

Public Function ADOSQLAnonymousDSN(ByVal DataDSN As String) As String
  DataDSN = SQLConnectReplace(DataDSN, SQLUSERID_DSN, "*****", True)
  ADOSQLAnonymousDSN = SQLConnectReplace(DataDSN, SQLPASSWORD_DSN, "*****", True)
End Function

Public Function ADOSQLConnectReplace(ByVal DataDSN As String, Optional ByVal Server As Variant, Optional ByVal UserID As Variant, Optional ByVal Password As Variant) As String
  Dim DSN As String
  
  DSN = DataDSN
  If Not IsMissing(Server) Then
    DSN = SQLConnectReplace(DSN, SQLSERVER_DSN, Server)
  End If
  If Not IsMissing(UserID) Then
    DSN = SQLConnectReplace(DSN, SQLUSERID_DSN, UserID)
  End If
  If Not IsMissing(Password) Then
    DSN = SQLConnectReplace(DSN, SQLPASSWORD_DSN, Password)
  End If
  ADOSQLConnectReplace = DSN
End Function

Private Function IsEmptyStringArray(v() As String) As Boolean
  IsEmptyStringArray = (UBound(v) - LBound(v)) < 0
End Function

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

