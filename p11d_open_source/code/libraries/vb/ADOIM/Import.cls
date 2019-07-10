VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ImportClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Option Compare Text
Implements ILibraryVersion

Public Enum IMPORT_UPDATE_TYPE
  [_NO_LOCK_UPDATES] = -1
  NO_UPDATES = 1
  UPDATE_ONLYFIRST = 2 ' Update first record matching the new record, if > 1 match Raise Error, if not present Addnew
  UPDATE_FIRST = 4      ' Update first record matching the new record, if not present Addnew
  UPDATE_ALL = 8        ' Update all records matching the new record, if not present Addnew
  UPDATEONLY_FLAG = 16  ' Do not AddNew if record is not present, no error
  UPDATEONLY_ERROR = UPDATEONLY_FLAG + 32 ' Do not AddNew if record is not present, Error raised
End Enum

Public Enum IMPORT_TYPE
  [_IMPORT_UNDEFINED] = 0
  IMPORT_FIXED
  IMPORT_DELIMITED
End Enum

Public Enum IMPORTFIELD_KEY
  [_KEY_INVALID] = 0
  KEY_FIELD
  KEY_SKIPFIELD
  KEY_BEGINMULTIPLE
  KEY_ENDMULTIPLE
  KEY_DATENOW
  KEY_FILENAME
  KEY_FILEPATH
  KEY_FILEDATE
  KEY_LINENUMBER
  KEY_CFGFILENAME
  KEY_USERNAME
  KEY_IMPDATE
  KEY_CONSTANT
  KEY_COPIEDFIELD   'MS
End Enum

Public Enum IMPORT_CONSTRAINT
  [_CONSTRAINT_INVALID] = 0
  CONSTRAINT_MASK = 2
  CONSTRAINT_MIN = 4
  CONSTRAINT_MAX = 8
  CONSTRAINT_FATAL = 16
  CONSTRAINT_INLIST = 32
  CONSTRAINT_EXCLUDELIST = 64
  CONSTRAINT_IGNORE_CONVERSIONERRORS = 128
End Enum

Public Enum IMPORTFIELD_ACTION
  [_ACTION_NOACTION] = 0
  ACTION_NEGATE
  ACTION_FACTOR
  ACTION_REFORMATFIXEDDATE = CONVERT_FIXEDDATE
  ACTION_REFORMATDELIMITED = CONVERT_DELIMITED
  ACTION_REFORMATFIXEDDATETIME = CONVERT_FIXEDDATETIME
  ACTION_REFORMATDATEGENERAL = CONVERT_DATETYPEUNKNOWN
End Enum

Private Enum IMPORT_DESTTYPE
  [_DEST_UNDEFINED] = 0
  DEST_RECORDSET
  DEST_ARRAY
End Enum

Private m_UseTransaction As Boolean
Private m_Connection As Connection

Private m_ImpParser As Parser
Private m_Icols As FieldSpecs
Private m_ErrorFilter As IErrorFilter
Private m_ProcessingLinesInError As Long
Private m_LinesOK As Long
Private m_UpdatedRecordCount As Long
Private m_AddedRecordCount As Long
Private m_AuditObj As IAudit
Private m_AuditNulls As Boolean

Private m_Header As String
Private m_HeaderCount As Long
Private m_Footer As String
Private m_FooterCount As Long
Private m_HeaderFooterOnly As Boolean
Private m_ImportFilePath As String
Private m_ImportFileName As String
Private m_ImportType As IMPORT_TYPE

Private m_UpdateType As IMPORT_UPDATE_TYPE

Private m_ImportFileDateTime As Date
Private m_ImportDateTime As Date
Private m_ImportCFGFilePath As String
Private m_ImportInitok As Boolean
Private m_ImpDate As Date
Private m_ImpUser As String
Private m_TestImport As Boolean
Private m_lAttemptedLines As Long

'* vars for delimited import
Private m_Delimiter As String       '* Field delimiter in delimited import
Private m_EscapeChar As String      '* Quote character in delimited imports
Private m_RightTrimFields As Boolean

'* vars for defaults
Private m_DefStr As Variant
Private m_DefDate As Variant
Private m_DefNumeric As Variant
Private m_DefFloat As Variant
Private m_DefBool As Variant
Private m_AbortImport As Boolean
Private m_TakeFieldNamesFromLine As Long
Private m_ImportName As String

Public Property Get ImportName() As String
  ImportName = m_ImportName
End Property

Friend Property Let ImportName(ByVal NewValue As String)
  m_ImportName = NewValue
End Property

Public Property Get UseTransaction() As Boolean
  UseTransaction = m_UseTransaction
End Property

Public Property Let UseTransaction(ByVal NewValue As Boolean)
  m_UseTransaction = NewValue
End Property

Public Property Get Connection() As Connection
  Set Connection = m_Connection
End Property

Public Property Set Connection(ByVal NewValue As Connection)
  Set m_Connection = NewValue
End Property

Public Property Get TakeFieldNamesFromLine() As Long
  TakeFieldNamesFromLine = m_TakeFieldNamesFromLine
End Property

Public Property Let TakeFieldNamesFromLine(ByVal NewValue As Long)
  m_TakeFieldNamesFromLine = NewValue
End Property

Public Property Get AttemptedLines() As Long
  AttemptedLines = m_lAttemptedLines
End Property
Public Property Let AttempedLineProcessed(ByVal NewValue As Long)
  AttempedLineProcessed = NewValue
End Property

' MPSPHD - Properties added so LoadSpec can access the columns
Friend Property Get Icols() As FieldSpecs
  Set Icols = m_Icols
End Property

Friend Property Set Icols(ByVal NewValue As FieldSpecs)
  Set m_Icols = NewValue
End Property

Public Property Get UpdatedRecordCount() As Long
  UpdatedRecordCount = m_UpdatedRecordCount
End Property

Public Property Let UpdatedRecordCount(ByVal NewValue As Long)
  m_UpdatedRecordCount = NewValue
End Property

Public Property Get AddedRecordCount() As Long
  AddedRecordCount = m_AddedRecordCount
End Property

Public Property Let AddedRecordCount(ByVal NewValue As Long)
  m_AddedRecordCount = NewValue
End Property

Public Property Get LinesOK() As Long
  LinesOK = m_LinesOK
End Property

Public Property Let LinesOK(ByVal NewValue As Long)
  m_LinesOK = NewValue
End Property

'* Note: ProcessingLinesInError = -1 if there was an error setting up the importer, otherwise
'* number of lines in the source file in error
Public Property Get ProcessingLinesInError() As Long
  ProcessingLinesInError = m_ProcessingLinesInError
End Property

Public Property Get DefaultString() As Variant
  DefaultString = m_DefStr
End Property

Public Property Let DefaultString(ByVal NewValue As Variant)
  m_DefStr = NewValue
End Property

Public Property Get DefaultDate() As Variant
  DefaultDate = m_DefDate
End Property

Public Property Let DefaultDate(ByVal NewValue As Variant)
  m_DefDate = NewValue
End Property

Public Property Get DefaultNumeric() As Variant
  DefaultNumeric = m_DefNumeric
End Property

Public Property Let DefaultNumeric(ByVal NewValue As Variant)
  m_DefNumeric = NewValue
End Property

Public Property Get DefaultFloat() As Variant
  DefaultFloat = m_DefFloat
End Property

Public Property Let DefaultFloat(ByVal NewValue As Variant)
  m_DefFloat = NewValue
End Property

Public Property Get DefaultBoolean() As Variant
  DefaultBoolean = m_DefBool
End Property

Public Property Let DefaultBoolean(ByVal NewValue As Variant)
  m_DefBool = NewValue
End Property
  
Public Property Let AuditNulls(ByVal NewValue As Boolean)
  m_AuditNulls = NewValue
End Property

Public Property Get AuditNulls() As Boolean
  AuditNulls = m_AuditNulls
End Property

Public Property Set Audit(ByVal NewObj As IAudit)
  Set m_AuditObj = NewObj
End Property

Public Property Get Audit() As IAudit
  Set Audit = m_AuditObj
End Property

Public Property Get Delimiter() As String
  Delimiter = m_Delimiter
End Property

Public Property Let Delimiter(ByVal NewValue As String)
  m_Delimiter = NewValue
End Property

Public Property Get EscapeChar() As String
  EscapeChar = m_EscapeChar
End Property

Public Property Let EscapeChar(ByVal NewValue As String)
  m_EscapeChar = NewValue
End Property

Public Property Let RightTrimFields(ByVal NewValue As Boolean)
  m_RightTrimFields = NewValue
End Property

Public Property Get RightTrimFields() As Boolean
  RightTrimFields = m_TestImport
End Property

Friend Property Let TestImport(ByVal NewValue As Boolean)
  m_TestImport = NewValue
End Property

Friend Property Get TestImport() As Boolean
  TestImport = m_TestImport
End Property

Public Property Get ErrorFilter() As IErrorFilter
  Set ErrorFilter = m_ErrorFilter
End Property

Public Property Set ErrorFilter(ByVal NewValue As IErrorFilter)
  Set m_ErrorFilter = NewValue
End Property

Public Property Get Header() As String
  Header = m_Header
End Property

Public Property Get Footer() As String
  Footer = m_Footer
End Property

Public Property Get HeaderCount() As Long
  HeaderCount = m_HeaderCount
End Property

Public Property Let HeaderCount(ByVal NewCount As Long)
  m_HeaderCount = NewCount
End Property

Public Property Get FooterCount() As Long
  FooterCount = m_FooterCount
End Property

Public Property Let FooterCount(ByVal NewCount As Long)
  m_FooterCount = NewCount
End Property

Public Property Get DebugHeaderFooter() As Boolean
  DebugHeaderFooter = m_HeaderFooterOnly
End Property

Public Property Let DebugHeaderFooter(ByVal bNewValue As Boolean)
  m_HeaderFooterOnly = bNewValue
End Property

Public Function NewColumn() As ImportSpec
  Set NewColumn = New ImportSpec
End Function

Public Function AddColumn(ByVal ispec As ImportSpec) As Boolean
  Call m_Icols.AddObject(ispec)
End Function

Public Property Get ImportType() As IMPORT_TYPE
  ImportType = m_ImportType
End Property

Public Property Let ImportType(ByVal NewValue As IMPORT_TYPE)
  m_ImportType = NewValue
End Property

Public Property Get UpdateType() As IMPORT_UPDATE_TYPE
  UpdateType = m_UpdateType
End Property

Public Property Let UpdateType(ByVal NewValue As IMPORT_UPDATE_TYPE)
  m_UpdateType = NewValue
  If (m_UpdateType <> NO_UPDATES) And _
     ((m_UpdateType And UPDATE_ONLYFIRST) <> UPDATE_ONLYFIRST) And _
     ((m_UpdateType And UPDATE_FIRST) <> UPDATE_FIRST) And _
     ((m_UpdateType And UPDATE_ALL) <> UPDATE_ALL) Then Call Err.Raise(380, "Let UpdateType")
End Property

Public Property Let AbortImport(ByVal NewValue As Boolean)
  m_AbortImport = NewValue
End Property

Public Property Get AbortImport() As Boolean
  AbortImport = m_AbortImport
End Property

Public Function ParseField(ByVal FieldSpec As String) As PARSESTATUS
  On Error GoTo ParseField_err
  ParseField = m_ImpParser.ParseLine(FieldSpec)
ParseField_end:
  Exit Function
  
ParseField_err:
  ParseField = PARSESTATUS_ABORTLINE
  Call ErrorMessage(ERR_ERROR, Err, "ImportParseField", "ERR_IMPORTPARSE", "There was an error initialising the Import for the file " & m_ImportFilePath & vbCr & "Unable to parse column " & FieldSpec)
  Resume ParseField_end
End Function

Public Sub RemoveParseField(ByVal DestFieldName As String, Optional ByVal RaiseErrors = True)
  Dim ispec As ImportSpec, i As Long
  
  For i = 1 To m_Icols.Count
    Set ispec = m_Icols(i)
    If StrComp(ispec.DestField, DestFieldName, vbTextCompare) = 0 Then
      Call m_Icols.Remove(i)
      Exit For
    End If
  Next i
  If RaiseErrors Then Err.Raise ERR_IMPORT, "RemoveParseField", "Unable to find field [" & DestFieldName & "] to remove"
End Sub

Public Sub ParseImportSpec(ByVal CFGFilename As String, Optional ByVal ParsePreProcess As IParsePreProcess)
  
  On Error GoTo ParseImportSpec_err
  Set m_Icols = Nothing
  Set m_Icols = New FieldSpecs
  Call m_ImpParser.ParseReset
  Set m_ImpParser.ParsePreProcess = ParsePreProcess
  Set m_ImpParser.ParseSettings.ParseParameter(1) = m_Icols
  Set m_ImpParser.ParseSettings.ParseParameter(2) = Me
  Set m_ImpParser.ParseSettings.ParseParameter(3) = m_ImportWizard
  Call ResetImportValues
    
  If Len(CFGFilename) > 0 Then
    m_ImportCFGFilePath = CFGFilename
    If m_ImpParser.ParseFile(CFGFilename) = PARSESTATUS_ABORT Then Err.Raise ERR_IMPORTINIT, "ImportInit", "Could not initialise import - Parse failed"
  End If
  
ParseImportSpec_end:
  Exit Sub
ParseImportSpec_err:
  Err.Raise Err.Number, ErrorSource(Err, "ParseImportSpec"), Err.Description
End Sub
Public Function InitImport(ByVal ImportFile As String, Optional ByVal ImportType As IMPORT_TYPE = [_IMPORT_UNDEFINED], Optional ByVal CFGFilename As String = "", Optional ByVal UpdateType As IMPORT_UPDATE_TYPE = NO_UPDATES, Optional ByVal ParsePreProcess As IParsePreProcess) As Boolean
  Dim s As String, sFile As String, sExt As String
  
On Error GoTo InitImport_err
  Call xSet("InitImport")
  m_ImportInitok = False
  m_ImportFilePath = ImportFile
  Call SplitPath(m_ImportFilePath, , sFile, sExt)
  m_ImportFileName = sFile & sExt
  m_ImportFileDateTime = FileDateTime(m_ImportFilePath)
  m_ImportDateTime = Now
  If ImportType <> [_IMPORT_UNDEFINED] Then m_ImportType = ImportType
  m_UpdateType = UpdateType
  
  ' reset import
  Call ParseImportSpec(CFGFilename, ParsePreProcess)
  
  m_LinesOK = 0: m_ProcessingLinesInError = 0
  m_AbortImport = False
  m_TestImport = False
  m_ImportInitok = True
  Call ResetIgnoreError

  
InitImport_end:
  InitImport = m_ImportInitok
  Call xReturn("InitImport")
  Exit Function
  
InitImport_err:
  m_ImportInitok = False
  If Len(CFGFilename) > 0 Then
    s = vbCr & "Configuration file: " & CFGFilename
  Else
    s = vbCr & "No Configuration file specified"
  End If
  If ErrorFilterOnly() Then
    Call ErrorMessage(ERR_ERROR, Err, "InitImport", ErrorFilterDescription(), "There was an error initialising the Import for the file " & m_ImportFilePath & s)
  Else
    Call ErrorMessage(ERR_ERROR, Err, "InitImport", "Initialising Import", "There was an error initialising the Import for the file " & m_ImportFilePath & s)
  End If
  Resume InitImport_end
End Function

Public Property Get ImportWizard() As ImportWizard
  If m_ImportWizard Is Nothing Then
    Set m_ImportWizard = New ImportWizard
    Set m_ImportWizard.ImpParent = Me
  End If
  If Not m_ImportWizard.ImpParent Is Me Then Call Err.Raise(ERR_IMPORTWIZARD, "ImportWizard", "Import wizard is already initialised by another Importer")
  Set ImportWizard = m_ImportWizard
End Property

Public Property Set ImportWizard(ByVal NewValue As ImportWizard)
  Static InSet As Boolean
  On Error Resume Next
  If InSet Then Exit Property
  InSet = True
  If NewValue Is Nothing Then
    Call KillWizard
  Else
    Set m_ImportWizard = NewValue
  End If
  InSet = False
End Property

'* Imports a file into a database given an open database, table name, CFG file and import file
'* Returns the number of errors encountered.
'* Can use tansactions so a failure during import will result in no lines being imported
'* Can pass a progress bar control for visual info.
'* uses as log database if one is defined LogDatabase
'* return value of <0 indicates fatal error where no lines were imported
Public Function ImportFile(ByVal rs As Recordset, Optional ByVal PBarObj As Object = Nothing) As Long
  
  ' use following settings
  ' lastmodified does not exist
  Set m_Connection = rs.ActiveConnection
  If m_Connection Is Nothing Then
    If rs.LockType <> adLockBatchOptimistic Then
      Call ErrorMessage(ERR_ERROR, Err, "ImportFile", "ImportFile", "The recordset to import into is disconnected, but its lock type is not batch optimistic.")
      Exit Function
    Else
      m_UseTransaction = False
    End If
  End If
  If (rs.CursorLocation <> adUseClient) Or (rs.CursorType <> adOpenStatic) Then
    Call ErrorMessage(ERR_ERROR, Err, "ImportFile", "ImportFile", "The recordset to import into must have a client-side cursor, and be a static recordset.")
  Else
    ImportFile = ImportAny("Import File to recordset", DEST_RECORDSET, PBarObj, , rs)
  End If
End Function

Public Function ImportLines(Values() As Variant, ByVal LinesToImport As Long) As Long
  ImportLines = ImportAny("Import Linese to array", DEST_ARRAY, Nothing, LinesToImport, , Values)
End Function

Private Function ImportAny(ByVal ImportType As String, ByVal Dest As IMPORT_DESTTYPE, ByVal PB As Object, Optional ByVal MaxImportLine As Long = &H7FFFFFFF, Optional ByVal rs As Recordset, Optional Values As Variant) As Long
  Dim FileNextIdle As Long
  Dim ImportFileName As String
  Dim cols() As ImportSpec, MaxCols As Long
  Dim rf As TCSFileread
  Dim buffer As String, linenum As Long, s As String
  Dim FilePos As Long
  Dim InMulti As Boolean, i As Long, lineindex As Long
  Dim OldErrorFilter As IErrorFilter
  Dim RecordsUpdated As Long, RecordsAdded As Long
  ' correct back CD/MS Dim tPrg As TCSProgressBar
  Dim tPrg As Object
   
On Error GoTo ImportAny_err
  Call xSet("ImportAny")
  Call SetCursor
  Set OldErrorFilter = ErrorFilter
  m_LinesOK = 0: m_ProcessingLinesInError = 0: m_lAttemptedLines = 0: m_AddedRecordCount = 0: m_UpdatedRecordCount = 0
  
  FileNextIdle = (8192& * 8)
  If Not m_ImportInitok Then Err.Raise ERR_IMPORT, ImportType, "Import not initialised correctly"
  If m_Icols Is Nothing Then Err.Raise ERR_IMPORT, ImportType, "Import specification invalid - Columns definition does not exist"
  If m_Icols.Count = 0 Then Err.Raise ERR_IMPORT, ImportType, "Import specification invalid - No columns defined"
  Select Case Dest
    Case DEST_RECORDSET
              If rs Is Nothing Then Err.Raise ERR_IMPORT, ImportType, "RS Destination invalid"
              If Not rs.Supports(adUpdate) Then Err.Raise ERR_IMPORT, ImportType, "Destination recordset is not updateable"
    Case DEST_ARRAY
              If IsMissing(Values) Then Err.Raise ERR_IMPORT, ImportType, "Values Destination invalid"
              If (MaxImportLine = &H7FFFFFFF) Or (MaxImportLine <= 0) Then Err.Raise ERR_IMPORT, ImportType, "Values Import must specify number of lines to import"
    Case Else
              Err.Raise ERR_IMPORT, ImportType, "Destination not specified"
  End Select
  Call CheckImportColumns(Dest = DEST_ARRAY, True)
  MaxCols = PrepareImport(rs, ImportFileName, cols)
  
  'Set up progress bar with values and find step value
  Set rf = New TCSFileread
  If Not rf.OpenFile(m_ImportFilePath) Then Err.Raise ERR_IMPORT, ImportType, "Unable to open file " & m_ImportFilePath
  If m_FooterCount > 0 Then
    If Dest = DEST_ARRAY Then
      MaxImportLine = MaxImportLine - m_FooterCount
    Else
      MaxImportLine = rf.LineCount - m_FooterCount
    End If
  End If
  If MaxImportLine <= 0 Then Err.Raise ERR_IMPORT, ImportType, "Import has no lines to import"
  
  If Dest = DEST_ARRAY Then
    ReDim Values(1 To MaxImportLine, 1 To MaxCols) As Variant
    MaxImportLine = MaxImportLine + m_HeaderCount
  End If
  If Dest = DEST_RECORDSET Then
    'MPSFEB
    'Call DBEngine.SetOption(dbMaxBufferSize, 10 * 1024)
    'Call DBEngine.SetOption(dbUserCommitSync, "No")
    If m_UseTransaction Then Call TransactionADO(m_Connection, TRANS_BEGIN)
  End If
    
  If Not PB Is Nothing Then
    Set tPrg = PB
    tPrg.Min = 0
    tPrg.Value = 0
    tPrg.Max = rf.LineCount
    If Len(tPrg.Caption) = 0 Then tPrg.Caption = "Importing Line"
    tPrg.Indicator = ValueOfMax
  
  End If
  
  If Not m_ErrorFilter Is Nothing And Not m_TestImport Then Set ErrorFilter = m_ErrorFilter
  linenum = 0: lineindex = 0
  Do While rf.GetLine(buffer)
    linenum = linenum + 1
    If linenum <= m_HeaderCount Then
      If linenum = 1 Then
        m_Header = buffer
      Else
        m_Header = m_Header & vbCrLf & buffer
      End If
      GoTo NextLine
    End If
    If linenum > MaxImportLine Then
      If Dest = DEST_ARRAY Then Exit Do
      If linenum = (MaxImportLine + 1) Then
        m_Footer = buffer
      Else
        m_Footer = m_Footer & vbCrLf & buffer
      End If
      GoTo NextLine
    End If
    If m_HeaderFooterOnly Then GoTo NextLine
    If m_AbortImport Then
      m_ProcessingLinesInError = -2
      GoTo ImportAny_end
    End If
    
    InMulti = False
process_multi:
    
    If Not ImportSingleLine(buffer, InMulti, MaxCols, cols, linenum) Then GoTo NextLine
    Select Case Dest
      Case DEST_RECORDSET
                If AddNewRecord(RecordsUpdated, RecordsAdded, rs, cols, MaxCols, linenum) Then
                  m_UpdatedRecordCount = m_UpdatedRecordCount + RecordsUpdated
                  m_AddedRecordCount = m_AddedRecordCount + RecordsAdded
                  m_LinesOK = m_LinesOK + RecordsAdded
                Else
                  GoTo NextLine
                End If
      Case DEST_ARRAY
                lineindex = lineindex + 1
                For i = 1 To MaxCols
                  Values(lineindex, i) = cols(i).Value
                Next i
      Case Else
                Err.Raise ERR_IMPORT, ImportType, "Destination invalid - not able to add record"
    End Select
    
    If InMulti Then GoTo process_multi
    
    FilePos = rf.CurrentPos
    If FilePos > FileNextIdle Then
      If (Dest = DEST_RECORDSET) And IsJetDSN(m_Connection.Provider) Then Call RefreshDAO(m_Connection)  'DBEngine.Idle
      FileNextIdle = FilePos + (8192& * 8)
    End If
    
NextLine:
    If Not tPrg Is Nothing Then Call tPrg.Step
    If InMulti Then
      InMulti = False
      For i = 1 To MaxCols
        cols(i).multioffset = 0&
      Next i
    End If
    If m_ProcessingLinesInError < 0 Then GoTo ImportAny_end
  Loop
 
ImportAny_end:
  If m_ProcessingLinesInError = 0 Then
    'apf cd errors on commit
    If m_UseTransaction Then Call TransactionADO(m_Connection, TRANS_COMMIT)
  Else
    If m_UseTransaction Then
      m_AddedRecordCount = 0
      m_UpdatedRecordCount = 0
      Call TransactionADO(m_Connection, TRANS_ROLLBACK)
    End If
    If (m_ProcessingLinesInError <> -1) And Not m_TestImport Then
      If m_AbortImport Then
        m_ProcessingLinesInError = 0
        s = vbCrLf & "Import Aborted" & vbCrLf
      End If
      If Not ErrorFilterOnly Then Call ErrorMessage(ERR_ERROR, Err, "Importing File", "Import File", "There was an error importing." & s & " Error Count=" & CStr(m_ProcessingLinesInError))
    End If
  End If
  Set ErrorFilter = OldErrorFilter
  If Not tPrg Is Nothing Then tPrg.Value = tPrg.Max
  Call ClearCursor
  Call xReturn("ImportAny")
  ImportAny = m_ProcessingLinesInError
  Exit Function
  
ImportAny_err:
  m_ProcessingLinesInError = -1
  If ErrorFilterOnly() Then
    Call ErrorMessage(ERR_ERROR, Err, "ImportAny", ErrorFilterDescription(), "There was an error setting up the import. Import aborted")
  Else
    Call ErrorMessage(ERR_ERROR, Err, "ImportAny", "Import File initialisation", "There was an error setting up the import. Import aborted")
  End If
  Resume ImportAny_end
  Resume
End Function


'* Raises as error if Invalid Import Columns used
Public Function CheckImportColumns(ByVal Preview As Boolean, ByVal RaiseError As Boolean) As Boolean
  Dim i As Long
  Dim ColumnList As New StringList
  Dim ispec As ImportSpec
  
  On Error GoTo CheckImportColumns_err
  If (m_UpdateType <> NO_UPDATES) And _
     ((m_UpdateType And UPDATE_ONLYFIRST) <> UPDATE_ONLYFIRST) And _
     ((m_UpdateType And UPDATE_FIRST) <> UPDATE_FIRST) And _
     ((m_UpdateType And UPDATE_ALL) <> UPDATE_ALL) Then Call Err.Raise(ERR_INVALID_UPDATETYPE, "CheckImportColumns", "Update type is not valid")

  ' NOTE: Array import cannot use UpdateType
  If m_UpdateType <> NO_UPDATES Then
    If Preview Then Call Err.Raise(ERR_PREVIEW_UPDATES, "CheckImportColumns", "Updates are not possible in preview mode")
    ' Check that fields with primary key exist!!
    For i = 1 To m_Icols.Count
      Set ispec = m_Icols(i)
      If ispec.KeyField And (Len(ispec.DestField) = 0) Then Call Err.Raise(ERR_PRIMARY_DEST, "CheckImportColumns", "Primary Key field " & ispec.FldName & " (" & CStr(i) & ") is not being imported")
    Next i
  End If
  
  For i = 1 To m_Icols.Count
    Set ispec = m_Icols(i)
    If (Len(ispec.DestField) > 0) Then
      If ColumnList.IsPresent(ispec.DestField) Then Err.Raise ERR_IMPORT, "CheckImportColumns", "Invalid import specification the Destination field [" & ispec.DestField & "] can only have one import column associated with it"
      Call ColumnList.Add(ispec.DestField)
    End If
  Next i
  
  For i = 1 To m_Icols.Count
    Set ispec = m_Icols(i)
    If Not Preview And (Len(ispec.DestField) = 0) Then
      If Not ((ispec.key = KEY_SKIPFIELD) Or _
              (ispec.key = KEY_BEGINMULTIPLE) Or _
              (ispec.key = KEY_ENDMULTIPLE)) Then Err.Raise ERR_IMPORT, "CheckImportCols", "Import Column specification invalid" & vbCr & "Column " & CStr(i) & vbCr & "No destination field name supplied"
    End If
    If ispec.key = [_KEY_INVALID] Then Err.Raise ERR_IMPORT, "CheckImportCols", "Import Column  specification invalid" & vbCr & "Column " & CStr(i) & " Destination Field=" & ispec.DestField & vbCr & "Column has invalid Import key field"
    If (ispec.vartype = [_TYPE_INVALID]) And _
       (Not ((ispec.key = KEY_CONSTANT) Or _
             (ispec.key = KEY_COPIEDFIELD) Or _
             (ispec.key = KEY_SKIPFIELD) Or _
             (ispec.key = KEY_BEGINMULTIPLE) Or _
             (ispec.key = KEY_ENDMULTIPLE))) Then Err.Raise ERR_IMPORT, "CheckImportCols", "Import Column  specification invalid" & vbCr & "Column " & CStr(i) & " Destination Field=" & ispec.DestField & vbCr & "Column has invalid import variable type"
  Next i
  CheckImportColumns = True
  
CheckImportColumns_end:
  Exit Function
  
CheckImportColumns_err:
  CheckImportColumns = False
  If RaiseError Then Err.Raise Err.Number, Err.Source, Err.Description
  Resume CheckImportColumns_end
End Function

Private Function ImportSingleLine(ByVal buffer As String, ByVal InMulti As Boolean, ByVal MaxCols As Long, cols() As ImportSpec, ByVal linenum As Long) As Boolean
  Dim multioffset As Long, BeginStart As Long
  Dim i As Long, j As Long, Offset As Long, Maxline As Long
  Dim ispec As ImportSpec, tmpISpec As ImportSpec, sItem As String, DestField As String
  Dim d0 As Date, f As Double, f1 As Double

  On Error GoTo ImportSingleLine_err
  
  ImportSingleLine = False: Offset = 1
  If Len(buffer) = 0 Then Exit Function
  m_lAttemptedLines = m_lAttemptedLines + 1
  
  For i = 1 To MaxCols
    Set ispec = cols(i)
    ispec.Value = ispec.DefaultValue
    Select Case ispec.key
      Case KEY_FIELD, KEY_SKIPFIELD
        If m_ImportType = IMPORT_DELIMITED Then
          Offset = GetDelimitedValue(sItem, buffer, Offset, m_Delimiter, m_EscapeChar)
          If m_RightTrimFields Then sItem = RTrim$(sItem)
        Else
          If (ispec.Offset + ispec.multioffset) <= 0 Then Err.Raise ERR_IMPORTFW, "ImportSingleLine", "Fixed width imports cannot have offsets < 1"
          sItem = Trim$(Mid$(buffer, ispec.Offset + ispec.multioffset, ispec.Length))
        End If
        If Len(sItem) > 0 Then
          If Not IsEmpty(ispec.DefaultNonNullValue) Then ispec.Value = ispec.DefaultNonNullValue
          Select Case ispec.vartype
            Case TYPE_STR
              If (ispec.Constraint And CONSTRAINT_MASK) = CONSTRAINT_MASK Then
                If Not (sItem Like ispec.Mask) Then
                  Call Err.Raise(ERR_IMPORTMASK, "ImportFile", "Parse item '" & sItem & "' on line " & CStr(linenum) & vbCrLf & "Could not import line as parse value does not match the mask '" & ispec.Mask & "'")
                End If
              End If
              If (ispec.Constraint And CONSTRAINT_INLIST) = CONSTRAINT_INLIST Then
                If Not ispec.ConstraintList.IsPresent(sItem) Then Call Err.Raise(ERR_IMPORTMASK, "ImportFile", "Parse item '" & sItem & "' on line " & CStr(linenum) & vbCrLf & "Could not import line as parse value is not within the list of allowed values")
              End If
              If (ispec.Constraint And CONSTRAINT_EXCLUDELIST) = CONSTRAINT_EXCLUDELIST Then
                If ispec.ConstraintList.IsPresent(sItem) Then Call Err.Raise(ERR_IMPORTMASK, "ImportFile", "Parse item '" & sItem & "' on line " & CStr(linenum) & vbCrLf & "Could not import line as parse value is within the list of forbidden values")
              End If
              If ispec.key = KEY_FIELD Then ispec.Value = sItem
            Case TYPE_LONG, TYPE_DOUBLE
              If Right$(sItem, 1) = "-" Then sItem = "-" & Left$(sItem, Len(sItem) - 1)
              If Not IsNumeric(sItem) Then
                If (ispec.Constraint And CONSTRAINT_IGNORE_CONVERSIONERRORS) <> CONSTRAINT_IGNORE_CONVERSIONERRORS Then Call Err.Raise(ERR_IMPORTMASK, "ImportFile", "Parse item '" & sItem & "' on line " & CStr(linenum) & vbCrLf & "Could not import line as parse value is not a number")
              Else
                f = CDbl(sItem)
                If ispec.vartype = TYPE_LONG Then
                  f1 = RoundN(f, 0)
                  If (f1 <> f) And ((ispec.Constraint And CONSTRAINT_IGNORE_CONVERSIONERRORS) <> CONSTRAINT_IGNORE_CONVERSIONERRORS) Then Call Err.Raise(ERR_IMPORTMASK, "ImportFile", "Parse item '" & sItem & "' on line " & CStr(linenum) & vbCrLf & "Could not import line as parse value is not a whole number")
                  f = f1
                End If
                If ispec.varaction = ACTION_NEGATE Then f = f * -1
                If ispec.varaction = ACTION_FACTOR Then f = f * CDbl(ispec.actionparam)
                If (ispec.Constraint And CONSTRAINT_MIN) = CONSTRAINT_MIN Then
                  If f < ispec.minvalue Then Call Err.Raise(ERR_IMPORTMASK, "ImportFile", "Parse item '" & sItem & "' on line " & CStr(linenum) & vbCrLf & "Could not import line as parse value is less than the minimum value allowed " & CStr(ispec.minvalue))
                End If
                If (ispec.Constraint And CONSTRAINT_MAX) = CONSTRAINT_MAX Then
                  If f > ispec.maxvalue Then Call Err.Raise(ERR_IMPORTMASK, "ImportFile", "Parse item '" & sItem & "' on line " & CStr(linenum) & vbCrLf & "Could not import line as parse value is greater than the maximum value allowed " & CStr(ispec.maxvalue))
                End If
                If ispec.vartype = TYPE_LONG Then
                  If ispec.key = KEY_FIELD Then ispec.Value = CLng(f)
                Else
                  If ispec.key = KEY_FIELD Then ispec.Value = f
                End If
              End If
            Case TYPE_DATE
              If ispec.varaction = ACTION_REFORMATDELIMITED Then
                d0 = ConvertDate(sItem, ispec.varaction, ispec.actionparam, ispec.DateDelimiter, ispec.TimeDelimiter)
              Else
                d0 = ConvertDate(sItem, ispec.varaction, ispec.actionparam)
              End If
              If (ispec.Constraint And CONSTRAINT_MIN) = CONSTRAINT_MIN Then
                If d0 < ispec.minvalue Then Call Err.Raise(ERR_IMPORTMASK, "ImportFile", "Parse item '" & sItem & "' on line " & CStr(linenum) & vbCrLf & "Could not import line as parse value is less than the minimum date allowed " & Format$(ispec.minvalue, "dd/mm/yyyy"))
              End If
              If (ispec.Constraint And CONSTRAINT_MAX) = CONSTRAINT_MAX Then
                If d0 > ispec.maxvalue Then Call Err.Raise(ERR_IMPORTMASK, "ImportFile", "Parse item '" & sItem & "' on line " & CStr(linenum) & vbCrLf & "Could not import line as parse value is greater than the maximum date allowed " & Format$(ispec.maxvalue, "dd/mm/yyyy"))
              End If
              If ispec.key = KEY_FIELD Then ispec.Value = d0
            Case TYPE_BOOL
              If (Len(ispec.BooleanTrue) = 0) Or (Len(ispec.BooleanFalse) = 0) Then
                If Not IsBoolean(sItem) Then Call Err.Raise(ERR_PARSEBOOLEAN, "ImportFile", "Could not parse " & sItem & " as type boolean")
                If ispec.key = KEY_FIELD Then ispec.Value = CBoolean(sItem)
              Else
                If StrComp(sItem, ispec.BooleanTrue, vbTextCompare) = 0 Then
                  If ispec.key = KEY_FIELD Then ispec.Value = True
                ElseIf StrComp(sItem, ispec.BooleanFalse, vbTextCompare) = 0 Then
                  If ispec.key = KEY_FIELD Then ispec.Value = False
                Else
                  Call Err.Raise(ERR_PARSEBOOLEAN, "ImportFile", "Could not parse " & sItem & " as type boolean")
                End If
              End If
            Case Else
              If ispec.key = KEY_FIELD Then Call ECASE("ImportFile - case invalid")
          End Select
        ElseIf InMulti Then
          Exit Function
        End If
        If (ispec.key = KEY_FIELD) And ispec.Required Then
          If IsNull(ispec.Value) Then Call Err.Raise(ERR_REQUIREDFIELD, "ImportFile", "Could not import line as required field '" & ispec.DestField & "' is Null")
        End If
      Case KEY_BEGINMULTIPLE
        InMulti = True
        multioffset = ispec.Offset
        BeginStart = i + 1
      Case KEY_ENDMULTIPLE
        If Not InMulti Then Call Err.Raise(ERR_IMPORT, "ImportFile", "Cannot do ENDMULTIPLE without BEGINMULTIPLE")
        For j = BeginStart To MaxCols
          Set tmpISpec = cols(j)
          If tmpISpec.key = KEY_ENDMULTIPLE Then Exit For
          tmpISpec.multioffset = tmpISpec.multioffset + multioffset
        Next j
      Case KEY_DATENOW 'Inserts the current date & time
        ispec.Value = m_ImportDateTime
      Case KEY_FILEDATE 'Inserts the import filedate
        ispec.Value = m_ImportFileDateTime
      Case KEY_FILENAME 'Inserts the import filename
        ispec.Value = m_ImportFileName
      Case KEY_FILEPATH 'Inserts the import filename full path
        ispec.Value = m_ImportFilePath
      Case KEY_CFGFILENAME 'Imports the configuration filename
        ispec.Value = m_ImportCFGFilePath
      Case KEY_LINENUMBER
        ispec.Value = linenum
      Case KEY_USERNAME
        ispec.Value = m_ImpUser
      Case KEY_IMPDATE
        ispec.Value = m_ImpDate
      Case KEY_CONSTANT
        ispec.Value = ispec.actionparam ' treat as per default value
      Case KEY_COPIEDFIELD
        ispec.Value = ispec.CopyField.Value
      Case Else
        Call ECASE("Unknown Import Key type")
    End Select
  Next i
  ImportSingleLine = True
    
ImportSingleLine_end:
  Exit Function
  
ImportSingleLine_err:
  m_ProcessingLinesInError = m_ProcessingLinesInError + 1
  ImportSingleLine = False
  DestField = ""
  If Not (ispec Is Nothing) Then
    DestField = ispec.DestField
    If (ispec.Constraint And CONSTRAINT_FATAL) = CONSTRAINT_FATAL Then
      m_ProcessingLinesInError = -1
      If ErrorFilterOnly() Then
        Call ErrorMessage(ERR_ERROR, Err, "ImportSingleLine", ErrorFilterDescription(sItem, linenum, i, DestField, buffer), "Fatal Error")
      Else
        Call ErrorMessage(ERR_ERROR, Err, "ImportSingleLine", "Import error on line: " & CStr(linenum), "There was a Fatal constraint error importing the file " & m_ImportFilePath & vbCr & "Current import value: " & sItem)
      End If
      Resume ImportSingleLine_end
    End If
  End If
  If ErrorFilterOnly() Then
    Call ErrorMessage(ERR_ERROR + ERR_ALLOWIGNORE, Err, "ImportSingleLine", ErrorFilterDescription(sItem, linenum, i, DestField, buffer), Err.Description)
  Else
    If ErrorMessage(ERR_ERROR + ERR_ALLOWIGNORE + ERR_ALLOWCANCEL, Err, "ImportSingleLine", "Import error on line: " & CStr(linenum), "There was an error importing the file " & m_ImportFilePath & vbCr) Then
      m_ProcessingLinesInError = -1
    End If
  End If
  Resume ImportSingleLine_end
  Resume
End Function

Private Function PrepareImport(ByVal rs As Recordset, ByVal ImportFileName As String, cols() As ImportSpec) As Long
  Dim i As Long, ispec As ImportSpec
  Dim sExt As String
  
  m_Header = "": m_Footer = ""
  m_ImpDate = Now
  m_ImpUser = GetNetUser(False)
  ReDim cols(1 To m_Icols.Count) As ImportSpec
  For i = 1 To m_Icols.Count
    Set cols(i) = m_Icols.Item(i)
    Set ispec = cols(i)
    If IsEmpty(ispec.DefaultValue) Then
      ispec.DefaultValue = Null
      Select Case ispec.vartype
        Case TYPE_STR
          If Not IsNull(m_DefStr) Then
            ispec.DefaultValue = CStr(m_DefStr)
          End If
        Case TYPE_LONG
          If Not IsNull(m_DefNumeric) Then
            ispec.DefaultValue = CLng(m_DefNumeric)
          End If
        Case TYPE_DOUBLE
          If Not IsNull(m_DefFloat) Then
            ispec.DefaultValue = CDbl(m_DefFloat)
          End If
        Case TYPE_DATE
          If Not IsNull(m_DefDate) Then
            ispec.DefaultValue = CDate(m_DefDate)
          End If
        Case TYPE_BOOL
          If Not IsNull(m_DefBool) Then
            ispec.DefaultValue = CBool(m_DefBool)
          End If
      End Select
    End If
    If ispec.key = KEY_CONSTANT Then ispec.actionparam = GetTypedValue(ispec.actionparam, ispec.vartype)
    If Not (rs Is Nothing) Then
      If Not ((cols(i).key = KEY_SKIPFIELD) Or (cols(i).key = KEY_BEGINMULTIPLE) Or (cols(i).key = KEY_ENDMULTIPLE) Or (cols(i).key = KEY_COPIEDFIELD)) Then
        If (Not rs Is Nothing) And (Not m_Connection Is Nothing) Then cols(i).DestTable = m_Connection.Properties("Data Source").Value
      End If
    End If
  Next i
  Call SplitPath(m_ImportFilePath, , ImportFileName, sExt)
  ImportFileName = ImportFileName & sExt
  PrepareImport = m_Icols.Count
End Function

Private Sub Class_Initialize()
  Set m_ImpParser = New Parser
  Set m_ImpParser.ParseSettings = New SpecParseSettings
  Call m_ImpParser.AddParseItem(New ParseBeginMultiple)
  Call m_ImpParser.AddParseItem(New ParseBoolean)
  Call m_ImpParser.AddParseItem(New ParseBooleanFalse)
  Call m_ImpParser.AddParseItem(New ParseBooleanTrue)
  Call m_ImpParser.AddParseItem(New ParseConstant)
  Call m_ImpParser.AddParseItem(New ParseCopyField)
  Call m_ImpParser.AddParseItem(New ParseDataSource)
  Call m_ImpParser.AddParseItem(New ParseDate)
  Call m_ImpParser.AddParseItem(New ParseDateDelimiter)
  Call m_ImpParser.AddParseItem(New ParseDateNow) 'MPSMarch2
  Call m_ImpParser.AddParseItem(New ParseDefault)
  Call m_ImpParser.AddParseItem(New ParseDefaultNonNull)
  Call m_ImpParser.AddParseItem(New ParseDelimitedImport)
  Call m_ImpParser.AddParseItem(New ParseDelimiter)
  Call m_ImpParser.AddParseItem(New ParseEndMultiple)
  Call m_ImpParser.AddParseItem(New ParseEscapeChar)
  Call m_ImpParser.AddParseItem(New ParseExList)
  Call m_ImpParser.AddParseItem(New ParseFactor)
  Call m_ImpParser.AddParseItem(New ParseFatal)
  Call m_ImpParser.AddParseItem(New ParseField)
  Call m_ImpParser.AddParseItem(New ParseFieldName)
  Call m_ImpParser.AddParseItem(New ParseFieldNames)
  Call m_ImpParser.AddParseItem(New ParseFileDate)
  Call m_ImpParser.AddParseItem(New ParseFileName)
  Call m_ImpParser.AddParseItem(New ParseFilePath)
  Call m_ImpParser.AddParseItem(New ParseFixedDate)
  Call m_ImpParser.AddParseItem(New ParseFixedDateTime)
  Call m_ImpParser.AddParseItem(New ParseFixedWidthImport)
  Call m_ImpParser.AddParseItem(New ParseFloat)
  Call m_ImpParser.AddParseItem(New ParseFooter)
  Call m_ImpParser.AddParseItem(New ParseHeader)
  Call m_ImpParser.AddParseItem(New ParseHideField)
  Call m_ImpParser.AddParseItem(New ParseIgnoreNumConvErrs)
  Call m_ImpParser.AddParseItem(New ParseInList)
  Call m_ImpParser.AddParseItem(New ParseImpDate)
  Call m_ImpParser.AddParseItem(New ParseKeyField)
  Call m_ImpParser.AddParseItem(New ParseLineNumber)
  Call m_ImpParser.AddParseItem(New ParseMask)
  Call m_ImpParser.AddParseItem(New ParseMaxValue)
  Call m_ImpParser.AddParseItem(New ParseMinValue)
  Call m_ImpParser.AddParseItem(New ParseNegate)
  Call m_ImpParser.AddParseItem(New ParseNumber)
  Call m_ImpParser.AddParseItem(New ParseReqField)
  Call m_ImpParser.AddParseItem(New ParseRightTrimFields)
  Call m_ImpParser.AddParseItem(New ParseSkipColumn)
  Call m_ImpParser.AddParseItem(New ParseSpecFileName)  'MPSMarch2
  Call m_ImpParser.AddParseItem(New ParseString)
  Call m_ImpParser.AddParseItem(New ParseTimeDelimiter)
  Call m_ImpParser.AddParseItem(New ParseUsername)
  m_Delimiter = ","
  m_EscapeChar = """"
  m_AuditNulls = False
  Call ResetImportValues
  Call RegisterLibrary(Me)
End Sub

Private Sub ResetImportValues()
  m_ImportInitok = False
  m_DefStr = Null
  m_DefDate = Null
  m_DefNumeric = Null
  m_DefFloat = Null
  m_DefBool = Null
  m_LinesOK = 0
  m_ProcessingLinesInError = 0
  m_ImportName = ""
End Sub

Private Sub Class_Terminate()
  Call KillImporter
End Sub

Private Sub KillWizard()
  If Not m_ImportWizard Is Nothing Then
    Call m_ImportWizard.KillWizard
  End If
  Set m_ImportWizard = Nothing
End Sub

Public Sub KillImporter()
  Call KillWizard
  Set m_AuditObj = Nothing
  If Not m_ImpParser Is Nothing Then
    Set m_ImpParser.ParsePreProcess = Nothing
    Set m_ImpParser = Nothing
  End If
  Set m_Icols = Nothing
End Sub

Private Function AddNewRecord(RecordsUpdated As Long, RecordsAdded As Long, ByVal rs As Recordset, cols() As ImportSpec, ByVal MaxCols As Long, ByVal linenum As Long) As Boolean
  Dim ispec As ImportSpec
  Dim i As Long, sRecordValues  As String
  Dim Criteria As String, OldFilter As Variant
  
  On Error GoTo AddNewRecord_err
  RecordsUpdated = 0: RecordsAdded = 0
  If m_UpdateType <> NO_UPDATES Then ' Find matching record
    For i = 1 To MaxCols
      Set ispec = cols(i)
      If ispec.KeyField Then
        If Len(Criteria) > 0 Then Criteria = Criteria & " AND "
        Criteria = Criteria & addcriteriaequal(ispec)
      End If
    Next i
    If Len(Criteria) = 0 Then Err.Raise ERR_IMPORT_NOCRITERIA, "AddNewRecord", "Cannot do update import with no Criteria set"
    OldFilter = rs.Filter
    rs.Filter = Criteria
  
NextUpdateRecord:
    If RecordsUpdated > 0 Then rs.MoveNext
        
    If rs.EOF Then
      If RecordsUpdated > 0 Then GoTo AddNewRecord_end
      If (m_UpdateType And UPDATEONLY_FLAG) = UPDATEONLY_FLAG Then
        If (m_UpdateType And UPDATEONLY_ERROR) = UPDATEONLY_ERROR Then Call Err.Raise(ERR_IMPORTUPDATEONLY, "AddNewRecord", "No records match the update key.")
        GoTo AddNewRecord_end
      End If
    Else
      If (RecordsUpdated > 0) And _
         ((m_UpdateType And UPDATE_ONLYFIRST) = UPDATE_ONLYFIRST) Then Call Err.Raise(ERR_UPDATEPRIMARYONLY, "AddNewRecord", "More than one record matches the primary key.")
      RecordsUpdated = RecordsUpdated + 1
    End If
  End If
  
  If (RecordsUpdated = 0) Or (m_UpdateType = NO_UPDATES) Then
    rs.AddNew
    RecordsAdded = RecordsAdded + 1
  End If
  
  For i = 1 To MaxCols
    Set ispec = cols(i)
    If Not ispec.key = KEY_SKIPFIELD Then
      If (Not IsValueEqual(rs.Fields(ispec.DestField).Value, ispec.Value)) And Not (IsNull(ispec.Value) And (RecordsAdded > 0)) Then
        rs.Fields(ispec.DestField).Value = ispec.Value
      End If
      If (Not m_AuditObj Is Nothing) And (m_AuditNulls Or Not IsNull(ispec.Value)) Then
        Call m_AuditObj.WriteAudit("" & ispec.Value, ispec.DestField, ispec.DestTable)
      End If
    End If
  Next i
  rs.Update
  If (RecordsUpdated > 0) And _
     (((m_UpdateType And UPDATE_ONLYFIRST) = UPDATE_ONLYFIRST) Or _
      ((m_UpdateType And UPDATE_ALL) = UPDATE_ALL)) Then GoTo NextUpdateRecord
  
AddNewRecord_end:
  If (m_UpdateType <> NO_UPDATES) And Not IsEmpty(OldFilter) Then rs.Filter = OldFilter
  AddNewRecord = (RecordsUpdated > 0) Or (RecordsAdded > 0)
  Exit Function
  
AddNewRecord_err:
  Call ErrorMessagePush(Err)
  If Not (rs.EOF Or rs.BOF) Then
    If (rs.EditMode = adEditAdd) Or (rs.EditMode = adEditInProgress) Then
      sRecordValues = DebugRecord(rs)
      rs.CancelUpdate
    End If
  End If
  If Not ((Err.Number = ERR_IMPORTUPDATEONLY) Or (Err.Number = ERR_UPDATEPRIMARYONLY)) Then
    If RecordsUpdated > 0 Then RecordsUpdated = RecordsUpdated - 1
    If RecordsAdded > 0 Then RecordsAdded = RecordsAdded - 1
  End If
  Call ErrorMessagePopErr(Err)
  If Err.Number = ERR_IMPORT_NOCRITERIA Then m_AbortImport = True
  If ErrorFilterOnly() Then
    Err.Description = "Error adding new record."
    Call ErrorMessage(ERR_ERROR + ERR_ALLOWIGNORE, Err, "AddNewRecord", ErrorFilterDescription(, linenum), Err.Description)
  Else
    Call ErrorMessage(ERR_ERROR + ERR_ALLOWIGNORE, Err, "AddNewRecord", "Import error on line: " & CStr(linenum), "There was an error importing the file " & m_ImportFilePath & vbCrLf & vbCrLf & "Unable to add new record" & vbCrLf & sRecordValues)
  End If
  m_ProcessingLinesInError = m_ProcessingLinesInError + 1
  Resume AddNewRecord_end
  Resume
End Function


Private Property Get ErrorFilterOnly() As Boolean
  ErrorFilterOnly = (Not m_TestImport) And (Not ErrorFilter Is Nothing)
End Property

Private Property Get ErrorFilterDescription(Optional ByVal Item As Variant, Optional ByVal linenum As Variant, Optional ByVal colnum As Variant, Optional ByVal DestField As Variant, Optional ByVal Line As Variant) As String
  If Not IsMissing(Item) Then ErrorFilterDescription = CStr(Item)
  ErrorFilterDescription = ErrorFilterDescription & ","
  If Not IsMissing(linenum) Then ErrorFilterDescription = ErrorFilterDescription & CStr(linenum)
  ErrorFilterDescription = ErrorFilterDescription & ","
  If Not IsMissing(colnum) Then ErrorFilterDescription = ErrorFilterDescription & CStr(colnum)
  ErrorFilterDescription = ErrorFilterDescription & ","
  If Not IsMissing(DestField) Then ErrorFilterDescription = ErrorFilterDescription & CStr(DestField)
  ErrorFilterDescription = ErrorFilterDescription & "="
  If Not IsMissing(Line) Then ErrorFilterDescription = ErrorFilterDescription & CStr(Line)
End Property

Private Property Get ILibraryVersion_Name() As String
  ILibraryVersion_Name = "Import Library"
End Property

Private Property Get ILibraryVersion_Version() As String
  ILibraryVersion_Version = App.Major & "." & App.Minor & "." & App.Revision
End Property
