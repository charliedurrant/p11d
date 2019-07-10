VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
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

Public Enum IMPORT_TYPE
  IMPORT_UNDEFINED = 0
  IMPORT_FIXED
  IMPORT_DELIMITED
End Enum

Public Enum IMPORTFIELD_KEY
  KEY_INVALID = 0
  KEY_FIELD
  KEY_SKIPFIELD
  KEY_BEGINMULTIPLE
  KEY_ENDMULTIPLE
  KEY_DATENOW
  KEY_FILENAME
  KEY_FILEDATE
  KEY_LINENUMBER
  KEY_CFGFILENAME
  KEY_USERNAME
  KEY_IMPDATE
  KEY_CONSTANT
  KEY_COPIEDFIELD   'MS
End Enum

Public Enum IMPORTFIELD_TYPE
  TYPE_INVALID = 0
  TYPE_STR
  TYPE_LONG
  TYPE_DOUBLE
  TYPE_DATE
  TYPE_BOOL
End Enum

Public Enum IMPORT_CONSTRAINT
  CONSTRAINT_INVALID = 0
  CONSTRAINT_MASK = 2
  CONSTRAINT_MIN = 4
  CONSTRAINT_MAX = 8
  CONSTRAINT_FATAL = 16
End Enum

Public Enum IMPORTFIELD_ACTION
  ACTION_NOACTION = 0
  ACTION_NEGATE
  ACTION_FACTOR
  ACTION_REFORMATFIXEDDATE = CONVERT_FIXEDDATE
  ACTION_REFORMATDELIMITED = CONVERT_DELIMITED
  ACTION_REFORMATFIXEDDATETIME = CONVERT_FIXEDDATETIME
End Enum

Private Enum IMPORT_DESTTYPE
  DEST_UNDEFINED = 0
  DEST_RDO
  DEST_RECORDSET
  DEST_ARRAY
End Enum

Private ImpParser As Parser
Private m_Icols As FieldSpecs
Private m_WorkSpace As WorkSpace
Private m_rdoConnection As RDOConnection
Private m_AuditObj As IAudit

Private m_Header As String
Private m_HeaderCount As Long
Private m_Footer As String
Private m_FooterCount As Long
Private m_HeaderFooterOnly As Boolean
Private m_ImportFilePath As String
Private m_ImportType  As IMPORT_TYPE
Private m_ImportFileDateTime As Date
Private m_ImportDateTime As Date
Private m_ImportCFGFilePath As String
Private m_ImportInitok As Boolean
Private m_ImpDate As Date
Private m_ImpUser As String

'* vars for delimited import
Private m_Delimiter As String       '* Field delimiter in delimited import
Private m_EscapeChar As String      '* Quote character in delimited imports

'* vars for defaults
Private m_DefStr As Variant
Private m_DefDate As Variant
Private m_DefNumeric As Variant
Private m_DefBool As Variant

'* Vars for ErrorFilter
Private m_ErrorFilter As IErrorFilter

' MPSPHD - Properties added so LoadSpec can access the columns
Friend Property Get Icols() As FieldSpecs
  Set Icols = m_Icols
End Property

Friend Property Set Icols(NewValue As FieldSpecs)
  Set m_Icols = NewValue
End Property

Public Property Get DefaultString() As String
  DefaultString = m_DefStr
End Property

Public Property Let DefaultString(NewValue As String)
  m_DefStr = NewValue
End Property

Public Property Get DefaultDate() As Date
  DefaultDate = m_DefDate
End Property

Public Property Let DefaultDate(ByVal NewValue As Date)
  m_DefDate = NewValue
End Property

Public Property Get DefaultNumeric() As Double
  DefaultNumeric = m_DefNumeric
End Property

Public Property Let DefaultNumeric(NewValue As Double)
  m_DefNumeric = NewValue
End Property

Public Property Get DefaultBoolean() As Boolean
  DefaultBoolean = m_DefBool
End Property

Public Property Let DefaultBoolean(ByVal NewValue As Boolean)
  m_DefBool = NewValue
End Property

Public Property Set Audit(NewObj As IAudit)
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

Public Property Get ErrorFilter() As IErrorFilter
  Set ErrorFilter = m_ErrorFilter
End Property

Public Property Set ErrorFilter(NewValue As IErrorFilter)
  Set m_ErrorFilter = NewValue
End Property

Public Property Get WorkSpace() As WorkSpace
  Set WorkSpace = m_WorkSpace
End Property

Public Property Set WorkSpace(NewValue As WorkSpace)
  Set m_WorkSpace = NewValue
End Property

Public Property Get RDOConnection() As RDOConnection
  Set RDOConnection = m_rdoConnection
End Property

Public Property Set RDOConnection(NewValue As RDOConnection)
  Set m_rdoConnection = NewValue
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

Public Function AddColumn(ispec As ImportSpec) As Boolean
  Call m_Icols.AddObject(ispec)
End Function

Public Property Get ImportType() As IMPORT_TYPE
  ImportType = m_ImportType
End Property

Public Property Let ImportType(ByVal NewValue As IMPORT_TYPE)
  m_ImportType = NewValue
End Property

Public Function ParseField(FieldSpec As String) As PARSESTATUS
  On Error GoTo ParseField_err
  
  ParseField = ImpParser.ParseLine(FieldSpec)
  
ParseField_end:
  Exit Function
  
ParseField_err:
  ParseField = PARSESTATUS_ABORTLINE
  Call ErrorMessage(ERR_ERROR, Err, "ImportParseField", "ERR_IMPORTPARSE", "There was an error initialising the Import for the file " & m_ImportFilePath & vbCr & "Unable to parse column " & FieldSpec)
  Resume ParseField_end
End Function

Public Function InitImport(ImportFile As String, Optional ImportType As IMPORT_TYPE = IMPORT_UNDEFINED, Optional CFGFilename As String = "") As Boolean
  Dim s As String
  
On Error GoTo InitImport_err
  Call xSet("InitImport")
  m_ImportInitok = False
  m_ImportFilePath = ImportFile
  m_ImportFileDateTime = FileDateTime(m_ImportFilePath)
  m_ImportDateTime = Now
  If ImportType <> IMPORT_UNDEFINED Then m_ImportType = ImportType
 
  ' reset import
  Set m_Icols = Nothing
  Set m_Icols = New FieldSpecs
  Call ImpParser.ParseReset
  Set ImpParser.ParseSettings.ParseParameter(1) = m_Icols
  Set ImpParser.ParseSettings.ParseParameter(2) = Me
  
  If Len(CFGFilename) > 0 Then
    m_ImportCFGFilePath = CFGFilename
    If ImpParser.ParseFile(CFGFilename) = PARSESTATUS_ABORT Then Call Err.Raise(ERR_IMPORTINIT, "ImportInit", "Could not initialise import - Parse failed")
  End If
  m_ImportInitok = True
  
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
  Call ErrorMessage(ERR_ERROR, Err, "ImportInit", "Initialising Import", "There was an error initialising the Import for the file " & m_ImportFilePath & s)
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

Public Property Set ImportWizard(NewValue As ImportWizard)
  Set m_ImportWizard = NewValue
End Property

'* Imports a file into a database given an open database, table name, CFG file and import file
'* Returns the number of errors encountered.
'* Can use tansactions so a failure during import will result in no lines being imported
'* Can pass a progress bar control for visual info.
'* uses as log database if one is defined LogDatabase
'* return value of <0 indicates fatal error where no lines were imported
Public Function ImportFile(rs As Recordset, Optional PBarObj As Object = Nothing) As Long
  ImportFile = ImportAny("Import File to recordset", DEST_RECORDSET, PBarObj, , rs)
End Function

Public Function ImportLines(Values() As Variant, ByVal LinesToImport As Long, Optional RaiseErrors As Boolean = False) As Long
  ImportLines = ImportAny("Import Linese to array", DEST_ARRAY, Nothing, LinesToImport, , , Values)
End Function

Public Function ImportFileRDO(rsRDO As rdoResultset, Optional PBarObj As Object = Nothing) As Long
  ImportFileRDO = ImportAny("Import File to RDO resultset", DEST_RDO, PBarObj, , , rsRDO)
End Function

Private Function ImportAny(ImportType As String, ByVal Dest As IMPORT_DESTTYPE, PB As ProgressBar, Optional ByVal MaxImportLine As Long = &H7FFFFFFF, Optional rs As Recordset, Optional rsRDO As rdoResultset, Optional Values As Variant) As Long
  Dim FileNextIdle As Long, ErrorCount As Long
  Dim ImportFileName As String
  Dim cols() As ImportSpec, MaxCols As Long
  Dim rf As CFileSeqRead
  Dim buffer As String, linenum As Long
  Dim PBInterVal As Double, PBNewVal As Long, FilePos As Long, InTrans As Boolean
  Dim InMulti As Boolean, i As Long, lineindex As Long
  Dim OldErrorFilter As IErrorFilter
  Const PBMAX As Long = 100
   
On Error GoTo ImportAny_err
  Call xSet("ImportAny")
  Call SetCursor
  Set OldErrorFilter = tcscoredll.ErrorFilter
  If Not m_ErrorFilter Is Nothing Then Set tcscoredll.ErrorFilter = m_ErrorFilter
  ErrorCount = 0:   FileNextIdle = (8192& * 8)
  If Not m_ImportInitok Then Err.Raise ERR_IMPORT, ImportType, "Import not initialised correctly"
  If m_Icols Is Nothing Then Err.Raise ERR_IMPORT, ImportType, "Import specification invalid - Columns definition does not exist"
  If m_Icols.Count = 0 Then Err.Raise ERR_IMPORT, ImportType, "Import specification invalid - No columns defined"
  Select Case Dest
    Case DEST_RDO
              If rsRDO Is Nothing Then Err.Raise ERR_IMPORT, ImportType, "RDO Destination invalid"
              If Not rsRDO.Updatable Then Err.Raise ERR_IMPORT, ImportType, "RDO Destination resultset is not updateable"
    Case DEST_RECORDSET
              If rs Is Nothing Then Err.Raise ERR_IMPORT, ImportType, "RS Destination invalid"
              If Not rs.Updatable Then Err.Raise ERR_IMPORT, ImportType, "Destination recordset is not updateable"
    Case DEST_ARRAY
              If IsMissing(Values) Then Err.Raise ERR_IMPORT, ImportType, "Values Destination invalid"
              If (MaxImportLine = &H7FFFFFFF) Or (MaxImportLine <= 0) Then Err.Raise ERR_IMPORT, ImportType, "Values Import must specify number of lines to import"
    Case Else
              Err.Raise ERR_IMPORT, ImportType, "Destination not specified"
  End Select
  Call CheckImportColumns(Dest = DEST_ARRAY, True)
  MaxCols = PrepareImport(rs, rsRDO, ImportFileName, cols)
  
  'Set up progress bar with values and find step value
  Set rf = New CFileSeqRead
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
    Call DBEngine.SetOption(dbMaxBufferSize, 10 * 1024)
    Call DBEngine.SetOption(dbUserCommitSync, "No")
    If Not m_WorkSpace Is Nothing Then
      m_WorkSpace.BeginTrans
      InTrans = True
    End If
  End If
  If Dest = DEST_RDO Then
    ' rdoengine/rdoconnetion properties
    If Not m_rdoConnection Is Nothing Then
      m_rdoConnection.BeginTrans
      InTrans = True
    End If
  End If
    
  If Not PB Is Nothing Then
    PB.Min = 0
    PB.Value = 0
    PB.Max = PBMAX
    PBInterVal = rf.Length / CDbl(PBMAX)
  End If
  
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
          
    InMulti = False
process_multi:
    If Not ImportSingleLine(buffer, InMulti, ErrorCount, MaxCols, cols, linenum) Then GoTo NextLine
    Select Case Dest
      Case DEST_RDO, DEST_RECORDSET
                If Not AddNewRecord(rs, rsRDO, cols, MaxCols, linenum) Then GoTo NextLine
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
    If Not PB Is Nothing Then
      PBNewVal = FilePos / PBInterVal
      If PBNewVal > PB.Value Then
        If Not (PB Is Nothing) Then PB.Value = PBNewVal
      End If
    End If
    If FilePos > FileNextIdle Then
      If Dest = DEST_RECORDSET Then DBEngine.Idle
      FileNextIdle = FilePos + (8192& * 8)
    End If
    
NextLine:
    If InMulti Then
      InMulti = False
      For i = 1 To MaxCols
        cols(i).multioffset = 0&
      Next i
    End If
    If ErrorCount < 0 Then GoTo ImportAny_end
  Loop
 
ImportAny_end:
  If ErrorCount = 0 Then
    If InTrans Then
      If Not m_WorkSpace Is Nothing Then m_WorkSpace.CommitTrans
      If Not m_rdoConnection Is Nothing Then m_rdoConnection.CommitTrans
    End If
  Else
    If InTrans Then
      If Not m_WorkSpace Is Nothing Then m_WorkSpace.Rollback
      If Not m_rdoConnection Is Nothing Then m_rdoConnection.RollbackTrans
    End If
    If ErrorCount <> -1 Then Call ErrorMessage(ERR_ERROR, Err, "Importing File", "Import File", "There was an error importing. Error Count=" & CStr(ErrorCount))
  End If
  Set tcscoredll.ErrorFilter = OldErrorFilter
  Call ClearCursor
  Call xReturn("ImportAny")
  ImportAny = ErrorCount
  Exit Function
  
ImportAny_err:
  ErrorCount = -1
  Call ErrorMessage(ERR_ERROR, Err, "Importing File", "Import File initialisation", "There was an error setting up the import. Import aborted")
  Resume ImportAny_end
End Function


'* Raises as error if Invalid Import Columns used
Public Function CheckImportColumns(Preview As Boolean, RaiseError As Boolean) As Boolean
  Dim i As Long
  Dim ispec As ImportSpec
  
  On Error GoTo CheckImportColumns_err
  
  For i = 1 To m_Icols.Count
    Set ispec = m_Icols(i)
    If Not Preview And (Len(ispec.DestField) = 0) Then
      If Not ((ispec.key = KEY_SKIPFIELD) Or _
              (ispec.key = KEY_BEGINMULTIPLE) Or _
              (ispec.key = KEY_ENDMULTIPLE)) Then
         Err.Raise ERR_IMPORT, "CheckImportCols", "Import Column specification invalid" & vbCr & "Column " & CStr(i) & vbCr & "No destination field name supplied"
       End If
    End If
    If ispec.key = KEY_INVALID Then Err.Raise ERR_IMPORT, "CheckImportCols", "Import Column  specification invalid" & vbCr & "Column " & CStr(i) & " Destination Field=" & ispec.DestField & vbCr & "Column has invalid Import key field"
    If (ispec.vartype = TYPE_INVALID) And _
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

Private Function ImportSingleLine(buffer As String, InMulti As Boolean, ErrorCount As Long, ByVal MaxCols As Long, cols() As ImportSpec, ByVal linenum As Long, Optional RaiseErrors As Boolean = False) As Boolean
  Static multioffset As Long
  Static BeginStart As Long
  Dim Starti As Long, i As Long, j As Long, Offset As Long, Maxline As Long
  Dim ispec As ImportSpec, tmpISpec As ImportSpec, sValue As String
  Dim d0 As Date, f As Double

On Error GoTo ImportSingleLine_err
  ImportSingleLine = False: Offset = 1
  If Len(buffer) = 0 Then Exit Function
  Starti = 1
  If InMulti Then Starti = BeginStart
  
  For i = 1 To MaxCols
    Set ispec = cols(i)
    Select Case ispec.vartype
      Case TYPE_STR
        ispec.Value = m_DefStr
      Case TYPE_LONG, TYPE_DOUBLE
        ispec.Value = m_DefNumeric
      Case TYPE_DATE
        ispec.Value = m_DefDate
      Case TYPE_BOOL
        ispec.Value = m_DefBool
      Case Else
        If Not ((ispec.key = KEY_CONSTANT) Or (ispec.key = KEY_COPIEDFIELD)) Then Call ECASE("Unknown Import Key type")
        ispec.Value = Null
    End Select
    Select Case ispec.key
      Case KEY_FIELD, KEY_SKIPFIELD
        If m_ImportType = IMPORT_DELIMITED Then
          Offset = GetDelimitedValue(sValue, buffer, Offset, m_Delimiter, m_EscapeChar)
        Else
          sValue = Trim$(Mid$(buffer, ispec.Offset + ispec.multioffset, ispec.Length))
        End If
        If Len(sValue) > 0 Then
          Select Case ispec.vartype
            Case TYPE_STR
              ispec.Value = m_DefStr
              If (ispec.constraint And CONSTRAINT_MASK) = CONSTRAINT_MASK Then
                If Not (sValue Like ispec.mask) Then Call Err.Raise(ERR_IMPORTMASK, "ImportFile", "Could not import line as parse value '" & sValue & "' does not match the mask '" & ispec.mask & "'")
              End If
              If ispec.key = KEY_FIELD Then ispec.Value = sValue
'              If Not ispec.MergeField Is Nothing Then
'                If ispec.MergeType = MERGE_INSERT Then
'                  ispec.MergeField.Value = ispec.Value & ispec.MergeSeparator & ispec.MergeField.Value
'                ElseIf ispec.MergeType = MERGE_APPEND Then
'                  ispec.MergeField.Value = ispec.MergeField.Value & ispec.MergeSeparator & ispec.Value
'                End If
'              End If
            Case TYPE_LONG, TYPE_DOUBLE
              f = 0#
              If Right$(sValue, 1) = "-" Then sValue = "-" & Left$(sValue, Len(sValue) - 1)
              If IsNumeric(sValue) Then f = CDbl(sValue)
              If ispec.varaction = ACTION_NEGATE Then f = f * -1
              If ispec.varaction = ACTION_FACTOR Then f = f * CDbl(ispec.actionparam)
              If (ispec.constraint And CONSTRAINT_MIN) = CONSTRAINT_MIN Then
                If f < ispec.minvalue Then Call Err.Raise(ERR_IMPORTMASK, "ImportFile", "Could not import line as parse value '" & sValue & "' is less than the minimum value allowed " & CStr(ispec.minvalue))
              End If
              If (ispec.constraint And CONSTRAINT_MAX) = CONSTRAINT_MAX Then
                If f < ispec.maxvalue Then Call Err.Raise(ERR_IMPORTMASK, "ImportFile", "Could not import line as parse value '" & sValue & "' is greater than the maximum value allowed " & CStr(ispec.maxvalue))
              End If
              If ispec.vartype = TYPE_LONG Then
                If ispec.key = KEY_FIELD Then ispec.Value = CLng(f)
              Else
                If ispec.key = KEY_FIELD Then ispec.Value = f
              End If
            Case TYPE_DATE
              If ispec.varaction = ACTION_REFORMATDELIMITED Then
                d0 = ConvertDate(sValue, ispec.varaction, ispec.actionparam, ispec.DateDelimiter, ispec.TimeDelimiter)
              Else
                d0 = ConvertDate(sValue, ispec.varaction, ispec.actionparam)
              End If
              If (ispec.constraint And CONSTRAINT_MIN) = CONSTRAINT_MIN Then
                If d0 < ispec.minvalue Then Call Err.Raise(ERR_IMPORTMASK, "ImportFile", "Could not import line as parse value '" & Format$(d0, "dd/mm/yyyy") & "' is less than the minimum date allowed " & Format$(ispec.minvalue, "dd/mm/yyyy"))
              End If
              If (ispec.constraint And CONSTRAINT_MAX) = CONSTRAINT_MAX Then
                If d0 > ispec.maxvalue Then Call Err.Raise(ERR_IMPORTMASK, "ImportFile", "Could not import line as parse value '" & Format$(d0, "dd/mm/yyyy") & "' is greater than the maximum date allowed " & Format$(ispec.maxvalue, "dd/mm/yyyy"))
              End If
              If ispec.key = KEY_FIELD Then ispec.Value = d0
            Case TYPE_BOOL
              If ispec.BooleanTrue = "" Or ispec.BooleanFalse = "" Then
                If (StrComp(sValue, "YES", vbTextCompare) = 0) Or (StrComp(sValue, "TRUE", vbTextCompare) = 0) Then
                  If ispec.key = KEY_FIELD Then ispec.Value = True
                ElseIf (StrComp(sValue, "NO", vbTextCompare) = 0) Or (StrComp(sValue, "FALSE", vbTextCompare) = 0) Then
                  If ispec.key = KEY_FIELD Then ispec.Value = False
                Else
                  Call Err.Raise(ERR_PARSEBOOLEAN, "ImportFile", "Could not parse " & sValue & " as type boolean")
                End If
              Else
                If StrComp(sValue, ispec.BooleanTrue, vbTextCompare) = 0 Then
                  If ispec.key = KEY_FIELD Then ispec.Value = True
                ElseIf StrComp(sValue, ispec.BooleanFalse, vbTextCompare) = 0 Then
                  If ispec.key = KEY_FIELD Then ispec.Value = False
                Else
                  Call Err.Raise(ERR_PARSEBOOLEAN, "ImportFile", "Could not parse " & sValue & " as type boolean")
                End If
              End If
            Case Else
              If ispec.key = KEY_FIELD Then Call ECASE("ImportFile - case invalid")
          End Select
        ElseIf InMulti Then
          Exit Function
        End If
        
      Case KEY_BEGINMULTIPLE
        InMulti = True
        multioffset = ispec.Offset
        BeginStart = i
        
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
        ispec.Value = ispec.actionparam
      Case KEY_COPIEDFIELD
        ispec.Value = ispec.CopyField.Value
      Case Else
        Call ECASE("Unknown Import Key type")
    End Select
    If ispec.Required And IsNull(ispec.Value) Then Call Err.Raise(ERR_REQUIREDFIELD, Err, "Required Field Error", "A required field cannot contain a null value")
  Next i
  ImportSingleLine = True
  
ImportSingleLine_end:
  Exit Function
  
ImportSingleLine_err:
  ErrorCount = ErrorCount + 1
  ImportSingleLine = False
  If Not (ispec Is Nothing) Then
    If (ispec.constraint And CONSTRAINT_FATAL) = CONSTRAINT_FATAL Then
      ErrorCount = -1
      If RaiseErrors Then Err.Raise Err, Err.Source, Err.Description
      Call ErrorMessage(ERR_ERROR, Err, "Importing File", "Import error on line: " & CStr(linenum), "There was a Fatal error importing the file " & m_ImportFilePath & vbCr & "Current import value: " & sValue)
      Resume ImportSingleLine_end
    End If
  End If
  If RaiseErrors Then Err.Raise Err, Err.Source, Err.Description
  Call ErrorMessage(ERR_ERROR, Err, "Importing File", "Import error on line: " & CStr(linenum), "There was an error importing the file " & m_ImportFilePath & vbCr & "Current import value: " & sValue)
  Resume ImportSingleLine_end
End Function

Private Function PrepareImport(rs As Recordset, rsRDO As rdoResultset, ImportFileName As String, cols() As ImportSpec) As Long
  Dim i As Long
  Dim sExt As String
  
  m_Header = "": m_Footer = ""
  m_ImpDate = Now
  m_ImpUser = GetNetUser(False)
  ReDim cols(1 To m_Icols.Count) As ImportSpec
  For i = 1 To m_Icols.Count
    Set cols(i) = m_Icols.Item(i)
    If Not ((rs Is Nothing) And (rsRDO Is Nothing)) Then
      If Not ((cols(i).key = KEY_SKIPFIELD) Or (cols(i).key = KEY_BEGINMULTIPLE) Or (cols(i).key = KEY_ENDMULTIPLE) Or (cols(i).key = KEY_COPIEDFIELD)) Then
        If Not rs Is Nothing Then cols(i).DestTable = rs.Fields(cols(i).DestField).SourceTable
        If Not rsRDO Is Nothing Then cols(i).DestTable = rsRDO.rdoColumns(cols(i).DestField).SourceTable
      End If
    End If
  Next i
  Call SplitPath(m_ImportFilePath, , ImportFileName, sExt)
  ImportFileName = ImportFileName & sExt
  PrepareImport = m_Icols.Count
End Function

Private Sub Class_Initialize()
  Set ImpParser = New Parser
  Set ImpParser.ParseSettings = New SpecParseSettings
  Call ImpParser.AddParseItem(New ParseBeginMultiple)
  Call ImpParser.AddParseItem(New ParseBoolean)
  Call ImpParser.AddParseItem(New ParseBooleanFalse)
  Call ImpParser.AddParseItem(New ParseBooleanTrue)
  Call ImpParser.AddParseItem(New ParseConstant)
  Call ImpParser.AddParseItem(New ParseCopyField)
  Call ImpParser.AddParseItem(New ParseDate)
  Call ImpParser.AddParseItem(New ParseDateDelimiter)
  Call ImpParser.AddParseItem(New ParseDelimitedImport)
  Call ImpParser.AddParseItem(New ParseDelimiter)
  Call ImpParser.AddParseItem(New ParseEndMultiple)
  Call ImpParser.AddParseItem(New ParseEscapeChar)
  Call ImpParser.AddParseItem(New ParseFactor)
  Call ImpParser.AddParseItem(New ParseFatal)
  Call ImpParser.AddParseItem(New ParseField)
  Call ImpParser.AddParseItem(New ParseFieldName)
  Call ImpParser.AddParseItem(New ParseFileDate)
  Call ImpParser.AddParseItem(New ParseFileName)
  Call ImpParser.AddParseItem(New ParseFixedDate)
  Call ImpParser.AddParseItem(New ParseFixedDateTime)
  Call ImpParser.AddParseItem(New ParseFixedWidthImport)
  Call ImpParser.AddParseItem(New ParseFloat)
  Call ImpParser.AddParseItem(New ParseFooter)
  Call ImpParser.AddParseItem(New ParseHeader)
  Call ImpParser.AddParseItem(New ParseImpDate)
  Call ImpParser.AddParseItem(New ParseKeyField)
  Call ImpParser.AddParseItem(New ParseLineNumber)
  Call ImpParser.AddParseItem(New ParseMask)
  Call ImpParser.AddParseItem(New ParseMaxValue)
  Call ImpParser.AddParseItem(New ParseMinValue)
  Call ImpParser.AddParseItem(New ParseNegate)
  Call ImpParser.AddParseItem(New ParseNumber)
  Call ImpParser.AddParseItem(New ParseReqField)
  Call ImpParser.AddParseItem(New ParseSkipColumn)
  Call ImpParser.AddParseItem(New ParseString)
  Call ImpParser.AddParseItem(New ParseTimeDelimiter)
  Call ImpParser.AddParseItem(New ParseUsername)
  m_Delimiter = ","
  m_EscapeChar = """"
  m_ImportInitok = False
  m_DefStr = Null
  m_DefDate = Null
  m_DefNumeric = Null
  m_DefBool = Null
End Sub

Private Sub Class_Terminate()
  Call KillImporter
End Sub

Public Sub KillImporter()
  If Not m_ImportWizard Is Nothing Then
    Call m_ImportWizard.KillWizard
    Set m_ImportWizard = Nothing
  End If
  Set m_AuditObj = Nothing
  Set ImpParser = Nothing
  Set m_Icols = Nothing
  Set m_WorkSpace = Nothing
  Set m_rdoConnection = Nothing
End Sub

Public Function GetLibraryVersion() As String
  GetLibraryVersion = App.Major & "." & App.Minor & "." & App.Revision
End Function

Private Function AddNewRecord(rs As Recordset, rsRDO As rdoResultset, cols() As ImportSpec, ByVal MaxCols As Long, ByVal linenum As Long) As Boolean
  Dim ispec As ImportSpec
  Dim i As Long
  
  On Error GoTo AddNewRecord_err
  If Not rs Is Nothing Then rs.AddNew
  If Not rsRDO Is Nothing Then rsRDO.AddNew
  For i = 1 To MaxCols
    Set ispec = cols(i)
    If Not ispec.key = KEY_SKIPFIELD Then
      If Not IsNull(ispec.Value) Then
        If Not rs Is Nothing Then rs.Fields(ispec.DestField).Value = ispec.Value
        If Not rsRDO Is Nothing Then rsRDO.rdoColumns(ispec.DestField).Value = ispec.Value
        If Not m_AuditObj Is Nothing Then
          Call m_AuditObj.WriteAudit(ispec.Value, ispec.DestField, ispec.DestTable)
        End If
      End If
    End If
    Next i
  If Not rs Is Nothing Then rs.Update
  If Not rsRDO Is Nothing Then rsRDO.Update
  AddNewRecord = True
AddNewRecord_end:
  Exit Function
  
AddNewRecord_err:
  If Not rs Is Nothing Then If rs.EditMode = dbEditAdd Then rs.CancelUpdate
  If Not rsRDO Is Nothing Then If rsRDO.EditMode = rdEditAdd Then rsRDO.CancelUpdate
  AddNewRecord = False
  Call ErrorMessage(ERR_ERROR, Err, "Importing File", "Import error on line: " & CStr(linenum), "There was an error importing the file " & m_ImportFilePath & vbCr & "Unable to add new record")
  Resume AddNewRecord_end
End Function
