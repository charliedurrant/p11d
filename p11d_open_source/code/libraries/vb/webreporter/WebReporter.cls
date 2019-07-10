VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WebReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public NumFlushRows As Integer
Public CalculateWidthsSampleRows As Integer

Public XMLLabelRoot As String
Public XMLLabelTitle As String
Public XMLLabelData As String
Public XMLLabelRow As String

'kk
Private mDBObjectTimeout As Long
'kk
Private WErrs As WebErrors
Private mDBObjectName As String
Private mDBObjectType As DBObjectType
Private mDBParameters As Collection

Private mOutputType As OutputType
Private mDelimiter As String
Private mQualifier As String

Private mDSN As String
Private mFilter As String
Private mOrderBy As String

Private mDescription As String
Private mTableHeader As String
Private mTableFooter As String
Private mColumns As Collection
Private mFields As Integer
Private mColumnWidths As ColumnWidths
Private mSum As Boolean
Private mCount As Boolean

Private Enum DBObjectType
  DBOBJTYPE_TABLE = 0
  DBOBJTYPE_VIEW
  DBOBJTYPE_STOREDPROCEDURE
  DBOBJTYPE_FUNCTION
  DBOBJTYPE_SQL
End Enum

Private Enum OutputType
  OPTYPE_HTML = 0
  OPTYPE_EXCEL
  OPTYPE_XML
  OPTYPE_CSV
End Enum

Private Enum ColumnWidths
  CW_EXPLICIT = 0
  CW_STYLES
  CW_AUTO
End Enum


' if WebErrors are passed in then the reporter just adds to them.
' Else it outputs any errors
Public Sub CreateReport(ByRef response As response, ByVal xml As String, ByVal apppath As String, Optional ByRef WErrsIn As WebErrors)
  Dim commandText As String
  Dim rs As Recordset
    
  On Error GoTo CreateReport_Err
  Call ReadSpecification(xml, apppath)
  commandText = GetCommandText
  
  ' create recordset
  If mDBObjectType = DBOBJTYPE_SQL Then
    Set rs = CreateRecordsetSQL(commandText)
  Else
    Set rs = CreateRecordset(commandText)
  End If
  
  ' Output
  Call Output(rs, response)
  

CreateReport_End:
  response.flush
  Call Kill
  Exit Sub
  
CreateReport_Err:
  If WErrsIn Is Nothing Then
    Set WErrs = New WebErrors
    Call WErrs.Add(Err, "The report cannot be created." & vbCrLf & vbCrLf)
    response.contentType = S_CT_HTML
    response.Write (WErrs.HTML)
  Else
    Call WErrsIn.Add(Err, "The report cannot be created." & vbCrLf & vbCrLf)
  End If
  Resume CreateReport_End
  Resume
End Sub

' parse the XML doc and set the modular variables
Private Sub ReadSpecification(ByVal xml As String, ByVal apppath As String)
  Dim xmldom As MSXML2.DOMDocument40, reporternode As IXMLDOMElement, el As IXMLDOMElement, elem As IXMLDOMElement
  Dim attr As IXMLDOMAttribute
  Dim nodes As IXMLDOMNodeList
  Dim col As Column
  Dim par As DBParameter
  Dim i As Integer
  Dim s As String
  
  On Error GoTo ReadSpecification_err
  'Call ValidateSpec(xml, apppath & "\" & S_SCHEMA)
  Set xmldom = New DOMDocument40
  xmldom.loadXML xml
  
  Set reporternode = xmldom.selectSingleNode("//reporter")   '.firstChild
  If reporternode Is Nothing Then Err.Raise ERR_INVALID_SPEC, "ReadSpecification", "The XML specification must have a <REPORTER> node"
  With reporternode
    ' Set global variables
    ' required variables
    Set el = .selectSingleNode("dbobject")
    mDBObjectName = el.getAttribute("name")
    'kk
    If IsNull(el.getAttribute("timeout")) Then
      mDBObjectTimeout = 30
    Else
      mDBObjectTimeout = CInt(el.getAttribute("timeout"))
    End If
    'kk
    Select Case UCase(el.getAttribute("type"))
      Case "TABLE"
        mDBObjectType = DBOBJTYPE_TABLE
      Case "VIEW"
        mDBObjectType = DBOBJTYPE_VIEW
      Case "STOREDPROC"
        mDBObjectType = DBOBJTYPE_STOREDPROCEDURE
      Case "FUNCTION"
        mDBObjectType = DBOBJTYPE_FUNCTION
      Case "SQL"
        mDBObjectType = DBOBJTYPE_SQL
    End Select
    Select Case UCase(reporternode.getAttribute("outputtype"))
      Case "HTML"
        mOutputType = OPTYPE_HTML
      Case "EXCEL"
        mOutputType = OPTYPE_EXCEL
      Case "XML"
        mOutputType = OPTYPE_XML
      Case "CSV"
        mOutputType = OPTYPE_CSV
    End Select
    mDSN = .selectSingleNode("dsn").Text
    Select Case .getAttribute("columnwidths")
      Case "EXPLICIT"
        mColumnWidths = CW_EXPLICIT
      Case "STYLES"
        mColumnWidths = CW_STYLES
      Case "AUTO"
        mColumnWidths = CW_AUTO
    End Select
    If IsNull(reporternode.getAttribute("delimiter")) Then
      mDelimiter = ","
    Else
      mDelimiter = reporternode.getAttribute("delimiter")
      If Len(mDelimiter) = 0 Then mDelimiter = ","
    End If
    If IsNull(reporternode.getAttribute("qualifier")) Then
      mQualifier = """"
    Else
      mQualifier = reporternode.getAttribute("qualifier")
      If Len(mQualifier) = 0 Then mQualifier = """"
    End If
    'optional variables
    Set el = .selectSingleNode("filtering")
    If Not el Is Nothing Then mFilter = el.Text
    Set el = .selectSingleNode("ordering")
    If Not el Is Nothing Then mOrderBy = el.Text
    Set el = .selectSingleNode("description")
    If Not el Is Nothing Then mDescription = el.Text
    Set el = .selectSingleNode("tableheader")
    If Not el Is Nothing Then mTableHeader = el.Text
    Set el = .selectSingleNode("tablefooter")
    If Not el Is Nothing Then mTableFooter = el.Text
  End With
  
  ' Set Column objects
  Set mColumns = New Collection
  mFields = 0
  Set nodes = reporternode.selectNodes("columns/column")
  For i = 0 To nodes.Length - 1
    Set elem = nodes.Item(i)
    With elem
      ' only include columns that are explicitly set to display
      If (IsNullEx(.getAttribute("display"), 1) = 1) Then
        Set col = New Column
        ' required
        col.Name = .getAttribute("name")
        ' optional
        Set el = .selectSingleNode("coldescription")
        If Not el Is Nothing Then col.Caption = el.Text
        col.Group = IsNullEx(.getAttribute("group"), 0)
        If col.Group = 0 Then mFields = mFields + 1 ' decide how many columns should be in the HTML table
        col.Sum = IsNullEx(.getAttribute("sum"), 0)
        col.Count = IsNullEx(.getAttribute("count"), 0)
        Set el = .selectSingleNode("colwidth")
        If mColumnWidths = CW_EXPLICIT Then
          If Not el Is Nothing Then
            If Len(el.Text) > 0 Then
              col.Width = el.Text
            Else
              Err.Raise ERR_INVALID_SPEC, , "Column width value missing[ " & col.Name & " ]"
            End If
          Else
              Err.Raise ERR_INVALID_SPEC, , "Column width node missing[ " & col.Name & " ]"
          End If
        End If
        mColumns.Add col, col.Name
      End If
    End With
  Next i
  
  're-order columns in group order
  For i = 1 To mColumns.Count - mFields
    For Each col In mColumns
      If col.Group = i Then
        mColumns.Remove (col.Name)
        mColumns.Add col, col.Name, i
        Exit For
      End If
    Next col
  Next i
  
  ' Set parameter objects
  If mDBObjectType = DBOBJTYPE_STOREDPROCEDURE Or mDBObjectType = DBOBJTYPE_FUNCTION Then
    Set mDBParameters = New Collection
    Set nodes = reporternode.selectNodes("parameters/parameter")
    For i = 0 To nodes.Length - 1
      Set elem = nodes.Item(i)
      With elem
        Set par = New DBParameter
        par.SetDatatype (.getAttribute("datatype"))
        par.SetValue (.selectSingleNode("paramvalue").Text)
        par.DTLength = IsNullEx(.getAttribute("length"), 0)
      End With
      mDBParameters.Add par
    Next i

  End If

ReadSpecification_end:
  Set col = Nothing
  Set reporternode = Nothing
  Set nodes = Nothing
  Set par = Nothing
  Set xmldom = Nothing
  Set el = Nothing
  Set elem = Nothing
  Set attr = Nothing
  
  Exit Sub
ReadSpecification_err:
  Err.Raise Err.Number, ErrorSourceComponent(Err, "ReadSpecification", App.EXEName), Err.Description
  Resume
End Sub

Private Sub ValidateSpec(ByVal xml As String, ByVal SchemaLocation As String)
  Dim xmlschema As MSXML2.XMLSchemaCache40
  Dim xmldom As MSXML2.DOMDocument40
  
  On Error GoTo ValidateSpec_err
  
  Set xmlschema = New MSXML2.XMLSchemaCache40
  Set xmldom = New MSXML2.DOMDocument40
  xmlschema.Add S_SCHEMA_NAMESPACE, SchemaLocation
  
  Set xmldom.schemas = xmlschema
  xmldom.async = False
  
  xmldom.validateOnParse = True
  xmldom.Load S_XML_START & xml

  If xmldom.parseError.errorCode <> 0 Then
     Err.Raise ERR_INVALID_SPEC, , xmldom.parseError.errorCode & vbCrLf & xmldom.parseError.reason
  End If

ValidateSpec_end:
  Set xmlschema = Nothing
  Set xmldom = Nothing
  Exit Sub
ValidateSpec_err:
  Err.Raise Err.Number, ErrorSourceComponent(Err, "ValidateSpec", App.EXEName), Err.Description
  Resume
End Sub


'Create string to call DB object dependent on object type
Private Function GetCommandText() As String
  
  On Error GoTo GetCommandText_err
  
  Select Case mDBObjectType
    Case DBOBJTYPE_TABLE
      GetCommandText = GetCommandText_Table
    Case DBOBJTYPE_VIEW
      GetCommandText = GetCommandText_View
    Case DBOBJTYPE_STOREDPROCEDURE
      GetCommandText = GetCommandText_StoredProcedure
    Case DBOBJTYPE_FUNCTION
      GetCommandText = GetCommandText_Function
    Case DBOBJTYPE_SQL
      GetCommandText = GetCommandText_SQL
  End Select
  
  Exit Function
GetCommandText_err:
  Err.Raise Err.Number, ErrorSourceComponent(Err, "GetCommandText", App.EXEName), Err.Description
  Resume
End Function

'Uses ADO Recordset object to provide query results
Private Function CreateRecordset(ByVal commandText As String) As Recordset
  Dim cmd As Command
  Dim rs As Recordset
  
  On Error GoTo CreateRecordset_err
  
  Set cmd = NewADOCommand
  ' set commandText
  cmd.commandText = commandText
  'kk
  cmd.CommandTimeout = mDBObjectTimeout
  'kk
  Call CreateParameters(cmd)
  ' execute the query.  This returns a recordset
  Set rs = cmd.Execute
  
  ' filter the recordset
  rs.Filter = mFilter
  ' order the recordset
  rs.Sort = mOrderBy
  
  Set CreateRecordset = rs
CreateRecordset_end:
  Set cmd = Nothing
  Set rs = Nothing
  Exit Function
CreateRecordset_err:
  Err.Raise Err.Number, ErrorSourceComponent(Err, "CreateRecordset", App.EXEName), Err.Description
  Resume
End Function

'Uses ADO Recordset object to provide query results
Private Function CreateRecordsetSQL(ByVal commandText As String) As Recordset
  Dim cn As Connection
  Dim rs As Recordset
  
  On Error GoTo CreateRecordsetSQL_err
    
  Set cn = New Connection
  cn.ConnectionTimeout = mDBObjectTimeout
  Call cn.Open(mDSN)
  cn.CommandTimeout = mDBObjectTimeout
  
  Set rs = New Recordset
  
  Call rs.Open(commandText, cn, adOpenForwardOnly, adLockReadOnly)
  
  Set CreateRecordsetSQL = rs
CreateRecordsetSQL_end:
  Set rs = Nothing
  Exit Function
CreateRecordsetSQL_err:
  Err.Raise Err.Number, ErrorSourceComponent(Err, "CreateRecordsetSQL", App.EXEName), Err.Description
  Resume
End Function

' command object used to query the DB
Private Function NewADOCommand() As Command
  Dim cmdType As Integer
  Dim cn As Connection
  
  On Error GoTo NewADOCommand_Err
  
  Dim cmd As Command
  
  ' create a command object using the DSN provided
  ' use existing library proc - ADOConnect
  ' set the CommandType property
  Set cmd = New Command
  Set cn = ADOConnect(mDSN, adUseClient)
  
  Select Case mDBObjectType
    Case DBOBJTYPE_TABLE, DBOBJTYPE_VIEW
      cmd.CommandType = adCmdTable
    Case DBOBJTYPE_STOREDPROCEDURE
      cmd.CommandType = adCmdStoredProc
    Case DBOBJTYPE_FUNCTION
      cmd.CommandType = adCmdUnknown 'function??
  End Select
    
  Set cmd.ActiveConnection = cn
  Set NewADOCommand = cmd
  
NewADOCommand_End:
  Set cn = Nothing
  
  Exit Function
NewADOCommand_Err:
  Err.Raise Err.Number, ErrorSourceComponent(Err, "NewADOCommand", App.EXEName), Err.Description
  Resume
End Function

'Subset of GetCommandText function
Private Function GetCommandText_View() As String
  On Error GoTo GetCommandText_View_Err
  
  GetCommandText_View = mDBObjectName
  
  Exit Function
GetCommandText_View_Err:
  Err.Raise Err.Number, ErrorSourceComponent(Err, "GetCommandText_View", App.EXEName), Err.Description
  Resume
End Function

'Subset of GetCommandText function
Private Function GetCommandText_StoredProcedure() As String
  On Error GoTo GetCommandText_StoredProcedure_Err
  
  GetCommandText_StoredProcedure = mDBObjectName
  
  Exit Function
GetCommandText_StoredProcedure_Err:
  Err.Raise Err.Number, ErrorSourceComponent(Err, "GetCommandText_StoredProcedure", App.EXEName), Err.Description
  Resume
End Function


'Subset of GetCommandText function
Private Function GetCommandText_Function() As String
  On Error GoTo GetCommandText_Function_Err
  
  GetCommandText_Function = mDBObjectName
  
  Exit Function
GetCommandText_Function_Err:
  Err.Raise Err.Number, ErrorSourceComponent(Err, "GetCommandText_Function", App.EXEName), Err.Description
  Resume
End Function

'Subset of GetCommandText function
Private Function GetCommandText_Table() As String
  On Error GoTo GetCommandText_Table_Err
  
  GetCommandText_Table = mDBObjectName
  
  Exit Function
GetCommandText_Table_Err:
  Err.Raise Err.Number, ErrorSourceComponent(Err, "GetCommandText_Table", App.EXEName), Err.Description
  Resume
End Function

'Subset of GetCommandText function
Private Function GetCommandText_SQL() As String
  On Error GoTo GetCommandText_SQL_Err
  
  GetCommandText_SQL = mDBObjectName
  
  Exit Function
GetCommandText_SQL_Err:
  Err.Raise Err.Number, ErrorSourceComponent(Err, "GetCommandText_SQL", App.EXEName), Err.Description
  Resume
End Function

'Creates output for each format type - HTML, Excel, XML, CSV

Private Sub Output(ByVal rs As Recordset, ByRef response As response)
  Dim i As Long
  Dim col As Column
  Dim qs As QString
  On Error GoTo Output_Err
  
  ' set type of output
  Call SetContentType(response)

  If rs.RecordCount = 0 Then
    Select Case mOutputType
      Case OPTYPE_XML
        response.Write ElementOpen(XMLLabelRoot) & (S_NODATA) & ElementClose(XMLLabelRoot)
      Case Else
        response.Write (S_NODATA)
    End Select
  Else
    ' Title/Header
    response.Write OutputTitle(rs)
    
    'Determine column widths
    If mColumnWidths = CW_AUTO Then
      Call CalculateWidths(rs)
    End If
    
    ' Table headings
    response.Write OutputHeadings(rs)
    If NumFlushRows <> 0 Then response.flush
    
    ' content
    i = 1
    While Not rs.EOF
      If i = 1 Then response.Write OutputSectionStartEnd(True)
      response.Write OutputDataLine(rs)
      rs.MoveNext
      If i = NumFlushRows Then
        i = 0
        response.Write OutputSectionStartEnd(False)
        response.flush
      End If
      i = i + 1
    Wend
    
    If i <> 1 Then response.Write OutputSectionStartEnd(False)
    
    ' Final section
    response.Write OutputSectionStartEnd(True)
    
    ' Sum and Count on their own row
    response.Write OutputTotals
    
    ' Footer
    response.Write OutputFooter()

  End If

Output_end:
  response.flush
  
  Set rs = Nothing
  Set col = Nothing
  Set qs = Nothing
  
  Exit Sub
Output_Err:
  Err.Raise Err.Number, ErrorSourceComponent(Err, "Output", App.EXEName), Err.Description
  Resume
End Sub

' outputs title and opens xml root
Private Function OutputTitle(ByVal rs As Recordset) As String
  Dim qs As QString
  
  On Error GoTo OutputTitle_Err
  Set qs = New QString
  Select Case mOutputType
    Case OPTYPE_HTML
      If Len(mDescription) > 0 Then
        qs.Append ElementOpen("DIV", Attrib("class", S_CLASS_TITLE))
        qs.Append mDescription
        qs.Append ElementClose("DIV")
      End If
    Case OPTYPE_EXCEL
    Case OPTYPE_XML
      qs.Append S_XML_START
      qs.Append ElementOpen(XMLLabelRoot)
      If Len(mDescription) > 0 Then
        qs.Append ElementOpen(XMLLabelTitle)
        qs.Append mDescription
        qs.Append ElementClose(XMLLabelTitle)
      End If
    Case OPTYPE_CSV
      qs.Append Quote(mDescription, mQualifier, mDelimiter) & vbCrLf
  End Select
  OutputTitle = qs
  Set rs = Nothing
  Set qs = Nothing
  
  Exit Function
OutputTitle_Err:
  Err.Raise Err.Number, ErrorSourceComponent(Err, "OutputTitle", App.EXEName), Err.Description
  Resume
End Function

' closes the table tags and outputs any footer
Private Function OutputFooter() As String
  Dim qs As QString
  
  On Error GoTo OutputFooter_Err
  Set qs = New QString
  Select Case mOutputType
    Case OPTYPE_HTML
      If Len(mTableFooter) > 0 Then qs.Append HTMLFooter()
    Case OPTYPE_EXCEL
    Case OPTYPE_XML
      qs.Append ElementClose(XMLLabelData)
      qs.Append ElementClose(XMLLabelRoot)
    Case OPTYPE_CSV
      qs.Append mTableFooter & vbCrLf
  End Select
  OutputFooter = qs
  
OutputFooter_end:
  Set qs = Nothing
  
  Exit Function
OutputFooter_Err:
  Err.Raise Err.Number, ErrorSourceComponent(Err, "OutputFooter", App.EXEName), Err.Description
  Resume
End Function

' creates output for current row in recordset
Private Function OutputDataLine(ByVal rs As Recordset) As String
  Dim qs As QString
  
  On Error GoTo OutputDataLine_err
  Set qs = New QString
  Select Case mOutputType
    Case OPTYPE_HTML
      qs.Append HTMLLine(rs)
    Case OPTYPE_EXCEL
    Case OPTYPE_XML
      qs.Append XMLLine(rs, XMLLabelRow)
    Case OPTYPE_CSV
    ' kk
      If mColumns.Count = 1 Then
        qs.Append (vbCrLf)
      End If
    'kk
      qs.Append CSVLine(rs)
  End Select
  
  Call SumLine(rs)
  Call CountLine

OutputDataLine_end:
  OutputDataLine = qs
  
  Set qs = Nothing
  Set rs = Nothing
  
  Exit Function
OutputDataLine_err:
  Err.Raise Err.Number, ErrorSourceComponent(Err, "OutputDataLine", App.EXEName), Err.Description
  Resume
End Function

Private Function OutputSectionStartEnd(ByVal bStart As Boolean) As String
  Dim qs As QString
  Dim i As Integer
  Dim sWidth As String
  Dim col As Column
  
  On Error GoTo OutputSectionStartEnd_err
  Set qs = New QString
  Select Case mOutputType
    Case OPTYPE_HTML
      'if first row then open new table
      If bStart Then
        qs.Append (ElementOpen("TABLE", Attrib("ID", mDescription) & Attrib("CLASS", S_CLASS_TABLE)))   ' rs.source should be name of rs
        qs.Append (ElementOpen("COLGROUP"))
        For i = 1 To mColumns.Count
          Set col = mColumns(i)
          If col.Group = 0 Then
            If mColumnWidths <> CW_STYLES Then
              sWidth = Attrib("WIDTH", col.Width)
            End If
            qs.Append (ElementOpen("COL", Attrib("ID", col.Name) & Attrib("CLASS", S_CLASS_COL) & sWidth, True))
          End If
        Next i
        qs.Append (ElementClose("COLGROUP"))
      Else
        ' close the last table
        qs.Append ElementClose("TABLE")
      End If
  End Select

OutputSectionStartEnd_end:
  OutputSectionStartEnd = qs
  
  Set qs = Nothing
  Set col = Nothing
  
  Exit Function
OutputSectionStartEnd_err:
  Err.Raise Err.Number, ErrorSourceComponent(Err, "OutputSectionStartEnd", App.EXEName), Err.Description
  Resume
End Function

' opens table tags, outputs any header, and outputs column headings
Private Function OutputHeadings(ByVal rs As Recordset) As String
  Dim qs As QString
  Dim i As Integer
  
  On Error GoTo OutputHeadings_err
  Set qs = New QString
  Select Case mOutputType
    Case OPTYPE_HTML
      qs.Append OutputSectionStartEnd(True)
      qs.Append HTMLHeadings()
      qs.Append OutputSectionStartEnd(False)
    Case OPTYPE_EXCEL
    Case OPTYPE_XML
      qs.Append ElementOpen(XMLLabelData)
    Case OPTYPE_CSV
      qs.Append CSVHeadings
  End Select
  
OutputHeadings_end:
  OutputHeadings = qs
  
  Set qs = Nothing
  Set rs = Nothing
  
  Exit Function
OutputHeadings_err:
  Err.Raise Err.Number, ErrorSourceComponent(Err, "OutputHeadings", App.EXEName), Err.Description
  Resume
End Function

Private Function OutputTotals() As String
  Dim qs As QString
  Dim col As Column
  Dim i As Integer
  
  On Error GoTo OutputTotals_err
  Set qs = New QString
  Select Case mOutputType
    Case OPTYPE_HTML
      qs.Append ElementOpen("TR", Attrib("CLASS", S_CLASS_TR_TOTALS))
      For i = 1 To mColumns.Count
        Set col = mColumns.Item(i)
        If col.Count = True Then
          qs.Append ElementOpen("TD", Attrib("CLASS", S_CLASS_TD_COUNT))
          qs.Append col.CountTotal
        ElseIf col.Sum = True Then
          qs.Append ElementOpen("TD", Attrib("CLASS", S_CLASS_TD_SUM))
          qs.Append col.SumTotal
        ElseIf col.Group = True Then
          qs.Append ElementOpen("TD", Attrib("CLASS", S_CLASS_TD_GROUP))
          qs.Append col.Group
        Else
          qs.Append ElementOpen("TD") & "&nbsp;"
        End If
        qs.Append ElementClose("TD")
      Next i
      qs.Append ElementClose("TR")
    Case OPTYPE_EXCEL
    Case OPTYPE_XML
    Case OPTYPE_CSV
  End Select
  
OutputTotals_end:
  OutputTotals = qs
 
  Set qs = Nothing
  Set col = Nothing
 
  Exit Function
OutputTotals_err:
  Err.Raise Err.Number, ErrorSourceComponent(Err, "OutputTotals", App.EXEName), Err.Description
  Resume
End Function


Private Sub CountLine()
  Dim col As Column, i As Long
  
  On Error GoTo CountLine_err
  For i = 1 To mColumns.Count
    Set col = mColumns.Item(i)
    If col.Count Then Call col.CountAdd(1)
  Next i
  
CountLine_end:
  Set col = Nothing
  Exit Sub
CountLine_err:
  Err.Raise Err.Number, ErrorSourceComponent(Err, "CountLine", App.EXEName), Err.Description
  Resume
End Sub

Private Sub SumLine(ByVal rs As Recordset)
  Dim col As Column, i As Long
  
  On Error GoTo SumLine_err
  For i = 1 To mColumns.Count
    Set col = mColumns.Item(i)
    If col.Sum Then Call col.SumAdd(CLng(IsNullEx(Trim(rs.fields(col.Name).Value), 0)))
  Next i
  
SumLine_end:
  Set col = Nothing
  Set rs = Nothing
  
  Exit Sub
SumLine_err:
  Err.Raise Err.Number, ErrorSourceComponent(Err, "SumLine", App.EXEName), Err.Description
  Resume
End Sub

Private Sub SetContentType(ByRef Resp As response)
  Select Case mOutputType
    Case OPTYPE_HTML
      Resp.contentType = "text/HTML"
    Case OPTYPE_EXCEL
    Case OPTYPE_XML
      Resp.contentType = "text/XML"
    Case OPTYPE_CSV
      Resp.contentType = "application/x-download"
      Resp.AddHeader "content-disposition", "attachment; filename=""" & mDescription & ".csv"""
  End Select
End Sub

'End report
Private Sub Kill()

  On Error GoTo Kill_Err
  Set mColumns = Nothing
  Set mDBParameters = Nothing
  
  Exit Sub
Kill_Err:
  Err.Raise Err.Number, ErrorSourceComponent(Err, "Kill", App.EXEName), Err.Description
  Resume
End Sub

Private Function Quote(ByVal s As String, Optional ByVal quotetype As String = """", Optional ByVal delimiter As String = ",") As String
  If Len(s) > 0 Then
    If InStr(s, quotetype) > 0 Or InStr(s, delimiter) Then
      s = Replace(s, quotetype, quotetype & quotetype)
      Quote = quotetype & s & quotetype
    Else
      Quote = s
    End If
  Else
    Quote = ""
  End If
End Function

Private Function XMLText(ByVal s As String) As String
  XMLText = "<![CDATA[" & Trim(s) & "]]>"
End Function

Private Sub CreateParameters(ByRef cmd As Command)
  Dim param As Parameter
  Dim dbParam As DBParameter
  Dim i As Integer
  Dim v As Variant
  
  On Error GoTo CreateParameters_err
  Select Case mDBObjectType
    Case DBOBJTYPE_STOREDPROCEDURE, DBOBJTYPE_FUNCTION
     For i = 1 To mDBParameters.Count
        Set param = New Parameter
        Set dbParam = mDBParameters(i)
        param.Type = dbParam.Datatype
        If param.Type = adVarChar Then param.Size = dbParam.DTLength
        param.Value = dbParam.Value
        cmd.Parameters.Append param
      Next i
    Case DBOBJTYPE_FUNCTION
    Case DBOBJTYPE_TABLE
      ' no parameters
    Case DBOBJTYPE_VIEW
      ' no parameters?
  End Select
  
CreateParameters_end:
  Set param = Nothing
  Set dbParam = Nothing
  Exit Sub
CreateParameters_err:
  Err.Raise Err.Number, ErrorSourceComponent(Err, "GetCommandText", App.EXEName), Err.Description
  Resume
End Sub

Private Function HTMLHeadings(Optional ByVal Attributes As String = "") As String
  Dim qs As QString
  Dim i As Long, col As Column
  
  On Error GoTo HTMLHeadings_ERR
  Set qs = New QString
  qs.Append ElementOpen("THEAD")
  qs.Append ElementOpen("TR", Attrib("CLASS", S_CLASS_TRH))
  qs.Append ElementOpen("TH", Attrib("colspan", mFields) & Attrib("CLASS", S_CLASS_TH_HEADER))
  qs.Append mTableHeader
  qs.Append ElementClose("TH")
  qs.Append ElementClose("TR")
  qs.Append ElementOpen("TR", Attrib("CLASS", S_CLASS_TRH))
  For i = 1 To mColumns.Count
    Set col = mColumns(i)
    ' if a grouping column then don't include in headings
    If col.Group = 0 Then
      qs.Append ElementOpen("TH", Attrib("ID", col.Name) & Attrib("CLASS", S_CLASS_TH))
      qs.Append col.Caption
      qs.Append ElementClose("TH")
    End If
  Next i
  qs.Append ElementClose("TR")
  qs.Append ElementClose("THEAD")
  HTMLHeadings = qs
  
HTMLHeadings_end:
  Set qs = Nothing
  Set col = Nothing
  
  Exit Function
  
HTMLHeadings_ERR:
  Err.Raise ERR_INCLUDE_STYLE, ErrorSourceComponent(Err, "HTMLHeadings", App.EXEName), Err.Description
  Resume
End Function

Private Function HTMLFooter() As String
  Dim qs As QString
  Dim i As Long

  On Error GoTo HTMLFooter_ERR
  Set qs = New QString
  qs.Append (ElementOpen("TFOOT"))
  qs.Append ElementOpen("TR") & ElementOpen("TD") & mTableFooter & ElementClose("TD") & ElementClose("TR")
  
  qs.Append (ElementClose("TFOOT"))
  qs.Append (ElementClose("TABLE"))
  HTMLFooter = qs
    
HTMLFooter_end:
  Set qs = Nothing
  
  Exit Function
  
HTMLFooter_ERR:
  Err.Raise ERR_INCLUDE_STYLE, ErrorSourceComponent(Err, "HTMLFooter", App.EXEName), Err.Description
  Resume
End Function

Private Function HTMLLine(ByVal rs As Recordset) As String
  Dim qs As QString
  Dim i As Long, col As Column
  Dim c As Column
  
  On Error GoTo HTMLLine_ERR
  Set qs = New QString
  If Not rs.EOF And Not rs.BOF Then
  
    qs.Append (ElementOpen("TR", Attrib("CLASS", S_CLASS_TRD)))
    For i = 1 To mColumns.Count
      Set col = mColumns(i)
      ' if a grouping column and is different from last groupvalue then output a row of its own
      If col.Group > 0 Then
        If StrComp(Trim(rs.fields(col.Name).Value), Trim(col.GroupLastValue)) <> 0 Then
          ' set last value of this column
          col.GroupLastValue = Trim(rs.fields(col.Name).Value)
          ' clear values from later groupings
          For Each c In mColumns
            If c.Group > col.Group Then c.GroupLastValue = ""
          Next c
          qs.Append (ElementOpen("TD", Attrib("CLASS", S_CLASS_TD_GROUP & CStr(col.Group)) & Attrib("COLSPAN", CStr(mFields))))
          qs.Append (Trim(col.GroupLastValue))
          qs.Append (ElementClose("TD"))
          qs.Append (ElementClose("TR"))
          qs.Append (ElementOpen("TR", Attrib("CLASS", S_CLASS_TRD)))
        End If
        ' if value hasn't changed then ignore
      Else
      qs.Append (ElementOpen("TD", Attrib("ID", rs.fields(col.Name).Name) & Attrib("CLASS", S_CLASS_TD)))
      qs.Append (IsNullEx(Trim(rs.fields(col.Name).Value), "") & "&nbsp;")
      qs.Append (ElementClose("TD"))
      End If
    
    Next i
    qs.Append (ElementClose("TR"))
  End If
  HTMLLine = qs
  
HTMLLine_end:
  Set rs = Nothing
  Set qs = Nothing
  Set col = Nothing
  Set c = Nothing
  
  Exit Function
  
HTMLLine_ERR:
  Err.Raise ERR_INCLUDE_STYLE, ErrorSourceComponent(Err, "HTMLLine", App.EXEName), Err.Description
  Resume
End Function

Private Function XMLLine(ByVal rs As Recordset, ByVal sRowName As String) As String

  Dim qs As QString
  Dim i As Long
     
  On Error GoTo XMLLine_ERR
  Set qs = New QString
  If Not rs.EOF And Not rs.BOF Then
    qs.Append ElementOpen(sRowName)
    For i = 1 To mColumns.Count
      qs.Append ElementOpen(rs.fields(mColumns(i).Name).Name) & XMLText(IsNullEx(rs.fields(mColumns(i).Name).Value, "")) & ElementClose(rs.fields(mColumns(i).Name).Name)
    Next i
    qs.Append ElementClose(sRowName)
  End If
  XMLLine = qs
  
XMLLine_end:
  Set rs = Nothing
  Set qs = Nothing
  
  Exit Function

XMLLine_ERR:
  Err.Raise ERR_INCLUDE_STYLE, ErrorSourceComponent(Err, "XMLLine", App.EXEName), Err.Description
  Resume
End Function

Private Function CSVLine(ByVal rs As Recordset) As String
  Dim qs As QString
  Dim col As Column, sValue As String
  Dim i As Integer

  On Error GoTo CSVLine_ERR
  Set qs = New QString
  For i = 1 To mColumns.Count
  Set col = mColumns(i)
    sValue = IsNullEx(Trim(rs.fields(col.Name).Value), 0)
      If i = 1 Then
        'first item
        qs.Append Quote(Trim(sValue), mQualifier, mDelimiter)
      ElseIf i = mColumns.Count Then
        'last item
        qs.Append mDelimiter & Quote(Trim(sValue), mQualifier, mDelimiter) & vbCrLf
      Else
        'middle item
        qs.Append mDelimiter & Quote(Trim(sValue), mQualifier, mDelimiter)
      End If
  Next i
  CSVLine = qs
    
CSVLine_end:
  Set col = Nothing
  Set qs = Nothing
  Set rs = Nothing
  
  Exit Function
  
CSVLine_ERR:
  Err.Raise ERR_INCLUDE_STYLE, ErrorSourceComponent(Err, "CSVLine", App.EXEName), Err.Description
  Resume
End Function

Private Function CSVHeadings() As String
  Dim qs As QString
  Dim col As Column, sValue As String
  Dim i As Integer

  On Error GoTo CSVHeadings_ERR
  Set qs = New QString
  qs.Append mTableHeader & vbCrLf  'pq would you want a header for csv?
  For i = 1 To mColumns.Count
    Set col = mColumns(i)
        If i = 1 Then
          'first item
          qs.Append Quote(col.Caption, mQualifier, mDelimiter)
        ElseIf i = mColumns.Count Then
          'last item
          qs.Append mDelimiter & Quote(col.Caption, mQualifier, mDelimiter) & vbCrLf
        Else
          'middle item
          qs.Append mDelimiter & Quote(col.Caption, mQualifier, mDelimiter)
        End If
  Next i

  CSVHeadings = qs
    
CSVHeadings_end:
  Set qs = Nothing
  Set col = Nothing
  
  Exit Function
  
CSVHeadings_ERR:
  Err.Raise ERR_INCLUDE_STYLE, ErrorSourceComponent(Err, "CSVHeadings", App.EXEName), Err.Description
  Resume
End Function

Private Sub CalculateWidths(ByRef rs As Recordset)
  Dim col As Column, i As Long, j As Long
  Dim lTotalChars As Long
  
  On Error GoTo CalculateWidths_ERR
  Select Case mOutputType
    Case OPTYPE_HTML
      ' Add up the total number of characters in each column
      i = 1
      While (Not rs.EOF) And (i < CalculateWidthsSampleRows)
        For j = 1 To mColumns.Count
          Set col = mColumns.Item(j)
          col.CharsAdd Len(IsNullEx(Trim(rs.fields(col.Name).Value), 0))
          lTotalChars = lTotalChars + Len(IsNullEx(Trim(rs.fields(col.Name).Value), 0))
        Next j
        rs.MoveNext
        i = i + 1
      Wend
      
      ' include titles in count
      For j = 1 To mColumns.Count
        Set col = mColumns.Item(j)
        col.CharsAdd Len(col.Caption)
        lTotalChars = lTotalChars + Len(col.Caption)
      Next j
      
      ' apportion percentages dependent on the fraction over total chars
      For j = 1 To mColumns.Count
        Set col = mColumns.Item(j)
        col.Width = FormatNumber(CStr(col.CharsTotal / lTotalChars) * 100, 0) & "%"
      Next j
      
      If Not (rs.BOF And rs.EOF) Then rs.MoveFirst
    Case OPTYPE_EXCEL
    Case OPTYPE_XML
    Case OPTYPE_CSV
  End Select
  
CalculateWidths_end:
  Set col = Nothing
  
  Exit Sub
  
CalculateWidths_ERR:
  Err.Raise ERR_INCLUDE_STYLE, ErrorSourceComponent(Err, "CalculateWidths", App.EXEName), Err.Description
  Resume
End Sub

Private Sub Class_Initialize()
  NumFlushRows = 5
  CalculateWidthsSampleRows = 20
  XMLLabelRoot = "report"
  XMLLabelTitle = "title"
  XMLLabelData = "records"
  XMLLabelRow = "record"
End Sub

