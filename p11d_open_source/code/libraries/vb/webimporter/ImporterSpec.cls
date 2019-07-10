VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ImporterSpec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim m_Delimiter As String
Dim m_Qualifier As String
Dim m_DestView As String
Dim m_LineCount As Long
Dim m_FileHeader As Boolean
Dim m_FileFooter As Boolean
Dim m_FileHashTotal As Boolean
Dim m_DBConnString As String
Dim m_Column As Collection
Dim m_Col As Collection
Dim m_Col2 As Collection
Dim m_ColFooter As Collection
Dim m_ColHeader As Collection
Dim m_DestTableName As String
Dim m_ErrorLogName As String
Dim m_AuditLogName As String
Dim m_ImportUsername As String
Dim m_ImportFileName As String
Dim m_ImportSpecName As String
Dim m_UseTransaction As Boolean
Dim m_DateFrom As String
Dim m_DateTo As String

' Need  specification defaults in this section.
Public Property Let DateFrom(val As String)
  m_DateFrom = val
End Property
Public Property Get DateFrom() As String
  DateFrom = m_DateFrom
End Property

Public Property Let DateTo(val As String)
  m_DateTo = val
End Property

Public Property Get DateTo() As String
  DateTo = m_DateTo
End Property

Public Property Let Delimiter(val As String)
  m_Delimiter = val
End Property

Public Property Get Delimiter() As String
  Delimiter = m_Delimiter
End Property

Public Property Let Qualifier(val As String)
  m_Qualifier = val
End Property

Public Property Get Qualifier() As String
  Qualifier = m_Qualifier
End Property

Public Property Let ImportDest(val As String)
  m_DestView = val
End Property

Public Property Get ImportDest() As String
  ImportDest = m_DestView
End Property

Public Property Let ErrorLog(val As String)
  m_ErrorLogName = val
End Property

Public Property Get ErrorLog() As String
  ErrorLog = m_ErrorLogName
End Property

Public Property Let AuditLog(val As String)
  m_AuditLogName = val
End Property

Public Property Get AuditLog() As String
  AuditLog = m_AuditLogName
End Property

Public Property Let LineCount(val As Long)
  m_LineCount = val
End Property

Public Property Get LineCount() As Long
  LineCount = m_LineCount
End Property
Public Property Let FileHeader(val As Boolean)
  m_FileHeader = val
End Property

Public Property Get FileHeader() As Boolean
  FileHeader = m_FileHeader
End Property
Public Property Let DBConnectionString(val As String)
  m_DBConnString = val
End Property

Public Property Get DBConnectionString() As String
  DBConnectionString = m_DBConnString
End Property

Public Property Let ColumnSpec(val As String)
  m_DBConnString = val
End Property

Public Function SetFileColumnProperties(ByVal xmlColList As IXMLDOMNodeList) As Collection
  
On Error GoTo SetFileColumnProperties_err
  Dim cSpec As ColumnSpec
  Dim tmpElem As IXMLDOMElement
  
  
  If Not xmlColList Is Nothing Then
    Set cSpec = New ColumnSpec
      cSpec.ColName = "GUID"
      cSpec.ColDataType = "String"
      m_Col.Add cSpec
    Set cSpec = New ColumnSpec
      cSpec.ColName = "LineNumber"
      cSpec.ColDataType = "Big integer"
      m_Col.Add cSpec
    For Each tmpElem In xmlColList
      Set cSpec = New ColumnSpec
      cSpec.ColName = IsNullEx(tmpElem.getAttribute("name"), "")
      cSpec.ColDataType = IsNullEx(tmpElem.getAttribute("datatype"), "String")
      cSpec.ColSkip = CBool(IsNullEx(tmpElem.getAttribute("skip"), False))
      cSpec.ColMap = IsNullEx(tmpElem.getAttribute("map"), 0)
      cSpec.ColNullable = IsNullEx(tmpElem.getAttribute("nullable"), True)
      cSpec.ColOrder = IsNullEx(tmpElem.getAttribute("order"), 0)
      cSpec.ColWidth = IsNullEx(tmpElem.getAttribute("width"), 0)
      m_Col.Add cSpec, cSpec.ColName
    Next tmpElem
  End If
  
  
SetFileColumnProperties_end:
  Set SetFileColumnProperties = m_Col
  Exit Function
  
SetFileColumnProperties_err:
  Err.Raise ERR_INVALID_COLUMN_PROPERTY, ErrorSource(Err, "ImportSpec.SetFileColumnProperties"), Err.Description & vbCrLf & "Error setting column properties."
  Resume
End Function

Private Sub Class_Initialize()
  Set m_Col = New Collection
  Set m_Col2 = New Collection
  Set m_ColFooter = New Collection
  Set m_ColHeader = New Collection
End Sub

Private Sub Class_Terminate()
  Set m_Col = Nothing
  Set m_Col2 = Nothing
  Set m_ColFooter = Nothing
  Set m_ColHeader = Nothing
End Sub
Public Function SetTableColumnProperties(ByVal xmlColList As IXMLDOMNodeList, Cn As ADODB.Connection) As Collection
Dim rs As ADODB.Recordset
Dim column As ADODB.Field

On Error GoTo SetTableColumnProperties_err

  Dim cSpec As ColumnSpec
  Dim tmpElem As IXMLDOMElement
  Dim lCounter As Long
  
  If Not xmlColList Is Nothing Then
    lCounter = 1
    If xmlColList.Length = 0 Then ' column information are not given
      If Not gADOHelper.IsTablePresentADO(Cn, m_DestView) Then
        Err.Raise ERR_INVALID_COLUMN_PROPERTY, ErrorSource(Err, "ImportSpec.SetTableColumnProperties"), Err.Description & vbCrLf & "Target table " & m_DestView & " not present."
      End If
    
      Set rs = New ADODB.Recordset
      rs.Open m_DestView, Cn
      'Collect constituent columns in a pack of columns
      For Each column In rs.Fields
        Set cSpec = New ColumnSpec
        cSpec.ColName = column.Name
        cSpec.ColDataType = column.Type
        cSpec.ColMap = lCounter
        cSpec.ColNullable = False
        cSpec.ColOrder = lCounter
        cSpec.ColWidth = column.DefinedSize
        m_Col2.Add cSpec, cSpec.ColName
        lCounter = lCounter + 1
        'Debug.Print tmpElem.Text
      Next column
   
    Else
      Set cSpec = New ColumnSpec
        cSpec.ColName = "GUID"
        m_Col2.Add cSpec, cSpec.ColName
      Set cSpec = New ColumnSpec
        cSpec.ColName = "LineNumber"
        m_Col2.Add cSpec, cSpec.ColName
      For Each tmpElem In xmlColList
        Set cSpec = New ColumnSpec
        cSpec.ColName = IsNullEx(tmpElem.getAttribute("name"), "")
        cSpec.ColDataType = IsNullEx(tmpElem.getAttribute("datatype"), "")
        cSpec.ColMap = IsNullEx(tmpElem.getAttribute("map"), 0)
        cSpec.ColNullable = IsNullEx(tmpElem.getAttribute("nullable"), True)
        cSpec.ColOrder = IsNullEx(tmpElem.getAttribute("order"), 0)
        cSpec.ColWidth = IsNullEx(tmpElem.getAttribute("width"), 0)
        m_Col2.Add cSpec, cSpec.ColName
        'Debug.Print tmpElem.Text
      Next tmpElem
    End If
  End If
  
  Set SetTableColumnProperties = m_Col2
  
SetTableColumnProperties_end:
    Set rs = Nothing
  Exit Function
  
SetTableColumnProperties_err:
  Err.Raise ERR_INVALID_COLUMN_PROPERTY, ErrorSource(Err, "ImportSpec.SetTableColumnProperties"), Err.Description & vbCrLf & "Error setting column properties."
  Resume
End Function
Public Property Let ImportUsername(val As String)
  m_ImportUsername = val
End Property

Public Property Get ImportUsername() As String
  ImportUsername = m_ImportUsername
End Property

Public Property Let ImportFileName(val As String)
  m_ImportFileName = val
End Property

Public Property Get ImportFileName() As String
  ImportFileName = m_ImportFileName
End Property

Public Property Let ImportSpecName(val As String)
  m_ImportSpecName = val
End Property

Public Property Get ImportSpecName() As String
  ImportSpecName = m_ImportSpecName
End Property

Public Property Let UseTransaction(val As Boolean)
  m_UseTransaction = val
End Property

Public Property Get UseTransaction() As Boolean
  UseTransaction = m_UseTransaction
End Property
Public Sub SetRowProperties(ByVal node As IXMLDOMElement)

On Error GoTo SetRowProperties_err

  Me.LineCount = IsNullEx(node.getAttribute("count"), 0)
  Me.FileHeader = CBool(IsNullEx(node.getAttribute("header"), 0))
  Me.FileFooter = CBool(IsNullEx(node.getAttribute("footer"), 0))
  
SetRowProperties_end:
    
  Exit Sub
  
SetRowProperties_err:
  Err.Raise ERR_INVALID_COLUMN_PROPERTY, ErrorSource(Err, "ImportSpec.SetRowProperties"), Err.Description & vbCrLf & "Error setting row properties."
  Resume
End Sub
Public Property Let FileFooter(val As Boolean)
  m_FileFooter = val
End Property

Public Property Get FileFooter() As Boolean
  FileFooter = m_FileFooter
End Property

Public Function SetHeaderAndFooter(ByVal xmlColList As IXMLDOMNodeList, Optional bfooter As Boolean = False) As Collection
Dim cSpec As ColumnSpec
Dim Elem As IXMLDOMElement
  
On Error GoTo SetHeaderAndFooter_err
  If Not xmlColList Is Nothing Then
      For Each Elem In xmlColList
      Set cSpec = New ColumnSpec
      cSpec.ColName = IsNullEx(Elem.getAttribute("name"), "")
      cSpec.ColDescription = IsNullEx(Elem.getAttribute("description"), "")
      If bfooter Then
        m_ColFooter.Add cSpec
      Else
        m_ColHeader.Add cSpec
      End If
    Next Elem
  End If
  
SetHeaderAndFooter_end:
  If bfooter Then
    Set SetHeaderAndFooter = m_ColFooter
  Else
    Set SetHeaderAndFooter = m_ColHeader
  End If
  
  Exit Function
  
SetHeaderAndFooter_err:
  Err.Raise ERR_INVALID_COLUMN_PROPERTY, ErrorSource(Err, "ImportSpec.SetHeaderAndFooter"), Err.Description & vbCrLf & "Error setting header and footer properties."
  Resume
End Function
