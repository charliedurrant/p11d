VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PayeOnlineXMLValidator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IVBSAXErrorHandler

Private m_Errors As Collection

Public Property Get Errors() As Collection
  Set Errors = m_Errors
End Property
Public Sub Validate(xmlToValidate As String, xmlSchemaFile As String)
  Dim oReader As SAXXMLReader60
  Dim oSC As MSXML2.XMLSchemaCache60
  
On Error GoTo err_Err
    
  
    
  Set oReader = New MSXML2.SAXXMLReader60
  Set oSC = New MSXML2.XMLSchemaCache60
  Set m_Errors = New Collection
  
  
  
  Call oSC.Add(p11d32.PAYEonline.namespace, xmlSchemaFile)
  
  
  
  'Configure the SAX reader to validate the XML document.
  oReader.putFeature "schema-validation", True
  oReader.putFeature "exhaustive-errors", True
  oReader.PutProperty "schemas", oSC

  'Assign an instance of the MyValidator class to
  'the errorHandler property of the SAX reader.
  Set oReader.ErrorHandler = Me
  'Parse and validate the file.
  oReader.parse xmlToValidate
  Debug.Print "End of parsing and validation"
 
err_End:
 Exit Sub
err_Err:
  Call MsgBox(Err.Description)
  Resume err_End
  Resume
End Sub
Private Sub ValidateByDOM(xmlToValidate As String, xmlSchemaFile As String)
  Dim dom As DOMDocument60
  Dim schemaCache As XMLSchemaCache60
  Dim error As IXMLDOMParseError
  
On Error GoTo err_Err

  Set dom = New DOMDocument60
  Call dom.setProperty("ProhibitDTD", False)
  dom.validateOnParse = True
  dom.resolveExternals = True
  dom.Async = False
  
  Set schemaCache = New XMLSchemaCache60
  Call schemaCache.Add(p11d32.PAYEonline.namespace, xmlSchemaFile)
  
  Set dom.schemas = schemaCache
  Call dom.loadXML(xmlToValidate)
  
   
  Set error = dom.Validate()
  If error.ErrorCode <> 0 Then
      Debug.Print "Error parser: " & error.ErrorCode & "; " & error.reason
  Else
      Debug.Print "No errors found"
  End If
  
err_End:
  Exit Sub
err_Err:
  Debug.Print (Err.Description)
  Resume err_End
  Resume
  
End Sub
Private Sub IVBSAXErrorHandler_error(ByVal oLocator As MSXML2.IVBSAXLocator, _
            strErrorMessage As String, ByVal nErrorCode As Long)
            
  Dim error As PayeOnlineXmlValidationError
  
  Set error = New PayeOnlineXmlValidationError
  
  error.Code = nErrorCode
  error.Description = strErrorMessage
  error.ColumnNumber = oLocator.ColumnNumber
  error.LineNumber = oLocator.LineNumber
  
  m_Errors.Add error
End Sub

Private Sub IVBSAXErrorHandler_fatalError(ByVal oLocator As MSXML2.IVBSAXLocator, _
            strErrorMessage As String, ByVal nErrorCode As Long)

   WriteErrorToDebugWindow "Fatal error or parsing error", strErrorMessage, _
         nErrorCode, oLocator.LineNumber, oLocator.ColumnNumber
End Sub

Private Sub IVBSAXErrorHandler_ignorableWarning(ByVal oLocator As MSXML2.IVBSAXLocator, _
            strErrorMessage As String, ByVal nErrorCode As Long)
End Sub


Private Sub WriteErrorToDebugWindow(strLabel As String, _
   strDescription As String, ByVal ErrCode As Long, _
   line As Long, column As Long)
   Debug.Print strLabel + ": (" + CStr(ErrCode) + ") " + _
        strDescription & "at " + "line " + _
        Str(line) + ", column " + _
        Str(column) + vbCrLf
End Sub
    

