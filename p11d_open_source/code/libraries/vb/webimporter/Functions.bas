Attribute VB_Name = "Functions"
Option Explicit
 
Public Function GetXmlFromFile(ByVal FName As String) As String
  Dim fs As FileSystemObject
  Dim ts As TextStream
  Dim f As file
  
  On Error GoTo GetXmlFromFile_Err
  
  Set fs = New FileSystemObject
  If fs.FileExists(FName) Then
    Set f = fs.GetFile(FName)
    Set ts = fs.OpenTextFile(FName, ForReading, False)
    GetXmlFromFile = ts.ReadAll
  Else
    Err.Raise ERR_NO_SPECIFICATION_FILE, ErrorSource(Err, "Functions.GetXmlFromFile"), "File not found. " & FName
  End If
  
GetXmlFromFile_End:
  Set fs = Nothing
  Set ts = Nothing
  Set f = Nothing
  Exit Function

GetXmlFromFile_Err:
   Err.Raise ERR_GET_XML_FROM_FILE, ErrorSource(Err, "Functions.GetXmlFromFile"), "Error in getting XML from file." & vbCrLf & Err.Description
   Resume
End Function

Public Function getXMLDoc(xmlString, NodeName) As DOMDocument30
  Dim xmlReturn As String
  Dim xmlDoc As DOMDocument30
  Dim posStart As Long
  Dim posEnd As Long
  
  Set getXMLDoc = New DOMDocument30
  posStart = InStr(1, xmlString, "<" & NodeName & ">", vbTextCompare)
  If posStart = 0 Then
    xmlReturn = "<" & NodeName & "/>"
  Else
    posEnd = InStr(1, xmlString, "</" & NodeName & ">", vbTextCompare)
    If posEnd = 0 Then posEnd = Len(xmlString)
    xmlReturn = Mid(xmlString, posStart, posEnd - posStart + Len(NodeName) + 3)
  End If
  Call getXMLDoc.loadXML(xmlReturn)
End Function


Public Function HTMLEncode(ByVal sAttrValue As String) As String
  If InStrAny(sAttrValue, "<>"" ", 1, vbBinaryCompare) > 0 Then
    sAttrValue = Replace(sAttrValue, "<", "&lt;")
    sAttrValue = Replace(sAttrValue, ">", "&gt;")
    sAttrValue = Replace(sAttrValue, """", "&quot;")
    sAttrValue = Replace(sAttrValue, " ", "&nbsp;")
  End If
  HTMLEncode = sAttrValue
End Function

Public Function HTMLOpen(element As String, Optional attributes As String = "")
  HTMLOpen = "<" & element & " " & attributes & ">"
End Function

Public Function HTMLClose(element As String)
  HTMLClose = "</" & element & ">"
End Function

Public Function HTMLAttrib(sAttrib As String, ByVal sValue As String) As String
  If InStr(sValue, S_QUOT) > 0 Then Err.Raise ERR_HTMLATTR, "HTMLAttrib", "Attribute [" & sValue & "] cannot contain embedded quotes"
  HTMLAttrib = sAttrib & "=" & S_QUOT & sValue & S_QUOT & " "
End Function

Public Function HTMLAttrValue(ByVal vValue As Variant, Optional ByVal ConvertLowerCase As Boolean = False) As String
  Dim s As String
  
  If VarType(vValue) = vbBoolean Then vValue = CLng(vValue)
  s = Replace(IsNullEx(vValue, ""), S_QUOT, "&quot;")
  If ConvertLowerCase Then s = LCase$(s)
  HTMLAttrValue = S_QUOT & s & S_QUOT
End Function
Public Function TypeName(TypeNum As Long) As String
  Select Case TypeNum
    Case adBigInt:            TypeName = "Big integer"
    Case adBinary:            TypeName = "Binary"
    Case adBoolean:           TypeName = "Boolean"
    Case adBSTR:              TypeName = "String"
    Case adChar:              TypeName = "String"
    Case adCurrency:          TypeName = "Currency"
    Case adDate:              TypeName = "Date Time"
    Case adDBDate:            TypeName = "Date"
    Case adDBTime:            TypeName = "Time"
    Case adDBTimeStamp:       TypeName = "Time stamp"
    Case adDecimal:           TypeName = "Decimal"
    Case adDouble:            TypeName = "Double"
    Case adEmpty:             TypeName = "Empty"
    Case adError:             TypeName = "Error"
    Case adGUID:              TypeName = "GUID"
    Case adIDispatch:         TypeName = "IDispatch"
    Case adInteger:           TypeName = "Integer"
    Case adIUnknown:          TypeName = "IUnknown"
    Case adLongVarBinary:     TypeName = "Binary"
    Case adLongVarChar:       TypeName = "String"
    Case adLongVarWChar:      TypeName = "String"
    Case adNumeric:           TypeName = "Numeric"
    Case adSingle:            TypeName = "Single"
    Case adSmallInt:          TypeName = "Small integer"
    Case adTinyInt:           TypeName = "Tiny integer"
    Case adUnsignedBigInt:    TypeName = "Big integer"
    Case adUnsignedInt:       TypeName = "Integer"
    Case adUnsignedSmallInt:  TypeName = "Small integer"
    Case adUnsignedTinyInt:   TypeName = "Tiny integer"
    Case adUserDefined:       TypeName = "User-defined"
    Case adVarBinary:         TypeName = "Binary"
    Case adVarChar:           TypeName = "String"
    Case adVariant:           TypeName = "Variant"
    Case adVarWChar:          TypeName = "String"
    Case adWChar:             TypeName = "String"
    Case Else:                TypeName = "???"
  End Select
End Function

Public Function PreProcessImportFile(ByVal sfileObj As String) As Long
Dim ptrEnd As Long
Dim sline As String
Dim frImportFile As FileSystemObject
Dim tsImport As TextStream

On Error Resume Next
Set frImportFile = New FileSystemObject
Set tsImport = frImportFile.OpenTextFile(sfileObj)

Do While Not tsImport.AtEndOfStream
sline = tsImport.ReadLine
  If Not Len(RTrimAny(RTrimAny(sline, " "), vbTab)) = 0 Then
    ptrEnd = tsImport.Line - 1
  End If
Loop

PreProcessImportFile = ptrEnd
PreProcessImportFile_end:
Set frImportFile = Nothing
Set tsImport = Nothing
End Function
