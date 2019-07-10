Attribute VB_Name = "PAYE_online"
Public Const PAYE_Yes As String = "yes"
Public Const PAYE_No As String = "no"


'PAYEonline - determines which form is being submitted
Public Enum PAYEonline_TYPES
  POT_P46
  POT_P11D
  POT_P11DB_ONLY
End Enum

Public Enum EFiler_Messages
    Efiler_Status
    Efiler_Message
    Efiler_Date
End Enum

Public Enum StatusForm_Types
    PO_Status = 0
    PO_File
    PO_Errors
End Enum
Public Function PAYEBoolYesNo(b As Boolean) As String
  If (b) Then
    PAYEBoolYesNo = PAYE_Yes
  Else
    PAYEBoolYesNo = PAYE_No
  End If
End Function
Public Function IndentXML(oXMLDoc As DOMDocument60, Optional bUnindent As Boolean, Optional bLeaveHeader As Boolean)
    Dim oXSLT       As DOMDocument60
    Dim XSL_FILE    As String
    Dim sResult     As String
    Dim sIndent     As String
    Const QT = """"
    
    Set oXSLT = New DOMDocument60
    
    If bUnindent Then
        sIndent = "no"
    Else
        sIndent = "yes"
    End If
    
    XSL_FILE = _
            "<?xml version=" & QT & "1.0" & QT & " encoding=" & QT & "UTF-8" & QT & "?>" & vbCrLf & _
            "<xsl:stylesheet version=" & QT & "1.0" & QT & " xmlns:xsl=" & QT & "http://www.w3.org/1999/XSL/Transform" & QT & ">" & vbCrLf & _
            "     <xsl:output method=" & QT & "xml" & QT & " version=" & QT & "1.0" & QT & " encoding=" & QT & "UTF-8" & QT & " indent=" & QT & sIndent & QT & "/>" & vbCrLf & _
            "     <xsl:template match=" & QT & "@* | node()" & QT & ">" & vbCrLf & _
            "          <xsl:copy>" & vbCrLf & _
            "               <xsl:apply-templates select=" & QT & "@* | node()" & QT & " />" & vbCrLf & _
            "          </xsl:copy>" & vbCrLf & _
            "     </xsl:template>" & vbCrLf & _
            "</xsl:stylesheet>"


    oXMLDoc.Async = False
    oXSLT.Async = False
    
    oXSLT.loadXML XSL_FILE
    
    If oXSLT.parseError.ErrorCode = 0 Then
        If oXSLT.ReadyState = 4 Then
            sResult = oXMLDoc.transformNode(oXSLT.documentElement)
            ' Get rid of the added header line
            If Not bLeaveHeader Then
                sResult = Replace$(sResult, "<?xml version=" & QT & "1.0" & QT & " encoding=" & QT & "UTF-16" & QT & "?>", vbNullString, , , vbTextCompare)
            End If
            oXMLDoc.loadXML sResult
            
        End If
    Else
        Err.Description = oXSLT.parseError.reason & vbCrLf & _
        "Line: " & oXSLT.parseError.line & vbCrLf & _
        "XML: " & oXSLT.parseError.srcText
        Err.Raise 1006
    End If
    
    Set oXSLT = Nothing
    IndentXML = Replace$(oXMLDoc.xml, vbTab, "  ")
    
End Function

Public Function PrettyFormatXML(xml As String)
  Dim xmlDoc As DOMDocument60
  
  Set xmlDoc = New DOMDocument60
  xmlDoc.loadXML (xml)
  
  PrettyFormatXML = IndentXML(xmlDoc, False, False)
End Function


Public Function PAYEAddXMLChild(XML_doc As DOMDocument60, top_node As IXMLDOMNode, parent_path As String, node_name As String, ByVal node_value As String, Optional UCASE As Boolean, Optional bTrim As Boolean = True, Optional MaxLength As Long = -1) As IXMLDOMNode
    Dim XMLchild As IXMLDOMNode
    Dim XMLparent As IXMLDOMNode
    Dim i As Long
    Dim sNewNodeValue As String
    Dim bInSpace As Boolean
    Dim c As String
    
    On Error GoTo PAYEAddXMLChild_Err
    Call xSet("PAYEAddXMLChild")
      
    'creating child node
    Set XMLchild = XML_doc.createElement(node_name)
    If Not UCASE Then
      node_value = LCase$(node_value)
    End If
    'remove repeated spaces, leave ones at end
    If bTrim Then node_value = TrimEx(node_value)
    For i = 1 To Len(node_value)
      c = Mid$(node_value, i, 1)
      If bInSpace Then
        If StrComp(c, " ") <> 0 Then
          bInSpace = False
        End If
      Else
        If StrComp(c, " ") = 0 Then
          sNewNodeValue = sNewNodeValue & c
          bInSpace = True
        End If
      End If
      If Not bInSpace Then
        sNewNodeValue = sNewNodeValue & c
      End If
    Next
    If (MaxLength <> -1) And (Len(sNewNodeValue) > MaxLength) Then
      sNewNodeValue = Left$(sNewNodeValue, MaxLength)
    End If

    XMLchild.Text = sNewNodeValue
    
    'setting parent
    If Len(parent_path) > 0 Then
        Set XMLparent = top_node.selectSingleNode(parent_path)
    Else
        Set XMLparent = top_node
    End If
    
    'adding child to parent
    XMLparent.appendChild XMLchild
    
    Set PAYEAddXMLChild = XMLchild
    
PAYEAddXMLChild_End:
  Call xReturn("PAYEAddXMLChild")
  Exit Function

PAYEAddXMLChild_Err:
  Call ErrorMessage(ERR_ERROR, Err, "PAYEAddXMLChild", "Error in PAYEAddXMLChild", "Undefined error.")
  Resume PAYEAddXMLChild_End
  Resume
End Function


Public Function PAYEAddXMLChildWithAttribute(XML_doc As DOMDocument60, top_node, parent_path As String, node_name As String, node_value As String, attribute_name As String, attribute_value As String) As IXMLDOMNode
    Dim XMLchild As IXMLDOMNode
    Dim XMLparent As IXMLDOMNode
    Dim att As IXMLDOMAttribute
    
    On Error GoTo PAYEAddXMLChild_Err
    Call xSet("PAYEAddXMLChild")
      
    'creating child
    Set XMLchild = XML_doc.createElement(node_name)
    XMLchild.Text = node_value

    'setting child's attribute
    
    Set att = XML_doc.createAttribute(attribute_name)
    att.Text = attribute_value
    XMLchild.Attributes.setNamedItem att

    'setting parent
    If Len(parent_path) > 0 Then
        Set XMLparent = top_node.selectSingleNode(parent_path)
    Else
        Set XMLparent = top_node
    End If

    'adding child to parent
    XMLparent.appendChild XMLchild
    
    Set PAYEAddXMLChildWithAttribute = XMLchild
PAYEAddXMLChild_End:
  Call xReturn("PAYEAddXMLChild")
  Exit Function

PAYEAddXMLChild_Err:
  Call ErrorMessage(ERR_ERROR, Err, "PAYEAddXMLChild", "Error in PAYEAddXMLChild", "Undefined error.")
  Resume PAYEAddXMLChild_End
  Resume
End Function




Public Function GetXMLAmount(amount, Optional b2DP As Boolean = False) As String
    Dim sFormatString As String
    
    If b2DP Then
      GetXMLAmount = Format$(Round(amount, 2), "###0.00##;")
    Else
      GetXMLAmount = Format$(amount, "##0.00;")
    End If
    
End Function



Public Function bMultipleRecs(bc As BEN_CLASS, ee As Employee, ey As Employer) As Boolean
Dim ben As IBenefitClass
Dim i As Long, j As Long

  Call xSet("bMultipleRecs")

  For i = 1 To ee.benefits.Count
    Set ben = ee.benefits(i)
    If Not (ben Is Nothing) Then
      If ben.BenefitClass = bc Then
        If ben.value(ITEM_BENEFIT_REPORTABLE) Then
        
          
            If IsNumeric(ben.value(ITEM_BENEFIT)) Then
        
              j = j + 1
              If j > 1 Then
                bMultipleRecs = True
                Exit For
              End If
        
            End If
        
          
        End If
      End If
    End If
  Next
  
End Function

Public Function HMITSectionInArray(HMITArr() As HMIT_SECTIONS, section As HMIT_SECTIONS) As Boolean
  Dim i As Long
  
  On Error GoTo HMITSectionInArray_ERR
  
  Call xSet("HMITSectionInArray")
  
  If IsArray(HMITArr) Then
    'If HMITArr(LBound(HMITArr)) Then GoTo HMITSectionInArray_END
    For i = LBound(HMITArr) To UBound(HMITArr)
      If section = HMITArr(i) Then
        HMITSectionInArray = True
        Exit For
      End If
    Next
  Else
    Call Err.Raise(ERR_NOT_ARRAY, "HMITSectionInArray", "The variable passed is not an array.")
  End If
  
HMITSectionInArray_END:
  Call xReturn("HMITSectionInArray")
  Exit Function
HMITSectionInArray_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "HMITSectionInArray", "HMIT Section In Array", "Error determining if HMIT Section is in HMITArr.")
  Resume HMITSectionInArray_END
  Resume
End Function

Public Function CheckBen(ey As IBenefitClass, ee As Employee, ben As IBenefitClass) As Boolean
  Dim sHMITSectionString As String
  Dim sBenefitFormCaption As String
  
  On Error GoTo CheckBen_ERR
  
  Call xSet("CheckBen")
      
  sHMITSectionString = p11d32.Rates.BenClassTo(ben.BenefitClass, BCT_HMIT_SECTION_STRING)
  sBenefitFormCaption = p11d32.Rates.BenClassTo(ben.BenefitClass, BCT_FORM_CAPTION)
    
  Call ben.Calculate
  If ben.value(ITEM_BENEFIT_REPORTABLE) Then
    If Not IsNumeric(ben.value(ITEM_BENEFIT)) Then
      Call Err.Raise(ERR_BEN_INCORRECT, "CheckBen", "Benefit is in error.")
    End If
    CheckBen = True
  Else
    If Not IsNumeric(ben.value(ITEM_BENEFIT)) Then 'Neg value thrown out here
      Call Err.Raise(ERR_BEN_INCORRECT, "CheckBen", "Benefit is in error.")
    End If
  End If
  
CheckBen_END:
  Call xReturn("CheckBen")
  Exit Function
CheckBen_ERR:
  Call ErrorMessagePush(Err)
  If Not ben Is Nothing Then
    If (Err.Number = ERR_BEN_NOT_REPORTABLE) Or (Err.Number = ERR_BEN_INCORRECT) Then
      Call ErrorMessagePop(ERR_ERROR, Err, "CheckBen", FilterMessageTitle(ey.Name, ee.PersonnelNumber, sHMITSectionString, ben.Name, sBenefitFormCaption), "")
    Else
      Call ErrorMessagePop(ERR_ERROR, Err, "CheckBen", FilterMessageTitle(ey.Name, ee.PersonnelNumber, sHMITSectionString, ben.Name, sBenefitFormCaption), "")
    End If
  Else
    Call ErrorMessagePop(ERR_ERROR, Err, "CheckBen", FilterMessageTitle(ey.Name, ee.PersonnelNumber, , "Unknown"), "")
  End If
  Resume CheckBen_END
  Resume
End Function

