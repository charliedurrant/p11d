Attribute VB_Name = "XMLTreeviews"
Option Explicit

Public Function RecordsetToXMLEx_org(ByVal rs As Recordset, ByVal bReOrder As Boolean, ByVal sParentID As String, ByVal sIDRSName As String, ByVal sParentIDRSName As String, ByVal XMLNames As StringList, ByVal XMLAttribs As StringList) As String
  Dim t0 As Long
  Dim qs As QString, nodesProcessed As Long, nodesProcessedList As Dictionary
  Dim AllFields As Boolean, rsfields As Collection, sNewParentID As String
  Dim fld As field, fldName As String
  
  On Error GoTo RecordsetToXMLEx_org_err
  t0 = GetTicks
  Set qs = New QString
  qs.Increment = QS_INCREMENT
  If (rs.EOF And rs.BOF) Then GoTo RecordsetToXMLEx_org_end
  AllFields = (XMLNames Is Nothing)
  If Not IsFieldPresentADO(rs.Fields, sIDRSName) Then Err.Raise ERR_FIELD_REQUIRED, "RecordsetToXMLEx_org", "Field (" & sIDRSName & ") is required"
  If Not IsFieldPresentADO(rs.Fields, sParentIDRSName) Then Err.Raise ERR_FIELD_REQUIRED, "RecordsetToXMLEx_org", "Field (" & sParentIDRSName & ") is required"
  Set nodesProcessedList = New Dictionary
  Set rsfields = New Collection
    
  For Each fld In rs.Fields
    fldName = LCase$(fld.Name)
    Select Case fldName
      Case sIDRSName
        Call rsfields.Add(AddRSField(fld, "id", fldName, True))
      Case sParentIDRSName
        Call rsfields.Add(AddRSField(fld, "parent_id", fldName, True))
      Case Else
        If AllFields Then
          Call rsfields.Add(AddRSField(fld, fldName, fldName, True))
        Else
          ' do not need to check XMLNames <> nothing as AllFields is false
          If XMLNames.IsPresent(fldName) Then
            Call rsfields.Add(AddRSField(fld, fldName, fldName, True))
          ElseIf XMLAttribs Is Nothing Then
            ' ignore item
          ElseIf XMLAttribs.IsPresent(fldName) Then
            Call rsfields.Add(AddRSField(fld, fldName, fldName, False))
          End If
        End If
    End Select
  Next fld
  
RecordsetToXMLEx_org_end:
  nodesProcessed = -1
  Call qs.Append(ElementOpenEx("nodes", "", False))
  If bReOrder Then
    ' add in all the nodes in the tree
    If Not (rs.BOF And rs.EOF) Then
      rs.Sort = sParentIDRSName
      nodesProcessed = nodesProcessed + NodeXMLEx_Sort(qs, sParentID, sIDRSName, sParentIDRSName, rsfields, True, rs, Nothing, nodesProcessedList)
    End If
    If nodesProcessed <> rs.RecordCount Then Err.Raise ERR_TREEVIEW, "RecordsetToXMLEx", "Unable to construct treeview from unsorted recordset.  Nodes processed does not match the number of records in the recordset"
  Else
    Do While Not rs.EOF
      Call NodeXMLEx(qs, sParentID, sIDRSName, sParentIDRSName, rsfields, True, rs, nodesProcessedList)
      If rs.EOF Then Exit Do
      rs.MoveNext
    Loop
  End If
  Call qs.Append(ElementCloseEx("nodes"))
  RecordsetToXMLEx_org = qs
  Debug.Print rs.Source & ", Size: " & rs.RecordCount & ", Time: " & (GetTicks - t0) & "ms"
  Exit Function

RecordsetToXMLEx_org_err:
  Err.Raise Err.Number, ErrorSourceComponentEx(Err, "RecordsetToXMLEx_org", COMPONENT_NAME), "Record source: " & rs.Source & vbCrLf & Err.Description
  Resume
End Function

Public Function RecordsetToXMLEx(ByVal rs As Recordset, ByVal bReOrder As Boolean, ByVal sParentID As String, ByVal sIDRSName As String, ByVal sParentIDRSName As String, ByVal XMLNames As StringList, ByVal XMLAttribs As StringList) As String
Attribute RecordsetToXMLEx.VB_Description = "Depricated use RecordsetToXML3"
  Dim t0 As Long
  
  Dim qs As QString, nodesProcessed As Long, nodesProcessedList As Dictionary
  Dim AllFields As Boolean, rsfields As Collection, sNewParentID As String
  Dim fld As field, fldName As String
  
  On Error GoTo RecordsetToXMLEx_err
  t0 = GetTicks
  Set qs = New QString
  qs.Increment = QS_INCREMENT
  If (rs.EOF And rs.BOF) Then GoTo RecordsetToXMLEx_end
  
  If Not IsFieldPresentADO(rs.Fields, sIDRSName) Then Err.Raise ERR_FIELD_REQUIRED, "RecordsetToXMLEx", "Field (" & sIDRSName & ") is required"
  If Not IsFieldPresentADO(rs.Fields, sParentIDRSName) Then Err.Raise ERR_FIELD_REQUIRED, "RecordsetToXMLEx", "Field (" & sParentIDRSName & ") is required"
  Set nodesProcessedList = New Dictionary
  Set rsfields = New Collection
    
  ' CAD Not sure why this only applies to bReOder = true and not all attributes used at client ?
  If bReOrder And Not XMLAttribs Is Nothing Then
    If XMLAttribs.IsPresent(S_NAV_ATTRIB_FIELD_CHILDREN) Then
        Call Err.Raise(ERR_NAVIGATOR, "RecordsetToXML", "If creating an unordered xml navigator you can not pass children atribute")
    End If
  End If
  
  For Each fld In rs.Fields
    fldName = LCase$(fld.Name)
    Select Case fldName
      Case S_NAV_ATTRIB_FIELD_OPEN, S_NAV_ATTRIB_FIELD_SELECTED, S_NAV_ATTRIB_FIELD_IMAGE_CLOSED, S_NAV_ATTRIB_FIELD_IMAGE_OPEN, S_NAV_ATTRIB_FIELD_IMAGE_LEAF         'cad
        Call rsfields.Add(AddRSField(fld, fldName, fldName, False))
      Case S_NAV_ATTRIB_FIELD_CHILDREN
        Call rsfields.Add(AddRSField(fld, fldName, fldName, False))
      Case S_NAV_NODE_FIELD_NAME, S_NAV_NODE_FIELD_TOOLTIP
        Call rsfields.Add(AddRSField(fld, fldName, fldName, True))
      Case sIDRSName
        Call rsfields.Add(AddRSField(fld, "id", fldName, True))
      Case sParentIDRSName
        Call rsfields.Add(AddRSField(fld, "parent_id", fldName, True))
      Case Else
        If Not XMLAttribs Is Nothing Then
          If XMLAttribs.IsPresent(fldName) Then
            Call rsfields.Add(AddRSField(fld, fldName, fldName, False))
          Else
            GoTo NAMES
          End If
        Else
NAMES:
          If Not XMLNames Is Nothing Then
            If XMLNames.IsPresent(fldName) Then
              Call rsfields.Add(AddRSField(fld, fldName, fldName, True))
            End If
          End If
        End If
    End Select
  Next fld
  
RecordsetToXMLEx_end:
  nodesProcessed = -1
  Call qs.Append(ElementOpenEx("nodes", "", False))
  If bReOrder Then
    'add in all the nodes in the tree
    If Not (rs.BOF And rs.EOF) Then
      rs.Sort = sParentIDRSName
      nodesProcessed = nodesProcessed + NodeXMLEx_Sort(qs, sParentID, sIDRSName, sParentIDRSName, rsfields, True, rs, Nothing, nodesProcessedList)
    End If
    If nodesProcessed <> rs.RecordCount Then Err.Raise ERR_TREEVIEW, "RecordsetToXMLEx", "Unable to construct treeview from unsorted recordset.  Nodes processed does not match the number of records in the recordset"
  Else
    Do While Not rs.EOF
      Call NodeXMLEx(qs, sParentID, sIDRSName, sParentIDRSName, rsfields, True, rs, nodesProcessedList)
      If rs.EOF Then Exit Do
      rs.MoveNext
    Loop
  End If
  Call qs.Append(ElementCloseEx("nodes"))
  RecordsetToXMLEx = qs
  Debug.Print rs.Source & ", Size: " & rs.RecordCount & ", Time: " & (GetTicks - t0) & "ms"
  Exit Function

RecordsetToXMLEx_err:
  Err.Raise Err.Number, ErrorSourceComponentEx(Err, "RecordsetToXMLEx", COMPONENT_NAME), "Record source: " & rs.Source & vbCrLf & Err.Description
  Resume
End Function

Private Function GetParentCriteria(ByVal sParentIDName As String, ByVal sParentID As String) As String
  If Len(sParentID) = 0 Then
    GetParentCriteria = sParentIDName & "=null"
  Else
    If IsNumeric(sParentID) Then
      GetParentCriteria = sParentIDName & "=" & sParentID
    Else
      GetParentCriteria = sParentIDName & "='" & (sParentID) & "'"
    End If
  End If
End Function

' The recordset is sorted so that all chilren of a particular parent follow that parent
Private Function NodeXMLEx_Sort(ByVal qs As QString, ByVal sParentID As String, ByVal sIDRSName As String, ByVal sParentIDRSName As String, ByVal rsfields As Collection, ByVal toplevel As Boolean, ByVal rsAllNodes As Recordset, ByVal rsChildren As Recordset, ByVal nodesProcessedList As Dictionary) As Long
  Dim sattribs As String, selements As String, sID As String, bRsStart As Boolean
  Dim rsfield As NavigatorRSField, id_fld As field, sDir As SearchDirectionEnum
  Dim bNodes As Boolean, rsNodeChildren As Recordset
  Dim nodesProcessed  As Long
  
  On Error GoTo NodeXMLEx_Sort_err
  nodesProcessed = 1
 
  ' deal with root node
  If rsChildren Is Nothing Then
    If nodesProcessedList.Exists(sParentID) Then Err.Raise ERR_TREEVIEW, "NodeXMLEx", "Node [" & sParentID & "] has already been processed,  cannot have a node appearing more than once in the treeview."
    Call nodesProcessedList.Add(sParentID, -1)
    Set rsChildren = rsAllNodes.Clone
    rsChildren.Filter = GetParentCriteria(sParentIDRSName, sParentID)
  End If
  If rsChildren.EOF Then GoTo NodeXMLEx_Sort_end
  
  Set id_fld = rsChildren.Fields(sIDRSName)
  Do While True
   If (Not toplevel) And (Not bNodes) Then
     qs.Append ElementOpenEx("nodes", "", False)
     bNodes = True
   End If
   sID = id_fld.Value
      
   sattribs = ""
   selements = ""
   For Each rsfield In rsfields
     If rsfield.AsNode Then
       selements = selements & ElementOpenEx(rsfield.Name, "", False)
       selements = selements & XMLTextEx("" & rsChildren.Fields(rsfield.FieldName))
       selements = selements & ElementCloseEx(rsfield.Name)
     Else
       sattribs = sattribs & XMLAttribEx(rsfield.Name, "" & rsChildren.Fields(rsfield.FieldName), True, False)
     End If
   Next rsfield
   
   ' does this node have any children ?
   If nodesProcessedList.Exists(sID) Then Err.Raise ERR_TREEVIEW, "NodeXMLEx_Sort", "Node [" & sParentID & "] has already been processed,  cannot have a node appearing more than once in the treeview."
   Call nodesProcessedList.Add(sID, -1)
   
   Set rsNodeChildren = rsAllNodes.Clone
   rsNodeChildren.Filter = GetParentCriteria(sParentIDRSName, sID)
      
   If rsNodeChildren.BOF Or rsNodeChildren.EOF Then
     sattribs = sattribs & XMLAttribEx("children", "0", True, False)
   Else
     sattribs = sattribs & XMLAttribEx("children", "-1", True, False)
   End If
   Call qs.Append(ElementOpenEx("NODE", sattribs, False))
   Call qs.Append(selements)
   If Not (rsNodeChildren.BOF Or rsNodeChildren.EOF) Then
     nodesProcessed = nodesProcessed + NodeXMLEx_Sort(qs, sID, sIDRSName, sParentIDRSName, rsfields, False, rsAllNodes, rsNodeChildren, nodesProcessedList)
   Else
     nodesProcessed = nodesProcessed + 1
   End If
   qs.Append ElementCloseEx("NODE")
   rsChildren.MoveNext
   If rsChildren.EOF Then Exit Do
  Loop
  If bNodes Then qs.Append ElementCloseEx("nodes")
  
NodeXMLEx_Sort_end:
  NodeXMLEx_Sort = nodesProcessed
  Exit Function
  
NodeXMLEx_Sort_err:
  Err.Raise Err.Number, ErrorSourceComponentEx(Err, "NodeXMLEx_Sort", COMPONENT_NAME), Err.Description
  Resume
End Function


' The recordset is sorted so that all chilren of a particular parent follow that parent
Private Sub NodeXMLEx(ByVal qs As QString, ByVal sParentID As String, ByVal sIDRSName As String, ByVal sParentIDRSName As String, ByVal rsfields As Collection, ByVal toplevel As Boolean, ByVal rs As Recordset, ByVal nodesProcessedList As Dictionary)
   Dim sattribs As String, selements As String
   Dim sID As String
   Dim rsfield As NavigatorRSField, id_fld As field, ParentID_fld As field
   Dim bNodes As Boolean
   
   On Error GoTo NodeXMLEx_ERR
   If nodesProcessedList.Exists(sParentID) Then Err.Raise ERR_TREEVIEW, "NodeXMLEx", "Node [" & sParentID & "] has already been processed,  cannot have a node appearing more than once in the treeview."
   Call nodesProcessedList.Add(sParentID, -1)
   
   If rs.EOF Then GoTo NodeXMLEx_END
   Set id_fld = rs.Fields(sIDRSName)
   Set ParentID_fld = rs.Fields(sParentIDRSName)
   Do While (StrComp(sParentID, "" & ParentID_fld.Value, vbBinaryCompare) = 0)
    If (Not toplevel) And (Not bNodes) Then
      qs.Append ElementOpenEx("nodes", "", False)
      bNodes = True
    End If
    sID = id_fld.Value
    
    sattribs = ""
    selements = ""
    For Each rsfield In rsfields
      If rsfield.AsNode Then
        selements = selements & ElementOpenEx(rsfield.Name, "", False)
        selements = selements & XMLTextEx("" & rsfield.field.Value)
        selements = selements & ElementCloseEx(rsfield.Name)
      Else
        sattribs = sattribs & XMLAttribEx(rsfield.Name, "" & rsfield.field.Value, True, False)
      End If
    Next
    
    Call qs.Append(ElementOpenEx("NODE", sattribs, False))
    Call qs.Append(selements)
    rs.MoveNext
    Call NodeXMLEx(qs, sID, sIDRSName, sParentIDRSName, rsfields, False, rs, nodesProcessedList)
    qs.Append ElementCloseEx("NODE")
    
    If rs.EOF Then Exit Do
   Loop
   If bNodes Then qs.Append ElementCloseEx("nodes")
   
NodeXMLEx_END:
  Exit Sub
  
NodeXMLEx_ERR:
  Err.Raise Err.Number, ErrorSourceComponentEx(Err, "NodeXMLEx", COMPONENT_NAME), Err.Description
  Resume
End Sub




