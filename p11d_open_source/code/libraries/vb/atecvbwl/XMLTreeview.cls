VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "XMLTreeview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Function RecordsetToXML(ByVal rs As Recordset, Optional ByVal sParentID As String = "", Optional sIDRSName As String = "id", Optional sParentIDRSName As String = "parent_id", Optional XMLNames As StringList = Nothing) As String
Attribute RecordsetToXML.VB_Description = "Depricated use RecordsetToXML3"
  RecordsetToXML = RecordsetToXMLEx(rs, False, sParentID, sIDRSName, sParentIDRSName, XMLNames, Nothing)
End Function

Public Function RecordSetToXML2(ByVal rs As Recordset, Optional ByVal sParentID As String = "", Optional sIDRSName As String = "id", Optional sParentIDRSName As String = "parent_id", Optional XMLNames As StringList = Nothing, Optional XMLAttribs As StringList = Nothing) As String
Attribute RecordSetToXML2.VB_Description = "Depricated use RecordSetToXML3"
  RecordSetToXML2 = RecordsetToXMLEx(rs, False, sParentID, sIDRSName, sParentIDRSName, XMLNames, XMLAttribs)
End Function
Public Function RecordSetToXML3(ByVal rs As Recordset, Optional ByVal ParentID As String = "-1", Optional ByVal bUnorderedRecordset As Boolean = False, Optional NodeIDFieldName As String = "node_id", Optional NodeParentIDFieldName As String = "parent_node_id", Optional FieldsAsAttributes As StringList = Nothing, Optional FieldsAsNodes As StringList = Nothing) As String
Attribute RecordSetToXML3.VB_Description = "Use to generate the XML required for the JavaScript treeview"
  If FieldsAsNodes Is Nothing Then Set FieldsAsNodes = New StringList
  
  Call FieldsAsNodes.Add(S_NAV_NODE_FIELD_NAME)
  Call FieldsAsNodes.Add(S_NAV_NODE_FIELD_TOOLTIP)
 
  If FieldsAsAttributes Is Nothing Then Set FieldsAsAttributes = New StringList
  If (Not bUnorderedRecordset) Then Call FieldsAsAttributes.Add(S_NAV_ATTRIB_FIELD_CHILDREN)
  
  Call FieldsAsAttributes.Add(S_NAV_ATTRIB_FIELD_IMAGE_CLOSED)
  Call FieldsAsAttributes.Add(S_NAV_ATTRIB_FIELD_IMAGE_OPEN)
  Call FieldsAsAttributes.Add(S_NAV_ATTRIB_FIELD_IMAGE_LEAF)
  Call FieldsAsAttributes.Add(S_NAV_ATTRIB_FIELD_OPEN)
  Call FieldsAsAttributes.Add(S_NAV_ATTRIB_FIELD_SELECTED)
  
  RecordSetToXML3 = XMLTreeviews.RecordsetToXMLEx(rs, bUnorderedRecordset, ParentID, NodeIDFieldName, NodeParentIDFieldName, FieldsAsNodes, FieldsAsAttributes)
End Function

