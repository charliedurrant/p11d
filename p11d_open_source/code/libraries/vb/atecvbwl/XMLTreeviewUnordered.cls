VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "XMLTreeviewUnordered"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Function RecordsetToXML(ByVal rs As Recordset, Optional ByVal sParentID As String = "", Optional sIDRSName As String = "id", Optional sParentIDRSName As String = "parent_id", Optional XMLNames As StringList = Nothing) As String
  RecordsetToXML = RecordsetToXMLEx_org(rs, True, sParentID, sIDRSName, sParentIDRSName, XMLNames, Nothing)
End Function

