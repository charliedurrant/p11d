VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "XMLNavigator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'CAD check
Public Function RecordsetToXML(ByVal rs As Recordset, Optional ByVal sParentID As String = "-1") As String
Attribute RecordsetToXML.VB_Description = "Depricated, Use XMLTreeview.RecordsetToXML3"
  'deprivated see XMLTreeView.RecordsetToXML3
  Dim xmt As XMLTreeview
  Set xmt = New XMLTreeview
  RecordsetToXML = xmt.RecordSetToXML3(rs, sParentID)
End Function
