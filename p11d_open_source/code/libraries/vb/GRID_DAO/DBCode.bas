Attribute VB_Name = "DBCode"
Option Explicit

' apf fix for Field not in Recordset
Public Sub GetPrimaryKeyDAO(aCols As Collection, db As Database, rs As Recordset, Optional ByVal CompleteKeysOnly As Boolean = False)
  Dim i As Long, j As Long, k As Long
  Dim aCol As AutoCol
  Dim TableList As StringList
  Dim rFld As field, tbl As TableDef, idx As Index
    
  On Error GoTo GetPrimaryKeyDAO_err
  If aCols.Count > 0 Then
    ' read all Fields
    Set TableList = New StringList
    For i = 1 To aCols.Count
      Set aCol = aCols.Item(i)
      aCol.PrimaryKey = False
      aCol.SourceField = ""
      aCol.SourceTable = ""
      If Not aCol.UnboundColumn Then
        Set rFld = rs.Fields(aCol.DataField)
        If Len(rFld.SourceField) > 0 Then
          aCol.SourceField = rFld.SourceField
          aCol.SourceTable = rFld.SourceTable
          If Not TableList.IsPresent(aCol.SourceTable) Then Call TableList.Add(aCol.SourceTable)
        End If
      End If
    Next i
    
    ' For each Table get Index Collection
    For i = 1 To TableList.Count
      Set tbl = db.TableDefs(TableList.Item(i))
      For j = 0 To (tbl.Indexes.Count - 1)
        Set idx = tbl.Indexes(j)
        If idx.Primary Then
          If CompleteKeysOnly Then
            For k = 0 To (idx.Fields.Count - 1)
              Set rFld = idx.Fields(k)
              Set aCol = GetAutoCol(aCols, rFld.Name, TableList.Item(i))
              If aCol Is Nothing Then GoTo skip_key
            Next k
          End If
          For k = 0 To (idx.Fields.Count - 1)
            Set rFld = idx.Fields(k)
            Set aCol = GetAutoCol(aCols, rFld.Name, TableList.Item(i))
            If Not aCol Is Nothing Then aCol.PrimaryKey = True
          Next k
skip_key:
          Exit For
        End If
      Next j
    Next i
  End If
  
GetPrimaryKeyDAO_end:
  Exit Sub
  
GetPrimaryKeyDAO_err:
  Call ErrorMessage(ERR_ERROR, Err, "GetPrimaryKeyDAO", "Get Primary Keys for recordset", "Unable to get the primary keys on the recordset")
  Resume GetPrimaryKeyDAO_end
  Resume
End Sub

Public Sub GetPrimaryKeyRDO(aCols As Collection, rsRDO As RDOResultset)
  Dim i As Long, j As Long, k As Long
  Dim aCol As AutoCol
  Dim rFld As rdoColumn
      
  On Error GoTo GetPrimaryKeyRDO_err
  If aCols.Count > 0 Then
    ' read all Fields
    For i = 1 To aCols.Count
      Set aCol = aCols.Item(i)
      aCol.PrimaryKey = False
      aCol.SourceField = ""
      aCol.SourceTable = ""
      
      Set rFld = rsRDO.rdoColumns(aCol.DataField)
      If Len(rFld.SourceColumn) > 0 Then
        aCol.SourceField = rFld.SourceColumn
        aCol.SourceTable = rFld.SourceTable
        'aCol.PrimaryKey = rFld.KeyColumn
      End If
    Next i
    
  End If
  
GetPrimaryKeyRDO_end:
  Exit Sub
  
GetPrimaryKeyRDO_err:
  Call ErrorMessage(ERR_ERROR, Err, "GetPrimaryKeyRDO", "Get Primary Keys for resultset", "Unable to get the primary keys on the resultset")
  Resume GetPrimaryKeyRDO_end
End Sub

Private Function GetAutoCol(ByVal aCols As Collection, ByVal FieldName As String, ByVal TableName As String) As AutoCol
  Dim i As Long, aCol As AutoCol
  
  For i = 1 To aCols.Count
    Set aCol = aCols.Item(i)
    If (StrComp(aCol.SourceTable, TableName, vbTextCompare) = 0) And _
       (StrComp(aCol.SourceField, FieldName, vbTextCompare) = 0) Then
      Set GetAutoCol = aCol
      Exit Function
    End If
  Next i
  Set GetAutoCol = Nothing
End Function

Public Function IsFieldUpdateableDAO(ByVal rs As Recordset, ByVal FieldName As String, Optional vSourceField As Variant) As Boolean
  On Error Resume Next
  IsFieldUpdateableDAO = rs.Fields(FieldName).DataUpdatable
  If IsFieldUpdateableDAO And Not IsMissing(vSourceField) Then IsFieldUpdateableDAO = (Len(vSourceField) > 0)
End Function

