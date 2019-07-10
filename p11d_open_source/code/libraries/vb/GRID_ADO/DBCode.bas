Attribute VB_Name = "DBCode"
Option Explicit

'Private Sub ADOListTables()
'  Dim cat As New ADOX.Catalog
'  Dim tbl As ADOX.table
'
'  ' Open the catalog
'  cat.ActiveConnection = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
'     "Data Source=.\NorthWind.mdb;"
'  ' Loop through the tables in the database and print their name
'  For Each tbl In cat.Tables
'     If tbl.Type <> "VIEW" Then Debug.Print tbl.Name
'  Next tbl
'End Sub

Private Sub DebugFieldProperties(ByVal fld As field)
  Dim i As Long
  
  For i = 0 To (fld.Properties.Count - 1)
    Debug.Print fld.Properties(i).Name & " = " & fld.Properties(i).Value
  Next i
End Sub

Public Sub GetPrimaryKeyADO(ByVal aCols As Collection, ByVal rs As Recordset, Optional ByVal CompleteKeysOnly As Boolean = False)
  Dim i As Long, j As Long, k As Long
  Dim aCol As AutoCol
  Dim TableList As StringList
  Dim fld As field
  
  On Error GoTo GetPrimaryKeyADO_err
  If aCols.Count > 0 Then
    Set TableList = New StringList
    For i = 1 To aCols.Count
      Set aCol = aCols.Item(i)
      aCol.PrimaryKey = False
      aCol.SourceField = ""
      aCol.SourceTable = ""
      If Not aCol.UnboundColumn Then
        Set fld = rs.Fields(aCol.DataField)
        'Debug.Print fld.Name & " " & IIf(fld.Attributes And adFldIsNullable, "", "NOT NULL")
        aCol.SourceTable = IsNullEx(fld.Properties("BASETABLENAME"), "")
        aCol.SourceField = IsNullEx(fld.Properties("BASECOLUMNNAME"), "")
        aCol.PrimaryKey = IsNullEx(fld.Properties("KEYCOLUMN"), "")
      End If
      'Call DebugFieldProperties(fld)
    Next i
    ' cd DEAL WITH VIEWS
    ' get view sql/ create rs of sql/retrieve keys ( order important )
  End If

GetPrimaryKeyADO_end:
  Exit Sub

GetPrimaryKeyADO_err:
  Call ErrorMessage(ERR_ERROR, Err, "GetPrimaryKeyADO", "Get Primary Keys for recordset", "Unable to get the primary keys on the recordset")
  Resume GetPrimaryKeyADO_end
  Resume
End Sub

Private Function IsFieldUpdateableADOEx(ByVal fld As field) As Boolean
  IsFieldUpdateableADOEx = (fld.Attributes() And (adFldUpdatable Or adFldUnknownUpdatable)) <> 0
End Function

Public Function IsFieldUpdateableADO(ByVal rs As Recordset, ByVal FieldName As String) As Boolean
  On Error Resume Next
  IsFieldUpdateableADO = IsFieldUpdateableADOEx(rs.Fields(FieldName))
End Function

Public Function IsNumberField(ByVal dType As DATABASE_FIELD_TYPES) As Boolean
  IsNumberField = (dType = TYPE_DOUBLE) Or (dType = TYPE_LONG)
End Function
