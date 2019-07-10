Attribute VB_Name = "DBCode"
Option Explicit

' This function must be run twice
' once for in memory properties and again for user defined props
Public Function CopyField(tdDest As TableDef, fdSrc As Field, ByVal mode As COPYTABLE_MODE, ByVal PropFilter As PROPERTIES_FILTER, ByVal UserPropsOnly As Boolean) As Boolean
  Dim fddest As Field
  On Error GoTo CopyField_Err
  Call xSet("CopyField")
  
  If Not UserPropsOnly Then
    If Not InCollection(tdDest.Fields, fdSrc.Name) Then
      Set fddest = tdDest.CreateField(fdSrc.Name, fdSrc.Type, fdSrc.Size)
      tdDest.Fields.Append fddest
    End If
  End If
    
  If (PropFilter And PROP_ALL) = PROP_ALL Then
    Call CopyFieldProperties(tdDest, fdSrc, UserPropsOnly)
  Else
    If Not UserPropsOnly Then
      If (PropFilter And PROP_DEFAULTVALUE) = PROP_DEFAULTVALUE Then Call TrySetDAOProperty(tdDest.Fields(fdSrc.Name).Properties("DefaultValue"), fdSrc.Properties("DefaultValue"))
      If (PropFilter And PROP_ALLOWZEROLENGTH) = PROP_ALLOWZEROLENGTH Then Call TrySetDAOProperty(tdDest.Fields(fdSrc.Name).Properties("AllowZeroLength"), fdSrc.Properties("AllowZeroLength"))
      If (PropFilter And PROP_REQUIRED) = PROP_REQUIRED Then Call TrySetDAOProperty(tdDest.Fields(fdSrc.Name).Properties("Required"), fdSrc.Properties("Required"))
    End If
  End If
  CopyField = True
  
CopyField_End:
  Set fddest = Nothing
  Call xReturn("CopyField")
  Exit Function

CopyField_Err:
  Call ErrorMessage(ERR_ERROR, Err, "CopyField", "Copying Database Field", "Error copying field " & fdSrc.Name & " into Table " & tdDest.Name)
  CopyField = False
  Resume CopyField_End
End Function

Public Sub CopyFieldProperties(tdDest As TableDef, fdSrc As Field, ByVal UserPropsOnly As Boolean)
  Dim prop As Property
   
  For Each prop In fdSrc.Properties
    If DAOPropertyExists(tdDest.Fields(fdSrc.Name).Properties, prop.Name) Then
      If Not UserPropsOnly Then
        Call TrySetDAOProperty(tdDest.Fields(fdSrc.Name).Properties(prop.Name), prop)
      End If
    ElseIf UserPropsOnly Then
      Call TryCreateDAOProperty(tdDest.Fields(fdSrc.Name), prop)
    End If
  Next prop
End Sub
  

Public Function AddIndexes(td As TableDef, cIndexes As Collection, Optional Overwrite As Boolean = True) As Boolean
  Dim idx As Index
  On Error GoTo AddIndexes_Err
  Call xSet("AddIndexes")
  
  AddIndexes = True
  For Each idx In cIndexes
    AddIndexes = AddIndexes And AddIndex(td.Indexes, idx, Overwrite)
  Next idx

AddIndexes_End:
  Call xReturn("AddIndexes")
  Exit Function

AddIndexes_Err:
  AddIndexes = False
  Call ErrorMessage(ERR_ERROR, Err, "AddIndexes", "ERR_UNDEFINED", "Undefined error.")
  Resume AddIndexes_End
End Function

Private Function AddIndex(cIndexes As Indexes, idx As Index, Overwrite As Boolean) As Boolean
  Dim prop As Property
  Dim newidx As Index
  Dim fld As Field
  Dim newfld As Field
  
  On Error GoTo AddIndex_err
  Call xSet("AddIndex")
  If InCollection(cIndexes, idx.Name) Then
    If Overwrite Then
      Call cIndexes.Delete(idx.Name)
    Else
      AddIndex = True
      GoTo AddIndex_end
    End If
  End If
  
  Set newidx = New Index
  For Each prop In idx.Properties
    Call TrySetDAOProperty(newidx.Properties(prop.Name), prop)
  Next prop
  For Each fld In idx.Fields
    Set newfld = newidx.CreateField(fld.Name)
    newidx.Fields.Append newfld
  Next fld
  Call cIndexes.Append(newidx)
  AddIndex = True
  
AddIndex_end:
  Call xReturn("Addindex")
  Exit Function
  
AddIndex_err:
  AddIndex = False
  Resume AddIndex_end
End Function

' CopyIndex copies/removes index
' if cIndexes is nothing then does not copy the index
' if idsFrom is nothing then does not remove the index
Public Function CopyIndex(cIndexes As Collection, idx As Index, idsFrom As Indexes) As Boolean
  Dim prop As Property
  Dim newidx As Index
  Dim fld As Field
  Dim newfld As Field
  On Error GoTo RemoveIndex_Err
  Call xSet("RemoveIndex")
  
  'Copy index into newidx and add to collection
  If Not cIndexes Is Nothing Then
    Set newidx = New Index
    For Each prop In idx.Properties
      Call TrySetDAOProperty(newidx.Properties(prop.Name), prop)
    Next prop
    For Each fld In idx.Fields
      Set newfld = newidx.CreateField(fld.Name)
      newidx.Fields.Append newfld
    Next fld
    Call cIndexes.Add(newidx)
  End If
  
  'Remove if idsFrom specified
  If Not idsFrom Is Nothing Then Call idsFrom.Delete(idx.Name)
  CopyIndex = True
      
RemoveIndex_End:
  Set newidx = Nothing
  Call xReturn("RemoveIndex")
  Exit Function

RemoveIndex_Err:
  CopyIndex = False
  Call ErrorMessage(ERR_ERROR, Err, "RemoveIndex", "ERR_DELIDX", "Error removing a table index.")
  Resume RemoveIndex_End
End Function

Private Sub TrySetDAOProperty(pDest As Property, pSrc As Property)
  On Error Resume Next
  pDest.Value = pSrc.Value
End Sub


'* Pass this function a property and it will return TRUE if
'* it is Read Only
Private Function isDAOPropertyReadOnly(prop As Property) As Boolean
  Dim v As Variant
  
  On Error GoTo isDAOPropertyReadOnly_Err
  Call xSet("isDAOPropertyReadOnly")
  isDAOPropertyReadOnly = True
  v = prop.Value
  prop.Value = v
  isDAOPropertyReadOnly = False
  
isDAOPropertyReadOnly_End:
  Call xReturn("isDAOPropertyReadOnly")
  Exit Function

isDAOPropertyReadOnly_Err:
  Resume isDAOPropertyReadOnly_End
End Function

Private Function DAOPropertyExists(daoProps As Properties, ByVal Name As String) As Boolean
  Dim s As String
  On Error GoTo DAOPropertyExists_Err
  
  s = daoProps(Name).Name
  DAOPropertyExists = True
DAOPropertyExists_End:
  Exit Function
  
DAOPropertyExists_Err:
  DAOPropertyExists = False
  Resume DAOPropertyExists_End
End Function

' note you cannot create an in memory user defined prop
' underlyinbg table must exist
Private Sub TryCreateDAOProperty(obj As Field, prop As Property)
  Dim propnew As Property
  
  On Error GoTo TryCreateDAOProperty_err
  Set propnew = obj.CreateProperty(prop.Name, prop.Type, prop.Value)
  obj.Properties.Append propnew
  
TryCreateDAOProperty_end:
  Exit Sub
  
TryCreateDAOProperty_err:
  Resume TryCreateDAOProperty_end
End Sub
      
Public Function IndexIdentical(idx1 As Index, idx2 As Index) As Boolean
  IndexIdentical = (idx1.IgnoreNulls = idx2.IgnoreNulls) And (idx1.Primary = idx2.Primary) And (idx1.Unique = idx2.Unique)
End Function

Public Function ConnectAllEx(DestDb As Database, SourceDb As Database, ByVal CopyQueries As Boolean, ByVal TablePrefix As String, ByVal FilterTablePrefix As String) As Boolean
  Dim FilterTablePrefixLen As Long
  Dim tdSrc As TableDef, tdDest As TableDef
  Dim sConnect As String
  
  On Error GoTo ConnectAllEx_Err
  FilterTablePrefixLen = Len(FilterTablePrefix)
  sConnect = ";DATABASE=" & SourceDb.Name
  For Each tdSrc In SourceDb.TableDefs
    If Not IsSysTableEx(tdSrc) Then
      If FilterTablePrefixLen > 0 Then
        If StrComp(Left$(tdSrc.Name, FilterTablePrefixLen), FilterTablePrefix, vbTextCompare) <> 0 Then GoTo next_table
      End If
      If Not InCollection(DestDb.TableDefs, TablePrefix & tdSrc.Name) Then
        Set tdDest = DestDb.CreateTableDef(TablePrefix & tdSrc.Name)
        tdDest.Connect = sConnect
        tdDest.SourceTableName = tdSrc.Name
        DestDb.TableDefs.Append tdDest
      End If
      Set tdDest = DestDb.TableDefs(TablePrefix & tdSrc.Name)
      If StrComp(tdDest.Connect, sConnect, vbTextCompare) <> 0 Then
        tdDest.Connect = sConnect
        tdDest.RefreshLink
      End If
    End If
next_table:
  Next tdSrc
  ConnectAllEx = True

ConnectAllEx_End:
  Exit Function

ConnectAllEx_Err:
  ConnectAllEx = False
  Resume ConnectAllEx_End
End Function

Public Function RemoveLinkedTablesEx(Db As Database, ByVal TablePrefix As String) As Boolean
  Dim td As TableDef, i As Long
  
  On Error GoTo RemoveLinkedTablesEx_Err
  
  For i = (Db.TableDefs.Count - 1) To 0 Step -1
    Set td = Db.TableDefs(i)
    If Not IsSysTableEx(td) Then
      If (Len(td.Connect) > 0) And _
         (InStr(1, td.Name, TablePrefix, vbTextCompare) = 1) Then
        Db.TableDefs.Delete td.Name
      End If
    End If
  Next i
  RemoveLinkedTablesEx = True

RemoveLinkedTablesEx_End:
  Exit Function

RemoveLinkedTablesEx_Err:
  RemoveLinkedTablesEx = False
  Resume RemoveLinkedTablesEx_End
End Function


Public Function IsSysTableEx(td As TableDef) As Boolean
  IsSysTableEx = (td.Attributes And dbSystemObject)
End Function

Public Function RepairCompactDBEx(ByVal DatabasePath As String, ByVal mode As REPAIRCOMPACT_MODE, ByVal ShowErrors As Boolean) As Boolean
  Dim tmpPath As String, dbDir As String, dbFileName As String
  Dim msg As String
    
  On Error GoTo RepairCompactDBEx_Err
  Call SetCursor
  If FileExists(DatabasePath) Then
    If IsDatabaseOpen(DatabasePath) Then Err.Raise ERR_DBOPEN, "RepairCompactDBEx", "Unable to Repair or Compact the database as it already open."
    Call SplitPath(DatabasePath, dbDir, dbFileName)
    tmpPath = GetTempFileName(dbDir, Left$(dbFileName & "XXX", 3))
    If Len(tmpPath) > 0 Then
      DoEvents
      Call DBEngine.Idle(dbFreeLocks)
      If (mode And MODE_COMPACT) = MODE_COMPACT Then msg = "Compact "
      If (mode And MODE_REPAIR) = MODE_REPAIR Then
        If Len(msg) > 0 Then msg = msg & "and "
        msg = msg & "Repair "
      End If
      Call DisplayMessagePopup(Nothing, "Performing " & msg & "on" & vbCrLf & "database " & DatabasePath, msg & "database")
      'DAO 3.6 no longer supports RepairDatabase() instead CompactDatabase performs both functions. RD 01/10/2002
      'If (mode And MODE_REPAIR) = MODE_REPAIR Then Call DBEngine.CompactDatabase(DatabasePath)
      'If (mode And MODE_COMPACT) = MODE_COMPACT Then
        Call DBEngine.CompactDatabase(DatabasePath, tmpPath)
        Call Kill(DatabasePath)
        Name tmpPath As DatabasePath
      'End If
      Call DBEngine.Idle(dbFreeLocks)
      RepairCompactDBEx = True
    End If
  End If

RepairCompactDBEx_End:
  Call DisplayMessageKill
  Call ClearCursor
  Exit Function

RepairCompactDBEx_Err:
  RepairCompactDBEx = False
  If ShowErrors Then Call ErrorMessage(ERR_ERROR, Err, "RepairCompactDBEx", "Repair/Compact database", "Error executing Repair/Compact on database:" & vbCrLf & DatabasePath & vbCrLf)
  Resume RepairCompactDBEx_End
End Function
