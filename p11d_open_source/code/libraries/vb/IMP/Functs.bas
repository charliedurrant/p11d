Attribute VB_Name = "Functions"
Option Explicit

Public Function TypeName(TypeNum As Long) As String
  Select Case TypeNum
    Case 16: TypeName = "Big integer"
    Case 9:  TypeName = "Binary"
    Case 1:  TypeName = "Boolean"
    Case 2:  TypeName = "Byte"
    Case 18: TypeName = "Char"
    Case 5:  TypeName = "Currency"
    Case 8:  TypeName = "Date/Time"
    Case 20: TypeName = "Decimal"
    Case 7:  TypeName = "Double"
    Case 21: TypeName = "Float"
    Case 15: TypeName = "GUID"
    Case 3:  TypeName = "Integer"
    Case 4:  TypeName = "Long"
    Case 11: TypeName = "Long Binary (OLE Object)"
    Case 12: TypeName = "Memo"
    Case 19: TypeName = "Numeric"
    Case 6:  TypeName = "Single"
    Case 10: TypeName = "Text"
    Case 22: TypeName = "Time"
    Case 23: TypeName = "Time Stamp"
    Case 17: TypeName = "VarBinary"
    Case Else: TypeName = "???"
  End Select
End Function



Public Function FieldName(FieldNum As Long, NSourceFields As Long) As String

Dim i As Long, j As Long
Dim s As String

i = FieldNum - NSourceFields
If i > 0 Then
  j = (i - 1) \ 26
  i = i - 26 * j
  If j = 0 Then
    s = "Field_" & Chr$(64 + i)
  Else
    s = "Field_" & Chr$(64 + j) & Chr$(64 + i)
  End If
Else
  s = "Field_" & CStr(Trim$(FieldNum))
End If

FieldName = s

End Function

Public Function ValidateDestination(cDest As DestRecordSet) As Boolean
  Dim fld As Field, FieldList As New StringList
  Dim dValue As DefaultValue, dStatic As DefaultStatic, dConst As DefaultConstraint
  Dim sErr As String, i As Long
  
  On Error GoTo ValidateDestination_Err
  Call xSet("ValidateDestination")
  For Each fld In cDest.rs.Fields
    Call FieldList.Add(fld.Name)
  Next fld
  For i = 1 To cDest.PrimaryKeys.Count
    If Not FieldList.IsPresent(cDest.PrimaryKeys.Item(i)) Then sErr = sErr & "Primary key field '" & cDest.PrimaryKeys.Item(i) & "' not found." & vbCrLf
  Next i
  For i = 1 To cDest.RequiredFields.Count
    If Not FieldList.IsPresent(cDest.RequiredFields.Item(i)) Then sErr = sErr & "Required field '" & cDest.RequiredFields.Item(i) & "' not found." & vbCrLf
  Next i
  For i = 1 To cDest.HiddenFields.Count
    If Not FieldList.IsPresent(cDest.HiddenFields.Item(i)) Then sErr = sErr & "Hidden field '" & cDest.HiddenFields.Item(i) & "' not found." & vbCrLf
  Next i
  For Each dValue In cDest.DefaultValues
    If Not FieldList.IsPresent(dValue.DestField) Then sErr = sErr & "Field '" & dValue.DestField & "' with default value not found." & vbCrLf
  Next dValue
  For Each dStatic In cDest.DefaultStatics
    If Len(dStatic.DestField) > 0 Then
      If Not FieldList.IsPresent(dStatic.DestField) Then sErr = sErr & "Field '" & dStatic.DestField & "' with static link not found." & vbCrLf
    End If
  Next dStatic
  For Each dConst In cDest.Constraints
    If Not FieldList.IsPresent(dConst.DestField) Then sErr = sErr & "Field '" & dConst.DestField & "' with constraint not found." & vbCrLf
  Next dConst
  If Len(sErr) > 0 Then Call Err.Raise(ERR_VALIDATEDEST, "ValidateDestination", vbCrLf & sErr)
  ValidateDestination = True
  
ValidateDestination_End:
  Call xReturn("ValidateDestination")
  Exit Function

ValidateDestination_Err:
  ValidateDestination = False
  Call ErrorMessage(ERR_ERROR, Err, "ValidateDestination", "Validate Import Wizard destination", "Import wizard destination '" & cDest.DisplayName & "' has errors.")
  Resume ValidateDestination_End
End Function


Public Function MakeFieldNamesUnique(FldSpecs As FieldSpecs) As Boolean
  Dim SList As StringList, CurrName As String, i As Long
  Dim j As Long
  
  On Error GoTo MakeFieldNamesUnique_Err
  Call xSet("MakeFieldNamesUnique")
  Set SList = New StringList
  
  For i = 1 To FldSpecs.Count
    j = 0
    CurrName = Trim$(FldSpecs(i).FldName)
    If Len(CurrName) = 0 Then CurrName = "Field_" & CStr(i) & "." & CStr(j)
    Do While SList.IsPresent(CurrName)
      j = j + 1
      CurrName = "Field_" & CStr(i) & "." & CStr(j)
    Loop
    FldSpecs(i).FldName = CurrName
    Call SList.Add(CurrName)
  Next i

MakeFieldNamesUnique_End:
  Call xReturn("MakeFieldNamesUnique")
  Exit Function

MakeFieldNamesUnique_Err:
  Call ErrorMessage(ERR_ERROR, Err, "MakeFieldNamesUnique", "Make Field Names Unique", "There was an error in ensuring that the delimited sorce field names were unique.")
  Resume MakeFieldNamesUnique_End
End Function


Public Function SpecialFieldTypeName(ByVal SpecialFieldType As IMPORTFIELD_KEY) As String 'MPSMarch2
  Select Case SpecialFieldType
    Case KEY_DATENOW
      SpecialFieldTypeName = "Date Now"
    Case KEY_FILENAME
      SpecialFieldTypeName = "File Name"
    Case KEY_FILEPATH
      SpecialFieldTypeName = "File Path"
    Case KEY_FILEDATE
      SpecialFieldTypeName = "File Date"
    Case KEY_LINENUMBER
      SpecialFieldTypeName = "Line Number"
    Case KEY_CFGFILENAME
      SpecialFieldTypeName = "Spec File Name"
    Case KEY_USERNAME
      SpecialFieldTypeName = "User"
    Case KEY_IMPDATE
      SpecialFieldTypeName = "Import Date"
    Case Else
      Call ECASE("FieldTypeName: Unknown SpecialFieldType value")
  End Select
End Function


Public Function IsSpecialFieldType(ByVal FieldType As IMPORTFIELD_KEY) As Boolean 'MPSMarch2
  Select Case FieldType
    Case KEY_DATENOW, KEY_FILENAME, KEY_FILEPATH, KEY_FILEDATE, _
         KEY_LINENUMBER, KEY_CFGFILENAME, KEY_USERNAME, KEY_IMPDATE
      IsSpecialFieldType = True
    Case Else
      IsSpecialFieldType = False
  End Select
End Function

