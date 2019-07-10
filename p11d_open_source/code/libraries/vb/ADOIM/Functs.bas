Attribute VB_Name = "Functions"
Option Explicit

'MPSFEB
Public Enum TRANSACTION_TYPE
  TRANS_BEGIN = 1
  TRANS_ROLLBACK
  TRANS_COMMIT
End Enum


Public Function TypeName(TypeNum As Long) As String
  Select Case TypeNum
    Case adBigInt:            TypeName = "Big integer"
    Case adBinary:            TypeName = "Binary"
    Case adBoolean:           TypeName = "Boolean"
    Case adBSTR:              TypeName = "String"
    Case adChar:              TypeName = "String"
    Case adCurrency:          TypeName = "Currency"
    Case adDate:              TypeName = "Date/Time"
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


'MPSFEB
Public Function TransactionADO(Conn As Connection, ByVal Action As TRANSACTION_TYPE) As Boolean
  On Error GoTo TransactionADO_err
  
  If Action = TRANS_BEGIN Then
    Conn.BeginTrans
  ElseIf Action = TRANS_COMMIT Then
    Conn.CommitTrans
  ElseIf Action = TRANS_ROLLBACK Then
    Conn.RollbackTrans
  Else
    Err.Raise 380
  End If
  TransactionADO = True
TransactionADO_end:
  Exit Function
  
TransactionADO_err:
  TransactionADO = False
  Resume TransactionADO_end
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



