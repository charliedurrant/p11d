VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ImportSpec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public key As IMPORTFIELD_KEY

Public Tempkey As IMPORTFIELD_KEY

Public vartype As DATABASE_FIELD_TYPES
Public varaction As IMPORTFIELD_ACTION
Public actionparam As String

' this is unique as it is the Name of the destination field in the recordset
Public DestField As String

Public FldName As String              'MS Field Name used for display
Public FldNameDefault As String
Public Hide As Boolean                ' "

Public KeyField As Boolean
Public Required As Boolean

Public CopyField As ImportSpec
Public CopyFieldKey  As Long
  
Public DefaultValue As Variant
Public DefaultNonNullValue As Variant
Public DateDelimiter As String        'MS
Public TimeDelimiter As String        'MS
Public BooleanTrue As String          'MS
Public BooleanFalse As String         'MS
  
' following for fixed format imports
Public Offset As Long
Public multioffset As Long
Public Length As Long

'multi record per line import
'constraints on imports
Public Constraint As IMPORT_CONSTRAINT
Public ConstraintList As StringList
Public Mask As String
Public minvalue As Variant
Public maxvalue As Variant

'calc fields
Public Value As Variant 'Value of this field in import
Public DestTable As String ' name of table to which field belongs
Public FieldKey As Long

Public Function Copy(ByVal ISpecTo As ImportSpec) As ImportSpec
  Dim i As Long
  ISpecTo.Hide = Hide
  ISpecTo.FieldKey = FieldKey
  ISpecTo.actionparam = actionparam
  ISpecTo.Constraint = Constraint
  ISpecTo.DestField = DestField
  ISpecTo.DestTable = DestTable
  ISpecTo.FldName = FldName
  ISpecTo.KeyField = KeyField
  ISpecTo.Required = Required
  ISpecTo.DateDelimiter = DateDelimiter
  ISpecTo.TimeDelimiter = TimeDelimiter
  ISpecTo.BooleanTrue = BooleanTrue
  ISpecTo.BooleanFalse = BooleanFalse
  ISpecTo.key = key
  ISpecTo.Length = Length
  ISpecTo.Mask = Mask
  ISpecTo.maxvalue = maxvalue
  ISpecTo.minvalue = minvalue
  ISpecTo.multioffset = multioffset
  ISpecTo.Offset = Offset
  ISpecTo.Value = Value
  ISpecTo.varaction = varaction
  ISpecTo.vartype = vartype
  ISpecTo.CopyFieldKey = CopyFieldKey
  ISpecTo.DefaultValue = DefaultValue
  ISpecTo.DefaultNonNullValue = DefaultNonNullValue
  ISpecTo.ConstraintList.RemoveAll
  For i = 1 To ConstraintList.Count
    Call ISpecTo.ConstraintList.Add(ConstraintList.Item(i))
  Next i
  Set ISpecTo.CopyField = Nothing
  'ISpecTo.required = required
  'ISpecTo.zerolen = zerolen
  'ISpecTo.fieldsize = fieldsize
  'Set ISpecTo.MergeField = MergeField
  'ISpecTo.MergeType = MergeType
  'ISpecTo.RealName = RealName
  'ISpecTo.MergeSeparator = MergeSeparator
  Set Copy = ISpecTo
End Function

Public Function GetColumn(ByVal ImportType As IMPORT_TYPE) As String
  Dim s As String, i As Long
  
  Select Case key
    Case KEY_FIELD, KEY_SKIPFIELD
      s = s & "(FIELD=" & DestField
      If ImportType = IMPORT_FIXED Then s = s & "," & Offset & "," & Length
      s = s & ")"
      If key = KEY_SKIPFIELD Or (Len(DestField) = 0) Then s = s & "(SKIPCOLUMN)"
    Case KEY_CONSTANT
      s = s & "(CONSTANT=" & DestField & "," & AddEscapeChars(actionparam) & ")"
    Case KEY_COPIEDFIELD
      s = s & "(COPYFIELD=" & DestField & "," & CopyFieldKey & ")"
    'MPSMarch2
    Case KEY_CFGFILENAME
      s = s & "(SPECFILENAME=" & DestField
      If ImportType = IMPORT_FIXED Then s = s & "," & Offset & "," & Length
      s = s & ")"
    Case KEY_DATENOW
      s = s & "(DATENOW=" & DestField
      If ImportType = IMPORT_FIXED Then s = s & "," & Offset & "," & Length
      s = s & ")"
    Case KEY_FILEDATE
      s = s & "(IMPFILEDATE=" & DestField
      If ImportType = IMPORT_FIXED Then s = s & "," & Offset & "," & Length
      s = s & ")"
    Case KEY_FILENAME
      s = s & "(IMPFILENAME=" & DestField
      If ImportType = IMPORT_FIXED Then s = s & "," & Offset & "," & Length
      s = s & ")"
    Case KEY_FILEPATH
      s = s & "(IMPFILEPATH=" & DestField
      If ImportType = IMPORT_FIXED Then s = s & "," & Offset & "," & Length
      s = s & ")"
    Case KEY_IMPDATE
      s = s & "(IMPDATE=" & DestField
      If ImportType = IMPORT_FIXED Then s = s & "," & Offset & "," & Length
      s = s & ")"
    Case KEY_LINENUMBER
      s = s & "(LINENUMBER=" & DestField
      If ImportType = IMPORT_FIXED Then s = s & "," & Offset & "," & Length
      s = s & ")"
    Case KEY_USERNAME
      s = s & "(USERNAME=" & DestField
      If ImportType = IMPORT_FIXED Then s = s & "," & Offset & "," & Length
      s = s & ")"
  End Select
  s = s & "(FIELDNAME=" & AddEscapeChars(FldName) & ")"
  
  Select Case vartype
    Case TYPE_BOOL
      s = s & "(BOOLEAN)"
      If Len(BooleanTrue) > 0 Then s = s & "(BOOLEANTRUE=" & BooleanTrue & ")"
      If Len(BooleanFalse) > 0 Then s = s & "(BOOLEANFALSE=" & BooleanFalse & ")"
    Case TYPE_DATE
      s = s & "(DATE)"
      If varaction = ACTION_REFORMATDELIMITED Then
        If Len(DateDelimiter) > 0 Then
          s = s & "(DATEDELIMITER=" & DateDelimiter
          If Len(actionparam) > 0 Then s = s & "," & actionparam
          s = s & ")"
        End If
        If Len(TimeDelimiter) > 0 Then s = s & "(TIMEDELIMITER=" & TimeDelimiter & ")"
      End If
      If varaction = ACTION_REFORMATFIXEDDATE Then
        s = s & "(FIXEDDATE=" & actionparam & ")"
      End If
      If varaction = ACTION_REFORMATFIXEDDATETIME Then
        s = s & "(FIXEDDATETIME=" & actionparam & ")"
      End If
    Case TYPE_DOUBLE
      s = s & "(FLOAT)"
      If varaction = ACTION_FACTOR Then s = s & "(FACTOR=" & CStr(actionparam) & ")"  'MPSMarch2
    Case TYPE_LONG
      s = s & "(NUMBER)"
      If varaction = ACTION_FACTOR Then s = s & "(FACTOR=" & CStr(actionparam) & ")"  'MPSMarch2
    Case TYPE_STR
      s = s & "(STRING)"
    Case [_TYPE_INVALID]
      ' MPSPHD ?
  End Select
  
   
  If (Constraint And CONSTRAINT_MIN) Then
    If vartype = TYPE_DATE Then ' see ParseMax/Min
      s = s & "(MINVALUE=" & Format$(GetTypedValue(minvalue, vartype), "DDMMYYYY") & ")"
    Else
      s = s & "(MINVALUE=" & GetValueAsString(minvalue, vartype) & ")"
    End If
  End If
  If (Constraint And CONSTRAINT_MAX) Then
    If vartype = TYPE_DATE Then ' see ParseMax/Min
      s = s & "(MAXVALUE=" & Format$(GetTypedValue(maxvalue, vartype), "DDMMYYYY") & ")"
    Else
      s = s & "(MAXVALUE=" & GetValueAsString(maxvalue, vartype) & ")"
    End If
  End If
    
  If (Constraint And CONSTRAINT_MASK) Then s = s & "(MASK=" & Mask & ")"
  If (Constraint And CONSTRAINT_EXCLUDELIST) Or (Constraint And CONSTRAINT_INLIST) Then
    If (Constraint And CONSTRAINT_INLIST) Then
      s = s & "(INLIST="
    Else
      s = s & "(EXLIST="
    End If
    For i = 1 To ConstraintList.Count
      s = s & AddEscapeChars(ConstraintList.Item(i))
      If i <> ConstraintList.Count Then s = s & ","
    Next i
    s = s & ")"
  End If
  If (Constraint And CONSTRAINT_IGNORE_CONVERSIONERRORS) Then s = s & "(IGNORE_CONVERSION_ERRORS)"
  
  If KeyField Then s = s & "(KEYFIELD)"
  If Required Then s = s & "(REQUIRED)"
  
  ' default save needs to be type safe need opposite GetTypedValue
  If IsValidValue(DefaultValue) Then s = s & "(DEFAULT=" & GetValueAsString(DefaultValue, vartype) & ")"
  If IsValidValue(DefaultNonNullValue) Then s = s & "(DEFAULTNONNULL=" & GetValueAsString(DefaultNonNullValue, vartype) & ")"
  If Hide Then s = s & "(HIDEFIELD)"
  GetColumn = s
End Function

Private Sub Class_Initialize()
  KeyField = False
  DateDelimiter = "/"
  TimeDelimiter = ":"
  DefaultValue = Empty
  DefaultNonNullValue = Empty
  Set ConstraintList = New StringList
End Sub

'Public Enum MERGE_TYPE
'  MERGE_NONE = 0
'  MERGE_APPEND
'  MERGE_INSERT
'End Enum


'Public required As Boolean 'PC maps to field property Required
'Public zerolen As Boolean  'PC maps to field property Allow Zero Length
'Public fieldsize As Long   'PC maps to field property Field Size
'Public MergeField As ImportSpec       ' field to merge into
'Public MergeType As MERGE_TYPE        ' apf checkimport cols
'Public RealName As String             ' apf original name
'Public MergeSeparator As String       'MS


' copy Ispec

Private Sub Class_Terminate()

End Sub
