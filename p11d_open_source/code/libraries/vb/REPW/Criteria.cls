VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Criterion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public Enum CRITERION_COMPARISONS
  CRITERION_COMPARISON_EQ = 1
  CRITERION_COMPARISON_NE
  CRITERION_COMPARISON_LT
  CRITERION_COMPARISON_LE
  CRITERION_COMPARISON_GT
  CRITERION_COMPARISON_GE
  CRITERION_COMPARISON_BW
  CRITERION_COMPARISON_BWEQ
  CRITERION_COMPARISON_LIKE
  CRITERION_COMPARISON_UNDATED
  CRITERION_COMPARISON_DATED
End Enum
Public Comparison As CRITERION_COMPARISONS
Public Value As Variant
Public Value2 As Variant
Public CaseSensitive As Boolean

Private Sub Class_Initialize()
  CaseSensitive = False
End Sub

Public Function ComparisonName(Comparison As CRITERION_COMPARISONS) As String
  Select Case Comparison
    Case CRITERION_COMPARISON_EQ
      ComparisonName = "="
    Case CRITERION_COMPARISON_NE
      ComparisonName = "<>"
    Case CRITERION_COMPARISON_LT
      ComparisonName = "<"
    Case CRITERION_COMPARISON_LE
      ComparisonName = "<="
    Case CRITERION_COMPARISON_GT
      ComparisonName = ">"
    Case CRITERION_COMPARISON_GE
      ComparisonName = ">="
    Case CRITERION_COMPARISON_LIKE
      ComparisonName = "LIKE"
    Case CRITERION_COMPARISON_UNDATED
      ComparisonName = "UNDATED"
    Case CRITERION_COMPARISON_DATED
      ComparisonName = "DATED"
    Case Else
      ComparisonName = "Invalid"
      Call ECASE("Error: Comparison Invalid")
  End Select
End Function

Public Function MeetsCriterion(ByVal TestValue As Variant) As Boolean
  Dim TmpValue As Variant, TmpValue2 As Variant
  
  MeetsCriterion = False

  If (VarType(TestValue) = vbString) And Not CaseSensitive Then
    TestValue = UCase(TestValue)
    TmpValue = UCase(Value)
    TmpValue2 = UCase(Value2)
  Else
    TmpValue = Value
    TmpValue2 = Value2
  End If
  Select Case Comparison
    Case CRITERION_COMPARISON_EQ
      If TestValue = TmpValue Then MeetsCriterion = True
    Case CRITERION_COMPARISON_NE
      If TestValue <> TmpValue Then MeetsCriterion = True
    Case CRITERION_COMPARISON_LT
      If TestValue < TmpValue Then MeetsCriterion = True
    Case CRITERION_COMPARISON_LE
      If TestValue <= TmpValue Then MeetsCriterion = True
    Case CRITERION_COMPARISON_GT
      If TestValue > TmpValue Then MeetsCriterion = True
    Case CRITERION_COMPARISON_GE
      If TestValue >= TmpValue Then MeetsCriterion = True
    Case CRITERION_COMPARISON_BW
      If (TestValue > TmpValue) And (TestValue < TmpValue2) Then MeetsCriterion = True
    Case CRITERION_COMPARISON_BWEQ
      If (TestValue >= TmpValue) And (TestValue <= TmpValue2) Then MeetsCriterion = True
    Case CRITERION_COMPARISON_LIKE
      If TestValue Like TmpValue Then MeetsCriterion = True
      'If TestValue Like TmpValue & "*" Then MeetsCriterion = True
      'If LikeX(TestValue, TmpValue) Then MeetsCriterion = True
    Case CRITERION_COMPARISON_UNDATED
      If TestValue = UNDATED Then MeetsCriterion = True
    Case CRITERION_COMPARISON_DATED
      If TestValue <> UNDATED Then MeetsCriterion = True
    Case Else
  End Select
End Function


Private Function LikeX(ByVal Value1 As Variant, ByVal Value2 As Variant) As Boolean
  If right(Value2, 1) = "|" Then
    Value2 = left(Value2, Len(Value2) - 1)
  Else
    Value2 = Value2 & "*"
  End If
  If Value1 Like Value2 Then LikeX = True
End Function


Public Function AsString()
  Dim sCrit As String
  
  sCrit = ""
  Select Case Comparison
    Case Is < CRITERION_COMPARISON_BW
      sCrit = ComparisonName(Comparison) & CStr(Value)
    Case CRITERION_COMPARISON_BW
      sCrit = "(" & CStr(Value) & " , " & CStr(Value2) & ")"
    Case CRITERION_COMPARISON_BWEQ
      sCrit = "[" & CStr(Value) & " , " & CStr(Value2) & "]"
    Case CRITERION_COMPARISON_LIKE
      sCrit = ComparisonName(Comparison) & " '" & CStr(Value) & "'"
    Case CRITERION_COMPARISON_UNDATED
      sCrit = ComparisonName(Comparison)
    Case CRITERION_COMPARISON_DATED
      sCrit = ComparisonName(Comparison)
    Case Else
  End Select
  'If Not CaseSensitive Then sCrit = sCrit & " ^"
  AsString = sCrit
End Function

Public Function SaveString(Index As Long) As String
  Dim s As String
  
  s = "(CRITERION"
  s = s & "=" & Trim$(CStr(Comparison))
  s = s & "," & AddEscapeCharsWithBrackets(Value)
  s = s & "," & AddEscapeChars(Value2)
  s = s & "," & Trim$(CStr(CaseSensitive))
  s = s & "," & Trim$(CStr(Index))
  s = s & ")"
  
  SaveString = s
End Function

Public Function AddEscapeCharsWithBrackets(ByVal s As String) As String
  If (InStr(s, ")") > 0) Or (InStr(s, ",") > 0) Or (InStr(s, """") > 0) Then
    If Not ((StrComp(left$(s, 1), """") = 0) And (StrComp(Mid$(s, Len(s), 1), """") = 0)) Then
      s = """" & StrDupChar(s, """") & """"
    End If
  End If
  AddEscapeCharsWithBrackets = s
End Function

Public Function CopyTo(CritTo As Criterion) As Boolean
  CritTo.CaseSensitive = CaseSensitive
  CritTo.Comparison = Comparison
  CritTo.Value = Value
  CritTo.Value2 = Value2
End Function

Public Function CriterionSQLString(ByVal Name As String) As String
  Dim SQLValue As String, SQLValue2 As String
  Dim ValueType As DATABASE_FIELD_TYPES
  Dim FirstComparison As String
  Dim SecondComparison As String
  
  ValueType = VarTypetoDatatype(VarType(Value))
  If (ValueType = TYPE_STR) And Not CaseSensitive Then
    SQLValue = GetSQLValue(UCase$(Value), ValueType)
    SQLValue2 = GetSQLValue(UCase$(Value2), ValueType)
  Else
    SQLValue = GetSQLValue(Value, ValueType)
    SQLValue2 = GetSQLValue(Value2, ValueType)
  End If
  If (Comparison = CRITERION_COMPARISON_BW) Or (Comparison = CRITERION_COMPARISON_BWEQ) Then
    If Comparison = CRITERION_COMPARISON_BW Then
      FirstComparison = ComparisonName(CRITERION_COMPARISON_GT)
      SecondComparison = ComparisonName(CRITERION_COMPARISON_LT)
    Else
      FirstComparison = ComparisonName(CRITERION_COMPARISON_GE)
      SecondComparison = ComparisonName(CRITERION_COMPARISON_LE)
    End If
    CriterionSQLString = "([" & Name & "] " & FirstComparison & " " & SQLValue & " AND [" & Name & "] " & SecondComparison & " " & SQLValue2 & ")"
  Else
    CriterionSQLString = "[" & Name & "] " & ComparisonName(Comparison) & " " & SQLValue
  End If
End Function
