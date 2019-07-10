Attribute VB_Name = "Common"
Option Explicit

Public Sub UpdateFieldOrders(rFields As Collection)
  Dim rFieldsSorted As Collection
  Dim rFld As ReportField, rFld0 As ReportField
  Dim MaxOrder As Long, i As Long
  
  Set rFieldsSorted = New Collection
  MaxOrder = rFields.Count
  For i = 1 To MaxOrder
    Set rFld0 = Nothing
    For Each rFld In rFields
      If rFld0 Is Nothing Then
        Set rFld0 = rFld
      ElseIf (rFld.Order < rFld0.Order) Then
        Set rFld0 = rFld
      End If
    Next rFld
    rFld0.Order = i
    Call rFieldsSorted.Add(rFld0, rFld0.KeyString)
    Call rFields.Remove(rFld0.KeyString)
no_reorder:
  Next i
  If rFields.Count > 0 Then Call ECASE("UpdateFieldOrders - Field Collection should be empty")
  For Each rFld In rFieldsSorted
    Call rFields.Add(rFld, rFld.KeyString)
  Next rFld
  Set rFieldsSorted = Nothing
End Sub

Public Function PercentDP(ByVal Value As Variant, ByVal dp As Long) As Double
  PercentDP = Min(100, Max(0, RoundN(CDbl(Value), dp)))
End Function

Public Function SortingName(ByVal Sort As SORT_TYPE) As String
  Select Case Sort
    Case SORT_ASCENDING
      SortingName = "Ascending"
    Case SORT_DESCENDING
      SortingName = "Descending"
    Case Else
      SortingName = "None"
      If Sort <> SORT_NONE Then Call ECASE("Error: Sorting Name Invalid")
  End Select
End Function

Public Function AlignmentName(ByVal AlignmentType As ALIGNMENT_TYPE) As String
  Select Case AlignmentType
    Case ALIGN_LEFT
      AlignmentName = "Left"
    Case ALIGN_RIGHT
      AlignmentName = "Right"
    Case ALIGN_CENTER
      AlignmentName = "Centre"
    Case Else
      Call ECASE("Error: AlignmentType Invalid")
  End Select
End Function

Public Function SumName(ByVal SumType As SUM_TYPE) As String
  Select Case SumType
    Case TYPE_NOSUM
      SumName = "None"
    Case TYPE_SUM
      SumName = "Sum"
    Case TYPE_MEAN
      SumName = "Mean"
    Case Else
      Call ECASE("Error: SumType Invalid")
  End Select
End Function

Public Function isNullEx(ByVal v As Variant, Optional ByVal NullValue As Variant = "(Null)") As Variant
  If IsNull(v) Then
    isNullEx = NullValue
  Else
    isNullEx = v
  End If
End Function

Public Function isLongEx(ByVal v As Variant, Optional ByVal DefaultValue As Long = 0) As Long
  Dim l As Long
  
  On Error Resume Next
  isLongEx = DefaultValue
  If IsNumeric(v) Then isLongEx = CLng(v)
End Function


Public Function YesNo(ByVal bValue As Boolean) As String
  If bValue Then
    YesNo = "Yes"
  Else
    YesNo = "No"
  End If
End Function

Public Sub ClearCollection(col As Collection)
  Do While col.Count > 0
    Call col.Remove(1)
  Loop
End Sub

Public Function ProperName(ByVal Name As String) As String
  Name = Trim$(Name)
  ProperName = UCase$(left$(Name, 1)) & LCase$(Mid$(Name, 2))
End Function

Public Sub FrameEnable(frm As Form, fContainer As Frame, ByVal Enable As Boolean)
  Dim Ctrl As Control
  
  On Error Resume Next
  For Each Ctrl In frm
    If (Not TypeOf Ctrl Is CommonDialog) And (Not TypeOf Ctrl Is Menu) Then
      If Ctrl.Container Is fContainer Then
        Ctrl.Enabled = Enable
        If TypeOf Ctrl Is Frame Then Call FrameEnable(frm, Ctrl, Enable)
      End If
    End If
  Next Ctrl
End Sub

Public Function GetScreenTab(ByVal lReportScreen As REPORT_SCREEN) As Long
  'RK 26/10/04
  'Returns relevant tab index for Frm_RepWiz.SSTab
  Select Case lReportScreen
    Case SCREEN_SELECTION
      GetScreenTab = 0
    Case SCREEN_CRITERIA, SCREEN_FORMATS 'Both on Fra_Format
      GetScreenTab = 1
    Case SCREEN_REPORT
      GetScreenTab = 2
    Case SCREEN_FILE_GROUPS
      GetScreenTab = 3
  End Select
End Function

'RK This functionality may be in libraries, but couldn't locate it
Public Function CValDataTypes(FieldDataType As DATABASE_FIELD_TYPES) As datatypes
  Select Case FieldDataType
    Case TYPE_DATE
      CValDataTypes = VT_DATE
    Case TYPE_DOUBLE
      CValDataTypes = VT_DOUBLE
    Case TYPE_LONG
      CValDataTypes = VT_LONG
    Case TYPE_STR
      CValDataTypes = VT_STRING
    Case Else
      CValDataTypes = VT_USER
  End Select
End Function

'RK TAKEN FROM P11D
Public Function DateValReadToScreen(ByVal v As Variant) As String
  Dim s As String
  Dim sDay As String
  Dim sMonth As String
    
  If (IsDate(v)) Then
    sDay = DatePart("d", v)
    If (Len(sDay) < 2) Then sDay = "0" + sDay
    sMonth = DatePart("m", v)
    If (Len(sMonth) < 2) Then sMonth = "0" + sMonth
    s = sDay & "/" & sMonth & "/" & DatePart("yyyy", v)
  Else
    s = DateStringEx(v, v)
  End If
  
  DateValReadToScreen = s
End Function


