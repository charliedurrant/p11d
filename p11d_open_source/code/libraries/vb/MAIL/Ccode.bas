Attribute VB_Name = "Criteria_code"
Option Explicit

Public Function OutputLineMeetsCriteria(ByVal MaxCol As Long, ByVal MaxCriteria As Long, ReportFields() As ReportField, PrintLine As Variant) As Boolean
  Dim RowTrue As Boolean
  Dim rFld As ReportField, Crit As Criterion
  Dim i As Long, j As Long
  
  On Error GoTo OutputLineMeetsCriteria_Err
  Call xSet("OutputLineMeetsCriteria")
  RowTrue = True
  For i = 1 To MaxCriteria
    RowTrue = True
    For j = 1 To MaxCol
      Set rFld = ReportFields(j)
      If rFld.Criteria.Count >= i Then
        Set Crit = rFld.Criteria(i)
        If Not Crit Is Nothing Then
          If Not Crit.MeetsCriterion(PrintLine(j)) Then
            RowTrue = False
            ' Move to check whether next criterion is met
            Exit For
          End If
        End If
      End If
    Next j
    If Frm_RepWiz.Chk_AllCriteriaRequired = vbUnchecked Then
      ' At least one of the criterion is met so row will be outputted
      If RowTrue Then Exit For
    End If
  Next i
  
  OutputLineMeetsCriteria = RowTrue
    
OutputLineMeetsCriteria_End:
  Call xReturn("OutputLineMeetsCriteria")
  Exit Function

OutputLineMeetsCriteria_Err:
  Call ErrorMessage(ERR_ERROR, Err, "OutputLineMeetsCriteria", "Check Line meets Criteria", "Error Checking line for criteria.")
  Resume OutputLineMeetsCriteria_End
  Resume
End Function

Public Function GetMaxCriteria(rFields As Collection) As Long
  Dim rFld As ReportField, MaxCriteria As Long, i As Long
  
  MaxCriteria = 0
  For Each rFld In rFields
    Call rFld.Criteria.CompactTop
    If rFld.Criteria.Count > MaxCriteria Then MaxCriteria = rFld.Criteria.Count
  Next rFld
  GetMaxCriteria = MaxCriteria
End Function

Private Function IsEmptyCriteriaRow(rFields As Collection, ByVal row As Long) As Boolean
  Dim rFld As ReportField
  
  IsEmptyCriteriaRow = True
  For Each rFld In rFields
    If rFld.Criteria.Count >= row Then
      If Not rFld.Criteria.Item(row) Is Nothing Then
        IsEmptyCriteriaRow = False
        Exit Function
      End If
    End If
  Next rFld
End Function

Private Sub RemoveCriteriaRow(rFields As Collection, ByVal row As Long)
  Dim rFld As ReportField, i As Long
  
  For Each rFld In rFields
    If rFld.Criteria.Count >= row Then
      For i = row To (rFld.Criteria.Count - 1)
        Call rFld.Criteria.Remove(i)
        Call rFld.Criteria.AddIndex(rFld.Criteria.Item(i + 1), i)
      Next i
      Call rFld.Criteria.Remove(rFld.Criteria.Count)
    End If
  Next rFld
End Sub

Public Function CompactCriteria(rFields As Collection) As Long
  Dim rFld As ReportField, row As Long

  row = 1
  Do While row <= GetMaxCriteria(rFields)
    If IsEmptyCriteriaRow(rFields, row) Then
      Call RemoveCriteriaRow(rFields, row)
    Else
      row = row + 1
    End If
  Loop
  CompactCriteria = GetMaxCriteria(rFields)
End Function

Public Function CriteriaSQLEx(ByVal Fields As Collection) As String
  Dim rFld As ReportField
  Dim Crit As Criterion
  Dim i As Long, j As Long
  Dim MaxCriteria As Long
  Dim CriteriaString As String
  Dim FieldString As String
  
  On Error GoTo CriteriaSQLEx_Err
  Call xSet("CriteriaSQLEx")
  
  MaxCriteria = GetMaxCriteria(Fields)
  For i = 1 To MaxCriteria
    FieldString = ""
    For j = 1 To Fields.Count
      Set rFld = Fields(j)
      If rFld.Criteria.Count >= i Then
        Set Crit = rFld.Criteria(i)
        If Not Crit Is Nothing Then
          'Add to Criteria String
          If Len(FieldString) = 0 Then
            FieldString = "("
          Else
            FieldString = FieldString & " AND "
          End If
          FieldString = FieldString & Crit.CriterionSQLString(rFld.Name)
        End If
      End If
    Next j
    If Len(FieldString) > 0 Then
      FieldString = FieldString & ")"
    Else
      FieldString = "(1=1)"
    End If
    If Len(CriteriaString) = 0 Then
      CriteriaString = FieldString
    Else
      CriteriaString = CriteriaString & " OR " & FieldString
    End If
  Next i
  CriteriaSQLEx = CriteriaString
    
CriteriaSQLEx_End:
  Call xReturn("CriteriaSQLEx")
  Exit Function

CriteriaSQLEx_Err:
  Call ErrorMessage(ERR_ERROR, Err, "CriteriaSQLEx", "Create Criteria SQL", "Error generating Criteria SQL.")
  Resume CriteriaSQLEx_End
  Resume
End Function

