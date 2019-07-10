Attribute VB_Name = "SQLParse"
Option Explicit

Private Function isWhiteSpace(ByVal s As String) As Boolean
  isWhiteSpace = (s = " ") Or (s = vbTab) Or (s = vbCr) Or (s = vbLf)
End Function

Private Function SkipSpaces(s As String, ByVal Offset As Long) As Long
  Do While isWhiteSpace(Mid$(s, Offset, 1))
    Offset = Offset + 1
  Loop
  SkipSpaces = Offset
End Function

Private Function RemoveExtraBrackets(ByVal s As String) As String
  Dim ch As String, i As Long
  Dim p0 As Long, p1 As Long, pLen As Long, bCount As Long
  ' ((A) or (B)) to (A) or (B)
  
  p0 = SkipSpaces(s, 1)
  ch = Mid$(s, p0, 1)
  pLen = Len(s)
  If ch = CLAUSE_BEGIN Then
    bCount = 1
    For i = (p0 + 1) To pLen
      ch = Mid$(s, i, 1)
      If ch = CLAUSE_BEGIN Then
        bCount = bCount + 1
      ElseIf ch = CLAUSE_END Then
        bCount = bCount - 1
      End If
      If bCount = 0 Then
        p1 = i
        Exit For
      End If
    Next i
    If i > p1 Then Err.Raise ERR_UNMATCHED, "RemoveExtraBrackets", "The sub clause " & s & " does not have matching brackets"
    If SkipSpaces(s, p1 + 1) >= Len(s) Then
      s = RemoveExtraBrackets(Mid$(s, p0 + 1, p1 - p0 - 1))
    End If
  End If
  RemoveExtraBrackets = s
End Function

Private Function FindOperator(Operator As TCSWHERE_LOGICAL_OPERATOR, ClauseString As String, ByVal Offset As Long) As Long
  Dim cItem As String
  Dim ch As String
  Dim i As Long
    
    'CONDITION_SQL_OR
    'CONDITION_SQL_AND
  For i = Offset To Len(ClauseString)
    ch = Mid$(ClauseString, i, 1)
    If IsAlpha(Asc(ch)) Then
      cItem = cItem & ch
    End If
    If (isWhiteSpace(ch) Or (ch = CLAUSE_BEGIN)) And (Len(cItem) > 0) Then
      If StrComp(cItem, CONDITION_SQL_AND, vbTextCompare) = 0 Then
        Operator = LOGICAL_AND
      ElseIf StrComp(cItem, CONDITION_SQL_OR, vbTextCompare) = 0 Then
        Operator = LOGICAL_OR
      Else
        Err.Raise ERR_OPERATOR, "FindOperator", "Expected operator found " & cItem
      End If
      FindOperator = SkipSpaces(ClauseString, i)
      Exit Function
    End If
  Next i
End Function

Private Function FindClause(ConditionName As String, ClauseString As String, ByVal Offset As Long) As Long
  Dim pAlphaChar As Long, bCount As Long
  Dim ch As String
  Dim i As Long
    
  pAlphaChar = 0
  For i = Offset To Len(ClauseString)
    ch = Mid$(ClauseString, i, 1)
    If pAlphaChar = 0 Then
      If IsAlpha(Asc(ch)) Then pAlphaChar = i
    End If
    If ch = CLAUSE_BEGIN Then
      bCount = bCount + 1
    ElseIf ch = CLAUSE_END Then
      bCount = bCount - 1
    End If
    If (bCount = 0) And (pAlphaChar <> 0) Then
      ConditionName = Trim$(Mid$(ClauseString, Offset, i - Offset + 1))
      FindClause = i + 1
      Exit Function
    End If
  Next i
End Function

Private Function IsCondition(ConditionName As String, col As Collection, ClauseString As String, ByVal Offset As Long) As Long
  Dim bCount As Long, pAlphaChar As Long
  Dim nextch As String
  Dim p0 As Long, p1 As Long, i As Long
    
  ConditionName = ""
  pAlphaChar = 0
  For i = Offset To Len(ClauseString)
    nextch = Mid$(ClauseString, i, 1)
    If IsAlpha(Asc(nextch)) Then
      If pAlphaChar = 0 Then pAlphaChar = i
      ConditionName = ConditionName & nextch
      If Len(ConditionName) <> (i - pAlphaChar + 1) Then Exit For
    End If
    If nextch = CLAUSE_BEGIN Then
      bCount = bCount + 1
    ElseIf nextch = CLAUSE_END Then
      bCount = bCount - 1
    End If
    If (isWhiteSpace(nextch) Or (i = Len(ClauseString))) And (pAlphaChar <> 0) Then
      If bCount = 0 Then ' condition found
        ConditionName = Trim$(ConditionName)
        If Len(ConditionName) > 0 Then
          If InCollection(col, ConditionName) Then
            IsCondition = SkipSpaces(ClauseString, i)
          End If
        End If
        Exit For
      End If
    End If
  Next i
End Function

Private Function ProcessClause(ByVal cStack As Stack, col As Collection, ClauseString As String, ByVal Offset As Long) As Long
  Dim wc As whereCondition
  Dim cItem As String
  Dim p1 As Long
    
  p1 = IsCondition(cItem, col, ClauseString, Offset)
  If p1 > 0 Then
    Set wc = col.Item(cItem)
    Call cStack.Push(wc)
  Else
    p1 = FindClause(cItem, ClauseString, Offset)
    If p1 = 0 Then Err.Raise ERR_FINDLHS, "CreateConditionTreeSQL", "Unable to find LHS of clause " & vbCrLf & Mid$(ClauseString, Offset)
    Call cStack.Push(cItem)
  End If
  ProcessClause = p1
End Function


Public Function CreateConditionTreeSQL(col As Collection, ClauseString As String) As whereClause
  Dim cStack As New Stack
  Dim RootClause As whereClause
     
  On Error GoTo CreateConditionTreeSQL_err
  Set RootClause = New whereClause
  Call cStack.Push(ClauseString)
  Call ProcessStack(RootClause, cStack, col)
  
  
CreateConditionTreeSQL_end:
  Set cStack = Nothing
  Set CreateConditionTreeSQL = RootClause
  Exit Function
  
CreateConditionTreeSQL_err:
  Call ErrorMessage(ERR_ERROR, Err, "CreateConditionTreeSQL", "Create SQL Condition", "Unable to construct SQL condition from the Clause " & ClauseString)
  If Not RootClause Is Nothing Then
    Call RootClause.Kill
    Set RootClause = Nothing
  End If
  Resume CreateConditionTreeSQL_end
End Function

Private Sub ProcessStack(ByVal ParentClause As whereClause, ByVal cStack As Stack, col As Collection)
  Dim vStackItem As Variant
  Dim tmpStack As Stack
  Dim lhs As whereClause, rhs As whereClause
  
  Do
    Call cStack.Popv(vStackItem)
    If VarType(vStackItem) = vbObject Then
      If TypeOf vStackItem Is whereCondition Then
        Set ParentClause.Value = vStackItem
      End If
    End If
    If VarType(vStackItem) = vbLong Then
      ' stack item is operator
      ParentClause.Operator = CLng(vStackItem)
      Set lhs = New whereClause
      Set rhs = New whereClause
      Set tmpStack = New Stack
      Call tmpStack.Push(cStack.Pop)
      Call ProcessStack(lhs, tmpStack, col)
      Call tmpStack.Push(cStack.Pop)
      Call ProcessStack(rhs, tmpStack, col)
      Set ParentClause.LHTree = lhs
      Set ParentClause.RHTree = rhs
    End If
    If VarType(vStackItem) = vbString Then
      Call ParseSQL(cStack, col, vStackItem)
    End If
  Loop Until cStack.IsEmpty
End Sub
  


Public Function ParseSQL(ByVal cStack As Stack, col As Collection, ByVal ClauseString As String) As Long
  Dim p0 As Long, vItemRHS As Variant, vItemLHS As Variant
  Dim op As TCSWHERE_LOGICAL_OPERATOR
    
  p0 = 1
  ClauseString = RemoveExtraBrackets(ClauseString)
  ' process lhs
  p0 = ProcessClause(cStack, col, ClauseString, p0)  ' on stack
  If p0 <> Len(ClauseString) Then
    ' process op
    p0 = FindOperator(op, ClauseString, p0)
    ' process rhs
    p0 = ProcessClause(cStack, col, ClauseString, p0) ' on stack
    Call cStack.Popv(vItemRHS)
    Call cStack.Popv(vItemLHS)
            
    ' stack order should be op, lhs, rhs
    Call cStack.Push(vItemRHS)
    Call cStack.Push(vItemLHS)
    Call cStack.Push(op)
  End If
  If p0 <> Len(ClauseString) Then Err.Raise ERR_PARSESQL, "ParseSQL", "Error parsing the sub clause " & ClauseString
End Function


