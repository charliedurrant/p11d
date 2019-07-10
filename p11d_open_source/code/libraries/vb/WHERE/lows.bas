Attribute VB_Name = "lows"
Option Explicit

Public Function SQLFieldName(ByVal FieldName As String) As String
  If DatabaseTarget = DB_TARGET_JET Then
    SQLFieldName = "[" & FieldName & "]"
  Else
    SQLFieldName = FieldName
  End If
End Function

Public Function ValidateValue(op As TCSWHERE_CONDITIONS, Val As Variant) As Boolean
  
  On Error GoTo ValidateValue_err
    
  ValidateValue = True
  Select Case op
    Case STR_BEGINS, STR_CONTAINS, STR_ENDS, STR_EQUALS, STR_NOT_INCLUDE
      If Len(Val) = 0 Then Call Err.Raise(ERR_VALID_FAIL, "ValidateValue", "You have not entered a value for the criteria.")
    Case NUM_EQUAL_TO, NUM_GREATER_OR_EQUAL, NUM_GREATER_THAN, NUM_ISEMPTY, NUM_LESS_OR_EQUAL, NUM_LESS_THAN, NUM_NOT_EQUAL
      If Len(Val) = 0 Then Call Err.Raise(ERR_VALID_FAIL, "ValidateValue", "You have not entered a value for the criteria.")
      If Not IsNumeric(Val) Then Call Err.Raise(ERR_VALID_FAIL, "ValidateValue", "The value that you have entered is not a valid number.")
    Case DT_AFTER, DT_BEFORE, DT_NOT_ON, DT_ON
      If Len(Val) = 0 Then Call Err.Raise(ERR_VALID_FAIL, "ValidateValue", "You have not entered a value for the criteria.")
      If TryConvertDate(Val) = UNDATED Then Call Err.Raise(ERR_VALID_FAIL, "ValidateValue", "The value that you entered is not a valid date.")
  End Select

ValidateValue_end:
  Exit Function
ValidateValue_err:
  ValidateValue = False
  If Err.Number = ERR_VALID_FAIL Then
    Call ErrorMessage(ERR_ERROR, Err, "ValidateValue", "Validate condition", "The condition which you are trying to add is not valid.")
  Else
    Call ErrorMessage(ERR_ERROR, Err, "ValidateValue", "Validate condition", "An unexpected error occurred whilst trying to validate the condition.")
  End If
  Resume ValidateValue_end
  Resume
End Function


Public Function MergeClauses(ByVal RootClause As whereClause, LHClause As whereClause, RHClause As whereClause, ByVal Operator As TCSWHERE_LOGICAL_OPERATOR) As whereClause
  Dim NewRootClause As whereClause
  Dim NewClause As whereClause
  Dim LParent As whereClause, RParent As whereClause
  
  If LHClause Is RHClause Then Err.Raise ERR_MERGECLAUSES, "MergeClauses", "Cannot merge Left and Right clauses as they are the same"
  If (RootClause Is Nothing) And ((LHClause Is Nothing) Or (RHClause Is Nothing)) Then
    'No root clause - Set root clause to L/R
    If Not LHClause Is Nothing Then Set NewRootClause = LHClause
    If Not RHClause Is Nothing Then Set NewRootClause = RHClause
    If NewRootClause Is Nothing Then Err.Raise ERR_MERGECLAUSES, "MergeClauses", "Both clauses to merge are Nothing"
  ElseIf (RootClause Is Nothing) Or (RootClause Is LHClause) Or (RootClause Is RHClause) Then
    'One of the clauses is the root clause or there is no root clause
    Set NewRootClause = New whereClause
    Set NewRootClause.LHTree = LHClause
    Set NewRootClause.RHTree = RHClause
    NewRootClause.Operator = Operator
  Else
    ' Everything else
    LParent = FindParent(RootClause, LHClause)
    RParent = FindParent(RootClause, RHClause)
    If (LParent Is Nothing) And (RParent Is Nothing) Then Err.Raise ERR_NOTINROOT, "MergeClauses", "Neither of the two clauses are children of the Root Clause, they therefore cannot be merged."
    If (Not LParent Is Nothing) And (Not RParent Is Nothing) Then Err.Raise ERR_NOTINROOT, "MergeClauses", "Both of the two clauses are children of the Root Clause, they therefore cannot be merged."
    
    ' Just one clause has been found create new intermediate clause
    Set NewClause = New whereClause
    Set NewClause.LHTree = LHClause
    Set NewClause.RHTree = RHClause
    NewClause.Operator = Operator
    If Not LParent Is Nothing Then
      If LParent.LHTree Is LHClause Then
        Set LParent.LHTree = NewClause
      Else
        Set LParent.RHTree = NewClause
      End If
    Else
      If RParent.LHTree = LHClause Then
        Set RParent.LHTree = NewClause
      Else
        Set RParent.RHTree = NewClause
      End If
    End If
    Set NewRootClause = RootClause
  End If
  Set MergeClauses = NewRootClause
End Function

'Returns the parent whereClause which contains testClause as either LHS or RHS
Private Function FindParent(ParentClause As whereClause, SubClause As whereClause) As whereClause
  If (ParentClause.LHTree Is SubClause) Or (ParentClause.RHTree Is SubClause) Then
    Set FindParent = ParentClause
  Else
    If Not ParentClause.LHTree Is Nothing Then
      Set FindParent = FindParent(ParentClause.LHTree, SubClause)
    End If
    If (Not ParentClause.RHTree Is Nothing) And (FindParent Is Nothing) Then
      Set FindParent = FindParent(ParentClause.RHTree, SubClause)
    End If
  End If
End Function

Private Function FindClause(ParentClause As whereClause, ByVal ConditionName As String) As whereClause
  If Not ParentClause.Value Is Nothing Then
    If StrComp(ParentClause.Value.Name, ConditionName, vbBinaryCompare) = 0 Then
      Set FindClause = ParentClause
    End If
  Else
    If Not ParentClause.LHTree Is Nothing Then
      Set FindClause = FindClause(ParentClause.LHTree, ConditionName)
    End If
    If (Not ParentClause.RHTree Is Nothing) And (FindClause Is Nothing) Then
      Set FindClause = FindClause(ParentClause.RHTree, ConditionName)
    End If
  End If
End Function

Public Sub FillConditions(ParentClause As whereClause, Conds As Collection)
  If Not ParentClause.Value Is Nothing Then
    If Not InCollection(Conds, ParentClause.Value.Name) Then
      Call Conds.Add(ParentClause.Value, ParentClause.Value.Name)
    End If
  Else
    If Not ParentClause.LHTree Is Nothing Then
      Call FillConditions(ParentClause.LHTree, Conds)
    End If
    If Not ParentClause.RHTree Is Nothing Then
      Call FillConditions(ParentClause.RHTree, Conds)
    End If
  End If
End Sub
Private Function SkipChars(s As String, ByVal p As Long, ByVal CharsToSkip As String) As Long
  Dim nextchar As String
  
  Do While p < Len(s)
    nextchar = Mid$(s, p, 1)
    If InStr(1, CharsToSkip, nextchar, vbTextCompare) = 0 Then Exit Do
    p = p + 1
  Loop
  SkipChars = p
End Function
    
Public Function NextClauseIndex(ByVal NextIndex As String) As String
  Dim LastChar As String, FirstChar As String
  Dim i As Long, j As Long
  
  LastChar = Right$(NextIndex, 1)
  If LastChar <> "Z" Then
    NextIndex = Left$(NextIndex, Len(NextIndex) - 1) & Chr$(Asc(LastChar) + 1)
  Else
    For i = Len(NextIndex) To 1 Step -1
      FirstChar = Mid$(NextIndex, i, 1)
      If FirstChar <> "Z" Then
        Mid$(NextIndex, i, 1) = Chr$(Asc(FirstChar) + 1)
        Exit For
      End If
    Next i
    For j = (i + 1) To Len(NextIndex)
      Mid$(NextIndex, j, 1) = "A"
    Next j
    If i = 0 Then NextIndex = NextIndex & "A"
  End If
  NextClauseIndex = NextIndex
End Function

Public Function DeleteCondition(ByVal ConditionName As String, ByVal RootClause As whereClause) As whereClause
  Dim wCL As whereClause, wCLParent As whereClause, wCLGrandParent As whereClause
  Dim wCLBranch As whereClause
  
  Set wCL = FindClause(RootClause, ConditionName)
  If wCL Is Nothing Then Err.Raise ERR_DELETECONDITION, "DeleteCondition", "Unable to delete condition " & ConditionName & " as the condition could not be found."
  If RootClause Is wCL Then
    Call wCL.Kill
    Set DeleteCondition = Nothing
  Else
    Set wCLParent = FindParent(RootClause, wCL)
    If wCLParent Is RootClause Then
      If wCL Is RootClause.LHTree Then
        Set wCLBranch = RootClause.RHTree
        Set RootClause.RHTree = Nothing
      End If
      If wCL Is RootClause.RHTree Then
        Set wCLBranch = RootClause.LHTree
        Set RootClause.LHTree = Nothing
      End If
      Call RootClause.Kill
      Set DeleteCondition = wCLBranch
    Else
      Set wCLGrandParent = FindParent(RootClause, wCLParent)
      If wCL Is wCLParent.LHTree Then
        Set wCLBranch = wCLParent.RHTree
        Set wCLParent.RHTree = Nothing
      End If
      If wCL Is wCLParent.RHTree Then
        Set wCLBranch = wCLParent.LHTree
        Set wCLParent.LHTree = Nothing
      End If
      Call wCLParent.Kill
      
      If wCLParent Is wCLGrandParent.LHTree Then
        Set wCLGrandParent.LHTree = wCLBranch
      End If
      If wCLParent Is wCLGrandParent.RHTree Then
        Set wCLGrandParent.RHTree = wCLBranch
      End If
      Set DeleteCondition = RootClause
    End If
  End If
End Function

Private Function CreateConditionTreeInternal(Parent As whereClause, col As Collection, ClauseString As String, ByVal cOffset As Long, ByVal EndChar As String) As Long
  Dim cItem As String, p As Long
  Dim wc As whereCondition
  Dim lhs As whereClause, rhs As whereClause
  
  If InStr(cOffset, ClauseString, CONDITION_OR, vbBinaryCompare) = cOffset Then
    Set lhs = New whereClause
    Set rhs = New whereClause
    Set Parent.LHTree = lhs
    Set Parent.RHTree = rhs
    Parent.Operator = LOGICAL_OR
    cOffset = cOffset + Len(CONDITION_OR)
    cOffset = CreateConditionTreeInternal(lhs, col, ClauseString, cOffset, CLAUSE_SEP)
    cOffset = CreateConditionTreeInternal(rhs, col, ClauseString, cOffset, CLAUSE_END)
  ElseIf InStr(cOffset, ClauseString, CONDITION_AND, vbBinaryCompare) = cOffset Then
    Set lhs = New whereClause
    Set rhs = New whereClause
    Set Parent.LHTree = lhs
    Set Parent.RHTree = rhs
    Parent.Operator = LOGICAL_AND
    cOffset = cOffset + Len(CONDITION_AND)
    cOffset = CreateConditionTreeInternal(lhs, col, ClauseString, cOffset, CLAUSE_SEP)
    cOffset = CreateConditionTreeInternal(rhs, col, ClauseString, cOffset, CLAUSE_END)
  Else
    If Len(EndChar) = 0 Then     ' deal single with condition
      cItem = Mid$(ClauseString, cOffset)
      p = cOffset
    Else
      p = InStr(cOffset, ClauseString, EndChar, vbBinaryCompare)
      If p = 0 Then Err.Raise ERR_CREATECLAUSETREE, "CreateConditionTreeInternal", "Unable to create clause tree from clause."
      cItem = Mid$(ClauseString, cOffset, p - cOffset)
    End If
    Set wc = col.Item(cItem)
    Set Parent.Value = wc
    cOffset = SkipChars(ClauseString, p, CLAUSE_SEP & CLAUSE_END)
  End If
  CreateConditionTreeInternal = cOffset
End Function


Public Function CreateClauseFromInternal(InternalFormat As String) As whereClause
  Dim RootClause As whereClause
  Dim Conds As Collection, wc As whereCondition
  Dim Conditions As String, Clauses As String
  Dim p0 As Long, p1 As Long
  
  ' parse conditions
  On Error GoTo CreateClauseFromInternal_err
  If Len(InternalFormat) = 0 Then GoTo CreateClauseFromInternal_end
  Set Conds = New Collection
  p0 = 1
  Do
    If StrComp(Mid$(InternalFormat, p0, Len(CONDITION_SEP)), CONDITION_SEP, vbBinaryCompare) = 0 Then
      p1 = 0
    Else
      p1 = InStr(p0, InternalFormat, CONDITION_FN, vbBinaryCompare)
      If p1 = 0 Then Err.Raise ERR_PARSEINTERNALFORMAT, "CreateClauseFromInternal", "Error parsing internal format string at position " & p0 & "." & vbCrLf & InternalFormat
      p0 = p1 + Len(CONDITION_FN)
      
      Set wc = New whereCondition
      p1 = InStr(p0, InternalFormat, CLAUSE_SEP, vbBinaryCompare)
      wc.Name = Mid$(InternalFormat, p0, p1 - p0)
      
      p0 = p1 + 1
      p1 = InStr(p0, InternalFormat, CLAUSE_SEP, vbBinaryCompare)
      wc.Field = Mid$(InternalFormat, p0, p1 - p0)
      
      p0 = p1 + 1
      p1 = InStr(p0, InternalFormat, CLAUSE_SEP, vbBinaryCompare)
      wc.DataType = Mid$(InternalFormat, p0, p1 - p0)
      
      p0 = p1 + 1
      p1 = InStr(p0, InternalFormat, CLAUSE_SEP, vbBinaryCompare)
      wc.Operator = Mid$(InternalFormat, p0, p1 - p0)
      
      p0 = p1 + 1
      If Mid$(InternalFormat, p0, Len(VALUE_BEGIN)) <> VALUE_BEGIN Then Err.Raise ERR_PARSEINTERNALFORMAT, "CreateClauseFromInternal", "Error parsing internal format string at position " & p0 & "." & vbCrLf & "Value not valid" & vbCrLf & InternalFormat
      p0 = p0 + Len(VALUE_BEGIN)
      p1 = InStr(p0, InternalFormat, VALUE_END, vbBinaryCompare)
      wc.Value = Mid$(InternalFormat, p0, p1 - p0)
      Call Conds.Add(wc, wc.Name)
      p0 = p1 + 1 + Len(VALUE_END)
    End If
  Loop Until p1 = 0
  p0 = p0 + Len(CONDITION_SEP)
  Set RootClause = New whereClause
  Call CreateConditionTreeInternal(RootClause, Conds, InternalFormat, p0, "")
  Set CreateClauseFromInternal = RootClause
  
CreateClauseFromInternal_end:
  Exit Function
  
CreateClauseFromInternal_err:
  Call ErrorMessage(ERR_ERROR, Err, "CreateClauseFromInternal", "Unable to create where clause", "Error in clause creation")
  Resume CreateClauseFromInternal_end
  Resume
End Function
  
