Attribute VB_Name = "MergeCode"
'  This code may be of use if field merging is to be implemented

''''Option Explicit
''''
''''Public Sub SplitOK()
''''
''''  On Error GoTo SplitOK_Err
''''  Call xSet("SplitOK")
''''
''''
''''Dim LeftCol As Long
''''Dim RightCol As Long
''''Dim Concat As Long
''''Dim FieldType As Long
''''Dim i As Long
''''Dim LColField As Long
''''Dim LColFieldDisp As Long
''''Dim PrevField As Long
''''Dim MergedName As String
''''
''''With Frm_Misc.FlG_Source
''''
''''LColFieldDisp = .Col
''''
''''If LColFieldDisp <> .ColSel Then
''''  MsgBox "More than one field is selected !"
''''  Exit Sub
''''End If
''''
''''''LColField = mvarDisplayCols(LColFieldDisp)
''''
''''If m_FieldSpecs(LColField).RealName = "" Then
''''  MsgBox "Selected field is not merged !"
''''  Exit Sub
''''End If
''''
''''For i = 1 To m_FieldSpecs.Count
''''  ''If Abs(m_FieldSpecs(i).Merge) = LColField Then
''''  ''  m_FieldSpecs(i).Merge = 0
''''  ''End If
''''Next i
''''
''''m_FieldSpecs(LColField).FldName = m_FieldSpecs(LColField).RealName
''''m_FieldSpecs(LColField).RealName = ""
''''
''''Call ReCalc_Misc
''''
''''.LeftCol = LColFieldDisp
''''.Col = LColFieldDisp
''''.Row = 2
''''.ColSel = LColFieldDisp
''''.RowSel = .Rows - 1
''''
''''End With
''''
''''
''''SplitOK_End:
''''  Call xReturn("SplitOK")
''''  Exit Sub
''''
''''SplitOK_Err:
''''  Call ErrorMessage(ERR_ERROR, Err, "SplitOK", "ERR_UNDEFINED", "Undefined error.")
''''  Resume SplitOK_End
''''End Sub
''''
''''Public Sub MergeColumns(Separator As String)
''''
''''    On Error GoTo MergeColumns_Err
''''    Call xSet("MergeColumns")
''''
''''Dim LeftCol As Long
''''Dim RightCol As Long
''''Dim Concat As Long
''''Dim FieldType As Long
''''Dim i As Long
''''Dim LColField As Long
''''Dim LColFieldDisp As Long
''''Dim PrevField As Long
''''Dim MergedName As String
''''
''''With Frm_Misc.FlG_Source
''''
''''LeftCol = .Col
''''RightCol = .ColSel
''''
''''If LeftCol = RightCol Then
''''  MsgBox "Only one field is selected !"
''''  GoTo MergeColumns_End
''''End If
''''
''''LColFieldDisp = 0
''''LColField = m_FieldSpecs.Count
''''For i = LeftCol To RightCol
''''  ''If m_FieldSpecs(mvarDisplayCols(i)).vartype <> TYPE_STR Then
''''  ''  MsgBox "Selected fields are not all of type STRING !"
''''  ''  GoTo MergeColumns_End
''''  ''End If
''''  ''If m_FieldSpecs(mvarDisplayCols(i)).RealName <> "" Then
''''  ''  MsgBox "Some of the fields selected are already merged !"
''''  ''  GoTo MergeColumns_End
''''  ''End If
''''  ''If mvarDisplayCols(i) < LColField Then
''''  ''  LColFieldDisp = i
''''  ''  LColField = mvarDisplayCols(LColFieldDisp)
''''  ''End If
''''
''''Next i
''''
''''MergedName = m_FieldSpecs(LColField).FldName
''''m_FieldSpecs(LColField).RealName = MergedName
''''
''''PrevField = LColField
''''''For i = LColFieldDisp - 1 To LeftCol Step -1
''''''    If mvarDisplayCols(i) <= PrevField Then
''''''      MsgBox "Merging with Field " & m_FieldSpecs(LColField).FldName & vbCrLf _
''''''       & "Fields cannot be merged in the selected order !" & vbCrLf _
''''''       & "Field " & m_FieldSpecs(mvarDisplayCols(i)).FldName & " cannot be inserted before Field " & m_FieldSpecs(PrevField).FldName
''''''      GoTo MergeColumns_End
''''''    End If
''''''    PrevField = mvarDisplayCols(i)
''''''  MergedName = m_FieldSpecs(mvarDisplayCols(i)).FldName & "+" & MergedName
''''''Next i
''''
''''PrevField = LColField
''''For i = LColFieldDisp + 1 To RightCol Step 1
''''''    If mvarDisplayCols(i) <= PrevField Then
''''''      MsgBox "Merging with Field " & m_FieldSpecs(LColField).FldName & vbCrLf _
''''''       & "Fields cannot be merged in the selected order !" & vbCrLf _
''''''       & "Field " & m_FieldSpecs(mvarDisplayCols(i)).FldName & " cannot be appended to Field " & m_FieldSpecs(PrevField).FldName
''''''      GoTo MergeColumns_End
''''''    End If
''''''    PrevField = mvarDisplayCols(i)
''''''  MergedName = MergedName & "+" & m_FieldSpecs(mvarDisplayCols(i)).FldName
''''Next i
''''
''''For i = LeftCol To RightCol
''''''  If i = LColFieldDisp Then
''''''    m_FieldSpecs(mvarDisplayCols(i)).FldName = MergedName
''''''  Else
''''''    'm_FieldSpecs(mvarDisplayCols(i)).Merge = LColField * Sgn(i - LColFieldDisp)
''''''    m_FieldSpecs(mvarDisplayCols(i)).MergeSeparator = Separator
''''''  End If
''''Next i
''''
''''Call ReCalc_Misc
''''
''''If .ColIsVisible(LColFieldDisp) = False Then
''''  .LeftCol = LColFieldDisp
''''End If
''''.Col = LColFieldDisp
''''.Row = 2
''''.ColSel = LColFieldDisp
''''.RowSel = .Rows - 1
''''
''''End With
''''
''''
''''MergeColumns_End:
''''    Call xReturn("MergeColumns")
''''    Exit Sub
''''
''''MergeColumns_Err:
''''    Call ErrorMessage(ERR_ERROR, Err, "MergeColumns", "ERR_UNDEFINED", "Undefined error.")
''''    Resume MergeColumns_End
''''End Sub
''''
