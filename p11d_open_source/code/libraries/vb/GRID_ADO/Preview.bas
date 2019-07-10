Attribute VB_Name = "Preview"
Option Explicit
Private mLines() As Variant
Private mCurLine As Long
Private mMaxLine As Long
Private LINE_INCREMENT As Long
Private Const MAX_LINE_INCREMENT As Long = (16384& * 8)

Public Sub SortPreviewLines(aCols As Variant, ByVal lmin As Long, ByVal lmax As Long)
  Dim aCol As AutoCol, aSort As AutoSort
  Dim DoSort As Boolean, i As Long
  
  On Error GoTo SortPreviewLines_err
  DoSort = False
  For i = lmin To lmax
    Set aCol = aCols(i)
    If aCol.SortType <> SORT_NONE Then
      DoSort = True
      Exit For
    End If
  Next i
  If DoSort And (mCurLine > 0) Then
    Set aSort = New AutoSort
    aSort.Columns = aCols
    aSort.MinCol = lmin
    aSort.MaxCol = lmax
    Call QSortEx(mLines, 1, mCurLine, aSort)
  End If
  
SortPreviewLines_end:
  Set aCol = Nothing
  Set aSort = Nothing
  Exit Sub
  
SortPreviewLines_err:
  Call ErrorMessage(ERR_ERROR, Err, "SortPreviewLines", "Sort report lines", "Error sorting report lines")
  Resume SortPreviewLines_end
End Sub

Public Sub ClearPreviewLines(aCols As Collection)
  Dim ac As AutoCol, i As Long
#If Not DEBUGVER Then
  On Error Resume Next
#End If
  Call xSet("ClearPreviewLines")
  For Each ac In aCols
    ac.MaxWidth = 0
    ac.pOffset = MINLEFTMARGIN
    ac.SumTotal = 0
    ac.SumLast = 0
    Call ac.RedimSumLevels(0)
    ac.SumGroupCount(0) = 0
    ac.SumGroup(0) = 0
    ac.DoGroup = False
    ac.FirstHeader = False
  Next ac
  For i = mMaxLine To 1 Step -1
    mLines(i) = Empty
  Next i
  mCurLine = 0
  Call xReturn("ClearPreviewLines")
End Sub

Public Sub AddPreviewLine(rLine As Variant)
  mCurLine = mCurLine + 1
  If mCurLine > mMaxLine Then
    If mCurLine = 1 Then LINE_INCREMENT = 4096
    If LINE_INCREMENT < MAX_LINE_INCREMENT Then LINE_INCREMENT = LINE_INCREMENT * 2
    mMaxLine = mMaxLine + LINE_INCREMENT
    ReDim Preserve mLines(1 To mMaxLine) As Variant
  End If
  mLines(mCurLine) = rLine
End Sub

Public Property Get MaxPreviewLine() As Long
  MaxPreviewLine = mCurLine
End Property

Public Sub GetPreviewLine(rLine As Variant, ByVal Index As Long)
  rLine = mLines(Index)
End Sub

Public Function GetForwardOnlyPreviewLine(ByVal Index As Long) As Variant
  If Index > 2 Then mLines(Index - 1) = Empty
  GetForwardOnlyPreviewLine = mLines(Index)
End Function


