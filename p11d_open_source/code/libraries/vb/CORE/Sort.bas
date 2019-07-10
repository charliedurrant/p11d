Attribute VB_Name = "Sorting"
Option Explicit
Private Declare Function SafeArrayGetElemsize Lib "oleaut32.dll" (ByVal pTRSafeArray As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, Source As Any, ByVal numBytes As Long)
Private Const vtByRef As Long = &H4000
Private Const VAR_OFFSET As Long = 8
Private Const SA_DATA_OFFSET As Long = 12
Private Const MAX_ELEMENT_SIZE As Long = 128
Public SortComparisons As Long
Public Enum TCS_SORTTYPE
  QUICK_SORT
  COMB_SORT
End Enum

Public Sub SortAny(v As Variant, ByVal lmin As Long, ByVal lmax As Long, SortFn As ISortFunction, ByVal st As TCS_SORTTYPE)
  Dim SizeOfElement As Long
  Dim pSafeArrayData As Long
  Dim pSA As Long
  Dim LBoundOfSA As Long

  SortComparisons = 0
  If lmax > lmin Then
    pSA = GetPtrToSA(v)
    If pSA = 0 Then Call Err.Raise(ERR_SORT, "SortAny", "Cannot sort a variant that is not an array")
    SizeOfElement = SafeArrayGetElemsize(pSA)
    If SizeOfElement > MAX_ELEMENT_SIZE Then Call Err.Raise(ERR_SORT, "SortAny", "Array element size must be less than " & CStr(MAX_ELEMENT_SIZE))
    pSafeArrayData = GetPtrToSAData(pSA)
    If pSafeArrayData <> 0 Then
      LBoundOfSA = LBound(v)
      If st = QUICK_SORT Then
        Call QSortFn(v, lmin, lmax, IsObject(v(lmax)), SortFn, pSafeArrayData, SizeOfElement, LBoundOfSA)
      ElseIf st = COMB_SORT Then
        Call CombSort(v, lmin, lmax, IsObject(v(lmax)), SortFn, pSafeArrayData, SizeOfElement, LBoundOfSA)
      Else
        Call ECASE_SYS("Unknown Sort type: " & CStr(st))
      End If
    End If
  End If
End Sub

Public Sub CombSort(v As Variant, ByVal lmin As Long, ByVal lmax As Long, ByVal UseObjects As Boolean, SortFn As ISortFunction, ByVal pSafeArrayData As Long, ByVal SizeOfElement As Long, ByVal LBoundOfSA As Long)
  Dim i As Long, j As Long
  Dim gap As Long, switches As Boolean
   
  gap = lmax - lmin + 1
  Do
    gap = (gap * 10) \ 13
    If gap < 1 Then
      gap = 1
    ElseIf (gap = 9) Or (gap = 10) Then
      gap = 11
    End If
    switches = False
    For i = lmin To (lmax - gap)
      j = i + gap
      #If DEBUGVER Then
      SortComparisons = SortComparisons + 1
      #End If
      If SortFn.CompareItems(v(i), v(j)) > 0 Then
        Call SwapValues(i, j, pSafeArrayData, SizeOfElement, LBoundOfSA)
        switches = True
      End If
    Next i
  Loop While switches Or (gap > 1)
End Sub

Public Sub QSortFn(v As Variant, ByVal lmin As Long, ByVal lmax As Long, ByVal UseObjects As Boolean, SortFn As ISortFunction, ByVal pSafeArrayData As Long, ByVal SizeOfElement As Long, ByVal LBoundOfSA As Long)
  Dim vmax As Variant, i As Long, j As Long, num As Long
  Dim vtmp As Variant
  
  If lmax > lmin Then
    num = lmax - lmin + 1
    If num = 2 Then
      #If DEBUGVER Then
      SortComparisons = SortComparisons + 1
      #End If
      If SortFn.CompareItems(v(lmin), v(lmax)) > 0 Then
        Call SwapValues(lmin, lmax, pSafeArrayData, SizeOfElement, LBoundOfSA)
      End If
    ElseIf num = 3 Then
      #If DEBUGVER Then
      SortComparisons = SortComparisons + 3
      #End If
      If SortFn.CompareItems(v(lmin), v(lmin + 1)) > 0 Then
        Call SwapValues(lmin, lmin + 1, pSafeArrayData, SizeOfElement, LBoundOfSA)
      End If
      If SortFn.CompareItems(v(lmin + 1), v(lmin + 2)) > 0 Then
        Call SwapValues(lmin + 1, lmin + 2, pSafeArrayData, SizeOfElement, LBoundOfSA)
      End If
      If SortFn.CompareItems(v(lmin), v(lmin + 1)) > 0 Then
        Call SwapValues(lmin, lmin + 1, pSafeArrayData, SizeOfElement, LBoundOfSA)
      End If
    Else
      num = lmin + (num \ 2)
      Call SwapValues(num, lmax, pSafeArrayData, SizeOfElement, LBoundOfSA)
      If UseObjects Then
        Set vmax = v(lmax)
      Else
        vmax = v(lmax)
      End If
      i = lmin
      j = lmax - 1
      
      Do While True
        #If DEBUGVER Then
        SortComparisons = SortComparisons + 1
        #End If
        Do While SortFn.CompareItems(v(i), vmax) < 0
          #If DEBUGVER Then
          SortComparisons = SortComparisons + 1
          #End If
          i = i + 1
        Loop
        If i = lmax Then GoTo do_sort
        #If DEBUGVER Then
        SortComparisons = SortComparisons + 1
        #End If
        Do While SortFn.CompareItems(v(j), vmax) >= 0
          #If DEBUGVER Then
          SortComparisons = SortComparisons + 1
          #End If
          j = j - 1
          If j <= i Then GoTo swap_imax
        Loop
        If j <= i Then GoTo swap_imax
        Call SwapValues(i, j, pSafeArrayData, SizeOfElement, LBoundOfSA)
      Loop
swap_imax:
      Call SwapValues(i, lmax, pSafeArrayData, SizeOfElement, LBoundOfSA)
do_sort:
      Call QSortFn(v, lmin, i - 1, UseObjects, SortFn, pSafeArrayData, SizeOfElement, LBoundOfSA)
      Call QSortFn(v, i + 1, lmax, UseObjects, SortFn, pSafeArrayData, SizeOfElement, LBoundOfSA)
    End If
  End If
End Sub

Private Function GetPtrToSA(v As Variant) As Long
  Dim pSA As Long, pSARef As Long, pSAAddr As Long
  Dim vType As Long
  
  CopyMemory ByVal VarPtr(vType), ByVal VarPtr(v), 2
  vType = vType And &HFFFF&
  If (vType And vbArray) = vbArray Then
    pSAAddr = VarPtr(v) + VAR_OFFSET
    If (vType And vtByRef) = vtByRef Then
      CopyMemory ByVal VarPtr(pSARef), ByVal pSAAddr, 4
      pSAAddr = pSARef
    End If
    CopyMemory ByVal VarPtr(pSA), ByVal pSAAddr, 4
    GetPtrToSA = pSA
  End If
End Function

Private Function GetPtrToSAData(ByVal ptrSA As Long) As Long
  Dim l As Long
  Dim ptrSAData As Long
  
  ptrSAData = ptrSA + SA_DATA_OFFSET
  CopyMemory ByVal VarPtr(l), ByVal ptrSAData, 4
  GetPtrToSAData = l
End Function

Private Function SwapValues(ByVal ArrayIndexDest As Long, ByVal ArrayIndexSrc As Long, ByVal pSafeArrayData As Long, ByVal SizeOfElement As Long, ByVal LBoundOfSA As Long)
  Dim tmp(0 To MAX_ELEMENT_SIZE) As Byte
  Dim pSADest As Long, pSASrc As Long
  
  pSADest = pSafeArrayData + ((ArrayIndexDest - LBoundOfSA) * SizeOfElement)
  pSASrc = pSafeArrayData + ((ArrayIndexSrc - LBoundOfSA) * SizeOfElement)
  
  CopyMemory ByVal VarPtr(tmp(0)), ByVal pSADest, ByVal SizeOfElement
  CopyMemory ByVal pSADest, ByVal pSASrc, ByVal SizeOfElement
  CopyMemory ByVal pSASrc, ByVal VarPtr(tmp(0)), ByVal SizeOfElement
End Function



