Attribute VB_Name = "WrapFunctions"
Option Explicit

Private Function GetWrapSubString(WrapString As String, ByVal tw As Single, ByVal WrapX As Single, ByVal Start As Long, ByVal Length As Long) As Long
  Dim SelLen As Long, Direction As Long
  
  If Length > 0 Then
    If tw > WrapX Then
      Direction = 0
      SelLen = (Length * WrapX) / tw
      If SelLen = 0 Then SelLen = 1
      tw = GetTextWidth(Mid$(WrapString, Start, SelLen))
      If tw > WrapX Then
        If SelLen > 1 Then Direction = -1
      ElseIf tw < WrapX Then
        Direction = 1
      End If
     
      ' this could be improved by making bigger jumps in Direction
      Do While Direction <> 0
        SelLen = SelLen + Direction
        tw = GetTextWidth(Mid$(WrapString, Start, SelLen))
        If (tw > WrapX) And (Direction = 1) Then Exit Do
        If (tw < WrapX) And (Direction = -1) Then Exit Do
        If tw = WrapX Then Direction = 0
      Loop
      If Direction = 1 Then SelLen = SelLen - 1  ' break char is last char
      GetWrapSubString = Start + SelLen - 1      ' return offset of break char in string (1 based)
      Exit Function
    End If
  End If
  GetWrapSubString = Start + Length
End Function

Public Function WrapTextToWidthEx(WrapString As String, ByVal WrapX As Single, Optional ByVal BreakChars As String = " ,;:-=") As Long
  Dim tw As Single
  Dim p0 As Long, p1 As Long, q0 As Long, Length As Long
  
  tw = GetTextWidth(WrapString)
  If tw > WrapX Then
    p0 = 1
    Do
      Length = Len(WrapString) - p0 + 1
      If Length <> Len(WrapString) Then tw = GetTextWidth(Mid$(WrapString, p0, Length))
      p1 = GetWrapSubString(WrapString, tw, WrapX, p0, Length)
      If p1 >= Len(WrapString) Then Exit Do
      q0 = InStrAnyRev(WrapString, BreakChars, p1)
      If q0 >= p0 Then p1 = q0
      WrapString = Left$(WrapString, p1) & vbCrLf & LTrim$(Mid$(WrapString, p1 + 1))
      WrapTextToWidthEx = WrapTextToWidthEx + 1
      p0 = p1 + 3
    Loop Until False
  Else
    WrapTextToWidthEx = 0
  End If
End Function

