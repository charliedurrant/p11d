Attribute VB_Name = "Functions"
Option Explicit

Public Function LowWord(ByVal l As Long) As Long
  LowWord = l And &HFFFF&
End Function

Public Function FontCopy(ByVal F As StdFont) As StdFont
  Dim fNew As StdFont
  
  Set fNew = New StdFont
  With fNew
    .Name = F.Name
    .Bold = F.Bold
    .Charset = F.Charset
    .Italic = F.Italic
    .SIZE = F.SIZE
    .Strikethrough = F.Strikethrough
    .Underline = F.Underline
    .Weight = F.Weight
  End With
  
  Set FontCopy = fNew
End Function

Public Function LowWordToHiWord(ByVal l As Long) As Long
  'first take off the hiword portion
  l = LowWord(l)
  'now multiply up
  LowWordToHiWord = l * (2 ^ 16)
End Function

Public Function TwoLongsToHiAndLow(ByVal HiWordLong As Long, ByVal LowWordLong As Long) As Long
  TwoLongsToHiAndLow = LowWordToHiWord(HiWordLong) + LowWord(LowWordLong)
End Function

Public Function GetCenterTextPosition(ByVal hdc As Long, ByVal sText As String, BoundingRect As RECT) As POINTAPI
  Dim sZRect As SIZE
  Dim sZText As SIZE
  
  Call GetTextExtentPoint32(hdc, sText, Len(sText), sZText)
  sZRect = GetRectDimensions(BoundingRect)
  GetCenterTextPosition.x = BoundingRect.Left + ((sZRect.cx - sZText.cx) / 2)
  GetCenterTextPosition.y = BoundingRect.Top + ((sZRect.cy - sZText.cy) / 2)
End Function

Public Function GetRectDimensions(r As RECT) As SIZE
  GetRectDimensions.cx = r.Right - r.Left
  GetRectDimensions.cy = r.Bottom - r.Top
End Function

Public Function Draw3DRect(ByVal hdc As Long, r As RECT, ByVal RT3D As Appearance)
  Dim lPenDark As Long, lPenLight As Long, lPenOriginal As Long
  Dim pT As POINTAPI
  
  lPenDark = CreatePen(PS_SOLID, 0, BOX_3D_DARK)
  lPenLight = CreatePen(PS_SOLID, 0, BOX_3D_LIGHT)
  lPenOriginal = SelectObject(hdc, lPenDark)
  
  Select Case RT3D
    Case Appearance.Up3D
      Call MoveToEx(hdc, r.Right - 1, r.Top - 1, pT)
      Call LineTo(hdc, r.Right - 1, r.Bottom - 1)
      Call LineTo(hdc, r.Left, r.Bottom - 1)
      Call DeleteObject(SelectObject(hdc, lPenLight))
      Call LineTo(hdc, r.Left, r.Top)
      Call LineTo(hdc, r.Right - 1, r.Top)
    Case Appearance.Down3D
      Call MoveToEx(hdc, r.Right - 1, r.Top, pT)
      Call LineTo(hdc, r.Left, r.Top)
      Call LineTo(hdc, r.Left, r.Bottom - 1)
      Call DeleteObject(SelectObject(hdc, lPenLight))
      Call LineTo(hdc, r.Right - 1, r.Bottom - 1)
      Call LineTo(hdc, r.Right - 1, r.Top)
    Case Else
  End Select

  Call DeleteObject(SelectObject(hdc, lPenOriginal))
End Function

