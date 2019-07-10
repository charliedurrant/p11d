Attribute VB_Name = "DateHelper"
Option Explicit

Public Function DateSerialEx(ByVal Year As Long, ByVal Month As Long, ByVal Day As Long) As Date
  Dim DayMax As Long

  On Error GoTo DateSerialEx_err
  If (Month < 1) Or (Month > 12) Then Err.Raise ERR_DATESERIAL, "DateSerialEx", "Invalid Month Given In Date"
  If (Year < 100) Or (Year > 9999) Then Err.Raise ERR_DATESERIAL, "DateSerialEx", "Invalid Year Given In Date"
  
  Select Case Month
    Case 4, 6, 9, 11
      DayMax = 30
    Case 2
      If IsLeapYear(Year) Then
        DayMax = 29
      Else
        DayMax = 28
      End If
    Case Else
      DayMax = 31
  End Select
  If (Day < 1) Or (Day > DayMax) Then Err.Raise ERR_DATESERIAL, "DateSerialEx", "Invalid Day Given In Date"
  DateSerialEx = DateSerial(Year, Month, Day)
  Exit Function
  
DateSerialEx_err:
  Err.Raise Err.Number, ErrorSourceEx(Err, "DateSerialEx"), "Invalid Date '" & Day & "/" & Month & "/" & Year & "' (DMY) " & vbCrLf & Err.Description
End Function

Private Function IsLeapYear(ByVal Year As Long) As Boolean
  If (Year Mod 100) = 0 Then Year = Year / 100
  IsLeapYear = (Year Mod 4 = 0)
End Function

Public Function GetFullYear_CD(ByVal nyear As Integer, ConvStr As String) As Integer
  If nyear < 100 Then
    If nyear > YEAR1900CONV Then
      nyear = nyear + 1900
    ElseIf nyear < YEAR2000CONV Then
      nyear = nyear + 2000
    Else
      Err.Raise ERR_CONVERTDATE, "ConvertDate", "Format not complete: " & ConvStr & vbCrLf & "Two digit years greater than " & CStr(YEAR2000CONV) & " and less than " & CStr(YEAR1900CONV) & " must be in the format YYYY"
    End If
  End If
  GetFullYear_CD = nyear
End Function


