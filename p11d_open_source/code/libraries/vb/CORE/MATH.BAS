Attribute VB_Name = "Math"
Option Explicit
Public Const DBL_EPSILON As Double = 0.000000005 ' cf Abacus,  not too small as VB dbl EPSILON in order 10^-13
Private Const LOW_POW As Long = -10
Private Const HIGH_POW As Long = 10
Private Powers(LOW_POW To HIGH_POW) As Double

Public Enum ROUND_TYPE
  R_NORMAL
  R_UP
  R_DOWN
  R_BANKERS
End Enum

Public Function RoundDouble(ByVal Number As Double, ByVal DecimalPlaces As Long, ByVal rType As ROUND_TYPE) As Double
  Dim d As Double
  Dim TenPow As Double
  
  If (DecimalPlaces >= LOW_POW) And (DecimalPlaces <= HIGH_POW) Then
    If Powers(1) = 0 Then Call MathInit
    TenPow = Powers(DecimalPlaces)
  Else
    TenPow = 10 ^ DecimalPlaces
  End If
  
  Select Case rType
    Case R_NORMAL
      RoundDouble = Int((Number * TenPow) + 0.5 + DBL_EPSILON) / TenPow
    Case R_UP, R_DOWN
      d = Number * TenPow
      If Int(d) <> d Then
        If rType = R_UP Then d = d + 1
        RoundDouble = Int(d + DBL_EPSILON) / TenPow
      Else
        RoundDouble = d
      End If
    Case R_BANKERS
      RoundDouble = CLng(Number * TenPow) / TenPow
  End Select
End Function

Public Sub MathInit()
  Dim i As Long
  For i = LOW_POW To HIGH_POW
    Powers(i) = 10 ^ i
  Next i
End Sub


