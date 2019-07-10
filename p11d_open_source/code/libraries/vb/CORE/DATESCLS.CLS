VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DatesClass"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'* modify date by changing any specific field
'* passing a -1 means leave field alone
'*
'* ChangeDate           Date to be modified
'* [Day]                Overide day (Default -1)
'* [Month]              Overide month (Default -1)
'* [Year]               Overide year (Default -1)
'* [Hour]               Overide hours (Default -1)
'* [Minute]             Overide minutes (Default -1)
'* [Seconds]            Overide seconds (Default -1)
'*
'* Return Value:
'* The date type comprising the date & hours, minutes and seconds
'*
Public Function SetDateTime(ByVal ChangeDate As Date, Optional Day As Integer = -1, Optional Month As Integer = -1, Optional Year As Integer = -1, Optional Hour As Integer = 0, Optional Minute As Integer = 0, Optional Seconds As Integer = 0) As Date
Dim nday0 As Integer, nmonth0 As Integer, nyear0 As Integer
  
  Call gTracer.xSet("SetDateTime")
  nday0 = DatePart("d", ChangeDate)
  If Day <> -1 Then nday0 = Day
  nmonth0 = DatePart("m", ChangeDate)
  If Month <> -1 Then nmonth0 = Month
  nyear0 = DatePart("yyyy", ChangeDate)
  If Year <> -1 Then nyear0 = Year
  
  SetDateTime = DateSerial(nyear0, nmonth0, nday0) + TimeSerial(Hour, Minute, Seconds)
  Call gTracer.xReturn("SetDateTime")
End Function

Public Sub DatesStub()
  Dim l As Long
  l = 5
End Sub

