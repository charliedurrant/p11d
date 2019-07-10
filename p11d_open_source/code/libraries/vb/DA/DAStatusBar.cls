VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DAStatusBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

Private mStat As TCSStatus
Private min As Long
Private max As Long

Public Property Set StatusBar(newval As TCSStatus)
  Set mStat = newval
End Property
Public Property Let StatMin(newval As Long)
  min = newval
End Property
Public Property Let StatMax(newval As Long)
  max = newval
End Property

Public Property Get StatusBar() As TCSStatus
  Set StatusBar = mStat
End Property
Public Property Get StatMin() As Long
  StatMin = min
End Property
Public Property Get StatMax() As Long
  StatMax = max
End Property

Private Sub Class_Terminate()
  Set mStat = Nothing
End Sub
