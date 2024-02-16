VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsFPCSBand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Scheme As String
Public BandName As String
Public AboveMileage As Long
Public AboveEngineSize As Long
Public Rate As Double
Public Order As Long


Private Sub Class_Initialize()
  Order = -1
End Sub
