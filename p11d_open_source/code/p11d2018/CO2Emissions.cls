VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CO2Emissions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public CO2Emissions As ObjectList
Public Name As String
Public Dirty As Boolean
Public Approved As Boolean

Private Sub Class_Initialize()
  Set CO2Emissions = New ObjectList
End Sub

Public Sub Kill()
  Dim i As Long
  Dim Emission As CO2Emission
  
  On Error GoTo Kill_Err
  Call xSet("Kill")

  For i = 1 To CO2Emissions.Count
    Set Emission = CO2Emissions(i)
    Set Emission = Nothing
  Next

Kill_End:
  Call xReturn("Kill")
  Exit Sub

Kill_Err:
  Call ErrorMessage(ERR_ERROR, Err, "Kill", "Kill", "Error killing a CO2Emission.")
  Resume Kill_End
End Sub

Private Sub Class_Terminate()
 'Debug.Print"CO2Emission - terminate"
End Sub
