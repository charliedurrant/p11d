VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FunctionItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Name As String
Public CallCount As Long
Public TotalTime As Long  ' total time spent in this function
Public DeltaTime As Long  ' total time spent in this function excluding
                          ' time spent in sub routines

