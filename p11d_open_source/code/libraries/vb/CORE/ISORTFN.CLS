VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ISortFunction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' -1 if v0 < v1, 1 if v0 > v1, 0 if v0 = v1
Public Function CompareItems(v0 As Variant, v1 As Variant) As Long
Attribute CompareItems.VB_Description = "Compare 2 sort items, return -1 if v0 < v1, 1 if v0 > v1, 0 if the same"
  
End Function
