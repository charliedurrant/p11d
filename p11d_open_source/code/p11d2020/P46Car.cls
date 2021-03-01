VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "P46Car"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public CompanyCar As IBenefitClass
Public FirstProvidedWithCar As Variant
Public CarProvidedReplaced As Variant
Public SecondCar As Variant
Public WithdrawnWithoutReplacement As Variant
Public Property Get PrintMake() As Boolean
  PrintMake = WithdrawnWithoutReplacement And (Not CarProvidedReplaced) And (Not FirstProvidedWithCar) And SecondCar
End Property

Private Sub Class_Initialize()
  
  FirstProvidedWithCar As Variant
  Public CarProvidedReplaced As Variant
  Public SecondCar As Variant
  Public WithdrawnWithoutReplacement As Variant
End Sub
