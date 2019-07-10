VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "VanData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Public FuelValue As Variant
Public VanValue As Variant
Public MadeGood As Variant

Public Function VanBenefit() As Variant
 'made good is a total not split
 Dim v As Variant
 
 If (MadeGood > VanValue) Then
  v = 0
 Else
  v = VanValue - MadeGood
 End If
 VanBenefit = v
 
End Function
Public Function FuelBenefit() As Variant
 'made good is a total not split
 Dim v As Variant
 Dim vFuelMadeGood As Variant
 
 If (MadeGood > VanValue) Then
  vFuelMadeGood = MadeGood - VanValue
 
 Else
  vFuelMadeGood = 0
 End If
  
 If (vFuelMadeGood > FuelValue) Then
  v = 0
 Else
  v = FuelValue - vFuelMadeGood
 End If
 
 
 FuelBenefit = v
End Function


Public Function TotalBenefit() As Variant
  Dim v As Variant
  v = Me.VanBenefit() + Me.FuelBenefit()
  TotalBenefit = v
End Function

Public Function TotalValue() As Variant
  Dim v As Variant
  v = Me.VanValue + Me.FuelValue
  TotalValue = v
End Function


Public Function TotalFuelValue() As Variant
  Dim v As Variant
  v = Me.FuelValue
  TotalFuelValue = v
End Function

Public Function TotalVanValue() As Variant
  Dim v As Variant
  v = Me.VanValue
  TotalVanValue = v
End Function




Private Sub Class_Initialize()
  Me.FuelValue = 0
  Me.VanValue = 0
  Me.MadeGood = 0

End Sub
