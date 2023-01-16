VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExternalOutput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Public Printing As Printing
'Public EDI As EDI
'
'Private Sub Class_Initialize()
'  Set Printing = New Printing
'  Set EDI = New EDI
'End Sub
'Private Function DateInRange(dDate As Date, dFrom As Date, dTo As Date) As Boolean
'  If dDate >= dFrom And dDate <= dTo Then DateInRange = True
'End Function
'
'
'Public Function GetP46Cars(p46Cars As ObjectList, Employee As Employee, ByVal dDateFrom As Date, ByVal dDateTo As Date) As Boolean
'  Dim i As Long
'  Dim CompanyCar As IBenefitClass
'  Dim bForcePrintingP46 As Boolean, bDateFromPass, bDateToPass As Boolean
'  Dim dTaxYearStart As Date, dTaxYearEnd As Date
'
'  On Error GoTo GetP46Cars_Err
'
'  Call xSet("GetP46Cars")
'
'  dTaxYearStart = P11d32.Rates.GetItem(taxyearstart)
'  dTaxYearEnd = P11d32.Rates.GetItem(taxyearend)
'
'  If p46Cars Is Nothing Then
'    Set p46Cars = New ObjectList
'  Else
'    p46Cars.RemoveAll
'  End If
'
'  For i = 1 To Employee.benefits.count
'    Set CompanyCar = Employee.benefits(i)
'    If Not CompanyCar Is Nothing Then
'      With CompanyCar
'        If .BenefitClass = BC_COMPANYCAR Then
'          bForcePrintingP46 = .GetItem(car_ForceP46)
'          bDateFromPass = P46DateInRange(.GetItem(car_Availablefrom), dDateFrom, dDateTo, bForcePrintingP46, dTaxYearStart, dTaxYearEnd)
'          bDateToPass = P46DateInRange(.GetItem(car_AvailableTo), dDateFrom, dDateTo, bForcePrintingP46, dTaxYearStart, dTaxYearEnd)
'
'          If bDateFromPass Or bDateToPass Then
'            Call SetP46CarDefaults(CompanyCar)
'            If bDateFromPass Then
'              If ben.GetItem(car_Replacement) Then
'                Call .SetItem(car_P46CarProvidedReplaced, True)
'              ElseIf ben.GetItem(car_Second) Then
'                Call .SetItem(car_P46SecondCar, True)
'              Else
'                Call .SetItem(car_P46FirstProvidedWithCar, True)
'              End If
'            End If
'
'            If bDateToPass Then
'              If ben.GetItem(car_Replaced) Then
'                Set P46Car = Nothing
'              Else
'                Call .SetItem(car_P46WithdrawnWithoutReplacement, True)
'              End If
'            End If
'
'            If Not P46Car Is Nothing Then
'              Call .SetItem(car_P46PrintCarDetails, (.GetItem(car_P46WithdrawnWithoutReplacement)) And (Not .GetItem(car_P46CarProvidedReplaced)) And (Not .GetItem(car_P46FirstProvidedWithCar)) And (.GetItem(car_P46SecondCar)))
'              GetP46Cars = True
'              Call p46Cars.Add(CompanyCar)
'            End If
'          End If
'
'        End If
'      End With
'    End If
'  Next
'
'GetP46Cars_End:
'  Call xReturn("GetP46Cars")
'  Exit Function
'
'GetP46Cars_Err:
'  Call ErrorMessage(ERR_ERROR, Err, "GetP46Cars", "ERR_UNDEFINED", "Undefined error.")
'  Resume GetP46Cars_End
'End Function
'
'
'
'
'
'Public Function SetP46CarPrintStrings(CompanyCar As IBenefitClass) As Boolean
'
'  On Error GoTo SetP46CarPrintStrings_Err
'  Call xSet("SetP46CarPrintStrings")
'
'  'add item here then add in SetP46CarDefaults
'  With CompanyCar
'    Call .SetItem(car_P46CarProvidedReplaced, IIf(.GetItem(car_P46CarProvidedReplaced), "ü", ""))
'    Call .SetItem(car_P46FirstProvidedWithCar, IIf(.GetItem(car_P46FirstProvidedWithCar), "ü", ""))
'    Call .SetItem(car_P46SecondCar, IIf(.GetItem(car_P46SecondCar), "ü", ""))
'    Call .SetItem(car_P46WithdrawnWithoutReplacement, IIf(.GetItem(car_P46WithdrawnWithoutReplacement), "ü", ""))
'  End With
'
'SetP46CarPrintStrings_End:
'  Call xReturn("SetP46CarPrintStrings")
'  Exit Function
'
'SetP46CarPrintStrings_Err:
'  Call ErrorMessage(ERR_ERROR, Err, "SetP46CarPrintStrings", "ERR_UNDEFINED", "Undefined error.")
'  Resume SetP46CarPrintStrings_End
'End Function
'
