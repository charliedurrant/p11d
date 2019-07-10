VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Printing_old"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public CurrentPrintEmployee As IBenefitClass
Private m_Reporter As Reporter

 Public Sub P46Car()
   
  Dim xsize As Integer
  Dim ysize As Integer
  Dim ipvtfuel As Integer 'private fuel supplied
  Dim irmakegood As Integer
  Dim ifuel As Integer 'fuel type
  Dim iengine As Integer 'engne size
  Dim s As String, sTmp As String
  
  Dim imileband As Integer
  
  Dim iFirst As Integer, iReplaced As Integer, iReplacement As Integer, isecond As Integer, imake As Integer
  Dim iwithdraw As Integer
  Dim p46Cars As ObjectList
  Dim P46Car As IBenefitClass
  Dim bCarDetails As Boolean
  Dim i As Long
  Dim ee As Employee
  
  On Error GoTo P46Car_err
  
  If Not P11d32.ExternalOutput.GetP46Cars(p46Cars, CurrentPrintEmployee, P11d32.Rates.GetItem(taxyearstart), P11d32.Rates.GetItem(taxyearend)) Then
   Exit Sub
  End If
  
  For i = 1 To p46Cars.count
    Set P46Car = p46Cars(i)
    With P46Car
      Call m_Reporter.InitReport("P46", PREPARE_REPORT)
      s = s & "{Arial=10,n}"
      s = s & vbCrLf & "{Arial=10,n}{x=0}Inland" & vbCrLf
      s = s & "{Arial=30,b}{x=0}Revenue{x=60}{Arial=11,b}Notification of a car provided for the" & vbCrLf
      s = s & "{x=60}private use of an employee or a director" & vbCrLf & vbCrLf
      s = s & "{x=0}{WBTEXTBOXL=99,7,}"
      s = s & "{Arial=7,i}" & vbCrLf & "{Arial=10,n}{X=2}Employer's name  ____________________________________________"
      s = s & "{x=17}{Times=10,BI}" & P11d32.CurrentEmployer.Name
      s = s & "{x=65}{Arial=10,n}PAYE reference  ___________________{x=97}{Times=10,RBI}" & P11d32.CurrentEmployer.Payeref & vbCrLf & vbCrLf & "{Arial=7,i}" & vbCrLf
      s = s & "{Arial=10,n}{x=2}Employee's/Director's name  ___________________________________"
      Set ee = CurrentPrintEmployee
      s = s & "{x=25}{Times=10,BI}" & ee.FullName & "{Arial=10,n}{x=65}NI number  _______________________{x=97}{Times=10,RBI}" & CurrentPrintEmployee.GetItem(ee_NINumber) & vbCrLf & vbCrLf & vbCrLf
      s = s & "{Arial=10,b}Part 1" & vbCrLf & vbCrLf
      s = s & "{Arial=10,n}You are required to make a return on this form for an employee earning at the rate of £8,500 "
      s = s & "a year or more or a " & vbCrLf & "director for whom a car is made available for private use. The completed form is required within 28 days of the end of" & vbCrLf
      s = s & "the quarter to 5 July, 5 October, 5 January or 5 April in which any of the following takes place." & vbCrLf
      'tick boxes
      s = s & "{Arial=9,I}{x=81}Tick whichever applies" & vbCrLf & vbCrLf
      s = s & "{Arial=10,n}1.{x=5}The employee/director is first provided with a car which is available for private use"
      s = s & "{x=94}{WBTEXTBOXL=" & xsize & "," & ysize & "}"
      s = s & "{x=94}{xrel=60}{yrel=60}{Wding=12,b}" & TickOut(P46Car.GetItem(car_P46FirstProvidedWithCar)) & "{xrel=-60}{yrel=-60}"
      s = s & "{Arial=10,n}" & vbCrLf & vbCrLf & vbCrLf
      s = s & "{Arial=10,n}2.{x=5}A car provided to the employee/director is replaced by another car which is available for"
      s = s & "{x=94}{WBTEXTBOXL=" & xsize & "," & ysize & "}"
      s = s & "{x=94}{xrel=60}{yrel=60}{Wding=12,b}" & TickOut(P46Car.GetItem(car_P46CarProvidedReplaced)) & "{xrel=-60}{yrel=-60}"
      s = s & "{Arial=10,n}" & vbCrLf & "{x=5}private use" & vbCrLf & vbCrLf
      s = s & "3.{x=5}The employee/director is provided with a second or further car, which is available for private use"
      s = s & "{x=94}{WBTEXTBOXL=" & xsize & "," & ysize & "}"
      s = s & "{x=94}{xrel=60}{yrel=60}{Wding=12,b}" & TickOut(P46Car.GetItem(car_P46SecondCar)) & "{xrel=-60}{yrel=-60}"
      s = s & "{Arial=10,n}" & vbCrLf & vbCrLf & vbCrLf
      s = s & "{Arial=10,n}4.{x=5}The employee starts to earn at a rate of £8,500 a year or more or becomes a director"
      s = s & "{x=94}{WBTEXTBOXL=" & xsize & "," & ysize & "}" & vbCrLf & vbCrLf & vbCrLf
      s = s & "5.{x=5}A car provided to the employee/director is withdrawn without replacement"
      s = s & "{x=94}{WBTEXTBOXL=" & xsize & "," & ysize & "}"
      s = s & "{x=94}{xrel=60}{yrel=60}{Wding=12,b}" & TickOut(P46Car.GetItem(car_P46WithdrawnWithoutReplacement)) & "{xrel=-60}{yrel=-60}"
      s = s & "{Arial=10,n}" & vbCrLf & vbCrLf & vbCrLf
      'end tick boxes
      s = s & "{Arial=10,b}Part 2 Details of car provided" & vbCrLf & vbCrLf
      
      'car details
      bCarDetails = .GetItem(car_P46PrintCarDetails)
      s = s & "{x=0}{Arial=10,n}Make and model   __________________________________________ " & IIf(bCarDetails, "{x=15}{Times=10,BI}" & .GetItem(car_make) & " " & .GetItem(car_model), "") & "{Arial=10,r}{x=97}Date first registered   _________________{x=97}{Times=10,RBI}" & IIf(bCarDetails, .GetItem(car_Registrationdate), "") & vbCrLf & vbCrLf & vbCrLf
      s = s & "{Arial=10,n}Price of car (normally the list price at date of registration){Arial=10,r}{x=97}£   _________________{x=97}{Times=10,RBI}" & IIf(bCarDetails, formatworkingnumber(.GetItem(car_Price)), "") & vbCrLf & vbCrLf & vbCrLf
      s = s & "{Arial=10,n}Price of accessories not included in price of car{Arial=10,r}{x=97}£   _________________{x=97}{Times=10,RBI}" & IIf(bCarDetails, formatworkingnumber(.GetItem(car_accessories)), "") & vbCrLf & vbCrLf & vbCrLf
      s = s & "{Arial=10,n}Date car first made available to employee{Arial=10,r}{x=97}_________________{x=97}{Times=10,RBI}" & IIf(bCarDetails, .GetItem(car_Availablefrom), "") & vbCrLf & vbCrLf & vbCrLf
      s = s & "{Arial=10,n}Capital contribution (if any) made by employee to cost of the car and for accessories{Arial=10,r}{x=97}£   _________________{x=97}{Times=10,RBI}" & IIf(bCarDetails, formatworkingnumber(.GetItem(car_capitalcontribution)), "") & vbCrLf & vbCrLf & vbCrLf
      s = s & "{Arial=10,n}Sum payable (if any) made by the employee for private use of the car{Arial=10,r}{x=97} £   _________________{Arial=10,r}{x=97}{Times=10,RBI}" & IIf(bCarDetails, formatworkingnumber(.GetItem(car_usagecontribution)), "") & vbCrLf & vbCrLf
      
      
      s = s & "{Arial=10,n}Is fuel for private use provided with this car?{x=38}{Arial=10,i}yes {WBTEXTBOXL=" & xsize & "," & ysize & "}{xrel=60}{yrel=60}{Wding=12,b}" & TickOut(.GetItem(car_privatefuel)) & "{xrel=-60}{yrel=-60}{Arial=10,i}" & "{x=47}  no {WBTEXTBOXL=" & xsize & "," & ysize & "}{xrel=60}{yrel=60}{Wding=12,b}" & TickOut(.GetItem(car_privatefuel)) & "{xrel=-60}{yrel=-60}{Arial=10,i}{Arial=10,n}{x=57}If so, is the employee required to make" & vbCrLf
      
      irmakegood = 0 'IIf(IsNull(("makegood")), False, ("makegood"))
      
      s = s & vbCrLf & "{Arial=10,n}good the cost of all fuel used for private motoring and do you expect him/her to continue to do so?{Arial=10,i}{x=82}yes{x=86}{WBTEXTBOXL=" & xsize & "," & ysize & "}" & IIf(ipvtfuel And irmakegood, "{xrel=60}{yrel=60}{Wding=12,b}ü{xrel=-60}{yrel=-60}{Arial=10,i}", "") & "{x=92}no {WBTEXTBOXL=" & xsize & "," & ysize & "}" & IIf(ipvtfuel And Not irmakegood, "{xrel=60}{yrel=60}{Wding=12,b}ü{xrel=-60}{yrel=-60}{Arial=10,i}", "") & vbCrLf
      ifuel = 0 'IIf(IsNull(("diesel")), False, ("diesel"))
      s = s & "{Arial=10,n}" & vbCrLf & "If the answer to either question is no please indicate the type of fuel{x=62}petrol{x=68}{WBTEXTBOXL=" & xsize & "," & ysize & "}" & IIf((ipvtfuel And Not irmakegood) And Not ifuel, "{xrel=60}{yrel=60}{Wding=12,b}ü{xrel=-60}{yrel=-60}{Arial=10,i}", "") & "{x=80}diesel{x=86}{WBTEXTBOXL=" & xsize & "," & ysize & "}" & IIf((ipvtfuel And Not irmakegood) And ifuel, "{xrel=60}{yrel=60}{Wding=12,b}ü{xrel=-60}{yrel=-60}{Arial=10,i}", "") & vbCrLf
    
      iengine = 0 'IIf(IsNull(("cc")), 0, ("cc"))
      s = s & vbCrLf & "{Arial=10,n}and the cylinder capacity{x=40}up to 1400cc {WBTEXTBOXL=" & xsize & "," & ysize & "}" & IIf(ipvtfuel And Not irmakegood And (iengine <= 1400), "{xrel=60}{yrel=60}{Wding=12,b}ü{xrel=-60}{yrel=-60}{Arial=10,i}", "") & "{x=56}1401 - 2000cc{x=68}{WBTEXTBOXL=" & xsize & "," & ysize & "}" & IIf(ipvtfuel And Not irmakegood And ((iengine > 1400) And (iengine <= 2000)), "{xrel=60}{yrel=60}{Wding=12,b}ü{xrel=-60}{yrel=-60}{Arial=10,i}", "") & "{x=75}2001 or more{x=86}{WBTEXTBOXL=" & xsize & "," & ysize & "}" & IIf(ipvtfuel And Not irmakegood And (iengine > 2000), "{xrel=60}{yrel=60}{Wding=12,b}ü{xrel=-60}{yrel=-60}{Arial=10,i}", "") & vbCrLf & vbCrLf & vbCrLf
      s = IIf(IsNull(("busmiles")), "18000 plus", ("busmiles"))
      imileband = IIf(s = "18000 plus", 3, IIf(s = "2500 to 17999", 2, 1))
      If Not (iFirst Or iReplaced Or isecond) Then
        imileband = 4 ' don't report milage
      End If
      s = s & "{Arial=10,n}If you have ticked box 1, 2, 3 or 4 in Part 1 please show{x=46} less than 2500{x=65}2500 - 17999{x=81}18000 or more" & vbCrLf
      s = s & "{Arial=10,n}the expected level of annual business mileage for this car{x=50} {WBTEXTBOXL=" & xsize & "," & ysize & "}" & IIf(imileband = 1, "{xrel=60}{yrel=60}{Wding=12,b}ü{xrel=-60}{yrel=-60}{Arial=10,i}", "") & "{x=68}{WBTEXTBOXL=" & xsize & "," & ysize & "}" & IIf(imileband = 2, "{xrel=60}{yrel=60}{Wding=12,b}ü{xrel=-60}{yrel=-60}{Arial=10,i}", "") & "{x=86}{WBTEXTBOXL=" & xsize & "," & ysize & "}" & IIf(imileband = 3, "{xrel=60}{yrel=60}{Wding=12,b}ü{xrel=-60}{yrel=-60}{Arial=10,i}", "") & vbCrLf & vbCrLf
    
      imake = 0 'IIf(IsNull(("mmrequired")), False, ("mmrequired"))
      s = s & "{Arial=10,n}If you have ticked box 2 in Part 1 but the employee has more than one" & vbCrLf
      s = s & "{Arial=10,n}car available for private use please provide details of the car replaced {x=97}{Arial=10,r}Make and model   __________________________"
      If P46Car.PrintMake Then
       If (Not IsNull(("carreplaced"))) Then
         s = s & "{Times=10,RBI}" & Left("carreplaced", 35)
       End If
      End If
      s = s & vbCrLf & "{Arial=6,N}" & vbCrLf
      s = s & "{Arial=10,n}If you have ticked box 5 in Part 1 please provide details of the car{Arial=10,r}{x=97}Date withdrawn   ____________{x=97}{Times=10,RBI}" & IIf(iwithdraw, IIf(IsNull(("availto")), "", ("availto")), "") & vbCrLf
      s = s & "{Arial=10,n}withdrawn{x=70}(where appropriate)" & vbCrLf
      s = s & "{WBTEXTBOXL=99,12}" & vbCrLf & "{Arial=11,b}{x=3}Declaration" & vbCrLf
      s = s & "{Arial=11,n}{x=3}I declare that all particulars required are fully and truly stated according to the best of my knowledge and" & vbCrLf & "{x=3}belief"
      s = s & "{Arial=11,i}" & vbCrLf & "{x=12}Signature{x=40}{Arial=11,i}____________________________________________________" & vbCrLf
      s = s & "{Arial=6,N}" & vbCrLf & "{Arial=11,i}{x=12}Capacity in which signed{x=40}{Arial=11,i}____________________________________________________" & vbCrLf  '{x=45} & (SignCap)
      s = s & "{Arial=6,N}" & vbCrLf & "{Arial=11,i}{x=12}Date{x=40}____________"
      s = s & "{y=98}{x=0}{Arial=8,b}P46(Car)(" & "!!!YEAR" & ")(Substitute)(Arthur Andersen)"
    End With
    Call m_Reporter.EndReport
    Call m_Reporter.PreviewReport
  Next


P46Car_end:
  Set p46Cars = Nothing
  Call xReturn("P46Car")
  Exit Sub
P46Car_err:
  Call ErrorMessage(ERR_ERROR, Err, "P46Car", "ERR_UNDEFINED", "Undefined error.")
  Resume P46Car_end
  Resume
End Sub

Public Function InitPrint() As Boolean

  On Error GoTo InitPrint_Err
  Call xSet("InitPrint")
  
  'which screen are we on
  F_Print.Show 1 'all init inside the load event of the form

InitPrint_End:
  Call xReturn("InitPrint")
  Exit Function

InitPrint_Err:
  Call ErrorMessage(ERR_ERROR, Err, "InitPrint", "ERR_UNDEFINED", "Undefined error.")
  Resume InitPrint_End
End Function


Public Function Reports(PrintCode As Long) As Boolean
  
  Dim i As Long
  
  On Error GoTo Reports_Err
  Call xSet("Reports")

  Set m_Reporter = New Reporter
  
  For i = 1 To P11d32.CurrentEmployer.Employees.count
    Set CurrentPrintEmployee = P11d32.CurrentEmployer.Employees(i)
    If CurrentPrintEmployee Is Nothing Then GoTo NEXT_EMPLOYEE
    Select Case PrintCode
      Case 1
        Call P46Car
        
    End Select
  
NEXT_EMPLOYEE:
  Next
Reports_End:
  Set m_Reporter = Nothing
  Call xReturn("Reports")
  Exit Function

Reports_Err:
  Call ErrorMessage(ERR_ERROR, Err, "Reports", "ERR_UNDEFINED", "Undefined error.")
  Resume Reports_End
End Function

