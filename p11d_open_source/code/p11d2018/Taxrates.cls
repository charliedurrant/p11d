VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Rates"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"

Option Explicit
Option Base 0
Implements ISortFunction

Private m_SterlingRates As ObjectList
Private m_FrancRates As ObjectList
Private m_YenRates As ObjectList
Private m_IRFPCSbands As ObjectList




Public Enum BEN_CLASS_TO
  BCT_DBCLASS
  BCT_HMIT_BOX_NUMBER
  BCT_HMIT_SECTION
  BCT_HMIT_SECTION_STRING
  BCT_DBCATEGORY
  BCT_FORM_CAPTION
  BCT_BENITEMS_LAST_ITEM
  BCT_UDM_BENEFIT_TITLE
  BCT_TABLE
  BCT_CLASS1A_ABLE
  BCT_BENEFIT_TABLE
End Enum


Public Enum taxConstants
  [_FIRST_TAX_RATE] = 1
  carCCLOW = 1  ' longs
  carCCHIGH
  carMAXLISTPRICE
  carMAXCAPCONTRIB
  carDIESELLOW
  carDIESELHIGH
  carPETROLLOW
  carPETROLMED
  carPETROLHIGH
  carLOWLIMIT
  carHIGHLIMIT
  vanHighCharge
  vanLowCharge
  vanMinDays
  
  vanDailyCharge
  vanRate 'CAD 2010
  'fpcsLOWCC
  'fpcsMIDCC
  'fpcsHIGHCC
  fpcsMILEAGELIMIT
  mphoScaleCharge
  mphoMinDays
  accomExpensiveChargeLimit

  
  'Doubles
  carlowmilered
  carhighmilered
  carBENPERCENT
  carBENPERCENTMAX  'AM
  carFOURYRFACTOR
  LoanMortgageRate
  fpcsLOWCCLOWRATE
  fpcsMEDLOWCCLOWRATE
  fpcsMEDHIGHCCLOWRATE
  fpcsHIGHCCLOWRATE
  fpcsLOWCCHIGHRATE
  fpcsMEDLOWCCHIGHRATE
  fpcsMEDHIGHCCHIGHRATE
  fpcsHIGHCCHIGHRATE
  
  goodLOWPERCENTAGE
  goodHIGHPERCENTAGE
  accomInterestRate
  accomJobRelatedPercentage
  vanElectricMultiplier
  'Strings
  carCARHIGH
  carCARMED
  carNICRate
  CarVATSCALEDIESELHIGH
  CarVATSCALEDIESELLOW
  CarVATSCALEPETROLHIGH 'JN
  CarVATSCALEPETROLLOW
  WebSite
  CarVATSCALEPETROLMED
  CarFUELBENSETFIGURE
  
  CarCO2QUALIFYINGLEVEL 'AM
  CarSTDFACTOR
  CarHYBRIDFACTOR
  
  'AM
  carNOCO2CUTOFFDATE
  carBIFUELCUTOFFDATE
  CarDIESELSUPPLEMENT
  CarELECTRICDISCOUNT
  CarHYBRIDELECTRICDISCOUNT
  VATRate
  
  CarSTDDISCOUNTSUPPLEMENT
  
  [_LAST_DB_TAX_RATE] = CarSTDDISCOUNTSUPPLEMENT
  
  'see function setupdates for the setting of the following
  TaxYearStart
  TaxYearEnd
  DaysInYearLeap
  carFOURYEAROLD
  AppYear
  DaysInYear
  IsLeapyYear
  LeapDate
  NextTaxYearStart
  NextTaxYearEnd
  LastTaxYearEnd
  LastTaxYearStart
  vanOldDate
  CarRegDateDef
  VanRegDateNew 'AM
  VanRegDateOld 'AM
  LoanDailyElectionDue
  goodDEFAULTDATE
  accomFirstOccupiedDef
  P46Quarter1Start
  P46Quarter1End
  P46Quarter2Start
  P46Quarter2End
  P46Quarter3Start
  P46Quarter3End
  P46Quarter4Start
  P46Quarter4End
  YearOfStart
  TaxFormYearNext
  
  EmpLetPersonalTaxSubmissionDeadline
  EmpLetPersonalTaxRevenueWillCalcDeadline
  EmpLetP11DKeepDetailsUntill
  
  EmpLetEmployeeResponseDate
  
  RelocationThisYear
  RelocationLastYear
  
  TaxFormYear
  [_LAST_TAX_RATE] = TaxFormYear
End Enum

Private Type TaxRates_Value
  value As Variant
  Field As String
End Type

Public Enum INTERESTRATE_TYPE
  IT_STERLING
  IT_SWISSFRANCS
  IT_YEN
End Enum



Private m_taxRatesValues([_FIRST_TAX_RATE] To [_LAST_TAX_RATE]) As TaxRates_Value
Private m_MaxBenefitDescriptionLength As Long
Public IRFPCS As FPCS

Private Sub Class_Initialize()
  Dim db As Database
  Dim i As Long
  Dim s As String
  
  On Error GoTo TaxRates_Initialise_ERR
  
  Call xSet("TaxRates_Initialise")
  Set m_SterlingRates = New ObjectList
  Set m_FrancRates = New ObjectList
  Set m_YenRates = New ObjectList
  
  Set db = p11d32.PDDB(p11d32.P11DWS)
  If Not SetUpRates(db) Then Err.Raise ERR_RATES_INITIALISE, "TaxRates_Initialise", "Unable to setup TaxRates"
  If Not SetUpLoanRates(db) Then Err.Raise ERR_RATES_INITIALISE, "TaxRates_Initialise", "Unable to setup LoanRates"
  
  Set IRFPCS = New FPCS
  Call ReadIRFPCS(IRFPCS, db)
    
TaxRates_Initialise_END:
  Call xReturn("TaxRates_Initialise")
  Exit Sub
  
TaxRates_Initialise_ERR:
   
  Err.Raise Err.Number, "TaxRates_Initialise", Err.Description
  Resume
End Sub
Public Function BenefitMenuCaption(ByVal bc As BEN_CLASS, Optional bSubItem As Boolean = False) As String
  If bSubItem Then
    BenefitMenuCaption = p11d32.Rates.BenClassTo(bc, BCT_FORM_CAPTION)
  Else
    BenefitMenuCaption = p11d32.Rates.BenClassTo(bc, BCT_HMIT_SECTION_STRING) & " - " & p11d32.Rates.BenClassTo(bc, BCT_FORM_CAPTION)
  End If
End Function
Property Get AppOldType() As String
  AppOldType = "P11D"
End Property
Property Get ApplicationType() As String
  ApplicationType = AppOldType & "32"
End Property
Property Get TemplateApp() As String
  TemplateApp = ApplicationType & " Template"
End Property
Property Get FileExtensionCurrent() As String
  FileExtensionCurrent = p11d32.AppYearShort & S_DB_FILE_EXTENSION
End Property
Property Get FileExtensionPrior() As String
  Dim l As Long
  
  l = p11d32.AppYear - 1
  FileExtensionPrior = Right$(l, 2) & S_DB_FILE_EXTENSION
  
End Property
Private Function SetUpRates(db As Database) As Boolean
  Dim rs As Recordset
  Dim i As Long

  On Error GoTo SetUpRates_Err
  Call xSet("SetUpRates")
  Call SetFieldNames
  If db Is Nothing Then Call Err.Raise(ERR_IS_NOTHING, "SetUpRates", "Database is Nothing")
  Set rs = db.OpenRecordset("T_NumbersAndDates", dbOpenTable, dbFailOnError)
  If rs Is Nothing Then Call Err.Raise(ERR_IS_NOTHING, "SetUpRates", "Fail to open T_NumbersAndDates")
  If rs.EOF And rs.BOF Then Call Err.Raise(ERR_NO_RECORDS, "SetUpRates", "No entries in T_NumbersAndDates")
  rs.Index = "PrimaryKey"
  For i = [_FIRST_TAX_RATE] To [_LAST_DB_TAX_RATE]
    rs.Seek "=", m_taxRatesValues(i).Field
      If rs.Fields("VarType") = vbLong Then
      m_taxRatesValues(i).value = CLng(rs.Fields("Number"))
    ElseIf rs.Fields("VarType") = vbDouble Then
      m_taxRatesValues(i).value = rs.Fields("Number")
    ElseIf rs.Fields("VarType") = vbString Then
      m_taxRatesValues(i).value = rs.Fields("string")
    ElseIf rs.Fields("VarType") = vbBoolean Then
      m_taxRatesValues(i).value = CBoolean(rs.Fields("Boolean"))
    ElseIf rs.Fields("VarType") = vbDate Then
      m_taxRatesValues(i).value = rs.Fields("Date")
    Else
      Call Err.Raise(ERR_RATES_TABLE, "Rates_Class_Initialize", "Unknown VarType for tax rate " & m_taxRatesValues(i).Field)
    End If
  Next i
  rs.MoveLast
  
  If [_LAST_DB_TAX_RATE] <> rs.RecordCount Then Call Err.Raise(ERR_RATES_TABLE, "Rates_Class_Initialize", "The rates class cannot be initialized because the rates database is out of date. " & "Expected " & [_LAST_DB_TAX_RATE] & " records, found " & rs.RecordCount & ".")
  
  Call SetUpDates

  SetUpRates = True
SetUpRates_End:
  Set rs = Nothing
  Call xReturn("SetUpRates")
  Exit Function
SetUpRates_Err:
  Call ErrorMessage(ERR_ERROR, Err, "SetUpRates", "Read Tax Rates", "Unable to load the tax rates.")
  Resume SetUpRates_End
  Resume
End Function
Private Sub SetUpDates()
  Dim lYear As Long
  
  On Error GoTo SetUpDates_ERR
  
  Call xSet("SetUpDates")
  
  m_taxRatesValues(TaxYearStart).value = TryConvertDate("06/04/" & p11d32.AppYear)
  m_taxRatesValues(TaxYearEnd).value = DateAdd("d", -1, DateAdd("YYYY", 1, m_taxRatesValues(TaxYearStart).value))
  
  m_taxRatesValues(NextTaxYearStart).value = DateAdd("d", 1, m_taxRatesValues(TaxYearEnd).value)
  
  m_taxRatesValues(DaysInYear).value = DateDiff("d", m_taxRatesValues(TaxYearStart).value, m_taxRatesValues(TaxYearEnd).value) + 1
  
  m_taxRatesValues(IsLeapyYear).value = IsLeapYear(m_taxRatesValues(TaxYearEnd).value)
  
  If m_taxRatesValues(IsLeapyYear).value Then
    m_taxRatesValues(DaysInYearLeap).value = m_taxRatesValues(DaysInYear).value - 1
    m_taxRatesValues(LeapDate).value = DateSerial(Year(m_taxRatesValues(TaxYearEnd).value), 2, 29)
  Else
    m_taxRatesValues(LeapDate).value = UNDATED
    m_taxRatesValues(DaysInYearLeap).value = m_taxRatesValues(DaysInYear).value
  End If
  
  m_taxRatesValues(NextTaxYearEnd).value = DateAdd("YYYY", 1, m_taxRatesValues(TaxYearEnd).value)
  
  m_taxRatesValues(LastTaxYearStart).value = DateAdd("YYYY", -1, m_taxRatesValues(TaxYearStart).value)
  m_taxRatesValues(LastTaxYearEnd).value = DateAdd("YYYY", -1, m_taxRatesValues(TaxYearEnd).value)
  
  m_taxRatesValues(CarRegDateDef).value = TryConvertDate("01/03/" & (Year(m_taxRatesValues(TaxYearStart).value)), UNDATED)
  
  'AM
  m_taxRatesValues(VanRegDateNew).value = TryConvertDate("01/03/" & (Year(m_taxRatesValues(TaxYearStart).value)), UNDATED)
  m_taxRatesValues(VanRegDateOld).value = TryConvertDate("01/08/" & (Year(m_taxRatesValues(TaxYearStart).value) - 4), UNDATED)
  
  m_taxRatesValues(accomFirstOccupiedDef).value = TryConvertDate("01/01/" & (Year(m_taxRatesValues(TaxYearStart).value) - 7), UNDATED)
  m_taxRatesValues(LoanDailyElectionDue).value = TryConvertDate("31/01/" & (Year(m_taxRatesValues(TaxYearStart).value) + 3), UNDATED)
  m_taxRatesValues(goodDEFAULTDATE).value = TryConvertDate("31/12/" & (Year(m_taxRatesValues(TaxYearStart).value) - 7), UNDATED)
  
  m_taxRatesValues(carFOURYEAROLD).value = DateAdd("YYYY", -3, m_taxRatesValues(TaxYearStart).value)
    
  m_taxRatesValues(TaxFormYear).value = Year(m_taxRatesValues(TaxYearStart).value) & "/" & Year(m_taxRatesValues(TaxYearEnd).value)
  
  '**********************
  m_taxRatesValues(TaxFormYearNext).value = Year(m_taxRatesValues(NextTaxYearStart).value) & "/" & Year(m_taxRatesValues(NextTaxYearEnd).value)
  
  
  m_taxRatesValues(EmpLetPersonalTaxSubmissionDeadline).value = TryConvertDate("31/01/" & Year(m_taxRatesValues(TaxYearEnd).value) + 1)
  m_taxRatesValues(EmpLetPersonalTaxRevenueWillCalcDeadline).value = TryConvertDate("30/09/" & Year(m_taxRatesValues(TaxYearEnd).value))
  m_taxRatesValues(EmpLetP11DKeepDetailsUntill).value = TryConvertDate("31/01/" & Year(m_taxRatesValues(NextTaxYearEnd).value) + 1)
  
  
  '***************** END
  m_taxRatesValues(AppYear).value = Right$(Year(m_taxRatesValues(TaxYearStart).value), 2)
  
  m_taxRatesValues(P46Quarter1Start).value = m_taxRatesValues(TaxYearStart).value
  m_taxRatesValues(P46Quarter1End).value = DateAdd("M", -9, m_taxRatesValues(TaxYearEnd).value)
  
  m_taxRatesValues(P46Quarter2Start).value = DateAdd("M", 3, m_taxRatesValues(TaxYearStart).value)
  m_taxRatesValues(P46Quarter2End).value = DateAdd("M", -6, m_taxRatesValues(TaxYearEnd).value)
  
  m_taxRatesValues(P46Quarter3Start).value = DateAdd("M", 6, m_taxRatesValues(TaxYearStart).value)
  m_taxRatesValues(P46Quarter3End).value = DateAdd("M", -3, m_taxRatesValues(TaxYearEnd).value)
  
  m_taxRatesValues(vanOldDate).value = DateAdd("yyyy", -3, m_taxRatesValues(TaxYearStart).value)
  
  m_taxRatesValues(P46Quarter4Start).value = DateAdd("M", 9, m_taxRatesValues(TaxYearStart).value)
  m_taxRatesValues(P46Quarter4End).value = m_taxRatesValues(TaxYearEnd).value
  
  m_taxRatesValues(YearOfStart).value = Format$(m_taxRatesValues(TaxYearStart).value, "yy")
  
  m_taxRatesValues(RelocationThisYear).value = Year(m_taxRatesValues(TaxYearStart).value) & "/" & Right$(Year(m_taxRatesValues(TaxYearEnd).value), 2)
  m_taxRatesValues(RelocationLastYear).value = Year(m_taxRatesValues(LastTaxYearStart).value) & "/" & Right$(Year(m_taxRatesValues(LastTaxYearEnd).value), 2)
  
  m_taxRatesValues(EmpLetEmployeeResponseDate).value = TryConvertDate("30/06/" & Year(m_taxRatesValues(TaxYearEnd).value))
  
SetUpDates_END:
  Call xReturn("SetUpDates")
  Exit Sub
SetUpDates_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "SetUpDates", "Set Up Dates", "Error setting up date information.")
  Resume SetUpDates_END
  Resume
End Sub
Private Function SetUpLoanRates(db As Database) As Boolean
  
  Dim rs As Recordset
  Dim li As LoanInterest
  On Error GoTo SetUpLoanRates_Err
  Call xSet("SetUpLoanRates")
  
  Set m_SterlingRates = New ObjectList
  Set m_FrancRates = New ObjectList
  Set m_YenRates = New ObjectList
  Set rs = db.OpenRecordset(sql.Queries(SELECT_LOAN_RATES), dbOpenForwardOnly)
  Do While Not rs.EOF
    Set li = New LoanInterest
    li.dtStartDate = rs.Fields("InterestDate")
    li.dRate = rs.Fields("Sterling")
    Call m_SterlingRates.Add(li)
    
    Set li = New LoanInterest
    li.dtStartDate = rs.Fields("InterestDate")
    li.dRate = rs.Fields("Yen")
    Call m_YenRates.Add(li)
    
    Set li = New LoanInterest
    li.dtStartDate = rs.Fields("InterestDate")
    li.dRate = rs.Fields("Franc")
    Call m_FrancRates.Add(li)
    rs.MoveNext
  Loop
  SetUpLoanRates = True
  
SetUpLoanRates_End:
  Set li = Nothing
  Set rs = Nothing
  Call xReturn("SetUpLoanRates")
  Exit Function
  
SetUpLoanRates_Err:
  Call ErrorMessage(ERR_ERROR, Err, "SetUpLoanRates", "ERR_LOAN_RATES", "Unable to initialise the loan interest rates.")
  Resume SetUpLoanRates_End
  Resume
End Function

Public Function ReadIRFPCS(scheme As FPCS, db As Database) As Long
  Dim i As Long
  Dim rs As Recordset
  Dim csb As FPCSBand, CSB2 As FPCSBand, lGreaterThanCC As Long
  
  On Error GoTo ReadIRFPCS_Err
  Call xSet("ReadIRFPCS_Err")
  
  If Not db Is Nothing Then
    Set rs = db.OpenRecordset(sql.Queries(SELECT_IR_FPCS), dbOpenForwardOnly)
    scheme.Name = S_IRFPCS
    
    Do While Not rs.EOF
      Set csb = New FPCSBand
      Set CSB2 = New FPCSBand
      
      csb.Name = "" & rs.Fields("CCBand")
      ' Call SplitFPCSBand(lGreaterThanCC, csb.Name) EK no longer need to split a band
      CSB2.Name = csb.Name
      
      csb.GreaterThanCC = lGreaterThanCC
      csb.GreaterThanMiles = 0
      If Not IsNull(rs.Fields("Lowrate")) Then csb.Rate = rs.Fields("Lowrate")
      
      CSB2.GreaterThanCC = lGreaterThanCC
      CSB2.GreaterThanMiles = m_taxRatesValues(fpcsMILEAGELIMIT).value
      If Not IsNull(rs.Fields("HighRate")) Then CSB2.Rate = rs.Fields("HighRate")
      
      Call scheme.Bands.Add(csb)
      Call scheme.Bands.Add(CSB2)
      rs.MoveNext
      ReadIRFPCS = ReadIRFPCS + 1
    Loop
  End If
  
ReadIRFPCS_End:
  Call xReturn("ReadIRFPCS_Err")
  Exit Function
  
ReadIRFPCS_Err:
  
  Call ErrorMessage(ERR_ERROR, Err, "ReadIRFPCS", "Read IR FPCS", "Error reading the HMRC FPCS from the database.")
  Resume ReadIRFPCS_End
End Function

Public Function IRCO2Makes() As Recordset

Dim rs As Recordset
Dim db As Database

  On Error GoTo IRCO2Makes_Err
  Call xSet("IRCO2Makes_Err")

  Set db = p11d32.PDDB(p11d32.P11DWS)
  Set IRCO2Makes = db.OpenRecordset(sql.Queries(SELECT_IR_CO2EMISSIONS_MAKES), dbOpenForwardOnly)

IRCO2Makes_End:
  Call xReturn("ReadIRFPCS_Err")
  Exit Function
  
IRCO2Makes_Err:
  Call ErrorMessage(ERR_ERROR, Err, "IRCO2Makes", "Read IR CO2 Makes", "Error reading the HMRC CO2 Emissions by Car Make.")
  Resume IRCO2Makes_End
End Function

Public Function GetInterestRatesInYear(ByVal IT As INTERESTRATE_TYPE) As ObjectList
  Dim OL As ObjectList
  Dim CurrencyRates As ObjectList
  Dim dStart As Date, dEnd As Date
  Dim i As Long
  Dim lNoOfRates As Long
  Dim li As LoanInterest
  Dim liLast As LoanInterest
  Dim b As Boolean
  
  On Error GoTo GetInterestRatesInYear_ERR
  Call xSet("GetInterestRatesInYear")
  
  Set OL = New ObjectList
  
  Set CurrencyRates = GetCurrencyInterestRates(IT)
  If CurrencyRates Is Nothing Then Call Err.Raise(ERR_IS_NOTHING, "GetInterestRatesInYear", "The curreny interest rates are nothing.")
  
  dStart = p11d32.Rates.value(TaxYearStart)
  dEnd = p11d32.Rates.value(TaxYearEnd)
  lNoOfRates = CurrencyRates.Count
  For i = 1 To CurrencyRates.Count
    Set li = CurrencyRates(i)
    If Not li Is Nothing Then
      If li.dtStartDate <= dEnd And li.dtStartDate >= dStart Or (i = lNoOfRates) Then
        If (li.dtStartDate > dStart) And Not b Then
          If Not liLast Is Nothing Then Call OL.Add(liLast)
        End If
        b = True
        Call OL.Add(li)
      End If
      Set liLast = li
    End If
  Next
  
  Set GetInterestRatesInYear = OL
  
GetInterestRatesInYear_END:
  Call xReturn("GetInterestRatesInYear")
  Exit Function
GetInterestRatesInYear_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "GetInterestRatesInYear", "Get Interest Rates In Year", "Error getting the currency dependant interest rates in the year.")
  Resume GetInterestRatesInYear_END
  Resume
End Function
Private Function GetCurrencyInterestRates(ByVal IT As INTERESTRATE_TYPE) As ObjectList
  Select Case IT
    Case INTERESTRATE_TYPE.IT_STERLING
      Set GetCurrencyInterestRates = m_SterlingRates
    Case INTERESTRATE_TYPE.IT_SWISSFRANCS
      Set GetCurrencyInterestRates = m_FrancRates
    Case INTERESTRATE_TYPE.IT_YEN
      Set GetCurrencyInterestRates = m_YenRates
    Case Else
      ECASE ("Invalid currency type in GetCurrencyInterestRates.")
  End Select
End Function
Public Function GetInterestRate(ByVal d0 As Date, IT As INTERESTRATE_TYPE) As Double
  Dim i As Long
  Dim li As LoanInterest
  Dim OL As ObjectList
  
  On Error GoTo GetInterestRate_Err
  Call xSet("GetInterestRate")
  
  Set OL = GetCurrencyInterestRates(IT)
  
  For i = OL.Count To 1 Step -1
    Set li = OL.Item(i)
    If d0 >= li.dtStartDate Then
      GetInterestRate = li.dRate
      Exit For
    End If
  Next i

GetInterestRate_End:
  Call xReturn("GetInterestRate")
  Exit Function
GetInterestRate_Err:
  GetInterestRate = 0#
  Resume GetInterestRate_End
End Function

'zzzz Only works with sterling at the moment
' ChangeDates array 1 to Count
Public Function RateChangeinPeriod(ByVal dFrom As Date, ByVal dTo As Date, ChangeDates() As Date) As Long
  Dim dRate As Double
  Dim li As LoanInterest
  Dim i As Long, StartI As Long
  Dim MaxIndex As Long, CurIndex As Long
  
  On Error GoTo RateChangeinPeriod_Err
  Call xSet("RateChangeinPeriod")
  MaxIndex = UBound(ChangeDates) - LBound(ChangeDates)
  CurIndex = 1
  
  ' Find the date before the beginning of the period
  'apf check whether inclusive/exclusive
  For i = m_SterlingRates.Count To 1 Step -1
    Set li = m_SterlingRates.Item(i)
    If dFrom > li.dtStartDate Then
      dRate = li.dRate
      StartI = i
      Exit For
    End If
  Next i
  
  ' Loop from the date found above to the end date recording any rate changes
  For i = StartI To m_SterlingRates.Count
    Set li = m_SterlingRates.Item(i)
    If dTo < li.dtStartDate Then
      Exit For
    End If
    If li.dRate <> dRate Then
      If CurIndex < MaxIndex Then
        If MaxIndex < 0 Then MaxIndex = 0
        MaxIndex = MaxIndex + 1
        ReDim ChangeDates(1 To MaxIndex) As Date
      End If
      ChangeDates(CurIndex) = li.dtStartDate
      CurIndex = CurIndex + 1
    End If
  Next i
  RateChangeinPeriod = (CurIndex - 1)
RateChangeinPeriod_End:
  Call xReturn("RateChangeinPeriod")
  Exit Function

RateChangeinPeriod_Err:
  Resume RateChangeinPeriod_End
End Function



Private Sub SetFieldNames()
  On Error GoTo SetFieldNames_Err
  Call xSet("SetFieldNames")
  'Initialise the field names
  m_taxRatesValues(carCCLOW).Field = "L_CCCARLOW"
  m_taxRatesValues(carCCHIGH).Field = "L_CCCARHIGH"
  m_taxRatesValues(carMAXLISTPRICE).Field = "L_MAXLISTPRICE"
  m_taxRatesValues(carMAXCAPCONTRIB).Field = "L_MAXCAPCONTRIB"
  m_taxRatesValues(carDIESELLOW).Field = "L_DIESELLOW"
  m_taxRatesValues(carDIESELHIGH).Field = "L_DIESELHIGH"
  m_taxRatesValues(carPETROLLOW).Field = "L_PETROLLOW"
  m_taxRatesValues(carPETROLMED).Field = "L_PETROLMED"
  m_taxRatesValues(carPETROLHIGH).Field = "L_PETROLHIGH"
  m_taxRatesValues(carLOWLIMIT).Field = "L_LOWLIMIT"
  m_taxRatesValues(carHIGHLIMIT).Field = "L_HIGHLIMIT"
  m_taxRatesValues(vanHighCharge).Field = "L_VANHIGH"
  m_taxRatesValues(vanLowCharge).Field = "L_VANLOW"
  m_taxRatesValues(vanMinDays).Field = "L_VANMINDAYS"
  
  
  m_taxRatesValues(vanDailyCharge).Field = "L_VANDAILY"
  m_taxRatesValues(fpcsMILEAGELIMIT).Field = "L_FPCSMILELIMIT"
  m_taxRatesValues(mphoScaleCharge).Field = "L_PHOSCALECHARGE"
  m_taxRatesValues(mphoMinDays).Field = "L_PHOMINDAYS"
  m_taxRatesValues(accomExpensiveChargeLimit).Field = "L_ACCOMLIMIT"
  
    
  m_taxRatesValues(carCARHIGH).Field = "S_CARHIGH"
  m_taxRatesValues(carCARMED).Field = "S_CARMED"
  m_taxRatesValues(TaxYearStart).Field = "DT_TAXYRSTART"
  m_taxRatesValues(carlowmilered).Field = "D_LOWMILERED"
  m_taxRatesValues(carhighmilered).Field = "D_HIGHMILERED"
  m_taxRatesValues(carBENPERCENT).Field = "D_BENPERCENT"
  m_taxRatesValues(carBENPERCENTMAX).Field = "D_BENPERCENTMAX"  'AM
  m_taxRatesValues(carFOURYRFACTOR).Field = "D_FOURYRFACTOR"

  m_taxRatesValues(LoanMortgageRate).Field = "D_MORTGAGE"
  m_taxRatesValues(fpcsLOWCCLOWRATE).Field = "D_FPCSLOWCCLOWRATE"
  m_taxRatesValues(fpcsMEDLOWCCLOWRATE).Field = "D_FPCSMEDLOWCCLOWRATE "
  m_taxRatesValues(fpcsMEDHIGHCCLOWRATE).Field = "D_FPCSMEDHIGHCCLOWRATE"
  m_taxRatesValues(fpcsHIGHCCLOWRATE).Field = "D_FPCSHIGHCCLOWRATE"
  m_taxRatesValues(fpcsLOWCCHIGHRATE).Field = "D_FPCSLOWCCHIGHRATE"
  m_taxRatesValues(fpcsMEDLOWCCHIGHRATE).Field = "D_FPCSMEDLOWCCHIGHRATE "
  m_taxRatesValues(fpcsMEDHIGHCCHIGHRATE).Field = "D_FPCSMEDHIGHCCHIGHRATE"
  m_taxRatesValues(fpcsHIGHCCHIGHRATE).Field = "D_FPCSHIGHCCHIGHRATE"
  m_taxRatesValues(goodHIGHPERCENTAGE).Field = "D_GOODHIGHPERCENTAGE"
  m_taxRatesValues(goodLOWPERCENTAGE).Field = "D_GOODLOWPERCENTAGE"
  m_taxRatesValues(accomInterestRate).Field = "D_ACCOMINTRATE"
  m_taxRatesValues(accomJobRelatedPercentage).Field = "D_ACCOMJOBRELPER"
  m_taxRatesValues(vanElectricMultiplier).Field = "D_VAN_ELECTRIC_MULTIPLIER"
  
  
  
  'this is really a general NIC Rate used for all benefits 98 only had NIC on car fuel
  m_taxRatesValues(carNICRate).Field = "D_CAR_NIC_RATE"
  m_taxRatesValues(WebSite).Field = "S_WEB_SITE"
  
  m_taxRatesValues(VATRate).Field = "D_VAT"
  
  
  m_taxRatesValues(CarVATSCALEDIESELLOW).Field = "L_VATSCALEDIESELLOW"
  m_taxRatesValues(CarVATSCALEDIESELHIGH).Field = "L_VATSCALEDIESELHIGH"
  m_taxRatesValues(CarVATSCALEPETROLLOW).Field = "L_VATSCALEPETROLLOW"
  m_taxRatesValues(CarVATSCALEPETROLMED).Field = "L_VATSCALEPETROLMED"
  m_taxRatesValues(CarVATSCALEPETROLHIGH).Field = "L_VATSCALEPETROLHIGH"
  m_taxRatesValues(CarFUELBENSETFIGURE).Field = "L_FUELBEN_SETFIGURE"
  
  m_taxRatesValues(CarCO2QUALIFYINGLEVEL).Field = "L_CO2_QUALIFYINGLEVEL" 'AM
  m_taxRatesValues(CarSTDFACTOR).Field = "L_STDFACTOR"
  m_taxRatesValues(CarHYBRIDFACTOR).Field = "L_HYBRIDFACTOR"
  
  'AM
  m_taxRatesValues(carNOCO2CUTOFFDATE).Field = "D_NOCO2_CUTOFF"
  m_taxRatesValues(carBIFUELCUTOFFDATE).Field = "D_BIFUEL_CUTOFF"

  m_taxRatesValues(CarDIESELSUPPLEMENT).Field = "D_DIESELSUPPLEMENT" 'AM
  m_taxRatesValues(CarELECTRICDISCOUNT).Field = "D_ELECTRICDISCOUNT" 'AM
  m_taxRatesValues(CarHYBRIDELECTRICDISCOUNT).Field = "D_HYBRIDELECTRICDISC" 'AM
  m_taxRatesValues(CarSTDDISCOUNTSUPPLEMENT).Field = "D_STANDARD_DISC_SUPP" 'AM
  
  
  m_taxRatesValues(vanRate).Field = "L_VAN_RATE"
    
SetFieldNames_End:
  Call xReturn("SetFieldNames")
  Exit Sub

SetFieldNames_Err:
  Call ErrorMessage(ERR_ERROR, Err, "SetFieldNames", "Setup TaxRate Field Names", "Unable to setup Field names.")
  Resume SetFieldNames_End
  Resume
End Sub

Public Function value(ByVal Item As taxConstants) As Variant
  value = m_taxRatesValues(Item).value
End Function

Public Function SetValue(ByVal Item As taxConstants, value As Variant) As Boolean
  m_taxRatesValues(Item).value = value
  SetValue = True
End Function

Public Function GetP46QuarterStartEnd(dStart As Date, dEnd As Date, lQuarter As Long) As Boolean

  On Error GoTo GetP46QuarterStartEnd_Err
  Call xSet("GetP46QuarterStartEnd")

  Select Case lQuarter
    Case 1
      dStart = p11d32.Rates.value(P46Quarter1Start)
      dEnd = p11d32.Rates.value(P46Quarter1End)
    Case 2
      dStart = p11d32.Rates.value(P46Quarter2Start)
      dEnd = p11d32.Rates.value(P46Quarter2End)
    Case 3
      dStart = p11d32.Rates.value(P46Quarter3Start)
      dEnd = p11d32.Rates.value(P46Quarter3End)
    Case 4
      dStart = p11d32.Rates.value(P46Quarter4Start)
      dEnd = p11d32.Rates.value(P46Quarter4End)
    Case Else
      dStart = p11d32.Rates.value(TaxYearStart)
      dEnd = p11d32.Rates.value(TaxYearEnd)
      
  
  End Select


GetP46QuarterStartEnd_End:
  Call xReturn("GetP46QuarterStartEnd")
  Exit Function

GetP46QuarterStartEnd_Err:
  Call ErrorMessage(ERR_ERROR, Err, "GetP46QuarterStartEnd", "Get P46 Quarter Start End", "Error finding the P46 Car quarter start and end dates.")
  Resume GetP46QuarterStartEnd_End
End Function

Public Function SplitFPCSBand(lGreaterThanCC As Long, ByVal Band As String) As Boolean
  Dim lPos As Long
  'EK problem with this function
  On Error GoTo SplitFPCSBand_Err
  Call xSet("SplitFPCSBand")

  'we need to split "1001 to 2000" and deal with "over 3000"
  Band = Trim$(Band)
  If Len(Band) > 0 Then
    lPos = InStr(1, Band, "To", vbTextCompare)
    If lPos Then
      
      lGreaterThanCC = CLng(Val(Band))
      If lGreaterThanCC <> 0 Then lGreaterThanCC = lGreaterThanCC - 1
    Else ' deal with over 9999
      lPos = InStr(0, Band, "over", vbTextCompare)
      If lPos = 0 Then Call Err.Raise(ERR_FPCSBAND, "SplitFPCSBand", "Unable to decode FPCS band: " & Band)
      Band = Mid$(Band, lPos + 4)
      lGreaterThanCC = CLng(Val(Band))
    End If
  End If
  SplitFPCSBand = True
  
SplitFPCSBand_End:
  Call xReturn("SplitFPCSBand")
  Exit Function

SplitFPCSBand_Err:
  SplitFPCSBand = False
  Call ErrorMessage(ERR_ERROR, Err, "SplitFPCSBand", "Split FPCS Band", "Error splitting the FPCS Band text.")
  Resume SplitFPCSBand_End
End Function

Public Function BenClassTo(bc As BEN_CLASS, bCToType As BEN_CLASS_TO) As Variant

On Error GoTo BenClassTo_Err

  Call xSet("BenClassTo")
  
  If (bCToType = BCT_BENEFIT_TABLE) Then
    BenClassTo = p11d32.BenDataLinkBenfitTable(bc)
    GoTo BenClassTo_End
  End If
  
  
  Select Case bc
    Case BC_ASSETSTRANSFERRED_A
      Select Case bCToType
        Case BCT_HMIT_BOX_NUMBER
          BenClassTo = "13"
        Case BCT_HMIT_SECTION
          BenClassTo = HMIT_A
        Case BCT_HMIT_SECTION_STRING
          BenClassTo = HMITSectionToHMITCode(BenClassTo(bc, BCT_HMIT_SECTION))
        Case BCT_FORM_CAPTION, BCT_UDM_BENEFIT_TITLE
          BenClassTo = "Assets transferred"
        Case BCT_BENITEMS_LAST_ITEM
          BenClassTo = trans_LASTITEM
        Case BCT_TABLE
          BenClassTo = "T_BenTrans"
        Case BCT_CLASS1A_ABLE
          BenClassTo = True
       Case Else
          ECASE ("Invalid ben class to type, benclass=" & bc & ",ben class to type=" & bCToType)
      End Select
    Case BC_TAX_NOTIONAL_PAYMENTS_B
      Select Case bCToType
        Case BCT_DBCLASS
          BenClassTo = S_TAX_NOTIONAL_PAYMENTS_B
        Case BCT_HMIT_BOX_NUMBER
          BenClassTo = "15"
        Case BCT_HMIT_SECTION
          BenClassTo = HMIT_B
        Case BCT_HMIT_SECTION_STRING
          BenClassTo = HMITSectionToHMITCode(BenClassTo(bc, BCT_HMIT_SECTION))
        Case BCT_FORM_CAPTION, BCT_UDM_BENEFIT_TITLE
          BenClassTo = "Tax on notional payments"
        Case BCT_DBCATEGORY
          BenClassTo = "MISC"
        Case BCT_BENITEMS_LAST_ITEM
          BenClassTo = oth_LASTITEM
        Case BCT_TABLE
          BenClassTo = "T_BenOther"
        Case BCT_CLASS1A_ABLE
          BenClassTo = False
        Case Else
          ECASE ("Invalid ben class to type, benclass=" & bc & ",ben class to type=" & bCToType)
      End Select
    Case BC_PAYMENTS_ON_BEFALF_B
      Select Case bCToType
        Case BCT_DBCLASS
          BenClassTo = S_PAYMENTS_ON_BEFALF_B
        Case BCT_HMIT_BOX_NUMBER
          BenClassTo = "15"
        Case BCT_HMIT_SECTION
          BenClassTo = HMIT_B
        Case BCT_HMIT_SECTION_STRING
          BenClassTo = HMITSectionToHMITCode(BenClassTo(bc, BCT_HMIT_SECTION))
        Case BCT_FORM_CAPTION
          BenClassTo = "Payments made on behalf of employee"
        Case BCT_DBCATEGORY
          BenClassTo = "MISC"
        Case BCT_BENITEMS_LAST_ITEM
          BenClassTo = oth_LASTITEM
        Case BCT_UDM_BENEFIT_TITLE
          BenClassTo = "Payments made"
        Case BCT_TABLE
          BenClassTo = "T_BenOther"
        Case BCT_CLASS1A_ABLE
          BenClassTo = False
        Case Else
          ECASE ("Invalid ben class to type, benclass=" & bc & ",ben class to type=" & bCToType)
      End Select
    Case BC_VOUCHERS_AND_CREDITCARDS_C
      Select Case bCToType
        Case BCT_DBCLASS
          BenClassTo = S_VOUCHERS_AND_CREDITCARDS_C
        Case BCT_HMIT_BOX_NUMBER
          BenClassTo = "12"
        Case BCT_HMIT_SECTION
          BenClassTo = HMIT_C
        Case BCT_HMIT_SECTION_STRING
          BenClassTo = HMITSectionToHMITCode(BenClassTo(bc, BCT_HMIT_SECTION))
        Case BCT_FORM_CAPTION
          BenClassTo = "Vouchers and credit cards"
        Case BCT_DBCATEGORY
          BenClassTo = "SUBS"
        Case BCT_BENITEMS_LAST_ITEM
          BenClassTo = oth_LASTITEM
        Case BCT_UDM_BENEFIT_TITLE
          BenClassTo = "Vouchers/Credit"
        Case BCT_TABLE
          BenClassTo = "T_BenOther"
        Case BCT_CLASS1A_ABLE
          BenClassTo = False
        Case Else
          ECASE ("Invalid ben class to type, benclass=" & bc & ",ben class to type=" & bCToType)
      End Select
    Case BC_LIVING_ACCOMMODATION_D
      Select Case bCToType
        Case BCT_HMIT_BOX_NUMBER
          BenClassTo = "14"
        Case BCT_HMIT_SECTION
          BenClassTo = HMIT_D
        Case BCT_HMIT_SECTION_STRING
          BenClassTo = HMITSectionToHMITCode(BenClassTo(bc, BCT_HMIT_SECTION))
        Case BCT_FORM_CAPTION
          BenClassTo = "Living accommodation"
        Case BCT_BENITEMS_LAST_ITEM
          BenClassTo = Accom_LASTITEM
        Case BCT_UDM_BENEFIT_TITLE
          BenClassTo = "Accommodation"
        Case BCT_TABLE
          BenClassTo = "T_BenAccom"
        Case BCT_CLASS1A_ABLE
          BenClassTo = True
        Case Else
          ECASE ("Invalid ben class to type, benclass=" & bc & ",ben class to type=" & bCToType)
      End Select
    Case BC_EMPLOYEE_CAR_E
      Select Case bCToType
        Case BCT_HMIT_BOX_NUMBER
          BenClassTo = "12"
        Case BCT_HMIT_SECTION
          BenClassTo = HMIT_E
        Case BCT_HMIT_SECTION_STRING
          BenClassTo = HMITSectionToHMITCode(BenClassTo(bc, BCT_HMIT_SECTION))
        Case BCT_FORM_CAPTION
          BenClassTo = "Mileage allowance / Employee cars"
        Case BCT_BENITEMS_LAST_ITEM
          BenClassTo = eecar_lastitem
        Case BCT_UDM_BENEFIT_TITLE
          BenClassTo = "Employee car"
        Case BCT_TABLE
          BenClassTo = "T_BenEECar"
        Case BCT_CLASS1A_ABLE
          BenClassTo = False
        Case Else
          ECASE ("Invalid ben class to type, benclass=" & bc & ",ben class to type=" & bCToType)
      End Select
    Case BC_COMPANY_CARS_F
      Select Case bCToType
        Case BCT_HMIT_BOX_NUMBER
          BenClassTo = "9"
        Case BCT_HMIT_SECTION
          BenClassTo = HMIT_F
        Case BCT_HMIT_SECTION_STRING
          BenClassTo = HMITSectionToHMITCode(BenClassTo(bc, BCT_HMIT_SECTION))
        Case BCT_FORM_CAPTION, BCT_UDM_BENEFIT_TITLE
          BenClassTo = "Company cars"
        Case BCT_BENITEMS_LAST_ITEM
          BenClassTo = car_lastitem
        Case BCT_TABLE
          BenClassTo = "T_BenCar"
        Case BCT_CLASS1A_ABLE
          BenClassTo = True
        Case Else
          ECASE ("Invalid ben class to type, benclass=" & bc & ",ben class to type=" & bCToType)
      End Select
    Case BC_FUEL_F
      Select Case bCToType
        Case BCT_HMIT_BOX_NUMBER
          BenClassTo = "10"
        Case BCT_HMIT_SECTION
          BenClassTo = HMIT_F
        Case BCT_HMIT_SECTION_STRING
          BenClassTo = HMITSectionToHMITCode(BenClassTo(bc, BCT_HMIT_SECTION))
        Case BCT_FORM_CAPTION
          BenClassTo = "Company car fuel"
        Case BCT_BENITEMS_LAST_ITEM
          BenClassTo = [_ITEM_LAST_ITEM]
        Case BCT_UDM_BENEFIT_TITLE
          BenClassTo = "Car fuel"
        Case BCT_TABLE
          BenClassTo = "T_BenCar"
        Case BCT_CLASS1A_ABLE
          BenClassTo = True
        Case Else
          ECASE ("Invalid ben class to type, benclass=" & bc & ",ben class to type=" & bCToType)
      End Select
    Case BC_nonSHAREDVAN_G, BC_NONSHAREDVANS_G, BC_SHAREDVAN_G, BC_SHAREDVANs_G
      Select Case bCToType
        Case BCT_HMIT_BOX_NUMBER
          BenClassTo = "9"
        Case BCT_HMIT_SECTION
          BenClassTo = HMIT_G
        Case BCT_HMIT_SECTION_STRING
          BenClassTo = HMITSectionToHMITCode(BenClassTo(bc, BCT_HMIT_SECTION))
        Case BCT_FORM_CAPTION, BCT_UDM_BENEFIT_TITLE
          BenClassTo = "Vans"
        Case BCT_BENITEMS_LAST_ITEM
          Select Case bc
            Case BC_nonSHAREDVAN_G
              BenClassTo = nsvan_LASTITEM
            Case BC_NONSHAREDVANS_G
              BenClassTo = nsvans_LASTITEM
            Case BC_SHAREDVAN_G
              BenClassTo = shvan_LASTITEM
            Case BC_SHAREDVANs_G
              BenClassTo = 1
            Case Else
              ECASE ("Invalid ben class to benefits last item, benclass=" & bc & ",ben class to type=" & bCToType)
          End Select
        Case BCT_TABLE
          BenClassTo = "T_BenVan"
        Case BCT_CLASS1A_ABLE
          BenClassTo = True
        Case Else
          ECASE ("Invalid ben class to type, benclass=" & bc & ",ben class to type=" & bCToType)
      End Select
    Case BC_NONSHAREDVANS_FUEL_G
      Select Case bCToType
        Case BCT_HMIT_BOX_NUMBER
          BenClassTo = "10"
        Case BCT_HMIT_SECTION
          BenClassTo = HMIT_G
        Case BCT_HMIT_SECTION_STRING
          BenClassTo = HMITSectionToHMITCode(BenClassTo(bc, BCT_HMIT_SECTION))
        Case BCT_FORM_CAPTION, BCT_UDM_BENEFIT_TITLE
          BenClassTo = "Vans - fuel"
        Case BCT_CLASS1A_ABLE
          BenClassTo = True
        Case Else
          ECASE ("Invalid ben class to type, benclass=" & bc & ",ben class to type=" & bCToType)
      End Select
    Case BC_LOAN_OTHER_H, BC_LOANS_H
      Select Case bCToType
        Case BCT_HMIT_BOX_NUMBER
          BenClassTo = "15"
        Case BCT_HMIT_SECTION
          BenClassTo = HMIT_H
        Case BCT_HMIT_SECTION_STRING
          BenClassTo = HMITSectionToHMITCode(BenClassTo(bc, BCT_HMIT_SECTION))
        Case BCT_FORM_CAPTION, BCT_UDM_BENEFIT_TITLE
          Select Case bc
            Case BC_LOAN_OTHER_H
'              If p11d32.AppYear <= 2001 Then 'km
'                BenClassTo = "Loans other than home loans"
'              Else
                BenClassTo = "Beneficial loans"
'              End If
'            Case BC_LOAN_HOME_H
'              BenClassTo = "Home loans"
            Case BC_LOANS_H
              BenClassTo = "Loans"
          End Select
        Case BCT_BENITEMS_LAST_ITEM
          If bc <> BC_LOANS_H Then
            BenClassTo = LoanItems.ln_LASTITEM
          Else
            BenClassTo = 1
          End If
        Case BCT_TABLE
          Select Case bc
            Case BC_LOAN_OTHER_H
              BenClassTo = "T_BenLoan"
            Case Else
              ECASE ("Invalid ben class to type, benclass=" & bc & ",ben class to type=" & bCToType)
          End Select
        Case BCT_CLASS1A_ABLE
          BenClassTo = True
        Case Else
          ECASE ("Invalid ben class to type, benclass=" & bc & ",ben class to type=" & bCToType)
      End Select
    
    Case BC_PRIVATE_MEDICAL_I
      Select Case bCToType
        Case BCT_DBCLASS
          BenClassTo = S_PRIVATE_MEDICAL_I
        Case BCT_HMIT_BOX_NUMBER
          BenClassTo = "11"
        Case BCT_HMIT_SECTION
          BenClassTo = HMIT_I
        Case BCT_HMIT_SECTION_STRING
          BenClassTo = HMITSectionToHMITCode(BenClassTo(bc, BCT_HMIT_SECTION))
        Case BCT_FORM_CAPTION
          BenClassTo = "Private medical treatment or insurance"
        Case BCT_UDM_BENEFIT_TITLE
          BenClassTo = "Medical"
        Case BCT_DBCATEGORY
          BenClassTo = "MED"
        Case BCT_BENITEMS_LAST_ITEM
          BenClassTo = oth_LASTITEM
        Case BCT_TABLE
          BenClassTo = "T_BenOther"
        Case BCT_CLASS1A_ABLE
          BenClassTo = True
        Case Else
          ECASE ("Invalid ben class to type, benclass=" & bc & ",ben class to type=" & bCToType)
      End Select
    Case BC_QUALIFYING_RELOCATION_J
      Select Case bCToType
        Case BCT_HMIT_BOX_NUMBER
          BenClassTo = "15"
        Case BCT_HMIT_SECTION
          BenClassTo = HMIT_J
        Case BCT_HMIT_SECTION_STRING
          BenClassTo = HMITSectionToHMITCode(BenClassTo(bc, BCT_HMIT_SECTION))
        Case BCT_FORM_CAPTION
          BenClassTo = "Qualifying relocation expenses payments and benefits"
        Case BCT_UDM_BENEFIT_TITLE
          BenClassTo = "Qualifying relocation"
        Case BCT_BENITEMS_LAST_ITEM
          BenClassTo = reloc_LASTITEM
        Case BCT_TABLE
          BenClassTo = "T_BenReloc"
        Case BCT_CLASS1A_ABLE
          BenClassTo = True
        Case Else
          ECASE ("Invalid ben class to type, benclass=" & bc & ",ben class to type=" & bCToType)
      End Select
    Case BC_SERVICES_PROVIDED_K
      Select Case bCToType
        Case BCT_HMIT_BOX_NUMBER
          BenClassTo = "15"
        Case BCT_HMIT_SECTION
          BenClassTo = HMIT_K
        Case BCT_HMIT_SECTION_STRING
          BenClassTo = HMITSectionToHMITCode(BenClassTo(bc, BCT_HMIT_SECTION))
        Case BCT_FORM_CAPTION
          BenClassTo = "Services supplied"
        Case BCT_UDM_BENEFIT_TITLE
          BenClassTo = "Services"
        Case BCT_BENITEMS_LAST_ITEM
          BenClassTo = ServicesProvided_LASTITEM
        Case BCT_TABLE
          BenClassTo = "T_BenGoods"
        Case BCT_CLASS1A_ABLE
          BenClassTo = True
        Case Else
          ECASE ("Invalid ben class to type, benclass=" & bc & ",ben class to type=" & bCToType)
      End Select
    Case BC_ASSETSATDISPOSAL_L
      Select Case bCToType
        Case BCT_HMIT_BOX_NUMBER
          BenClassTo = "13"
        Case BCT_HMIT_SECTION
          BenClassTo = HMIT_L
        Case BCT_HMIT_SECTION_STRING
          BenClassTo = HMITSectionToHMITCode(BenClassTo(bc, BCT_HMIT_SECTION))
        Case BCT_FORM_CAPTION
          BenClassTo = "Assets placed at employees disposal"
        Case BCT_UDM_BENEFIT_TITLE
          BenClassTo = "Assets at disposal"
        Case BCT_BENITEMS_LAST_ITEM
          BenClassTo = AssetsAtDisposal_LASTITEM
        Case BCT_TABLE
          BenClassTo = "T_BenGoods"
        Case BCT_CLASS1A_ABLE
          BenClassTo = True
        Case Else
          ECASE ("Invalid ben class to type, benclass=" & bc & ",ben class to type=" & bCToType)
      End Select
    Case BC_CLASS_1A_M
      Select Case bCToType
        Case BCT_DBCLASS
          BenClassTo = S_Other_M_Class1a
        Case BCT_HMIT_BOX_NUMBER
          BenClassTo = "15"
        Case BCT_HMIT_SECTION
          BenClassTo = HMIT_M
        Case BCT_HMIT_SECTION_STRING
          BenClassTo = HMITSectionToHMITCode(BenClassTo(bc, BCT_HMIT_SECTION))
        Case BCT_FORM_CAPTION
'          If p11d32.AppYear > 2000 Then 'km
            BenClassTo = "Other items - Class 1A"
'          Else
'            BenClassTo = "Subscriptions and professional fees - employer contracts"
'          End If
        Case BCT_UDM_BENEFIT_TITLE
'          If p11d32.AppYear > 2000 Then 'km
            BenClassTo = "Other items - Class 1A"
'          Else
'            BenClassTo = "Subscriptions - employer contracts"
'          End If
        Case BCT_DBCATEGORY
          BenClassTo = "SUBS"
        Case BCT_BENITEMS_LAST_ITEM
          BenClassTo = oth_LASTITEM
        Case BCT_TABLE
          BenClassTo = "T_BenOther"
        Case BCT_CLASS1A_ABLE
          BenClassTo = True
        Case Else
          ECASE ("Invalid ben class to type, benclass=" & bc & ",ben class to type=" & bCToType)
      End Select
    Case BC_NON_CLASS_1A_M
      Select Case bCToType
        Case BCT_DBCLASS
          BenClassTo = S_Other_M_NonClass1a
        Case BCT_HMIT_BOX_NUMBER
          BenClassTo = "15"
        Case BCT_HMIT_SECTION
          BenClassTo = HMIT_M
        Case BCT_HMIT_SECTION_STRING
          BenClassTo = HMITSectionToHMITCode(BenClassTo(bc, BCT_HMIT_SECTION))
        Case BCT_FORM_CAPTION
'          If p11d32.AppYear > 2000 Then 'km
            BenClassTo = "Other items - Non Class 1A"
'          Else
'            BenClassTo = "Nursery and reimbursed Education/Subscriptions"
'          End If
        Case BCT_UDM_BENEFIT_TITLE
'          If p11d32.AppYear > 2000 Then 'km
            BenClassTo = "Other items - Non Class 1A"
'          Else
'            BenClassTo = "Nursery and reimbursed Education/Subscriptions"
'          End If
        Case BCT_DBCATEGORY
          BenClassTo = "MED"
        Case BCT_BENITEMS_LAST_ITEM
          BenClassTo = oth_LASTITEM
        Case BCT_TABLE
          BenClassTo = "T_BenOther"
        Case BCT_CLASS1A_ABLE
          BenClassTo = False
        Case Else
          ECASE ("Invalid ben class to type, benclass=" & bc & ",ben class to type=" & bCToType)
      End Select
    Case BC_INCOME_TAX_PAID_NOT_DEDUCTED_M
    
      Select Case bCToType
        Case BCT_DBCLASS
          BenClassTo = S_INCOME_TAX_PAID_NOT_DEDUCTED_N
        Case BCT_HMIT_BOX_NUMBER
          BenClassTo = "15"
        Case BCT_HMIT_SECTION
          BenClassTo = HMIT_M
        Case BCT_HMIT_SECTION_STRING
          BenClassTo = HMITSectionToHMITCode(BenClassTo(bc, BCT_HMIT_SECTION))
        Case BCT_FORM_CAPTION
          BenClassTo = "Income tax paid but not deducted"
        Case BCT_UDM_BENEFIT_TITLE
          BenClassTo = "Income tax paid"
        Case BCT_DBCATEGORY
          BenClassTo = "MISC"
        Case BCT_BENITEMS_LAST_ITEM
          BenClassTo = oth_LASTITEM
        Case BCT_TABLE
          BenClassTo = "T_BenOther"
        Case BCT_CLASS1A_ABLE
          BenClassTo = False
        Case Else
          ECASE ("Invalid ben class to type, benclass=" & bc & ",ben class to type=" & bCToType)
      End Select
    Case BC_TRAVEL_AND_SUBSISTENCE_N
      Select Case bCToType
        Case BCT_DBCLASS
          BenClassTo = S_TRAVEL_N
        Case BCT_HMIT_BOX_NUMBER
          BenClassTo = "16"
        Case BCT_HMIT_SECTION
          BenClassTo = HMIT_N
        Case BCT_HMIT_SECTION_STRING
          BenClassTo = HMITSectionToHMITCode(BenClassTo(bc, BCT_HMIT_SECTION))
        Case BCT_FORM_CAPTION
          BenClassTo = "Travel and subsistence payments"
        Case BCT_UDM_BENEFIT_TITLE
          BenClassTo = "Travel and subsistence"
        Case BCT_DBCATEGORY
          BenClassTo = "GEN"
        Case BCT_BENITEMS_LAST_ITEM
          BenClassTo = oth_LASTITEM
        Case BCT_TABLE
          BenClassTo = "T_BenOther"
        Case BCT_CLASS1A_ABLE
          BenClassTo = False
        Case Else
          ECASE ("Invalid ben class to type, benclass=" & bc & ",ben class to type=" & bCToType)
      End Select
    Case BC_ENTERTAINMENT_N
      Select Case bCToType
        Case BCT_DBCLASS
          BenClassTo = S_ENTERTAINMENT_N
        Case BCT_HMIT_BOX_NUMBER
          BenClassTo = "16"
        Case BCT_HMIT_SECTION
          BenClassTo = HMIT_N
        Case BCT_HMIT_SECTION_STRING
          BenClassTo = HMITSectionToHMITCode(BenClassTo(bc, BCT_HMIT_SECTION))
        Case BCT_FORM_CAPTION, BCT_UDM_BENEFIT_TITLE
          BenClassTo = "Entertainment"
        Case BCT_DBCATEGORY
          BenClassTo = "GEN"
        Case BCT_BENITEMS_LAST_ITEM
          BenClassTo = oth_LASTITEM
        Case BCT_TABLE
          BenClassTo = "T_BenOther"
        Case BCT_CLASS1A_ABLE
          BenClassTo = False
        Case Else
          ECASE ("Invalid ben class to type, benclass=" & bc & ",ben class to type=" & bCToType)
      End Select
    Case BC_GENERAL_EXPENSES_BUSINESS_N
      Select Case bCToType
        Case BCT_DBCLASS
          BenClassTo = S_GENERAL_EXPENSES_N
        Case BCT_HMIT_BOX_NUMBER
          BenClassTo = "16"
        Case BCT_HMIT_SECTION
          BenClassTo = HMIT_N
        Case BCT_HMIT_SECTION_STRING
          BenClassTo = HMITSectionToHMITCode(BenClassTo(bc, BCT_HMIT_SECTION))
        Case BCT_FORM_CAPTION
          BenClassTo = "General expenses allowed for business travel"
        Case BCT_UDM_BENEFIT_TITLE
          BenClassTo = "General expenses"
        Case BCT_DBCATEGORY
          BenClassTo = "GEN"
        Case BCT_BENITEMS_LAST_ITEM
          BenClassTo = oth_LASTITEM
        Case BCT_TABLE
          BenClassTo = "T_BenOther"
        Case BCT_CLASS1A_ABLE
          BenClassTo = False
        Case Else
          ECASE ("Invalid ben class to type, benclass=" & bc & ",ben class to type=" & bCToType)
      End Select
    Case BC_CHAUFFEUR_OTHERO_N
      Select Case bCToType
        Case BCT_DBCLASS
          BenClassTo = S_CHAUFFEUR_N_OTHER
        Case BCT_HMIT_BOX_NUMBER
          BenClassTo = "16"
        Case BCT_HMIT_SECTION
          BenClassTo = HMIT_N
        Case BCT_HMIT_SECTION_STRING
          BenClassTo = HMITSectionToHMITCode(BenClassTo(bc, BCT_HMIT_SECTION))
        Case BCT_FORM_CAPTION
          BenClassTo = "Chauffeur expenses - N Other"
        Case BCT_UDM_BENEFIT_TITLE
          BenClassTo = "Chauffeur"
        Case BCT_DBCATEGORY
          BenClassTo = "MISC"
        Case BCT_BENITEMS_LAST_ITEM
          BenClassTo = oth_LASTITEM
        Case BCT_TABLE
          BenClassTo = "T_BenOther"
        Case BCT_CLASS1A_ABLE
          BenClassTo = False
        Case Else
          ECASE ("Invalid ben class to type, benclass=" & bc & ",ben class to type=" & bCToType)
      End Select
    Case BC_PHONE_HOME_N
      Select Case bCToType
        Case BCT_DBCLASS
          BenClassTo = S_HOME
        Case BCT_HMIT_BOX_NUMBER
          BenClassTo = "16"
        Case BCT_HMIT_SECTION
          BenClassTo = HMIT_N
        Case BCT_HMIT_SECTION_STRING
          BenClassTo = HMITSectionToHMITCode(BenClassTo(bc, BCT_HMIT_SECTION))
        Case BCT_FORM_CAPTION
          BenClassTo = "Payments for use of home phone"
        Case BCT_UDM_BENEFIT_TITLE
          BenClassTo = "Home phone"
        Case BCT_BENITEMS_LAST_ITEM
          BenClassTo = Pho_LASTITEM
        Case BCT_TABLE
          BenClassTo = "T_BenPhone"
        Case BCT_CLASS1A_ABLE
          BenClassTo = False
        Case Else
          ECASE ("Invalid ben class to type, benclass=" & bc & ",ben class to type=" & bCToType)
      End Select
    Case BC_NON_QUALIFYING_RELOCATION_N
      Select Case bCToType
        Case BCT_HMIT_BOX_NUMBER
          BenClassTo = "16"
        Case BCT_HMIT_SECTION
          BenClassTo = HMIT_N
        Case BCT_HMIT_SECTION_STRING
          BenClassTo = HMITSectionToHMITCode(BenClassTo(bc, BCT_HMIT_SECTION))
        Case BCT_FORM_CAPTION
          BenClassTo = "Non-qualifying relocation expenses"
        Case BCT_UDM_BENEFIT_TITLE
          BenClassTo = "Non qual relocation"
        Case BCT_BENITEMS_LAST_ITEM
          BenClassTo = [_ITEM_LAST_ITEM]
        Case BCT_TABLE
          BenClassTo = "T_BenReloc"
        Case BCT_CLASS1A_ABLE
          BenClassTo = False
        Case Else
          ECASE ("Invalid ben class to type, benclass=" & bc & ",ben class to type=" & bCToType)
      End Select
    Case BC_OOTHER_N
      Select Case bCToType
        Case BCT_DBCLASS
          BenClassTo = S_N
        Case BCT_HMIT_BOX_NUMBER
          BenClassTo = "16"
        Case BCT_HMIT_SECTION
          BenClassTo = HMIT_N
        Case BCT_HMIT_SECTION_STRING
          BenClassTo = HMITSectionToHMITCode(BenClassTo(bc, BCT_HMIT_SECTION))
        Case BCT_FORM_CAPTION, BCT_UDM_BENEFIT_TITLE
          BenClassTo = "Other expenses"
        Case BCT_DBCATEGORY
          BenClassTo = "GEN"
        Case BCT_BENITEMS_LAST_ITEM
          BenClassTo = oth_LASTITEM
        Case BCT_TABLE
          BenClassTo = "T_BenOther"
        Case BCT_CLASS1A_ABLE
          BenClassTo = False
        Case Else
          ECASE ("Invalid ben class to type, benclass=" & bc & ",ben class to type=" & bCToType)
      End Select
    Case BC_EMPLOYEE
      Select Case bCToType
        Case BCT_FORM_CAPTION, BCT_UDM_BENEFIT_TITLE
          BenClassTo = "Employees"
        Case BCT_BENITEMS_LAST_ITEM
          BenClassTo = ee_LASTITEM
        Case BCT_TABLE
          BenClassTo = "T_Employees"
        Case Else
          ECASE ("Invalid ben class to type, benclass=" & bc & ",ben class to type=" & bCToType)
      End Select
    Case BC_ALL
      Select Case bCToType
        Case BCT_FORM_CAPTION
          BenClassTo = "All benefits"
        Case Else
          ECASE ("Invalid ben class to type, benclass=" & bc & ",ben class to type=" & bCToType)
      End Select
    Case BC_CDB
      Select Case bCToType
        Case BCT_FORM_CAPTION
          BenClassTo = "Company defined benefits"
        Case BCT_BENITEMS_LAST_ITEM
          BenClassTo = BenClassTo(BC_OOTHER_N, bCToType)
        Case Else
          ECASE ("Invalid ben class to type, benclass=" & bc & ",ben class to type=" & bCToType)
      End Select
    Case BC_EMPLOYER
      Select Case bCToType
        Case BCT_FORM_CAPTION, BCT_UDM_BENEFIT_TITLE
          BenClassTo = "Employers"
        Case BCT_BENITEMS_LAST_ITEM
          BenClassTo = employer_LAST_ITEM
        Case Else
          ECASE ("Invalid ben class to type, benclass=" & bc & ",ben class to type=" & bCToType)
      End Select
    Case Else
      ECASE ("Invalid benclass in BenClassToHMITSection")
  End Select
  
BenClassTo_End:
  Call xReturn("BenClassTo")
  Exit Function
BenClassTo_Err:
  Call ErrorMessage(ERR_ERROR, Err, "BenClassTo", "Ben Class To ", "Error mapping a benefit class.")
  Resume BenClassTo_End
End Function

Public Function HMITSectionToHMITCode(ByVal HMITS As HMIT_SECTIONS) As String
  HMITSectionToHMITCode = Chr$(Asc("A") - 1 + HMITS)
End Function
Public Function HMITSectionToHMITDescription(ByVal HMITS As HMIT_SECTIONS)
  Select Case HMITS
    Case HMIT_A
      HMITSectionToHMITDescription = BenClassTo(BC_ASSETSTRANSFERRED_A, BCT_FORM_CAPTION)
    Case HMIT_B
      HMITSectionToHMITDescription = BenClassTo(BC_PAYMENTS_ON_BEFALF_B, BCT_FORM_CAPTION)
    Case HMIT_C
      HMITSectionToHMITDescription = BenClassTo(BC_VOUCHERS_AND_CREDITCARDS_C, BCT_FORM_CAPTION)
    Case HMIT_D
      HMITSectionToHMITDescription = BenClassTo(BC_LIVING_ACCOMMODATION_D, BCT_FORM_CAPTION)
    Case HMIT_E
      HMITSectionToHMITDescription = BenClassTo(BC_EMPLOYEE_CAR_E, BCT_FORM_CAPTION)
    Case HMIT_F
      HMITSectionToHMITDescription = BenClassTo(BC_COMPANY_CARS_F, BCT_FORM_CAPTION)
    Case HMIT_G
      HMITSectionToHMITDescription = BenClassTo(BC_NONSHAREDVANS_G, BCT_FORM_CAPTION)
    Case HMIT_H
      HMITSectionToHMITDescription = BenClassTo(BC_LOANS_H, BCT_FORM_CAPTION)
    Case HMIT_I
      HMITSectionToHMITDescription = BenClassTo(BC_PRIVATE_MEDICAL_I, BCT_FORM_CAPTION)
    Case HMIT_J
      HMITSectionToHMITDescription = BenClassTo(BC_QUALIFYING_RELOCATION_J, BCT_FORM_CAPTION)
    Case HMIT_K
      HMITSectionToHMITDescription = BenClassTo(BC_SERVICES_PROVIDED_K, BCT_FORM_CAPTION)
    Case HMIT_L
      HMITSectionToHMITDescription = BenClassTo(BC_ASSETSATDISPOSAL_L, BCT_FORM_CAPTION)
'    Case HMIT_M
'      HMITSectionToHMITDescription = BenClassTo(BC_SHARES_M, BCT_FORM_CAPTION)
    Case HMIT_M
      HMITSectionToHMITDescription = "Other items" 'BenClassTo(BC_NON_CLASS_1A_M, BCT_FORM_CAPTION)
    Case HMIT_N
      HMITSectionToHMITDescription = BenClassTo(BC_OOTHER_N, BCT_FORM_CAPTION)
    Case Else
      ECASE ("Invalid HMIT section in HMITSEctionToHMITDescription")
  End Select
  HMITSectionToHMITDescription = HMITSectionToHMITCode(HMITS) & " - " & HMITSectionToHMITDescription
  
End Function
Public Function DBClassToBCClass(sClass As String) As BEN_CLASS
   Select Case UCASE$(sClass)
    Case UCASE$(S_N), UCASE$("O Other Expense") 'fix to ma when M shares was removed
      DBClassToBCClass = BC_OOTHER_N
    Case UCASE$(S_VOUCHERS_AND_CREDITCARDS_C)
      DBClassToBCClass = BC_VOUCHERS_AND_CREDITCARDS_C
    Case UCASE$(S_SUBSCRIPTION_N)
      DBClassToBCClass = BC_CLASS_1A_M
    Case UCASE$(S_PRIVATE_MEDICAL_I)
      DBClassToBCClass = BC_PRIVATE_MEDICAL_I
    Case UCASE$(S_NURSERY_N_OTHER)
      DBClassToBCClass = BC_NON_CLASS_1A_M
    Case UCASE$(S_PAYMENTS_ON_BEFALF_B)
      DBClassToBCClass = BC_PAYMENTS_ON_BEFALF_B
    Case UCASE$(S_TAX_NOTIONAL_PAYMENTS_B)
      DBClassToBCClass = BC_TAX_NOTIONAL_PAYMENTS_B
    'Case UCASE$(S_SHARES_M)
    '  DBClassToBCClass = BC_SHARES_M
    Case UCASE$(S_INCOME_TAX_PAID_NOT_DEDUCTED_N)
      DBClassToBCClass = BC_INCOME_TAX_PAID_NOT_DEDUCTED_M
    Case UCASE$(S_CHAUFFEUR_N_OTHER)
      DBClassToBCClass = BC_CHAUFFEUR_OTHERO_N
    Case UCASE$(S_ENTERTAINMENT_N)
      DBClassToBCClass = BC_ENTERTAINMENT_N
    Case UCASE$(S_TRAVEL_N)
      DBClassToBCClass = BC_TRAVEL_AND_SUBSISTENCE_N
    Case UCASE$(S_GENERAL_EXPENSES_N)
      DBClassToBCClass = BC_GENERAL_EXPENSES_BUSINESS_N
    Case UCASE$(S_HOME)
      DBClassToBCClass = BC_PHONE_HOME_N
 
    Case UCASE$(S_Other_M_Class1a)
      DBClassToBCClass = BC_CLASS_1A_M
    Case UCASE$(S_Other_M_NonClass1a)
      DBClassToBCClass = BC_NON_CLASS_1A_M
    Case Else
      ECASE ("Invalid p11d class string in db")
  End Select
End Function
Public Function BenClassesFromHMITSection(BenClasses() As BEN_CLASS, HS As HMIT_SECTIONS, benClassesRestrictions() As BEN_CLASS) As Long
  Dim bc As BEN_CLASS
  Dim i As Long
  
  On Error GoTo BenClassesFromHMITSection_ERR
  
  Call xSet("BenClassesFromHMITSection")
  
  For bc = BC_FIRST_ITEM To BC_UDM_BENEFITS_LAST_ITEM
    
      If BenClassTo(bc, BCT_HMIT_SECTION) = HS Then
        For i = 1 To UBound(benClassesRestrictions)
          If bc = benClassesRestrictions(i) Then GoTo NEXT_ITEM
        Next
        BenClassesFromHMITSection = BenClassesFromHMITSection + 1
        ReDim Preserve BenClasses(1 To BenClassesFromHMITSection)
        BenClasses(BenClassesFromHMITSection) = bc
      End If
NEXT_ITEM:
  Next
  
BenClassesFromHMITSection_END:
  Call xReturn("BenClassesFromHMITSection")
  Exit Function
BenClassesFromHMITSection_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "BenClassesFromHMITSection", "Get Ben Classes From HMIT Section", "Error getting the benclasses that relate to HMIT section " & HMITSectionToHMITCode(HS) & ".")
  Resume BenClassesFromHMITSection_END
End Function

Private Function ISortFunction_CompareItems(v0 As Variant, v1 As Variant) As Long
  'Compare 2 sort items, return -1 if v0 < v1, 1 if v0 > v1, 0 if the same
  If v0 < v1 Then
    ISortFunction_CompareItems = -1
  ElseIf v1 > v0 Then
    ISortFunction_CompareItems = 1
  Else
    ISortFunction_CompareItems = 0
  End If
End Function


