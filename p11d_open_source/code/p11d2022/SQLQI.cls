VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SQLQUERIES_IMP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum QUERY_NAMES_IMP
  IMP_DELETE_ERRORS = 1
  IMP_SELECT_ERRORS
  IMP_EMPLOYEES
  
  IMP_HOMEPHONES
  IMP_ENTERTAINMENT
  IMP_CHAUFFEUR
  IMP_CREDITCARDS_AND_VOUCHERS
  IMP_EDUCATION
  IMP_GENERAL
  IMP_MEDICAL
  IMP_NOTIONAL_PAYMENTS
  IMP_N_NON_CLASS_1A 'Non-class 1A
  IMP_N_CLASS_1A 'Class 1A
  IMP_TAXPAID
  IMP_TRAVEL_AND_SUBSISTENCE
  IMP_ACCOMMODATION
  IMP_COMPANYCARS
  IMP_COMPANY_CAR_MILES
  IMP_EMPLOYEE_CARS
  IMP_EMPLOYEE_CAR_MILEAGE
  IMP_HOME_LOANS
'MP DB  IMP_HOME_LOAN_BALANCES
  IMP_BENEFICIAL_LOANS
  IMP_BENEFICIAL_LOAN_BALANCES
  IMP_ADDRESSES
  IMP_ASSETS_TRANSEFERRED
  IMP_ASSETS_AT_DISPOSAL
  IMP_SERVICES_PROVIDED
  IMP_SHARES
  IMP_O_OTHER_EXPENSES
  IMP_RELOCATION_ITEM
  IMP_RELOCATION_DETAIL
  IMP_NON_SHARED_VANS
  IMP_SHARED_VANS
  IMP_PAYMENTS_ON_BEHALF_OF_EMPLOYEE
  
  POST_PROCESS_NON_SHARED_VANS_SELECT
  POST_PROCESS_SHARED_VANS_SELECT
  
  POST_PROCESS_RELOCATION_DETAIL_SELECT
  
  
'MP DB  POST_PROCESS_HOMELOAN_BALANCES_SELECT

  
  POST_PROCESS_BENEFICIAL_LOAN_BALANCES_SELECT

  
'MP DB  POST_PROCESS_COMPANYCAR_MILEAGE_SELECT

  POST_PROCESS_EMPLOYEECAR_MILEAGE_SELECT
  
  POST_PROCESS_OOTHER_O
  
  UPDATE_EMPLOYEES
'MP DB  UPDATE_HOME_LOAN_BALANCES_KEYS
  UPDATE_BENEFICIAL_LOAN_BALANCES_KEYS
  'UPDATE_COMPANY_CAR_KEYS
  UPDATE_EMPLOYEE_CAR_KEYS
  UPDATE_RELOCATION_KEYS
  UPDATE_COMPANY_CAR_FUEL_AVAILABLE_TO

End Enum
Private mLastQuery As QUERY_NAMES_IMP
Private Const S_GEN_IMP_SQL_BASE As String = "SELECT DISTINCTROW T_BenOther.P_Num AS [" & S_PNUM & "], T_BenOther.Item AS Description,  T_BenOther.Value AS [Value of benefit],T_BenOther.MadeGood AS [Amount made good], T_BenOther.MadeGoodIsTaxDeducted AS [Amount is tax deducted],T_BenOther.From AS [Available from], T_BenOther.To AS [Available to], T_BenOther.UDBCode, T_BenOther.Category, T_BenOther.Class"
Private Const S_GEN_IMP_SQL_NO_OPRA As String = S_GEN_IMP_SQL_BASE & " From T_BenOther"
Private S_GEN_IMP_SQL As String

Public Function Queries(ByVal QueryID As QUERY_NAMES_IMP, ParamArray P()) As String
  Dim v() As Variant, Paramcount As Long, i As Long
  
  Paramcount = UBound(P) - LBound(P) + 1
  If Paramcount > 0 Then
    ReDim v(1 To Paramcount)
    For i = 1 To Paramcount
      v(i) = P(i - 1 + LBound(P))
    Next i
  End If
  Queries = QueriesEx(QueryID, v, Paramcount, False, False)
End Function

Public Function Name(ByVal QueryID As QUERY_NAMES_IMP, Optional ByVal IgnoreEmpty As Boolean = False) As String
  Dim v() As Variant
  Name = QueriesEx(QueryID, v, 0, True, IgnoreEmpty)
End Function

Public Property Get LastQuery() As QUERY_NAMES_IMP
  LastQuery = mLastQuery
End Property

Private Function QueriesEx(ByVal QueryID As QUERY_NAMES_IMP, params() As Variant, ByVal Paramcount As Long, ByVal NameOnly As Boolean, ByVal IgnoreEmpty As Boolean) As String
  
  On Error GoTo QueriesEx_ERR
  
  Call xSet("Queries")
  
  mLastQuery = QueryID
  Select Case QueryID
    Case IMP_SELECT_ERRORS
        QueriesEx = "SELECT ImportErrors.Row, ImportErrors.Column, ImportErrors.Parent, ImportErrors.ErrorItem, ImportErrors.ErrorDescription, ImportErrors.Import, ImportErrors.DateTime " & _
                    "FROM ImportErrors;"
    
    Case IMP_DELETE_ERRORS
      QueriesEx = "DELETE * FROM ImportErrors"
    
    Case IMP_EMPLOYEES
      If NameOnly Then
        QueriesEx = "Employees"
      Else 'FC - Class1A
        QueriesEx = "SELECT DISTINCTROW T_Employees.P_Num AS " & S_IMP_PNUM_ALIAS & ", T_Employees.Surname,T_Employees.FirstName AS [First name], T_Employees.Initials, T_Employees.Title, T_Employees.NI AS [NI number],  T_Employees.Salutation, T_Employees.Email AS [Email address], T_Employees.Status AS [Status/Position], T_Employees.Joined AS [Date joined], T_Employees.Left AS [Date left],  T_Employees.Class1AEmployeeIsNotSubjectTo AS [Not Subject to Class1a], T_Employees.Group1, T_Employees.Group2, T_Employees.Group3, T_Employees.EEUsername AS [Intranet Username], T_Employees.EEPassword AS [Intranet Password], T_Employees.Comments, T_Employees.DOB, T_Employees.Gender  FROM T_Employees WHERE T_Employees.P_Num<>" & StrSQL(S_CDB_EMPLOYEE_NUMBER_PREFIX)
      End If
    
    Case IMP_HOMEPHONES
      If NameOnly Then
        QueriesEx = "Home phones"
      Else
        QueriesEx = "SELECT DISTINCTROW T_BenPhone.P_Num AS " & S_IMP_PNUM_ALIAS & ", T_BenPhone.Item AS Description,T_BenPhone.ValueCalls AS [Calls cost], T_BenPhone.ContribCalls AS [Contribution to calls], T_BenPhone.ValueRent AS [Rental cost],T_BenPhone.ContribRent AS [Contribution to rent],T_BenPhone.MadeGoodIsTaxDeducted as [Amount is tax deducted], T_BenPhone.Class " & OPRAAmountForegoneField("T_BenPhone") & " From T_BenPhone WHERE T_BenPhone.Class=" & StrSQL(p11d32.Rates.BenClassTo(BC_PHONE_HOME_N, BCT_DBCLASS))
      End If
    
    Case IMP_ENTERTAINMENT
      If NameOnly Then
        QueriesEx = "Entertainment"
      Else
        QueriesEx = S_GEN_IMP_SQL & BenOtherWhereClause(BC_ENTERTAINMENT_N)
      End If
    
    Case IMP_CHAUFFEUR
      If NameOnly Then
        QueriesEx = "Chauffeur"
      Else
        'CAD MAKE INTO STANDARD FUNCTION
        QueriesEx = "SELECT DISTINCTROW T_BenOther.P_Num AS " & S_IMP_PNUM_ALIAS & ", T_BenOther.IRDesc AS [" & S_IMP_FIELD_CATEGORY_DESCRIPTION & "], T_BenOther.Item AS Description,  T_BenOther.Value AS [Value of benefit],T_BenOther.MadeGood AS [Amount made good],T_BenOther.MadeGoodIsTaxDeducted as [Amount is tax deducted], T_BenOther.From AS [Available from], T_BenOther.To AS [Available to], T_BenOther.UDBCode, T_BenOther.Category, T_BenOther.Class " & OPRAAmountForegoneField("T_BenOther") & " From T_BenOther" & BenOtherWhereClause(BC_CHAUFFEUR_OTHERO_N)
      End If
    
    Case IMP_CREDITCARDS_AND_VOUCHERS
      If NameOnly Then
        QueriesEx = "Credit cards and vouchers"
      Else
        QueriesEx = S_GEN_IMP_SQL & BenOtherWhereClause(BC_VOUCHERS_AND_CREDITCARDS_C)
      End If
    
    Case IMP_GENERAL
      If NameOnly Then
        QueriesEx = "General"
      Else
        QueriesEx = S_GEN_IMP_SQL & BenOtherWhereClause(BC_GENERAL_EXPENSES_BUSINESS_N)
      End If
    Case IMP_MEDICAL
      If NameOnly Then
        QueriesEx = "Medical"
      Else
        QueriesEx = S_GEN_IMP_SQL & BenOtherWhereClause(BC_PRIVATE_MEDICAL_I)
      End If
    Case IMP_NOTIONAL_PAYMENTS
      If NameOnly Then
        QueriesEx = "Notional payments"
      Else
        QueriesEx = S_GEN_IMP_SQL_NO_OPRA & BenOtherWhereClause(BC_TAX_NOTIONAL_PAYMENTS_B)
      End If
    Case IMP_N_NON_CLASS_1A
      If NameOnly Then
          QueriesEx = "Other items - Non Class 1A"
      Else
        QueriesEx = "SELECT DISTINCTROW T_BenOther.P_Num AS [" & S_PNUM & "], T_BenOther.Item AS Description,  T_BenOther.IRDesc AS [" & S_IMP_FIELD_CATEGORY_DESCRIPTION & "], T_BenOther.Value AS [Value of benefit],T_BenOther.MadeGood AS [Amount made good], T_BenOther.MadeGoodIsTaxDeducted AS [Amount is tax deducted],T_BenOther.From AS [Available from], T_BenOther.To AS [Available to], T_BenOther.UDBCode, T_BenOther.Category, T_BenOther.Class " & OPRAAmountForegoneField("T_BenOther") & " From T_BenOther" & BenOtherWhereClause(BC_NON_CLASS_1A_M)
        ' QueriesEx = S_GEN_IMP_SQL & BenOtherWhereClause(BC_NON_CLASS_1A_M)
      End If
    Case IMP_N_CLASS_1A
      If NameOnly Then
          QueriesEx = "Other items - Class 1A"
      Else
        QueriesEx = "SELECT DISTINCTROW T_BenOther.P_Num AS [" & S_PNUM & "], T_BenOther.Item AS [Description],  T_BenOther.IRDesc AS [" & S_IMP_FIELD_CATEGORY_DESCRIPTION & "], T_BenOther.Value AS [Value of benefit],T_BenOther.MadeGood AS [Amount made good], T_BenOther.MadeGoodIsTaxDeducted AS [Amount is tax deducted],T_BenOther.From AS [Available from], T_BenOther.To AS [Available to], T_BenOther.UDBCode, T_BenOther.Category, T_BenOther.Class " & OPRAAmountForegoneField("T_BenOther") & " From T_BenOther" & BenOtherWhereClause(BC_CLASS_1A_M)
      End If
    Case IMP_TAXPAID
      If NameOnly Then
        QueriesEx = "Tax paid"
      Else
        QueriesEx = S_GEN_IMP_SQL & BenOtherWhereClause(BC_INCOME_TAX_PAID_NOT_DEDUCTED_M)
      End If
    Case IMP_TRAVEL_AND_SUBSISTENCE
      If NameOnly Then
        QueriesEx = "Travel and subsistence"
      Else
        QueriesEx = S_GEN_IMP_SQL & BenOtherWhereClause(BC_TRAVEL_AND_SUBSISTENCE_N)
      End If
    Case IMP_ACCOMMODATION
      If NameOnly Then
        QueriesEx = "Accommodation"
      Else
        QueriesEx = "SELECT DISTINCTROW T_BenAccom.P_Num AS " & S_IMP_PNUM_ALIAS & ", T_BenAccom.Item AS Address, T_BenAccom.Rent AS [Rent paid by employer],T_BenAccom.IsRent AS [Is amount paid by employee rent], T_BenAccom.Consideration AS [Made good],T_BenAccom.MadeGoodIsTaxDeducted as [Amount is tax deducted],T_BenAccom.AvailFrom AS [Available from], T_BenAccom.AvailTo AS [Available to], T_BenAccom.Business AS [Percentage of business use],T_BenAccom.JobRelated AS [Is accommodation job related],T_BenAccom.nemployees as [Number of other employees sharing], " & _
                    "T_BenAccom.ErOwn AS [Property owned by employer],T_BenAccom.Price AS [Purchase cost of employer owned accommodation], T_BenAccom.Improv AS [Cost of improvements to employer owned accommodation], T_BenAccom.CapContrib AS [Capital contribution by employee for employer owned],T_BenAccom.Tenancy AS [Payments for grant of tenancy for employer owned accommodation], T_BenAccom.FirstOcc AS [Date first occupied],  " & _
                    "T_BenAccom.SixYear AS [Does accommodation meet 6 year rule],T_BenAccom.MVFirstOcc AS [Market value first date occupied],T_BenAccom.RecentImprov AS [Improvements in tax year] " & _
                    OPRAAmountForegoneField("T_BenAccom") & " From T_BenAccom"
      End If
    Case IMP_COMPANYCARS
      If NameOnly Then
        QueriesEx = "Company cars"
      Else
          QueriesEx = "SELECT T_BenCar.P_Num AS " & S_IMP_PNUM_ALIAS & _
                      ", T_BenCar.Reg AS [Registration number], T_BenCar.Price AS [List price]" & _
                      ", T_BenCar.AvailFrom AS [Available from], T_BenCar.AvailTo AS [Available to]" & _
                      ", T_BenCar.Unavail AS [Days unavailable], T_BenCar.CapContrib AS [Capital contribution]" & _
                      ", T_BenCar.UseContrib AS [Contribution to private use], T_BenCar.MadeGoodIsTaxDeducted AS [Amount is tax deducted]" & _
                      ", T_BenCar.P46PaymentFrequency AS [Payment Frequency] " & _
                      ", T_BenCar.Make, T_BenCar.Model" & _
                      ", T_BenCar.RegDate AS [Registration date]" & _
                      ", T_BenCar.Acc AS [Accessories when first registered], T_BenCar.NewAcc AS [Accessories fitted after registration], T_BenCar.CheapAcc AS [Accessories - less than £100 included in accessories]" & _
                      ", T_BenCar.P46CarbonDioxide as [Grams of Carbon dioxide]" & _
                      ", T_BenCar.P46NoApprovedCO2Figure as [No CO2 Figure available] " & _
                      ", T_BenCar.cc AS [Engine size of car in cc]" & _
                      ", T_BenCar.NumOfUsers As [Number of users]" & _
                      ", T_BenCar.SecondCar AS [Is this a second car?]" & _
                      ", T_BenCar.Replaced AS [Is this car replaced?],T_BenCar.Replacement AS [Is this a replacement car?]" & _
                      ", T_BenCar.RegReplaced AS [Registration number of the car replaced]" & _
                      ", T_BenCar.CarReplaced AS [Make and model of car replaced], T_BenCar.DateCarReplaced AS [Date car replaced],T_BenCar.ForceP46 AS [Force printing of the P46 form?]" & _
                      ", T_BenCar.PvtFuel AS [Is fuel provided for private use?]" & _
                      ", T_BenCar.makeGood AS [Is employee required to make good private fuel costs?],T_BenCar.madegood AS [Was private fuel cost made good?] " & _
                      ", T_BenCar.P46FuelType as [Type of fuel]" & _
                      ", T_BenCar.FuelAvailTo AS [" & S_FIELD_CAR_FUEL_WITHDRAWN_DATE & "]" & _
                      ", T_BenCar.FuelReinstated AS [Fuel reinstated?]" & _
                      ", T_BenCar.HasFuelUnavail AS [" & S_COMPANY_CAR_DAYS_UNAVAILABLE_FUEL_DESCRIPTION & "]" & _
                      OPRAAmountForegoneField("T_BenCar") & OPRAAmountForegoneField("T_BenCar", "OPRAAmountForegoneFuel", " (fuel)")
                      
            QueriesEx = QueriesEx & _
                      ", T_BenCar." & S_FIELD_CAR_ELECTRIC_RANGE_MILES & " AS [" & S_COMPANY_CAR_ELECTFRIC_RANGE_MILES_DESCRIPTION & "]" & _
                      " From T_BenCar"
                      
      End If
    
    Case IMP_EMPLOYEE_CAR_MILEAGE
      If NameOnly Then
        QueriesEx = "Employee owned car mileage"
      Else
        'AM Remove import of (car key)
        'QueriesEx = "SELECT T_BenEECarDetails.P_Num AS " & S_IMP_PNUM_ALIAS & ", T_BenEECarDetails.Reg AS [Car], T_BenEECarDetails.Miles AS [Number of miles], T_BenEECarDetails.MilesItem AS [Description of mileage], T_BenEECarDetails.MilesDate AS [Date of car usage], T_BenEECarDetails.EECarKey AS [(Car key)] FROM T_BenEECarDetails;"
        QueriesEx = "SELECT T_BenEECarDetails.P_Num AS " & S_IMP_PNUM_ALIAS & ", T_BenEECarDetails.Reg AS [Car], T_BenEECarDetails.Miles AS [Number of miles], T_BenEECarDetails.MilesItem AS [Description of mileage], T_BenEECarDetails.MilesDate AS [Date of car usage] FROM T_BenEECarDetails;"
      End If
    Case IMP_EMPLOYEE_CARS
      If NameOnly Then
        QueriesEx = "Employee owned cars"
      Else
        QueriesEx = "SELECT T_BenEeCar.P_Num AS " & S_IMP_PNUM_ALIAS & ", T_BenEeCar.Item AS [Car], T_BenEeCar.EECarAllowance as [Amount received],T_BenEeCar.enginesize AS [Engine size in cc], T_BenEeCar.FPCS AS [Name of MARORS],T_BenEECar.UseFPCS as [Use Company mileage scheme?],T_BenEeCar.LUMPSUM AS [Lump sum payment], T_BenEeCar.MadeGood AS [Amount made good], T_BenEeCar.MadeGoodIsTaxDeducted as [Amount is tax deducted], T_BenEeCar.HireCost AS [Hire cost payment],T_BenEeCar.HCMadeGood AS [Hire cost made good] FROM T_BenEeCar;"
      End If
    
    Case IMP_BENEFICIAL_LOAN_BALANCES
      If NameOnly Then
        QueriesEx = "Beneficial loan balances"
      Else
        QueriesEx = "SELECT DISTINCTROW T_Loans.P_Num AS " & S_IMP_PNUM_ALIAS & ", T_Loans.Item AS [Loan description], T_Loans.From AS [Date from], T_Loans.Payment AS [Amount received/(repaid)], T_Loans.Type AS [Loan type] From T_Loans WHERE ((T_Loans.Type='" & S_BENEFICIAL_LOANS_TYPE & "'));"
      End If
      
    Case IMP_BENEFICIAL_LOANS
      If NameOnly Then
        QueriesEx = "Beneficial loans"
      Else
        'AM Remove "loan type" and "part of relocation"
        QueriesEx = "SELECT DISTINCTROW T_BenLoan.P_Num AS " & S_IMP_PNUM_ALIAS & ", T_BenLoan.Item AS [Loan description], T_BenLoan.Interest AS [Interest paid by employee],T_BenLoan.MadeGoodIsTaxDeducted as [Amount is tax deducted], T_BenLoan.Waived AS [Amount waived],T_BenLoan.Curr AS [Currency],  T_BenLoan.DailyOnly AS [Use daily calculation only]" & IIf(p11d32.AppYear > 1998, ", T_Benloan.DidLoanCommenceOnFirstDayOfTaxYear AS [" & S_DID_LOAN_COMMENCE_ON_FIRST_DAY_OF_TAX_YEAR_IMPORT_HEADER & "]", "") & ", CheapTaxableLoan as [" & S_FIELD_LOAN_CHEAP_TAXABLE & "]" & OPRAAmountForegoneField("T_BenLoan") & " FROM T_BenLoan;"
      End If
    
    Case IMP_ADDRESSES
      If NameOnly Then
        QueriesEx = "Addresses"
      Else
        QueriesEx = "SELECT DISTINCTROW T_Addresses.P_Num AS " & S_IMP_PNUM_ALIAS & ", T_Addresses.Address1 AS [Address line 1], T_Addresses.Address2 AS [Address line 2], T_Addresses.Address3 AS [Address line 3], T_Addresses.City, T_Addresses.County, T_Addresses.PostCode AS [Postcode], T_Addresses.Country FROM T_Addresses"
      End If
    
    'EK altered to add IR category - bring in text string which is translated to number in readDB
    Case IMP_ASSETS_TRANSEFERRED
      If NameOnly Then
        QueriesEx = "Assets transferred"
      Else
        QueriesEx = "SELECT DISTINCTROW T_BenTrans.P_Num AS " & S_IMP_PNUM_ALIAS & ", T_BenTrans.IRDesc AS [" & S_IMP_FIELD_CATEGORY_DESCRIPTION & "], T_BenTrans.Item AS [Asset description], T_BenTrans.IsCar AS [Is this a car or has this not been a benefit previously], T_BenTrans.MktVal AS [Market value when transferred], T_BenTrans.MadeGood AS [Made good], T_BenTrans.MadeGoodIsTaxDeducted as [Amount is tax deducted], T_BenTrans.FirstMV AS [Market value when first provided], T_BenTrans.SumBen AS [Benefit in respect of private use prior to transfer]" & OPRAAmountForegoneField("T_BenTrans") & " From T_BenTrans"
      End If
    
    'EK altered to add IR category - bring in text string which is translated to number in readDB
    Case IMP_ASSETS_AT_DISPOSAL
      If NameOnly Then
        QueriesEx = "Assets at disposal"
      Else
        QueriesEx = "SELECT DISTINCTROW T_BenGoods.P_Num AS " & S_IMP_PNUM_ALIAS & ", T_BenGoods.IRDesc AS [" & S_IMP_FIELD_CATEGORY_DESCRIPTION & "], T_BenGoods.Item AS [Asset description], T_BenGoods.MktVal AS [Market value when provided], T_BenGoods.Marginal AS [Equivalent annual marginal costs], T_BenGoods.MadeGood AS [Made good],T_BenGoods.MadeGoodIsTaxDeducted as [Amount is tax deducted], T_BenGoods.Rent AS [Annual rent paid by employer], T_BenGoods.AvailFrom AS [Available from], T_BenGoods.AvailTo AS [Available to], T_BenGoods.DateAvail AS [Date first available as a benefit], T_BenGoods.UDBCode" & OPRAAmountForegoneField("T_BenGoods") & " From T_BenGoods WHERE ((T_BenGoods.UDBCode=" & StrSQL(S_ASSETSATDISPOSAL_UDBCODE) & "));"
      End If
    
      
    Case IMP_PAYMENTS_ON_BEHALF_OF_EMPLOYEE
      If NameOnly Then
        QueriesEx = "Payments made on behalf of employee"
      Else
        QueriesEx = "SELECT DISTINCTROW T_BenOther.P_Num AS " & S_IMP_PNUM_ALIAS & ", T_BenOther.Item AS [Description], T_BenOther.IRDesc as [" & S_IMP_FIELD_CATEGORY_DESCRIPTION & "],  T_BenOther.Value AS [Value of benefit],T_BenOther.MadeGood AS [Amount made good],T_BenOther.MadeGoodIsTaxDeducted as [Amount is tax deducted], T_BenOther.From AS [Available from], T_BenOther.To AS [Available to], T_BenOther.UDBCode, T_BenOther.Category, T_BenOther.Class From T_BenOther " & BenOtherWhereClause(BC_PAYMENTS_ON_BEFALF_B)
      End If
      
    Case IMP_SERVICES_PROVIDED
      If NameOnly Then
        QueriesEx = "Services provided"
      Else
        QueriesEx = "SELECT DISTINCTROW T_BenGoods.P_Num AS " & S_IMP_PNUM_ALIAS & ", T_BenGoods.Item AS [Service description], T_BenGoods.MktVal AS [Market value when provided], T_BenGoods.Marginal AS [Equivalent annual marginal costs], T_BenGoods.MadeGood AS [Made good],T_BenGoods.MadeGoodIsTaxDeducted as [Amount is tax deducted],  T_BenGoods.UDBCode " & OPRAAmountForegoneField("T_BenGoods") & " From T_BenGoods WHERE ((T_BenGoods.UDBCode=" & StrSQL(S_SERVICESPROVIDED_UDBCODE) & "));"
      End If
    
    Case IMP_O_OTHER_EXPENSES
      If NameOnly Then
        QueriesEx = "Other expenses"
      Else
        QueriesEx = "SELECT DISTINCTROW T_BenOther.P_Num AS " & S_IMP_PNUM_ALIAS & ", T_BenOther.IRDesc AS [" & S_IMP_FIELD_CATEGORY_DESCRIPTION & "], T_BenOther.Item AS Description,  T_BenOther.Value AS [Value of benefit],T_BenOther.MadeGood AS [Amount made good],T_BenOther.MadeGoodIsTaxDeducted as [Amount is tax deducted], T_BenOther.From AS [Available from], T_BenOther.To AS [Available to], T_BenOther.UDBCode, T_BenOther.Category, T_BenOther.Class " & OPRAAmountForegoneField("T_BenOther") & " From T_BenOther" & BenOtherWhereClause(BC_OOTHER_N)
      End If
    
    Case IMP_RELOCATION_ITEM
      If NameOnly Then
        QueriesEx = "Relocations"
      Else
        QueriesEx = "SELECT DISTINCTROW T_BenReloc.P_Num AS " & S_IMP_PNUM_ALIAS & ", T_BenReloc.Item AS Address, T_BenReloc.LastYear AS [Amount used last year], T_BenReloc.TaxYear AS [Tax Year of move] From T_BenReloc"
      End If
    
    Case IMP_RELOCATION_DETAIL
      If NameOnly Then
        QueriesEx = "Relocation items"
      Else
        QueriesEx = "SELECT DISTINCTROW T_BenRelocDetails.P_Num AS " & S_IMP_PNUM_ALIAS & ", T_BenRelocDetails.Address, T_BenRelocDetails.Item, T_BenRelocDetails.Value, T_BenRelocDetails.MadeGood AS [Amount made good], T_BenRelocDetails.MadeGoodIsTaxDeducted AS [Amount is tax deducted],  T_BenRelocDetails.Qualify AS [Is this a qualifying relocation expense], T_BenRelocDetails.Comments From T_BenRelocDetails"
      End If
    
'    Case IMP_SHARES
'      If NameOnly Then
'          QueriesEx = "Shares"
'        Else
'          QueriesEx = "SELECT DISTINCTROW T_BenOther.P_Num AS " & S_IMP_PNUM_ALIAS & ", T_BenOther.Item AS Description,  T_BenOther.Value AS [Value of benefit],T_BenOther.MadeGood AS [Amount made good],T_BenOther.MadeGoodIsTaxDeducted as [Amount is tax deducted],T_BenOther.From AS [Available from], T_BenOther.To AS [Available to], T_BenOther.UDBCode, T_BenOther.Category, T_BenOther.Class From T_BenOther" & BenOtherWhereClause(BC_SHARES_M)
'        End If
    
    Case IMP_NON_SHARED_VANS
      If NameOnly Then
        QueriesEx = "Non shared vans"
      Else
        QueriesEx = "SELECT DISTINCTROW T_BenVan.P_Num AS " & S_IMP_PNUM_ALIAS & ", T_BenVan.Item AS Description, T_BenVan.AvailFrom AS [Available from],T_BenVan.AvailTo AS [Available to], T_BenVan.Unavail AS [Days unavailable], T_BenVan.PrivateContrib AS [Private contribution],T_BenVan.RegistrationDate as [Registration Date] ,T_BenVan.MadeGoodIsTaxDeducted as [Amount is tax deducted], T_BenVan.BusReqOnly as [Mainly available for business travel only?], T_BenVan.CommuterUse as [Substantially available for Ordinary Commuting only?], T_BenVan.fuel_available as [Was fuel available?]" & _
                  "From T_BenVan WHERE ((T_BenVan.P_Num<>" & StrSQL(S_SHAREDVAN_KEY) & "));"
      End If
    
    Case IMP_SHARED_VANS
      If NameOnly Then
        QueriesEx = "Shared vans"
      Else
        QueriesEx = "SELECT DISTINCTROW T_BenVan.P_Num AS " & S_IMP_PNUM_ALIAS & ", T_BenVan.Item AS Description, T_BenVan.AvailFrom AS [Available from],T_BenVan.AvailTo AS [Available to], T_BenVan.Unavail AS [Days unavailable], T_BenVan.PrivateContrib AS [Private contribution],T_BenVan.RegistrationDate as [Registration Date] ,T_BenVan.MadeGoodIsTaxDeducted as [Amount is tax deducted], T_BenVan.fuel_available as [Was fuel available?] " & _
                  "From T_BenVan WHERE ((T_BenVan.P_Num=" & StrSQL(S_SHAREDVAN_KEY) & "));"
      End If
      
    Case POST_PROCESS_NON_SHARED_VANS_SELECT
      QueriesEx = "SELECT DISTINCTROW T_BENVAN.P_Num AS " & S_IMP_PNUM_ALIAS & ", T_BENVAN.Item AS [Non shared van], T_BENVAN.P_Num " & _
                  "FROM T_BENVAN LEFT JOIN T_EMPLOYEES ON T_BENVAN.P_Num = T_EMPLOYEES.P_Num " & _
                  "WHERE ((T_BENVAN.P_Num<>" & StrSQL(S_SHAREDVAN_KEY) & ") AND (T_EMPLOYEES.P_Num Is Null));"
    Case POST_PROCESS_SHARED_VANS_SELECT
      QueriesEx = "SELECT DISTINCTROW T_BenVan.P_Num AS " & S_IMP_PNUM_ALIAS & ", T_BenVan.Item AS Description, T_BenVan.AvailFrom AS [Available from],T_BenVan.AvailTo AS [Available to], T_BenVan.Unavail AS [Days unavailable], T_BenVan.PrivateContrib AS [Private contribution],T_BenVan.RegistrationDate as [Registration Date] ,T_BenVan.MadeGoodIsTaxDeducted as [Amount is tax deducted] " & _
                  "From T_BenVan WHERE ((T_BenVan.P_Num Is Null));"
    
'MP DB T_CarMiles tbl deleted - chk qry still true
'    Case POST_PROCESS_COMPANYCAR_MILEAGE_SELECT
'      QueriesEx = "SELECT DISTINCTROW T_CarMiles.Reg as [Car], T_CarMiles.Item as [Description of Mileage] ,T_CarMiles.P_Num AS " & S_IMP_PNUM_ALIAS & _
'                  "FROM T_BenCar RIGHT JOIN T_CarMiles ON (T_BenCar.Reg = T_CarMiles.Reg) AND (T_BenCar.P_Num = T_CarMiles.P_Num) " & _
'                  "Where ((T_BenCar.P_Num Is Null)) "
    Case POST_PROCESS_EMPLOYEECAR_MILEAGE_SELECT
      QueriesEx = "SELECT DISTINCTROW T_BenEeCarDetails.Reg AS Car, T_BenEeCarDetails.MilesItem AS [Description of Mileage], T_BenEeCarDetails.P_Num as " & S_IMP_PNUM_ALIAS & ", T_BenEeCar.P_Num " & _
                  "FROM T_BenEeCarDetails LEFT JOIN T_BenEeCar ON (T_BenEeCarDetails.Reg = T_BenEeCar.Item) AND (T_BenEeCarDetails.P_Num = T_BenEeCar.P_Num) " & _
                  "Where ((T_BenEeCar.P_Num Is Null)) "
    
    
    
    Case POST_PROCESS_BENEFICIAL_LOAN_BALANCES_SELECT
      QueriesEx = "SELECT DISTINCTROW T_Loans.Item as [Loan description], T_Loans.Payment as [Amount received] , T_Loans.P_Num AS " & S_IMP_PNUM_ALIAS & " FROM T_Loans LEFT JOIN T_BenLoan ON (T_Loans.Item = T_BenLoan.Item) AND (T_Loans.P_Num = T_BenLoan.P_Num) Where ((T_BenLoan.P_Num Is Null) AND (T_Loans.Type=" & StrSQL(S_BENEFICIAL_LOANS_TYPE) & "))"
    Case POST_PROCESS_RELOCATION_DETAIL_SELECT
      QueriesEx = "SELECT DISTINCTROW T_BenRelocDetails.Address, T_BenRelocDetails.Item, T_BenRelocDetails.P_Num AS " & S_IMP_PNUM_ALIAS & " FROM T_BenRelocDetails LEFT JOIN T_BenReloc ON (T_BenRelocDetails.P_Num = T_BenReloc.P_Num) AND (T_BenRelocDetails.Address = T_BenReloc.Item) WHERE (T_BenReloc.P_Num Is Null);"
    Case POST_PROCESS_OOTHER_O
      QueriesEx = "SELECT DISTINCTROW T_BENOTHER.P_Num AS " & S_IMP_PNUM_ALIAS & ", T_BENOTHER.Item AS [Other expense] " & _
                "FROM T_BENOTHER LEFT JOIN T_EMPLOYEES ON T_BENOTHER.P_Num = T_EMPLOYEES.P_Num " & _
                "WHERE ((T_BENOTHER.P_Num Not Like " & StrSQL(S_CDB_EMPLOYEE_NUMBER_PREFIX_LIKE) & ") AND (T_EMPLOYEES.P_Num Is Null));"
    Case UPDATE_EMPLOYEES
      QueriesEx = "UPDATE DISTINCTROW T_Employees SET T_Employees.Name = [T_Employees].[Surname] & ' ' & [T_Employees].[Title] & ' ' & [T_Employees].[FirstName] & ' ' & [T_Employees].[Initials] " & _
                  "Where ((T_Employees.Name Is Null))"
    Case UPDATE_BENEFICIAL_LOAN_BALANCES_KEYS
      QueriesEx = "UPDATE DISTINCTROW T_Loans INNER JOIN T_BenLoan ON (T_Loans.Item = T_BenLoan.Item) AND (T_Loans.P_Num = T_BenLoan.P_Num) SET T_Loans.[Key] = [T_BenLoan].[BalKey] WHERE ((T_Loans.Key Is Null Or T_Loans.Key='') AND (T_Loans.Type=" & StrSQL(S_BENEFICIAL_LOANS_TYPE) & "));"
'MP DB - tbl deleted
'   Case UPDATE_COMPANY_CAR_KEYS
'      QueriesEx = "UPDATE DISTINCTROW T_CarMiles INNER JOIN T_BenCar ON (T_CarMiles.Reg = T_BenCar.Reg) AND (T_CarMiles.P_Num = T_BenCar.P_Num) SET T_CarMiles.CarKey = [T_BenCar].[CarKey] Where ((T_CarMiles.CarKey=0))"
    Case UPDATE_EMPLOYEE_CAR_KEYS
      QueriesEx = "UPDATE DISTINCTROW T_BenEECarDetails INNER JOIN T_BenEECar ON (T_BenEECarDetails.Reg = T_BenEeCar.Item) AND (T_BenEECarDetails.P_Num = T_BenEECar.P_Num) SET T_BenEECarDetails.EECarKey = [T_BenEECar].[EECarKey] Where ((isnull(T_BenEECarDetails.EECarKey)))"
    Case UPDATE_RELOCATION_KEYS
      QueriesEx = "UPDATE DISTINCTROW T_BenReloc INNER JOIN T_BenRelocDetails ON (T_BenReloc.Item = T_BenRelocDetails.Address) AND (T_BenReloc.P_Num = T_BenRelocDetails.P_Num) SET T_BenRelocDetails.RelocKey = [T_BenReloc].[RelocKey] Where ((T_BenRelocDetails.RelocKey Is Null Or T_BenRelocDetails.RelocKey = ''))"
    Case UPDATE_COMPANY_CAR_FUEL_AVAILABLE_TO
      QueriesEx = "UPDATE  T_BenCar set FuelAvailTo = AvailTo Where isnull(FuelAvailTo)"
    
  End Select
  If Not IgnoreEmpty And (Len(QueriesEx) = 0) Then Call ECASE("Queries " & CStr(QueryID) & "not found")
  
QueriesEx_END:
  Call xReturn("Queries")
  Exit Function
QueriesEx_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "QueriesEx", "Queries Ex", "Error getting the query string for an import, query id = " & QueryID & ".")
  Resume QueriesEx_END
  Resume
End Function


Private Function OPRAAmountForegoneField(ByVal tableName As String, Optional ByVal fieldName As String = "OPRAAmountForegone", Optional ByVal displayNameAdditional As String = "")
  Dim s As String
  
  s = ", " & tableName & "." & fieldName & " AS [" & S_UDM_OPRA_AMOUNT_FOREGONE & displayNameAdditional & "]"
  OPRAAmountForegoneField = s
  
End Function

Private Sub Class_Initialize()
  S_GEN_IMP_SQL = S_GEN_IMP_SQL_BASE & OPRAAmountForegoneField("T_BenOther") & " From T_BenOther"
End Sub
