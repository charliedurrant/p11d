VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SQLQUERIES"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum QUERY_NAMES
  SELECT_EMPLOYEES = 0
  SELECT_EMPLOYEES_NO_CDB
  SELECT_EMPLOYEE_READ
  SELECT_EMPLOYEE_WRITE
  SELECT_EMPLOYEECHECK
  SELECT_EMPLOYEES_COUNT
  SELECT_EMPLOYEE_NEW
  SELECT_CARS
  SELECT_OTHERS
  SELECT_VANS
  SELECT_SHARED_VANS
'MP DB  DELETE_SHARED_VANS
  SELECT_CDB_LINKS
  SELECT_FPCS_EECARS
'MP DB  SELECT_CARMILES
'MP DB  DELETE_CARMILES
  SELECT_EECARMILES
  DELETE_EECARMILES
  FIX_O_OTHER_EXPENSE
  FIX_SHARES
  FIX_MEDICAL
  FIX_CHAUFFEUR
  FIX_GENERAL
  FIX_ENTERTAINMENT
  FIX_TRAVEL
  FIX_TAX_PAID
  FIX_EDUCATION
  FIX_NURSERY
  FIX_SUBSCRIPTIONS
  FIX_O_OTHER
  FIX_SERVICES_PROVIDED
  FIX_ASSETS_DISPOSAL
  FIX_N_2004
  FIX_M_2004
  FIX_N_OTHER_2004
'MP RV TTP#320  SET_ACTUAL_MILES
  FIX_FREQUENCY_OF_PAYMENTS
  FIX_MOBILE_PHONES           'RH
  FIX_SHARES_REMOVE
  FIX_UPDATE_GUID
  FIX_BRING_CAR_FUEL_AVAIL_TO_THIS_YEAR
      
  UPDATE_LAST_APPLICATION
'MP DB  UPDATE_CDB_LINKS
  UPDATE_NAME_PARTS_TO_NULL
  UPDATE_FIX_LEVEL
  UPDATE_CDCS_OLD_TO_NEW
  UPDATE_PERSONNEL_NUMBER
  
  SELECT_LOAN_BALANCES
  SELECT_LOAN_RATES
  DELETE_LOANBAL
  SELECT_ID
'MP DB  SELECT_AUDIT
'MP DB  SELECT_AUDITITEM
  SELECT_BENLOANS
'MP DB  SELECT_HOMELOANS
  SELECT_ALLLOANBALS
  SELECT_PHONE
  SELECT_ASSETSATDISPOSAL
  SELECT_SERVICESPROVIDED
  SELECT_TRANS
  SELECT_ACCOM
  SELECT_RELOC
  SELECT_EECARS
  SELECT_CDC_EMPLOYEES
  SELECT_CDCS_EMPLOYEES
  
  UPDATE_CDB_LINKS
  COPY_CDB_LINKS
  
  SELECT_UNSPLIT_EMPLOYEES
  DELETE_EECAR
  SELECT_RELOCDETAILS
  DELETE_RELOCDETAILS
  SELECT_FPCS
  DELETE_ER_FPCS
  SELECT_IR_FPCS
  DELETE_IR_FPCS
  SELECT_IR_CO2EMISSIONS_MAKES
  
  DELETE_T_BEN_CDC_ITEMS
  
  
  DELETE_CDBS_DEBUG
  DELETE_CDB_LINKS_DEBUG
  DELETE_CDB_LINKS_EMPLOYEE
  DELETE_CDB_EMPLOYEE_BENEFITS
  DELETE_ADDRESS
  SELECT_CDCS
'MP DB  SELECT_CDC
  DELETE_CDC
'MP DB  DELETE_CDC_OOTHERS
  DELETE_CDCS
  'Address specific
  ADD_ADDRESS
  UPDATE_ADDRESS
  DELETE_MAGNETIC_MEDIA_ERRORS
  SELECT_MAGNETIC_MEDIA_ERRORS
  DELETE_COMPANYCAR_CHECKER_CARS_LOG
  SELECT_COMPANYCAR_CHECKER_CARS_LOG_ALL
  SELECT_COMPANYCAR_CHECKER_CARS_LOG_DATES
  SELECT_COMPANYCAR_CHECKER_CARS_LOG_OVERLAPS
  SELECT_COMPANYCAR_CHECKER_CARS_LOG_NOCARS
  ' SELECT_COMPANYCAR_CHECKER_CARS_LOG_FUELAVAILDATES
  SELECT_COMPANYCAR_CHECKER_CARS_LOG_AVAILDATES
  SELECT_COMPANYCAR_CHECKER_CARS_LOG_REGDATES
  SELECT_COMPANYCAR_CHECKER_CARS_DATES
  SELECT_COMPANYCAR_CHECKER_CARS_OVERLAPS
  SELECT_COMPANYCAR_CHECKER_CARS_NOCARS
  SELECT_COMPANYCAR_CHECKER_CARS_FUELAVAILDATES
  SELECT_COMPANYCAR_CHECKER_CARS_AVAILDATES
  SELECT_COMPANYCAR_CHECKER_CARS_REGDATES
  SELECT_COMPANYCAR_CHECKER_REG_DATES_NODE_CLICK
  SELECT_COMPANYCAR_CHECKER_CARS_AVAILDATES_TV
  SELECT_COMPANYCAR_CHECKER_CARS_OVERLAP
  SELECT_COMPANYCAR_CHECKER_CARS_OVERLAP_TV
  SELECT_COMPANYCAR_CHECKER_SEQUENTIAL_NOT_MARKED_AS_REPLACED_TV
  SELECT_COMPANYCAR_CHECKER_CARS_IN_USE_BY_MORE_THAN_ONE
  SELECT_COMPANYCAR_CHECKER_CARS_IN_USE_BY_MORE_THAN_ONE_TV
  SELECT_COMPANYCAR_CHECKER_CARS_IN_USE_BY_MORE_THAN_ONE_ANALYSE_TV
  SELECT_COMPANYCAR_CHECKER_CARS_EEE_AVAIL_DATES
  SELECT_COMPANYCAR_CHECKER_CARS_EEE_AVAIL_DATES_TV
  SELECT_EMPLOYEE_CHECK_NI_VALID_TV
  SELECT_EMPLOYEE_CHECK_PNUM_VALID_TV
  SELECT_EMPLOYEE_CHECK_INVALID_NAME_TV
  FIND_COMPANY_CAR
  DELETE_CDC_DEBUG
  DELETE_CDC_LINKS_DEBUG
  SELECT_DUPLICATE_NI_NUMBERS
  
  SELECT_PAYEONLINE_XML
  SELECT_PAYEONLINE_SUBMISSION_ERRORS
  SELECT_PAYEONLINE_VALIDATION_ERRORS
  DELETE_PAYEONLINE_VALIDATION_ERRORS
'MP DB  FIND_ACCOM_EXPENSES
'MP DB  DELETE_ACCOM_EXPENSES
  
  SELECT_HELPLINKS
End Enum

Public Function Queries(QueryID As QUERY_NAMES, ParamArray P()) As String
  Dim i As Long
  Dim s As String
  
  Call xSet("Queries")
  i = LBound(P)
  Select Case QueryID
    Case ADD_ADDRESS
      Queries = "INSERT INTO T_Addresses ( Address1, Address2, Address3, City, County, PostCode, Country ,P_Num) VALUES (" & StrSQL(P(0)) & "," & StrSQL(P(1)) & "," & StrSQL(P(2)) & "," & StrSQL(P(3)) & "," & StrSQL(P(4)) & "," & StrSQL(P(5)) & "," & StrSQL(P(6)) & "," & StrSQL(P(7)) & ")"
    Case UPDATE_ADDRESS
      Queries = "UPDATE DISTINCTROW T_Addresses LEFT JOIN T_Employees ON T_Addresses.P_Num = T_Employees.P_Num SET T_Addresses.Address1 = " & StrSQL(P(0)) & ", T_Addresses.Address2 = " & StrSQL(P(1)) & ", T_Addresses.Address3 = " & StrSQL(P(2)) & ", T_Addresses.City = " & StrSQL(P(3)) & ", T_Addresses.County = " & StrSQL(P(4)) & ", T_Addresses.PostCode = " & StrSQL(P(5)) & ", T_Addresses.Country = " & StrSQL(P(6)) & " Where T_Employees.P_Num =" & StrSQL(P(7))
    Case DELETE_ADDRESS
      Queries = "Delete * FROM T_Addresses Where P_Num =" & StrSQL(P(i)) & ";"
    Case SELECT_EMPLOYEES
      Queries = "SELECT DISTINCTROW T_Employees.*, T_Addresses.Address1, T_Addresses.Address2, T_Addresses.Address3, T_Addresses.City, T_Addresses.County, T_Addresses.PostCode, T_Addresses.Country FROM T_Employees LEFT JOIN T_Addresses ON T_Employees.P_Num = T_Addresses.P_Num;"
    Case SELECT_EMPLOYEES_NO_CDB
      Queries = "SELECT DISTINCTROW T_Employees.* From T_Employees WHERE ((T_Employees.P_Num <> " & StrSQL(S_CDB_EMPLOYEE_NUMBER_PREFIX) & "));"
    Case SELECT_EMPLOYEE_READ
      Queries = "SELECT DISTINCTROW T_Employees.*, T_Addresses.Address1, T_Addresses.Address2, T_Addresses.Address3, T_Addresses.City, T_Addresses.County, T_Addresses.PostCode, T_Addresses.Country FROM T_Employees LEFT JOIN T_Addresses ON T_Employees.P_Num = T_Addresses.P_Num  WHERE T_Employees.P_Num = " & StrSQL(P(i))
    Case SELECT_EMPLOYEE_WRITE
      Queries = "SELECT * FROM T_EMPLOYEES WHERE P_NUM = " & StrSQL(P(i))
'MP DB not in use
'    Case SELECT_EMPLOYEE_NEW
'      Queries = "T_Employees"
    Case SELECT_EMPLOYEECHECK
      Queries = "SELECT DISTINCTROW P_Num, DateEdit, UserEdit " & _
                "From T_Employees WHERE ((P_Num=" & StrSQL(P(i)) & "))"
    Case SELECT_EMPLOYEES_COUNT
      Queries = "SELECT DISTINCTROW (Count(T_Employees.P_Num) - 1) AS Count FROM T_Employees;" 'less 1 for CDB employee
    Case SELECT_CARS
      Queries = "SELECT T_BenCar.* From T_BenCar ORDER BY T_BenCar.AvailFrom;"
    Case SELECT_EECARS
      Queries = "SELECT * FROM T_BENEECAR"
    Case SELECT_FPCS_EECARS
      Queries = "SELECT DISTINCTROW T_BenEeCar.P_Num,T_BenEeCar.FPCS From T_BenEeCar WHERE ((T_BenEeCar.FPCS=" & StrSQL(P(i)) & "));"
    Case SELECT_OTHERS
      Queries = "SELECT * FROM T_BENOTHER where class <> " & StrSQL(S_SHARES_M)
    Case SELECT_VANS
      Queries = "SELECT DISTINCTROW T_BenVan.* From T_BenVan WHERE ((T_BenVan.P_Num<> " & StrSQL(S_SHAREDVAN_KEY) & "))"
    Case SELECT_SHARED_VANS
      Queries = "SELECT DISTINCTROW T_BenVan.* From T_BenVan WHERE ((T_BenVan.P_Num=" & StrSQL(S_SHAREDVAN_KEY) & "))"
'MP DB not used
'    Case DELETE_SHARED_VANS 'SO
'      Queries = "DELETE * T_BenVan"

    Case SELECT_CDB_LINKS
      Queries = "T_BenStd"
'MP DB table not used
'    Case SELECT_CARMILES
'      Queries = "SELECT * FROM T_CARMILES WHERE (CARKEY=" & NumSQL(P(i)) & ") ORDER BY MILESDATE"
    Case SELECT_UNSPLIT_EMPLOYEES
      Queries = "SELECT Name, Title, FirstName, Initials, Surname From T_Employees WHERE ((Title Is Null) or (FirstName Is Null) or (Initials Is Null) or (Surname Is Null)) and (P_Num<>" & StrSQL(S_CDB_EMPLOYEE_NUMBER_PREFIX) & ")"
'MP DB    Case DELETE_CARMILES
'      Queries = "DELETE FROM T_CARMILES WHERE (CARKEY=" & NumSQL(P(i)) & ")"
    Case SELECT_EECARMILES
      Queries = "SELECT * FROM T_BENEECARDETAILS WHERE (EECARKEY=" & NumSQL(P(i)) & ") ORDER BY MILESDATE"
    Case DELETE_EECARMILES
      Queries = "DELETE FROM T_BENEECARDETAILS WHERE (EECARKEY=" & NumSQL(P(i)) & ")"
'    Case FIX_LOAN_OTHER_BALANCES
'      Queries = "UPDATE DISTINCTROW T_BenLoan INNER JOIN T_Loans ON T_BenLoan.BalKey = T_Loans.Key SET T_Loans.P_Num = [T_BenLoan].[P_Num], T_Loans.Item = [T_BenLoan].[Item] " & _
'                "Where ((T_Loans.Type = " & StrSQL(S_BENEFICIAL_LOANS_TYPE) & "))"
'    Case FIX_RELOCATION_ITEMS
'      Queries = "UPDATE DISTINCTROW T_BenReloc, T_BenRelocDetails SET T_BenRelocDetails.P_Num = [T_BenReloc].[P_Num], T_BenRelocDetails.Address = [T_BenReloc].[Item]"
'    Case FIX_EECAR_MILES
'      Queries = "UPDATE DISTINCTROW T_BenEeCar INNER JOIN T_BenEeCarDetails ON T_BenEeCar.EeCarKey = T_BenEeCarDetails.EeCarKey SET T_BenEeCarDetails.P_Num = [T_BenEECar].[P_Num], T_BenEeCarDetails.Reg = [T_BenEECar].[Item]"
'    Case FIX_CAR_MILES
'      Queries = "UPDATE DISTINCTROW T_BenCar INNER JOIN T_CarMiles ON T_BenCar.CarKey = T_CarMiles.CarKey SET T_CarMiles.P_Num = [T_BenCar].[P_Num], T_CarMiles.Reg = [T_BenCar].[Reg] "
'    Case FIX_CARS_MAKE_MODEL
'      Queries = "UPDATE DISTINCTROW T_BenCar SET T_BenCar.Make = T_BenCar.Item WHERE ((T_BenCar.Model Is Null) AND (T_BenCar.Make Is Null));"
    Case UPDATE_LAST_APPLICATION
      Queries = "UPDATE DISTINCTROW id SET id.LastApplication = " & StrSQL(p11d32.Rates.ApplicationType)
'MP DB - not used anywhere
'    Case UPDATE_CDB_LINKS
'      Queries = "INSERT INTO T_BenStd (P_Num, BenCode) VALUES (" & StrSQL(P(0)) & "," & StrSQL(P(1)) & ")"
    Case UPDATE_FIX_LEVEL
      Queries = "UPDATE DISTINCTROW id SET id.FixLevel = " & NumSQL(P(0))
'    Case UPDATE_OOTHER
'      Queries = "INSERT INTO T_BenOther ( P_Num, UDBCode, UDB, Item, Benefit, [Value], MadeGood, [From], To, Category, Class )" & _
'                "SELECT DISTINCTROW T_BenCDC.P_Num, T_BenCDC.UDBCode, T_BenCDC.UDB, T_BenCDC.Item, T_BenCDC.Benefit, T_BenCDC.Value, T_BenCDC.MadeGood, T_BenCDC.From, T_BenCDC.To, '" & p11d32.Rates.BenClassTo(BC_OOTHER_N, BCT_DBCATEGORY) & "' AS Expr1 ," & StrSQL(p11d32.Rates.BenClassTo(BC_OOTHER_N, BCT_DBCLASS)) & " AS Expr2 " & _
'                "From T_BenCDC Where ((T_BenCDC.CDCKey = 1))"
'    Case UPDATE_OOTHER_CDC
'      Queries = "INSERT INTO T_BenOther ( P_Num, UDBCode, UDB, Item, Benefit, [Value], MadeGood, [From], To, CDCKey, Category, Class ) " & _
'                "SELECT DISTINCTROW T_BenCDC.P_Num, T_BenCDC.UDBCode, T_BenCDC.UDB, T_BenCDC.Item, T_BenCDC.Benefit, T_BenCDC.Value, T_BenCDC.MadeGood, T_BenCDC.From, T_BenCDC.To, T_BenCDC.CDCKey, '" & p11d32.Rates.BenClassTo(BC_OOTHER_N, BCT_DBCATEGORY) & "' AS Expr1 ," & StrSQL(p11d32.Rates.BenClassTo(BC_OOTHER_N, BCT_DBCLASS)) & " AS Expr2 " & _
'                "FROM T_CDC INNER JOIN T_BenCDC ON T_CDC.CDCKey = T_BenCDC.CDCKey " & _
'                "Where ((T_BenCDC.CDCKey > 1))"
    Case UPDATE_NAME_PARTS_TO_NULL
      Queries = "UPDATE DISTINCTROW T_Employees SET T_Employees.Initials = Null, T_Employees.FirstName = Null, T_Employees.Surname = Null, T_Employees.Title = Null;"
    Case UPDATE_PERSONNEL_NUMBER
      Queries = "UPDATE DISTINCTROW [" & P(i) & "] SET P_Num = " & StrSQL(P(i + 1)) & " WHERE ((P_Num=" & StrSQL(P(i + 2)) & "));"
    Case SELECT_LOAN_BALANCES
      If IsNumeric(P(i)) Then
        Queries = "SELECT DISTINCTROW [Key], [Type], [From], [Payment] FROM T_Loans " & _
                 "WHERE ((Key = " & NumSQL(P(i)) & ") And (T_Loans.Type = " & StrSQL(P(i + 1)) & ")) ORDER BY T_Loans.From;" 'so
      Else
        Queries = "SELECT DISTINCTROW [Key], [Type], [From], [Payment] FROM T_Loans " & _
                 "WHERE ((Key = " & StrSQL(P(i)) & ") And (T_Loans.Type = " & StrSQL(P(i + 1)) & ")) ORDER BY T_Loans.From;" 'so
      End If
    Case SELECT_LOAN_RATES
      Queries = "SELECT DISTINCTROW InterestDate, Sterling, Yen, Franc From T_InterestRates " & _
                "ORDER BY InterestDate"
    Case DELETE_LOANBAL
      Queries = "DELETE FROM T_LOANS WHERE (KEY=" & NumSQL(P(i)) & ") AND (TYPE=" & StrSQL(P(i + 1)) & ")"
'       Queries = "DELETE FROM T_LOANS WHERE (KEY=" & NumSQL(P(i)) & ") "
'    Case FIX_VANS_ROGUE
'      Queries = "DELETE DISTINCTROW T_BenVan.*, T_BenVan.P_Num From T_BenVan WHERE ((T_BenVan.P_Num=" & StrSQL("dsfsdf") & "));"
'    Case FIX_VANS_BLANK
'      Queries = "DELETE DISTINCTROW T_BenVan.Item From T_BenVan WHERE ((T_BenVan.Item Is Null)) OR ((T_BenVan.Item=''));"
'      'bad mistake CD
    Case SELECT_ID
      Queries = "Select * from ID"
    Case SELECT_BENLOANS
      Queries = "Select * from T_BENLOAN"
    Case SELECT_ALLLOANBALS
      Queries = "Select * from T_LOANS"
    Case SELECT_PHONE
      Queries = "SELECT * FROM T_BENPHONE"
    Case SELECT_SERVICESPROVIDED
      Select Case p11d32.AppYear
      Case 2000
        Queries = "SELECT DISTINCTROW T_BenGoods.* From T_BenGoods Where AvailFrom < #04/04/1999#" ' SO.  Charlie please check
      Case Else
        Queries = "SELECT DISTINCTROW T_BenGoods.* From T_BenGoods Where ((T_BenGoods.UDBCode = " & StrSQL(S_SERVICESPROVIDED_UDBCODE) & "))"
      End Select
    Case SELECT_ASSETSATDISPOSAL
      Select Case p11d32.AppYear
      Case 2000
        Queries = "SELECT DISTINCTROW T_BenGoods.* From T_BenGoods Where AvailFrom >= #04/04/1999#" ' SO.  Charlie please check
      Case Else
        Queries = "SELECT DISTINCTROW T_BenGoods.* From T_BenGoods Where ((T_BenGoods.UDBCode = " & StrSQL(S_ASSETSATDISPOSAL_UDBCODE) & "))"
      End Select
    Case SELECT_TRANS
      Queries = "SELECT * FROM T_BENTRANS"
    Case SELECT_ACCOM
      Queries = "SELECT * FROM T_BENACCOM"
    Case SELECT_RELOC
      Queries = "SELECT * FROM T_BENRELOC"
    Case SELECT_RELOCDETAILS
      Queries = "SELECT * FROM T_BENRELOCDETAILS WHERE (RELOCKEY=" & StrSQL(P(i)) & ")"
    Case SELECT_CDC_EMPLOYEES
      Queries = "SELECT DISTINCTROW T_BenOther.P_Num, T_BenOther.CDCKey " & _
                "From T_BenOther Where ((T_BenOther.CDCKey > 0)) "
    Case SELECT_CDCS_EMPLOYEES
      Queries = "SELECT DISTINCTROW T_BenOther.CDCKey, Count(T_BenOther.CDCKey) AS Count " & _
                "From T_BenOther GROUP BY T_BenOther.CDCKey Having ((T_BenOther.CDCKey > 0))"
    Case SELECT_DUPLICATE_NI_NUMBERS
      If Len(P(i)) > 0 Then
        P(i) = " WHERE " & P(i)
      End If
      Queries = "SELECT count(NI) AS CountOfNI, NI From T_Employees " & P(i) & " GROUP BY NI  HAVING (Count(NI)>1);"
    Case DELETE_RELOCDETAILS
      Queries = "DELETE * FROM T_BENRELOCDETAILS " & _
                "WHERE (RELOCKEY=" & StrSQL(P(i)) & ")"
    Case DELETE_T_BEN_CDC_ITEMS
      Queries = "DELETE * FROM T_BENCDC"
    Case UPDATE_CDCS_OLD_TO_NEW
      Queries = UPDATE_CDCS_OLD_TO_NEW
    Case SELECT_CDCS
      Queries = "Select * FROM T_CDC"
'MP DB not used anywhere
'    Case SELECT_CDC
'      Queries = "Select * FROM T_CDC Where CDCKey = " & NumSQL(P(0))
    Case DELETE_CDC
      Queries = "Delete * FROM T_CDC Where CDCKey = " & NumSQL(P(0))
    Case DELETE_CDCS
      Queries = "Delete * FROM T_CDC "
'MP DB not used anywhere
'    Case DELETE_CDC_OOTHERS
'      Queries = "Delete * FROM T_CDC where CDCKey = 1"
    Case SELECT_FPCS
      Queries = "SELECT * FROM T_CDRNEW ORDER BY FPCS"
    Case DELETE_ER_FPCS
      Queries = "DELETE * FROM T_CDRNEW"
    Case SELECT_IR_FPCS
      Queries = "SELECT * FROM T_FPCS"
    Case DELETE_IR_FPCS
      Queries = "DELETE * FROM T_FPCS"
    Case SELECT_IR_CO2EMISSIONS_MAKES 'km
      Queries = "SELECT Make FROM T_CO2Emissions group by Make order by MAKE"   'km
    Case DELETE_CDB_EMPLOYEE_BENEFITS
      Queries = "DELETE * From T_BenOther WHERE P_Num Like " & StrSQL(S_CDB_EMPLOYEE_NUMBER_PREFIX_LIKE)
    Case DELETE_CDB_LINKS_EMPLOYEE
      Queries = "DELETE DISTINCTROW T_BenStd.* From T_BenStd WHERE P_Num = " & StrSQL(P(i))
    Case DELETE_CDBS_DEBUG
      Queries = "Delete * from T_BenStd"
    Case DELETE_CDB_LINKS_DEBUG
      Queries = "DELETE DISTINCTROW T_BenOther.P_Num From T_BenOther WHERE ((T_BenOther.P_Num Like " & StrSQL(S_CDB_EMPLOYEE_NUMBER_PREFIX_LIKE) & "));"
    Case UPDATE_CDB_LINKS
      Queries = "UPDATE T_BenStd SET BenCode = " & StrSQL(P(i)) & _
                " WHERE BenCode =  " & StrSQL(P(i + 1))
    Case COPY_CDB_LINKS
      Queries = "INSERT INTO T_BenStd ( P_Num, BenCode) IN " & StrSQL(P(i)) & _
                " SELECT DISTINCTROW T_BenStd.P_Num, T_BenStd.BenCode " & _
                "FROM T_BenStd;"
    Case DELETE_MAGNETIC_MEDIA_ERRORS
      Queries = "DELETE * from MagneticMediaErrors"
    Case SELECT_MAGNETIC_MEDIA_ERRORS
      Queries = "SELECT MagneticMediaErrors.ErrorNumber AS [Error number], MagneticMediaErrors.ErrorDescription AS Description, MagneticMediaErrors.EmployerName AS Employer, MagneticMediaErrors.EmployeeName AS Employee, MagneticMediaErrors.BenefitName AS Benefit, MagneticMediaErrors.HMITSectionString AS [HMIT Section], MagneticMediaErrors.BenfitFormName AS [Benefit type] " & _
                "From MagneticMediaErrors"
    Case DELETE_PAYEONLINE_VALIDATION_ERRORS
      Queries = "DELETE * from ValidationErrors"
    Case SELECT_PAYEONLINE_VALIDATION_ERRORS
      Queries = "SELECT ValidationErrors.ErrorNumber AS [Error number], ValidationErrors.ErrorDescription AS Description, ValidationErrors.EmployerName AS Employer, ValidationErrors.EmployeeName AS Employee, ValidationErrors.BenefitName AS Benefit, ValidationErrors.HMITSectionString AS [HMIT Section], ValidationErrors.BenfitFormName AS [Benefit type] " & _
                "From ValidationErrors"
    Case SELECT_PAYEONLINE_XML
      Queries = "SELECT SubmissionBody AS [xml] From Submissions where SubmissionID like '" & P(0) & "'"
    Case DELETE_COMPANYCAR_CHECKER_CARS_LOG
      Queries = "DELETE * from CompanyCarChecker"
    Case SELECT_COMPANYCAR_CHECKER_CARS_LOG_ALL
      Queries = "CompanyCarChecker"
    Case SELECT_COMPANYCAR_CHECKER_CARS_LOG_DATES
      Queries = "Select Employee,PNum,Reg,[From] ,OldTo AS [To],OldDateReg as [RegDate], NewDateReg as [Suggested Reg Date], DateRegAmended As [Reg Amended], Replaced, Replacement, NewTo AS [Suggested To Date], ToAmended, Comments from CompanyCarChecker"
    Case SELECT_COMPANYCAR_CHECKER_CARS_LOG_OVERLAPS
      Queries = "Select Employee,PNum,Reg,[From] ,OldTo AS [To],OldDateReg as [RegDate],NewDateReg as [Suggested Reg Date], OldNumOfUsers As [Number of users], NewNumOfUsers As [Suggested number of users], Comments from CompanyCarChecker"
    Case SELECT_COMPANYCAR_CHECKER_CARS_LOG_NOCARS
      Queries = "Select Employee,PNum,Reg,[From] ,OldTo AS [To],OldDateReg as [RegDate],Comments from CompanyCarChecker"
'    Case SELECT_COMPANYCAR_CHECKER_CARS_LOG_FUELAVAILDATES 'IK
'      Queries = "Select Employee,PNum,Reg,[From] ,OldTo AS [To], FuelFrom As [Fuel Available From],OldFuelTo As [Fuel Available To],Comments from CompanyCarChecker"
    Case SELECT_COMPANYCAR_CHECKER_CARS_LOG_AVAILDATES
      Queries = "SELECT P_Num, Reg, AvailFrom, AvailTo  from T_BenCar WHERE AvailFrom >= AvailTo AND P_Num = '" & P(0) & "'"
    Case SELECT_COMPANYCAR_CHECKER_REG_DATES_NODE_CLICK
      Queries = "SELECT P_Num, Reg, AvailFrom, AvailTo, RegDate  from T_BenCar WHERE  Reg = '" & P(0) & "'"
    Case SELECT_COMPANYCAR_CHECKER_CARS_AVAILDATES_TV
      Queries = "SELECT * FROM (SELECT  DISTINCT  e.Surname, e.Surname + ', ' + e.Firstname AS FULLNAME, e.P_Num AS Data, e.P_Num  , e.NI from T_BenCar car INNER JOIN T_Employees e ON car.P_Num = e.P_Num  WHERE AvailFrom > AvailTo) ORDER BY " & CStr(P(0))
    Case SELECT_COMPANYCAR_CHECKER_CARS_OVERLAP
      Queries = "SELECT  P_Num, Reg, Reg as Displayname, AvailFrom, AvailTo,Replaced, Replacement, RegReplaced  FROM T_BenCar WHERE P_Num = '" & P(0) & "' order by AvailFrom"
    Case SELECT_COMPANYCAR_CHECKER_CARS_OVERLAP_TV
      'Get users who have more than one car - as o\nly they can have overlapping cars
      Queries = "SELECT * FROM (SELECT NI, Surname, Surname + ', ' + Firstname AS FULLNAME, P_Num as Data, P_Num  FROM T_Employees WHERE " & _
                "P_Num IN (SELECT c1.P_Num FROM T_BenCar c1 " & _
                "INNER JOIN T_BenCar c2 " & _
                "ON ((c1.P_Num = c2.P_Num) AND (NOT (c1.P_Num =c2.P_Num  AND c1.Reg=c2.Reg AND c1.AvailFrom=c2.AvailFrom AND c1.AvailTo=c2.AvailTo)) " & _
                "AND c2.AvailFrom<c1.AvailTo AND c2.AvailTo>=c1.AvailTo) where c1.SecondCar = 0 and c2.SecondCar = 0 )) ORDER BY " & CStr(P(0))
    Case SELECT_COMPANYCAR_CHECKER_SEQUENTIAL_NOT_MARKED_AS_REPLACED_TV
      'TV is the query to get the people with cars that have problems
       Queries = "SELECT DISTINCT * FROM " & _
                 "(SELECT * FROM (SELECT NI, Surname, Surname + ', ' + Firstname AS FULLNAME, P_Num as Data, P_Num  FROM T_Employees WHERE " & _
                 "P_Num IN " & _
                 "( " & _
                 "SELECT c1.P_Num " & _
                 "FROM T_BenCar AS c1 INNER JOIN T_BenCar AS c2 ON (c2.AvailTo=DateAdd('d',-1,c1.AvailFrom)) AND (NOT (c1.P_Num =c2.P_Num  AND c1.Reg=c2.Reg AND c1.AvailFrom=c2.AvailFrom AND c1.AvailTo=c2.AvailTo)) AND (c1.P_Num = c2.P_Num) " & _
                 "Where (c1.SecondCar = 0) " & _
                 "AND ( (c2.replaced = 0  ) ) " & _
                 ") " & _
                 ") " & _
                 " UNION " & _
                 "SELECT * FROM (SELECT NI, Surname, Surname + ', ' + Firstname AS FULLNAME, P_Num as Data, P_Num  FROM T_Employees WHERE " & _
                 "P_Num IN " & _
                 "( " & _
                 "SELECT c1.P_Num " & _
                 "FROM T_BenCar AS c1 INNER JOIN T_BenCar AS c2 ON (c2.AvailTo=c1.AvailFrom) AND (NOT (c1.P_Num =c2.P_Num  AND c1.Reg=c2.Reg AND c1.AvailFrom=c2.AvailFrom AND c1.AvailTo=c2.AvailTo)) AND (c1.P_Num = c2.P_Num) " & _
                 "Where (c1.SecondCar = 0) " & _
                 "AND ( (c2.replaced = 0 ) ) " & _
                 ") " & _
                 ") )"
                 
    Case SELECT_COMPANYCAR_CHECKER_CARS_IN_USE_BY_MORE_THAN_ONE_ANALYSE_TV
      Queries = "SELECT P_Num, Reg, AvailFrom, AvailTo FROM T_BenCar WHERE Reg = '" & P(0) & "'"
    Case SELECT_COMPANYCAR_CHECKER_CARS_IN_USE_BY_MORE_THAN_ONE_TV
      'Queries = "SELECT Reg  FROM T_BenCar GROUP BY Reg HAVING Count(Reg) > 1"
      
      Queries = "SELECT DISTINCT c1.Reg AS Data, c1.Reg AS DisplayName FROM T_BenCar c1 INNER JOIN T_BenCar c2 ON (c1.Reg = c2.Reg) AND (NOT (c1.P_Num =c2.P_Num AND c1.Reg=c2.Reg AND c1.AvailFrom=c2.AvailFrom AND c1.AvailTo=c2.AvailTo)) AND c2.AvailFrom<=c1.AvailTo AND c2.AvailTo>=c1.AvailTo group by c1.Reg"
      'Queries = "SELECT DISTINCT c1.Reg AS Data, c1.Reg AS DisplayName, c1.P_Num FROM T_BenCar c1 INNER JOIN T_BenCar c2 " & _
      '          "ON (c1.Reg = c2.Reg) AND (NOT (c1.P_Num =c2.P_Num AND c1.Reg=c2.Reg AND c1.AvailFrom=c2.AvailFrom AND c1.AvailTo=c2.AvailTo)) " & _
      '          "AND c2.AvailFrom<=c1.AvailTo AND c2.AvailTo>=c1.AvailTo"
    Case SELECT_COMPANYCAR_CHECKER_CARS_IN_USE_BY_MORE_THAN_ONE
      Queries = "SELECT ee.Surname + ',' + ee.Firstname + ' (' + ee.P_Num + ')'  as Displayname ,  c.P_Num,c.Reg, c.AvailFrom, c.AvailTo, c.Replaced, c.Replacement, c.RegReplaced  FROM T_BenCar c INNER JOIN T_Employees ee ON c.P_Num = ee.P_Num WHERE Reg = '" & P(0) & "'"
    Case SELECT_COMPANYCAR_CHECKER_CARS_LOG_REGDATES
      Queries = "Select Employee,PNum,Reg,[From] ,OldTo AS [To],OldDateReg as [RegDate],NewDateReg as [Suggested Reg Date],Comments from CompanyCarChecker"
    Case SELECT_COMPANYCAR_CHECKER_CARS_DATES, SELECT_COMPANYCAR_CHECKER_CARS_OVERLAPS, SELECT_COMPANYCAR_CHECKER_CARS_FUELAVAILDATES, SELECT_COMPANYCAR_CHECKER_CARS_AVAILDATES
      Queries = "SELECT T_Employees.SurName AS Employee, T_Employees.Joined AS [Joined], T_Employees.Left AS [Left], T_BenCar.P_Num AS [Personnel Number], T_BenCar.AvailFrom AS [Available from], T_BenCar.AvailTo AS [Available to], T_BenCar.Reg AS Registration, T_BenCar.Replacement, T_BenCar.SecondCar AS SecondCar, T_BenCar.Replaced, T_BenCar.RegReplaced AS [Registration Replaced], T_BenCar.DateCarReplaced, T_BenCar.CarReplaced AS MakeModelReplaced,T_BenCar.RegDate  , trim$([Make] & ' ' & [Model]) AS MakeAndModel, T_BenCar.NumOfUsers As [NumberOfUsers], T_BenCar.Unavail AS [Days unavailable], T_BenCar.FuelAvailTo AS [FuelTo]" & _
                "FROM T_BenCar INNER JOIN T_Employees ON T_BenCar.P_Num = T_Employees.P_Num " & _
                "ORDER BY T_BenCar.P_Num, T_BenCar.AvailFrom, T_BenCar.AvailTo, T_BenCar.Reg;"
    Case SELECT_COMPANYCAR_CHECKER_CARS_NOCARS
        Queries = CCarCheckGapsInCarUsage(CStr(P(0)))
    
   Case SELECT_COMPANYCAR_CHECKER_CARS_REGDATES
      Queries = "SELECT * FROM (SELECT c.Reg AS Data, ee.Surname + ', ' + ee.FirstName AS FULLNAME, ee.Surname, ee.NI, c.P_Num  FROM T_BenCar c INNER JOIN  T_Employees ee ON ee.P_Num = c.P_Num WHERE c.AvailFrom < c.RegDate) ORDER BY " & CStr(P(0))

    Case SELECT_COMPANYCAR_CHECKER_CARS_EEE_AVAIL_DATES
      Queries = "SELECT * FROM (SELECT DISTINCT c.P_Num AS DATA, c.P_Num "
      Queries = Queries & ", e.Surname as SURNAME,  e.NI as NI, e.Surname + ', ' + e.Firstname AS FULLNAME"
      Queries = Queries & "  FROM T_BenCar c INNER JOIN T_Employees e ON e.P_Num = c.P_Num WHERE (e.Joined > c.AvailFrom AND e.Joined <> #01/01/100#)  OR (e.Left < c.AvailTo AND  e.Left <>  #01/01/100#) ) temp ORDER BY " & CStr(P(0))
    Case SELECT_COMPANYCAR_CHECKER_CARS_EEE_AVAIL_DATES_TV
      Queries = "SELECT e.P_Num, c.Reg,  c.AvailFrom, c.AvailTo, " & ConvertUNDATEDDateSQL("e.Joined", False) & " as [EmployeeJoined] ," & ConvertUNDATEDDateSQL("e.Left", True) & " as [EmployeeLeft] FROM T_BenCar c INNER JOIN T_Employees e ON e.P_Num = c.P_Num WHERE ((e.Joined > c.AvailFrom AND e.Joined <> #01/01/100#)  OR (e.Left < c.AvailTo AND  e.Left <>  #01/01/100#)) AND e.P_Num = " & StrSQL(P(0))
    Case SELECT_EMPLOYEE_CHECK_NI_VALID_TV
      Queries = "SELECT P_Num, NI FROM T_Employees WHERE P_Num = " & StrSQL(P(0))
    Case SELECT_EMPLOYEE_CHECK_PNUM_VALID_TV
      Queries = "SELECT P_Num, P_Num as 'PNUM' FROM T_Employees WHERE P_Num = " & StrSQL(P(0))
    Case SELECT_EMPLOYEE_CHECK_INVALID_NAME_TV
      Queries = "SELECT P_Num, FirstName, Surname FROM T_Employees WHERE P_Num = " & StrSQL(P(0))
    Case FIND_COMPANY_CAR
      Queries = "SELECT T_BenCar.Reg, T_BenCar.P_Num From T_BenCar Where T_BenCar.Reg = " & StrSQL(P(i))
    Case FIX_O_OTHER_EXPENSE
      Queries = "UPDATE T_BenOther set Class = 'O Other Expense' where Class ='P Other Expense'"
    Case FIX_SHARES
      Queries = "UPDATE T_BenOther set UDBCode = 'M' WHERE Class = 'Shares'"
    Case FIX_MEDICAL
      Queries = "UPDATE T_BenOther set UDBCode = 'I' WHERE Class = 'Medical'"
    Case FIX_CHAUFFEUR
      Queries = "UPDATE T_BenOther set UDBCode = 'O' WHERE Class = 'Chauffeur'"
    Case FIX_GENERAL
      Queries = "UPDATE T_BenOther set UDBCode = 'O' WHERE Class = 'General'"
    Case FIX_ENTERTAINMENT
      Queries = "UPDATE T_BenOther set UDBCode = 'O' WHERE Class = 'Entertainment'"
    Case FIX_TRAVEL
      Queries = "UPDATE T_BenOther set UDBCode = 'O' WHERE Class = 'Travel and subsistence'"
    Case FIX_TAX_PAID
      Queries = "UPDATE T_BenOther set UDBCode = 'N' WHERE Class = 'Tax paid'"
    Case FIX_EDUCATION
      Queries = "UPDATE T_BenOther set UDBCode = 'N' WHERE Class = 'Education'"
    Case FIX_NURSERY
      Queries = "UPDATE T_BenOther set UDBCode = 'N' WHERE Class = 'Nursery'"
    Case FIX_SUBSCRIPTIONS
      Queries = "UPDATE T_BenOther set UDBCode = 'N' WHERE Class = 'Subscriptions'"
    Case FIX_O_OTHER
      Queries = "UPDATE T_BenOther set UDBCode = 'O' WHERE Class = 'O Other Expense'"
    Case FIX_SERVICES_PROVIDED
      Queries = "UPDATE T_BenGoods set Class = 'Services Provided', UDBCODE = 'K' WHERE AvailFrom < #04/04/1999#" ' Charlie, please check
    Case FIX_ASSETS_DISPOSAL
      Queries = "UPDATE T_BenGoods set Class = 'Assets at disposal', UDBCODE = 'L' WHERE AvailFrom >= #04/04/1999#"
    Case FIX_FREQUENCY_OF_PAYMENTS  'COMPANY CARS
      Queries = "UPDATE T_BenCar set P46PaymentFrequency = 'Actual'"
    Case FIX_MOBILE_PHONES          'RH
      Queries = "DELETE * FROM T_BENPHONE where class = 'MOBILE'"
    Case FIX_SHARES_REMOVE
      Queries = "Delete * from T_BenOther WHERE Class = " & StrSQL(S_SHARES_M)
    Case FIX_UPDATE_GUID
      Queries = "Update id set id.guid = """ & P(i) & """ where guid is null or guid = """""
    Case FIX_BRING_CAR_FUEL_AVAIL_TO_THIS_YEAR
      Queries = "UPDATE T_BenCar set FuelAvailTo = " & DateSQL(p11d32.Rates.value(TaxYearEnd)) & " Where FuelAvailTo < " & DateSQL(p11d32.Rates.value(TaxYearStart))
    Case SELECT_PAYEONLINE_SUBMISSION_ERRORS
        Queries = "Select Message As [Errors] From Submissions Where SubmissionID in "
    Case SELECT_HELPLINKS
        Queries = "SELECT T_HelpForms.*, T_HelpControls.*, T_Tooltips.tooltiptext AS tooltip " & _
                  "FROM T_Tooltips RIGHT JOIN " & _
                  "  (T_HelpForms LEFT JOIN T_HelpControls " & _
                  "   ON T_HelpForms.form_pid = T_HelpControls.form_fid) " & _
                  "ON T_Tooltips.tooltip_pid = T_HelpControls.tooltip_fid "
    Case FIX_M_2004
      Queries = "UPDATE T_BenOther set UDBCode = 'M' where UDBCode ='N'"
    Case FIX_N_2004
      Queries = "UPDATE T_BenOther set UDBCode = 'O' where UDBCode ='N'"
    Case FIX_N_OTHER_2004
      Queries = "UPDATE T_BenOther set Class = 'N Other Expense' where Class ='O Other Expense'"
    Case Else
      Call ECASE("Queries - query " & CStr(QueryID) & "not found")
  End Select
  
  Call xReturn("Queries")
End Function

Function CCarCheckGapsInCarUsage(sOrderBy As String) As String
  
  Dim sql As String
  Dim dtYearEnd As Date
  On Error GoTo err_err
  
  
  
  sql = "SELECT * FROM (SELECT DISTINCT "
  sql = sql & "    e.P_Num AS Data, e.P_Num, "
  sql = sql & "    Surname + ', ' + Firstname AS FULLNAME, Surname , e.NI as NI "
  sql = sql & "     "
  sql = sql & "FROM T_Employees e "
  sql = sql & "INNER JOIN "
  sql = sql & "    ( "
  sql = sql & "    SELECT "
  sql = sql & "        p_num "
  sql = sql & "    FROM "
  sql = sql & "        ( "
  sql = sql & "        SELECT "
  sql = sql & "            car1.reg,"
  sql = sql & "            car1.P_Num , "
  sql = sql & "            car1.AvailFrom, "
  sql = sql & "            car1.AvailFrom,"
  sql = sql & "            car1.AvailTo , "
  sql = sql & "            y.avail_to_1, "
  sql = sql & "            y.reg "
  sql = sql & "        FROM T_BenCar car1 "
  sql = sql & "        LEFT JOIN "
  sql = sql & "            ( "
  sql = sql & "            SELECT "
  sql = sql & "                car2.reg, "
  sql = sql & "                car2.P_Num, "
  sql = sql & "                car2.AvailTo + 1 AS avail_to_1 "
  sql = sql & "            FROM T_BenCar car2 "
  sql = sql & "            INNER JOIN "
  sql = sql & "                T_Employees e "
  sql = sql & "                ON e.P_Num = car2.P_Num "
  sql = sql & "            ) "
  sql = sql & "            y "
  sql = sql & "            ON y.P_Num = car1.P_Num "
  sql = sql & "            AND y.avail_to_1 = car1.AvailFrom "
  sql = sql & "        WHERE car1.Reg NOT IN "
  sql = sql & "            ( "
  sql = sql & "            SELECT "
  sql = sql & "                c1.Reg "
  sql = sql & "            FROM T_BenCar c1 "
  sql = sql & "            INNER JOIN "
  sql = sql & "                T_BenCar c2 "
  sql = sql & "                ON "
  sql = sql & "                ("
  sql = sql & "                    ("
  sql = sql & "                        c1.P_Num = c2.P_Num"
  sql = sql & "                    ) "
  sql = sql & "                    AND "
  sql = sql & "                    ("
  sql = sql & "                        NOT (c1.P_Num =c2.P_Num "
  sql = sql & "                        AND c1.Reg=c2.Reg "
  sql = sql & "                        AND c1.AvailFrom=c2.AvailFrom "
  sql = sql & "                        AND c1.AvailTo=c2.AvailTo)"
  sql = sql & "                    ) "
  sql = sql & "                    AND c2.AvailFrom<=c1.AvailTo "
  sql = sql & "                    AND c2.AvailTo>=c1.AvailTo"
  sql = sql & "                )"
  sql = sql & "            ) "
  sql = sql & "        ) "
  sql = sql & "        temp "
  sql = sql & "    GROUP BY P_Num, "
  sql = sql & "        y.reg "
  sql = sql & "    HAVING isnull(y.reg) "
  sql = sql & "        AND count(*) > 1 "
'  sql = sql & "    UNION "
'  sql = sql & "    SELECT "
'  sql = sql & "        P_Num "
'  sql = sql & "    FROM "
'  sql = sql & "        ("
'  sql = sql & "        SELECT "
'  sql = sql & "            c.P_Num, "
'  sql = sql & "            MAX(c.AvailTo) - MIN(c.AvailFrom) AS test , "
'  sql = sql & "            IIF(e.Left=#1/1/100#, #" & CStr(Format(p11d32.Rates.value(TaxYearEnd), "mm / DD / YY")) & "#, e.Left) - IIF(e.Joined<#" & CStr(Format(p11d32.Rates.value(TaxYearStart), "mm / DD / YY")) & "#, #" & CStr(Format(p11d32.Rates.value(TaxYearStart), "mm / DD / YY")) & "#, e.Joined) AS test1 "
'  sql = sql & "        FROM T_BenCar c "
'  sql = sql & "        INNER JOIN "
'  sql = sql & "            T_Employees e "
'  sql = sql & "            ON e.P_Num = c.P_Num "
'  sql = sql & "        GROUP BY c.P_Num, "
'  sql = sql & "            e.Joined, "
'  sql = sql & "            e.Left"
'  sql = sql & "        ) "
'  sql = sql & "    WHERE test < test1 "
'  sql = sql & "    AND test > 0 "
  sql = sql & "    ) "
  sql = sql & "    r "
  sql = sql & "    ON r.P_Num = e.P_Num) temp "
  sql = sql & "    ORDER BY " & sOrderBy
  
  CCarCheckGapsInCarUsage = sql

err_end:
  Exit Function
err_err:
  Call ErrorMessage(Err.Number, Err, "CCarCheckGapsInCarUsage", "CCarCheckGapsInCarUsage", Err.Description)
End Function




