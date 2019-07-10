VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SQLQUERIES_TRANSFER"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum QUERY_NAMES_TRANSFER
'MP DB - not used   HOME_LOAN_TRANSFER_DATA = 1
'MP DB - not used   HOME_LOAN_DELETE_DATA
'MP DB - not used   HOME_LOAN_KEYS
  HOME_LOAN_TRANSFER_KEYS = 1
  HOME_LOAN_DELETE_KEYS
  OTHER_LOAN_TRANSFER_DATA
  OTHER_LOAN_DELETE_DATA
  OTHER_LOAN_KEYS
  OTHER_LOAN_TRANSFER_KEYS
  OTHER_LOAN_DELETE_KEYS
  RELOC_TRANSFER_DATA
  RELOC_DELETE_DATA
  RELOC_KEYS
  RELOC_TRANSFER_KEYS
  RELOC_DELETE_KEYS
  COMPANY_CAR_TRANSFER_DATA
  COMPANY_CAR_DELETE_DATA
  COMPANY_CAR_KEYS
'MP DB  COMPANY_CAR_TRASNSFER_KEYScv
'MP DB  COMPANY_CAR_DELETE_KEYS
  EECAR_TRANSFER_DATA
  EECAR_DELETE_DATA
  EECAR_KEYS
  EECAR_TRASNSFER_KEYS
  EECAR_DELETE_KEYS
  GENERAL_TRANSFER
  GENERAL_DELETE
  OTHER_TRANSFER
End Enum


Public Function Queries(QueryID As QUERY_NAMES_TRANSFER, ParamArray P()) As String
  Dim i As Long

  Call xSet("Queries")

   i = LBound(P)
  Select Case QueryID
    Case HOME_LOAN_TRANSFER_KEYS
      Queries = "INSERT INTO T_Loans ( Type, [From], Payment, [Key], P_Num, Item ) IN " & StrSQL(P(i)) & " " & _
                "SELECT DISTINCTROW T_Loans.Type, T_Loans.From, T_Loans.Payment, " & P(i + 1) & " AS Expr1, T_Loans.P_Num, T_Loans.Item " & _
                "From T_Loans " & _
                "Where ((T_Loans.Type = " & StrSQL(S_HOME_LOANS_TYPE) & ") And (T_Loans.Key = " & NumSQL(P(i + 2)) & "))"
    Case HOME_LOAN_DELETE_KEYS
      Queries = "DELETE DISTINCTROW T_Loans.*, T_Loans.Type, T_Loans.Key " & _
                "From T_Loans Where ((T_Loans.Type = " & StrSQL(S_HOME_LOANS_TYPE) & ") And (T_Loans.Key = " & NumSQL(P(i)) & "))"
    Case OTHER_LOAN_TRANSFER_DATA
'MP DB - CheapTaxableLoan was not present in INSERT, present in SELECT
'       Queries = "INSERT INTO T_BenLoan ( P_Num, OLDBALKEY, Item, Interest,MadeGoodIsTaxDeducted, Waived, Curr, DailyOnly, LType, relocation,DidLoanCommenceOnFirstDayOfTaxYear ) IN " & StrSQL(P(i)) & " " & _
'                "SELECT DISTINCTROW T_BenLoan.P_Num, T_BenLoan.BalKey, T_BenLoan.Item, T_BenLoan.Interest,T_BenLoan.MadeGoodIsTaxDeducted,  T_BenLoan.Waived, T_BenLoan.Curr, T_BenLoan.DailyOnly, T_BenLoan.LType, T_BenLoan.relocation ,T_BenLoan.DidLoanCommenceOnFirstDayOfTaxYear, T_BenLoan.CheapTaxableLoan" & _
'                " From T_BenLoan WHERE ((T_BenLoan.P_Num=" & StrSQL(P(i + 1)) & "))"
      Queries = "INSERT INTO T_BenLoan  IN " & StrSQL(P(i)) & " " & _
                "SELECT * " & _
                " From T_BenLoan WHERE ((T_BenLoan.P_Num=" & StrSQL(P(i + 1)) & "))"
    Case OTHER_LOAN_DELETE_DATA
      Queries = "DELETE DISTINCTROW T_BenLoan.* From T_BenLoan WHERE ((T_BenLoan.P_Num=" & StrSQL(P(i)) & "));"
    Case OTHER_LOAN_KEYS
      Queries = "SELECT DISTINCTROW T_BenLoan.OLDBALKEY AS KeyOld, T_BenLoan.BalKey AS KeyNew " & _
                "From T_BenLoan Where ((T_BenLoan.OLDBALKEY > 0))"
    Case OTHER_LOAN_TRANSFER_KEYS
      Queries = "INSERT INTO T_Loans IN " & StrSQL(P(i)) & " " & _
                "SELECT  *" & _
                " From T_Loans WHERE ((T_Loans.P_Num=" & StrSQL(P(i + 1)) & "))"
    Case OTHER_LOAN_DELETE_KEYS
      Queries = "DELETE DISTINCTROW T_Loans.*, T_Loans.Type, T_Loans.Key " & _
                "From T_Loans Where ((T_Loans.Type = " & StrSQL(S_BENEFICIAL_LOANS_TYPE) & ") And (T_Loans.Key = " & NumSQL(P(i)) & "))"


    Case RELOC_TRANSFER_DATA
        Queries = "INSERT INTO T_BenReloc IN " & StrSQL(P(i)) & " " & _
                "SELECT * " & _
                "From T_BenReloc WHERE ((T_BenReloc.P_Num=" & StrSQL(P(i + 1)) & "))"
    Case RELOC_DELETE_DATA
      Queries = "DELETE DISTINCTROW T_BenReloc.* From T_BenReloc WHERE ((T_BenReloc.P_Num=" & StrSQL(P(i)) & "));"
    Case RELOC_KEYS
      Queries = "SELECT DISTINCTROW T_BenReloc.OLDRELOCKEY AS KeyOld, T_BenReloc.RelocKey AS KeyNew " & _
                "From T_BenReloc Where ((T_BenReloc.OLDRELOCKEY > 0))"
    Case RELOC_TRANSFER_KEYS
      Queries = "INSERT INTO T_BenRelocDetails IN " & StrSQL(P(i)) & " " & _
                "SELECT * " & _
                "From T_BenRelocDetails WHERE ((T_BenRelocDetails.P_Num=" & StrSQL(P(i + 1)) & "))"
    Case RELOC_DELETE_KEYS
          Queries = "DELETE DISTINCTROW T_BenRelocDetails.*, T_BenRelocDetails.RelocKey " & _
               "From T_BenRelocDetails Where ((T_BenRelocDetails.RelocKey = " & NumSQL(P(i)) & "))"
    Case COMPANY_CAR_TRANSFER_DATA
'MP DB removed BusMiles, Diesel, BusMilesActual
'            Queries = "INSERT INTO T_BenCar ( P_Num, Reg, Price, Acc, NewAcc, CapContrib, UseContrib, P46PaymentFrequency, MadeGoodIsTaxDeducted,P46FuelType, P46CarbonDioxide,P46NoApprovedCO2Figure, RegDate, AvailFrom, AvailTo, Unavail, Replacement, SecondCar, Replaced, RegReplaced, BusMiles, Diesel, PvtFuel, makeGood, MadeGood, cc, CheapAcc, OLDCARKEY,  BusMilesActual, ForceP46, CarReplaced, DateCarReplaced, Model, Make ,FuelAvailTo,NumOfUsers,HasFuelUnavail ) IN " & StrSQL(P(i)) & " " & _
'                "SELECT DISTINCTROW T_BenCar.P_Num, T_BenCar.Reg, T_BenCar.Price, T_BenCar.Acc, T_BenCar.NewAcc, T_BenCar.CapContrib, T_BenCar.UseContrib, T_BenCar.P46PaymentFrequency, T_BenCar.MadeGoodIsTaxDeducted,T_BenCar.P46FuelType, T_BenCar.P46CarbonDioxide,T_BenCar.P46NoApprovedCO2Figure, T_BenCar.RegDate, T_BenCar.AvailFrom, T_BenCar.AvailTo, T_BenCar.Unavail, T_BenCar.Replacement, T_BenCar.SecondCar, T_BenCar.Replaced,T_BenCar.RegReplaced, T_BenCar.BusMiles, T_BenCar.Diesel, T_BenCar.PvtFuel, T_BenCar.makeGood, T_BenCar.MadeGood, T_BenCar.cc, T_BenCar.CheapAcc, T_BenCar.CarKey, T_BenCar.BusMilesActual, T_BenCar.ForceP46, T_BenCar.CarReplaced, T_BenCar.DateCarReplaced, T_BenCar.Model, T_BenCar.Make,  T_BenCar.FuelAvailTo, T_BenCar.NumOfUsers,  T_BenCar.HasFuelUnavail " & _
'                "From T_BenCar " & _
'                "Where ((T_BenCar.P_Num = " & StrSQL(P(i + 1)) & ")) "
            Queries = "INSERT INTO T_BenCar ( P_Num, Reg, Price, Acc, NewAcc, CapContrib, UseContrib, P46PaymentFrequency, MadeGoodIsTaxDeducted,P46FuelType, P46CarbonDioxide,P46NoApprovedCO2Figure, RegDate, AvailFrom, AvailTo, Unavail, Replacement, SecondCar, Replaced, RegReplaced, PvtFuel, makeGood, MadeGood, cc, CheapAcc, OLDCARKEY,  ForceP46, CarReplaced, DateCarReplaced, Model, Make ,FuelAvailTo,NumOfUsers,HasFuelUnavail ) IN " & StrSQL(P(i)) & " " & _
                "SELECT DISTINCTROW T_BenCar.P_Num, T_BenCar.Reg, T_BenCar.Price, T_BenCar.Acc, T_BenCar.NewAcc, T_BenCar.CapContrib, T_BenCar.UseContrib, T_BenCar.P46PaymentFrequency, T_BenCar.MadeGoodIsTaxDeducted,T_BenCar.P46FuelType, T_BenCar.P46CarbonDioxide,T_BenCar.P46NoApprovedCO2Figure, T_BenCar.RegDate, T_BenCar.AvailFrom, T_BenCar.AvailTo, T_BenCar.Unavail, T_BenCar.Replacement, T_BenCar.SecondCar, T_BenCar.Replaced,T_BenCar.RegReplaced, T_BenCar.PvtFuel, T_BenCar.makeGood, T_BenCar.MadeGood, T_BenCar.cc, T_BenCar.CheapAcc, T_BenCar.CarKey, T_BenCar.ForceP46, T_BenCar.CarReplaced, T_BenCar.DateCarReplaced, T_BenCar.Model, T_BenCar.Make,  T_BenCar.FuelAvailTo, T_BenCar.NumOfUsers,  T_BenCar.HasFuelUnavail " & _
                "From T_BenCar " & _
                "Where ((T_BenCar.P_Num = " & StrSQL(P(i + 1)) & ")) "
    Case COMPANY_CAR_DELETE_DATA
      Queries = "DELETE DISTINCTROW T_BenCar.* From T_BenCar WHERE ((T_BenCar.P_Num=" & StrSQL(P(i)) & "));"
    Case COMPANY_CAR_KEYS
      Queries = "SELECT DISTINCTROW T_BenCar.OLDCARKEY AS KeyOld, T_BenCar.CArKey AS KeyNew " & _
                "From T_BenCar Where ((T_BenCar.OLDCARKEY > 0))"
'MP DB - T_CarMiles deleted
'    Case COMPANY_CAR_TRASNSFER_KEYS
'      Queries = "INSERT INTO T_CarMiles ( Item, Miles, MilesDate, CarKey, P_Num ) IN " & StrSQL(P(i)) & " " & _
'                "SELECT DISTINCTROW T_CarMiles.Item, T_CarMiles.Miles, T_CarMiles.MilesDate, " & P(i + 1) & " AS Expr1, T_CarMiles.P_Num " & _
'                "From T_CarMiles WHERE ((T_CarMiles.CarKey=" & NumSQL(P(i + 2)) & "));"
'MP DB - T_CarMiles deleted
'    Case COMPANY_CAR_DELETE_KEYS
'      Queries = "DELETE DISTINCTROW T_CarMiles.*, T_CarMiles.CarKey " & _
'                "From T_CarMiles Where ((T_CarMiles.CarKey = " & NumSQL(P(i)) & "))"
    Case EECAR_TRANSFER_DATA
         Queries = "INSERT INTO T_BenEeCar  IN " & StrSQL(P(i)) & " " & _
                "SELECT *" & _
                "From T_BenEeCar Where ((T_BenEeCar.P_Num = " & StrSQL(P(i + 1)) & "))"
    Case EECAR_DELETE_DATA
      Queries = "DELETE DISTINCTROW T_BenEeCar.* From T_BenEeCar WHERE ((T_BenEeCar.P_Num=" & StrSQL(P(i)) & "));"
'MP DB ToDo how to handle - OLDEECARKEY will never be >0 as never written to DB
'           EXCEPT via EECAR_TRANSFER_DATA where OLDEECARKEY = T_BenEeCar.EeCarKey
'           AND need to run EECAR_TRASNSFER_KEYS which is currently run after EECAR_KEYS
    Case EECAR_KEYS
      Queries = "SELECT DISTINCTROW T_BenEeCar.OLDEECARKEY AS KeyOld, T_BenEeCar.EECArKey AS KeyNew " & _
                "From T_BenEeCar Where ((T_BenEeCar.OLDEECARKEY > 0 ))"
    Case EECAR_TRASNSFER_KEYS
      Queries = "INSERT INTO T_BenEECarDetails  IN " & StrSQL(P(i)) & " " & _
                "SELECT *" & _
                "From T_BenEECarDetails Where ((T_BenEECarDetails.P_Num = " & StrSQL(P(i + 1)) & "))"
    Case EECAR_DELETE_KEYS
      Queries = "DELETE DISTINCTROW T_BenEECarDetails.*, T_BenEECarDetails.EECarKey " & _
                "From T_BenEECarDetails Where ((T_BenEECarDetails.EECarKey = " & StrSQL(P(i)) & "))"
    Case GENERAL_TRANSFER
      Queries = "INSERT INTO " & P(i) & " IN " & StrSQL(P(i + 1)) & " SELECT DISTINCTROW " & P(i) & ".* " & _
                "From " & P(i) & " WHERE P_Num=" & StrSQL(P(i + 2)) & ";"
    Case GENERAL_DELETE
      Queries = "DELETE DISTINCTROW " & P(i) & ".* " & _
                "From " & P(i) & " WHERE P_Num=" & StrSQL(P(i + 1)) & ";"
    Case OTHER_TRANSFER 'must not transfer CDC
      Queries = "INSERT INTO T_BenOther ( P_Num, UDBCode, UDB, Category, Class, Item, [Value], MadeGood, MadeGoodIsTaxDeducted, [From], To, IRDesc ) IN " & StrSQL(P(i)) & " " & _
                "SELECT DISTINCTROW T_BenOther.P_Num, T_BenOther.UDBCode, T_BenOther.UDB, T_BenOther.Category, T_BenOther.Class, T_BenOther.Item, T_BenOther.Value, T_BenOther.MadeGood, T_BenOther.MadeGoodIsTaxDeducted, T_BenOther.From, T_BenOther.To, T_BenOther.IRDesc " & _
                "From T_BenOther Where ((T_BenOther.P_Num = " & StrSQL(P(i + 1)) & "))"
  End Select

 'Debug.PrintQueries

  Call xReturn("Queries")
End Function

