VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BalanceItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Public DateFrom As Date
Public Payment As Double
Public Balance As Double
Public DateTo As Date
Public days As Long
Public RegularPayment As Boolean

Public Function Interest(IT As INTERESTRATE_TYPE) As Double
  Interest = RoundN((Balance * RoundN(InterestRate(IT), 4) * (days / p11d32.Rates.value(DaysInYearLeap))))
End Function
Public Function InterestRate(IT As INTERESTRATE_TYPE) As Double
  Dim dblInterestRate As Double
  
  If GetDailyInterestRate(dblInterestRate, IT) Then
    InterestRate = dblInterestRate
  End If
  
End Function
Private Function GetDailyInterestRate(dblDailyInterestRate As Variant, IT As INTERESTRATE_TYPE) As Boolean
  Dim l As Long
  Dim dblSumOfRates As Double
  
  
  On Error GoTo GetDailyInterestRate_Err
  Call xSet("GetDailyInterestRate")
  
  If days > 0 Then
    For l = 0 To days - 1
      dblSumOfRates = dblSumOfRates + p11d32.Rates.GetInterestRate(DateAdd("d", l, DateFrom), IT)
    Next
    dblDailyInterestRate = IRLoanRateAdjustmentDailyInterestRate(dblSumOfRates, days)
    GetDailyInterestRate = True
  Else
    dblDailyInterestRate = -1
  End If

GetDailyInterestRate_End:
  Call xReturn("GetDailyInterestRate")
  Exit Function

GetDailyInterestRate_Err:
  dblDailyInterestRate = -1
  Call ErrorMessage(ERR_ERROR, Err, "GetDailyInterestRate", "Get Daily Interest Rate", "Error obtaining the daily interest rate.")
  Resume GetDailyInterestRate_End
End Function

Private Sub Class_Initialize()
  DateFrom = UNDATED
End Sub

Private Sub Class_Terminate()
 'Debug.Print"Balance Item - terminate"
End Sub
