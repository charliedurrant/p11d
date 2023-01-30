VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestingTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sheet As String
Public Spec As TestingSpec

Public Property Get ImportQueryId() As QUERY_NAMES_IMP
  Select Case Me.Sheet
    Case "employees"
      ImportQueryId = IMP_EMPLOYEES
    Case "company_cars"
      ImportQueryId = IMP_COMPANYCARS
    Case "loans"
      ImportQueryId = IMP_BENEFICIAL_LOANS
    Case "loans_balances"
      ImportQueryId = IMP_BENEFICIAL_LOAN_BALANCES
    Case "relocation"
      ImportQueryId = IMP_RELOCATION_ITEM
    Case "relocation_items"
      ImportQueryId = IMP_RELOCATION_DETAIL
      
    Case Else
      Call Err.Raise(ERR_TESTING, "Testing", "Unrecognosed import sheet:" & Me.Sheet)
  End Select
End Property


