Attribute VB_Name = "Const"
Option Explicit

Public Enum TCSWhereErrors
  ERR_VALID_FAIL = TCSWHERE_ERROR + 1
  ERR_TOO_MANY_CONDITIONS
  ERR_MERGECLAUSES
  ERR_NOTINROOT
  ERR_DELETECONDITION
  ERR_PARSEINTERNALFORMAT
  ERR_LETVALUE
  ERR_CREATECLAUSETREE
  ERR_STACK_POP
  ERR_OPERATOR
  ERR_FINDLHS
  ERR_PARSESQL
  ERR_STACK_ISOBJECT
  ERR_UNMATCHED
End Enum

Public Const CONDITION_SEP As String = ":??:"
Public Const CONDITION_FN As String = "CONDITION("
Public Const VALUE_BEGIN As String = "«" 'Chr$(171)
Public Const VALUE_END As String = "»"   'Chr$(187)

Public Const CONDITION_OR As String = "OR("
Public Const CONDITION_AND As String = "AND("

Public Const CONDITION_SQL_OR As String = "OR"
Public Const CONDITION_SQL_AND As String = "AND"

Public Const CLAUSE_SEP As String = ","
Public Const CLAUSE_BEGIN As String = "("
Public Const CLAUSE_END As String = ")"

