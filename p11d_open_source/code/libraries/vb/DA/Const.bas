Attribute VB_Name = "Const"
Option Explicit
Option Compare Text

Public gOpenDatabases As Collection
Public gDebugMode As Boolean ' default false
Public gNotify As IBaseNotify
Public gCalc As Boolean

Public gShowPopMessages As Boolean

Public gStatusBar As DAStatusBar


Public Enum DAInternalErrors
  ERR_NODATABASES = TCSDA_ERROR + 1
  ERR_DBOPEN
  ERR_NO_CONNECTION
  ERR_NODB
  ERR_ADD_LIST
  ERR_REMOVE_LIST
  ERR_SQL_PARSE
  ERR_DAPARSE
  ERR_DB_LOCKED
  ERR_FIELD_NAME
  ERR_TABLE_NAME
  ERR_LINK_PARENTS
  ERR_NOT_IN_LIST
  ERR_INVALID_QUERY_LOAD
  ERR_NO_FIX_INTERFACE
  ERR_NOTHING
  ERR_OPEN_DATABASE
  ERR_LOAD_QUERY
  ERR_OPEN_DB
  ERR_ACCESS_MODE
  ERR_DIRTY_ALL
  ERR_CLEAN_ALL
  ERR_CREATE_SQL
  ERR_CREATE_QUERY
  ERR_RECALC
  ERR_RECALC_EX
  ERR_LOAD_QUERY_NAME
  ERR_SETUP_SQL_APPEND
  ERR_SETUP_SQL_UNION
  ERR_SETUP_SQL_UPDATE
  ERR_SETUP_SQL_SELECT
End Enum

Public Sub Main()
  Call DLLSetup("", True)
End Sub

