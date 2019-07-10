Attribute VB_Name = "const"
Option Explicit

Public Const DEFAULT_DELIMITER As String = ","
Public Const DEFAULT_QUALIFIER As String = """"
Public Const XML_SCHEMA_FILE As String = "\ImportSpec.xsd"
Public Const XML_CONNECTION_NODE As String = "//connection/server"
Public Const XML_TARGETTABLE_NODE As String = "//tables/target_db_name"
Public Const XML_ERRORLOG_NODE As String = "//tables/error_log"
Public Const XML_AUDITLOG_NODE As String = "//tables/audit_log"

Public Const XML_TABLECOLUMN_NODE As String = "//table/column"
Public Const XML_COLUMN_NODE As String = "//columns/column"
Public Const XML_ROW_NODE As String = "//rows"
Public Const XML_HEADER_NODE As String = "//header/column"
Public Const XML_FOOTER_NODE As String = "//footer/column"
Public Const XML_DSN_NODE As String = "//dsn"
Public Const XML_DELIMITER_NODE As String = "//delimiter"
Public Const XML_QUALIFIER_NODE As String = "//qualifier"
Public Const XML_DATE_FROM As String = "//datefrom"
Public Const XML_DATE_TO As String = "//dateto"

Public Const S_QUOT As String = """"

Public Enum WEBIMPORTER_ERRORS
  ERR_IMPORT = vbObjectError + 512
  ERR_NO_SOURCE_FILE
  ERR_GET_XML_FROM_FILE
  ERR_INVALID_SPECIFICATION_FILE_ATTRIBUTES
  ERR_INVALID_SPECIFICATION_FILE
  ERR_IMPORT_SPEC_READ_ERROR
  ERR_IMPORT_DATA
  ERR_INVALID_COLUMN_PROPERTY
  ERR_INVALID_ROW_PROPERTY
  ERR_MISSING_DEST_NAME
  ERR_MISSING_ERRORLOG_NAME
  ERR_MISSING_AUDITLOG_NAME
  ERR_NO_SPECIFICATION_FILE
  ERR_NO_SCHEMA_FILE
  ERR_COLUMN_MISMATCH
  ERR_BUILDING_SQL_STRING
  ERR_IMPORT_FILE_MISSING
  ERR_IMPORT_LOG_ERROR
  ERR_OBJECT_IS_NOTHING
  ERR_VALIDATING_SPEC_WITH_SCHEMA
  ERR_HTMLATTR
  ERR_VALIDATING_DATA_TYPE
  ERR_No_IMPORT_ID_FOUND
  ERR_MISSING_HEADER
  ERR_MISSING_FOOTER
  ERR_INVALID_TRANSACTION
  ERR_APPEND_DATA
  ERR_UNDO_IMPORT
  ERR_INVALID_CONNECTION
  ERR_CONVERTING_DATES
End Enum



Public gWarnings As Boolean
Public gErrHelp As ErrHelper
Public gDBHelper As DBHelper
Public gADOHelper As ADOHelper

