Attribute VB_Name = "Const"
Option Explicit

Public Enum ATECWEBREP_ERRORS
  ERR_INVALID_SPEC = TCSCORE_ERROR + 1
  ERR_INCLUDE_STYLE = TCSCORE_ERROR + 2
End Enum

Public Const S_STYLESHEET As String = "Styles\atecwebrep.css"
Public Const S_SCHEMA As String = "reports\atecwebrep.xsd"
Public Const S_SCHEMA_NAMESPACE As String = "atecwebrep"

Public Const S_WIDTH_AUTO As String = "AUTO"
Public Const S_NODATA As String = "There is no relevant data for this report"

Public Const S_CLASS_TABLE As String = "ATEC_TABLE"
Public Const S_CLASS_COL As String = "ATEC_COL"
Public Const S_CLASS_TRH As String = "ATEC_TRH"
Public Const S_CLASS_TRD As String = "ATEC_TRD"
Public Const S_CLASS_TR_TOTALS As String = "ATEC_TR_TOTALS"
Public Const S_CLASS_TD_COUNT As String = "ATEC_TD_COUNT"
Public Const S_CLASS_TD_SUM As String = "ATEC_TD_SUM"
Public Const S_CLASS_TD As String = "ATEC_TD"
Public Const S_CLASS_TD_GROUP As String = "ATEC_TD_GROUP_" ' for grouped values append group number on end eg "ATEC_TDG_2"
Public Const S_CLASS_THEAD As String = "ATEC_THEAD"
Public Const S_CLASS_TH_HEADER As String = "ATEC_TH_HEADER"
Public Const S_CLASS_TH As String = "ATEC_TH"
Public Const S_CLASS_TFOOT As String = "ATEC_TFOOT"
Public Const S_CLASS_TITLE As String = "ATEC_TITLE"

Public Const S_XML_START As String = "<?xml version=""1.0""?>"

Public Const S_CT_HTML As String = "text/HTML"
Public Const S_CT_CSV As String = "text/plain"
