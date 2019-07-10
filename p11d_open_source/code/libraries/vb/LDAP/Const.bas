Attribute VB_Name = "Const"
Option Explicit

Public Enum LDAP_ERRORS
  ERR_ENUMERATEPEOPLE = TCSCLIENT_ERROR 'apf change
  ERR_NOSERVERCONTEXT
End Enum


Public gDBHelp As DBHelper

Public Sub main()
  Set gDBHelp = New DBHelper
  gDBHelp.DatabaseTarget = DB_TARGET_SQLSERVER
End Sub
