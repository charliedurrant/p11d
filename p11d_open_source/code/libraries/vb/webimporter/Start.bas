Attribute VB_Name = "Start"
Option Explicit

Public Sub main()
  Debug.Print "App started"
  Set gErrHelp = New ErrHelper
  Set gDBHelper = New DBHelper
  Set gADOHelper = New ADOHelper
  gDBHelper.DatabaseTarget = DB_TARGET_SQLSERVER
End Sub

