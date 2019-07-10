Attribute VB_Name = "SetVersions"
Option Explicit

Public Function setDBVersion(db As Database, lNewVersionNumber As Long) As Long
  Dim rs As Recordset
  
  On Error GoTo setDBVersion_err
  
  Set rs = db.OpenRecordset("SELECT * FROM sys_Control", dbOpenDynaset, dbFailOnError)
  rs.FindFirst ("Flag=" & StrSQL("TPDBVersion"))
  If rs.NoMatch Then Err.Raise ERR_APPLY_FIXES, "initVars", "The internal version for the file cannot be found."
  rs.Edit
    rs!State = lNewVersionNumber
  rs.Update
  setDBVersion = lNewVersionNumber
  
setDBVersion_end:
  Set rs = Nothing
  Exit Function
  
setDBVersion_err:
  Err.Raise Err.Number, ErrorSource(Err, "setDBVersion"), "Error setting the database version."
End Function

Public Function setDBSubVersion(db As Database, lNewVersionNumber As Long) As Long
  Dim rs As Recordset
  
  On Error GoTo setDBSubVersion_err
  
  Set rs = db.OpenRecordset("SELECT * FROM sys_Control", dbOpenDynaset, dbFailOnError)
  rs.FindFirst ("Flag=" & StrSQL("TPDBSubVersion"))
  If rs.NoMatch Then Err.Raise ERR_APPLY_FIXES, "initVars", "The internal sub version for the file cannot be found."
  rs.Edit
    rs!State = lNewVersionNumber
  rs.Update
  setDBSubVersion = lNewVersionNumber
  
setDBSubVersion_end:
  Set rs = Nothing
  Exit Function
  
setDBSubVersion_err:
  Err.Raise Err.Number, ErrorSource(Err, "setDBSubVersion"), "Error setting the database version."
End Function

Public Function setFixLevel(db As Database, lNewFixLevel As Long) As Long
  Dim rs As Recordset
  
  On Error GoTo setFixLevel_err
  
  Set rs = db.OpenRecordset("SELECT * FROM sys_Control", dbOpenDynaset, dbFailOnError)
  rs.FindFirst ("Flag=" & StrSQL("TPFixLevel"))
  If rs.NoMatch Then Err.Raise ERR_APPLY_FIXES, "initVars", "The fix level for the file cannot be found."
  rs.Edit
    rs!State = lNewFixLevel
  rs.Update
  setFixLevel = lNewFixLevel
  
setFixLevel_end:
  Set rs = Nothing
  Exit Function
  
setFixLevel_err:
  Err.Raise Err.Number, ErrorSource(Err, "setFixLevel"), "Error setting the fix level."
End Function

Public Function setPostFixLevel(db As Database, lNewFixLevel As Long) As Long
  Dim rs As Recordset
  
  On Error GoTo setPostFixLevel_err
  
  Set rs = db.OpenRecordset("SELECT * FROM sys_Control", dbOpenDynaset, dbFailOnError)
  rs.FindFirst ("Flag=" & StrSQL("TPPostFixLevel"))
  If rs.NoMatch Then Err.Raise ERR_APPLY_FIXES, "initVars", "The post synchronise fix level for the file cannot be found."
  rs.Edit
    rs!State = lNewFixLevel
  rs.Update
  setPostFixLevel = lNewFixLevel
  
setPostFixLevel_end:
  Set rs = Nothing
  Exit Function
  
setPostFixLevel_err:
  Err.Raise Err.Number, ErrorSource(Err, "setPostFixLevel"), "Error setting the post synchronise fix level."
End Function

