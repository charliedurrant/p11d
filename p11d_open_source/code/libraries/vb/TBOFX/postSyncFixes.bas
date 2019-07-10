Attribute VB_Name = "postSyncFixes"
Option Explicit

Public Sub ApplyIndividualPostFixes(db As Database, dbTemplate As Database, mlFixLevel As Long, mlPostFixLevel As Long, mlDbVersion As Long, mlDBSubVersion As Long)
  
  On Error GoTo ApplyIndividualPostFixes_err
  
  '*************************************************************************************
  '         VERY IMPORTANT
  '         Make sure that any fixes are already in the template db and the FixLevel is
  '         incremented accordingly
  '*************************************************************************************
  
  Call ProvideFeedback(-1, -1, "Applying post synchronise fixes")
  
  If mlPostFixLevel < 1 Then
    If PostFix1(db, dbTemplate) Then mlPostFixLevel = setPostFixLevel(db, 1)
  End If
  
  If mlPostFixLevel < 2 Then
    If PostFix2(db, dbTemplate) Then mlPostFixLevel = setPostFixLevel(db, 2)
  End If

ApplyIndividualPostFixes_end:
  Exit Sub
  
ApplyIndividualPostFixes_err:
  Err.Raise Err.Number, ErrorSource(Err, "ApplyIndividualPostFixes"), "An error occurred applying a post synchronise database fix." & vbCrLf & vbCrLf & Err.Description
  Resume
End Sub

'Public Function Fix1(db As Database) As Boolean
'  On Error GoTo Fix1_Err:
'  'Fix code goes here
'
'  'End of fix code
'  Fix1 = True
'
'Fix1_end:
'  Exit Function
'
'Fix1_Err:
'  Err.Raise Err.Number, ErrorSource(Err, "Fix1"), "An error occurred running Fix 1" & vbcrlf & vbcrlf & err.Description
'End Function

Public Function PostFix1(db As Database, dbTemplate As Database) As Boolean
  On Error GoTo PostFix1_Err:
  
  Call SyncProcessHelp(db, dbTemplate)
  
  PostFix1 = True

PostFix1_end:
  Exit Function

PostFix1_Err:
  Err.Raise Err.Number, ErrorSource(Err, "PostFix1"), "An error occurred running PostFix 1" & vbCrLf & vbCrLf & Err.Description
End Function

Public Function PostFix2(db As Database, dbTemplate As Database) As Boolean
  On Error GoTo PostFix2_Err:
  
  Call SyncMultiSchedules(db, dbTemplate)
  
  PostFix2 = True

PostFix2_end:
  Exit Function

PostFix2_Err:
  Err.Raise Err.Number, ErrorSource(Err, "PostFix2"), "An error occurred running PostFix 2" & vbCrLf & vbCrLf & Err.Description
End Function

