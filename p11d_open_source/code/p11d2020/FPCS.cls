VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FPCS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Bands As ObjectList
Public Name As String
Public Dirty As Boolean
Public Approved As Boolean

Private Sub Class_Initialize()
  Set Bands = New ObjectList
End Sub


'cd apf check
Public Function CheckForDuplicates(FirstColIndexInError As Long, ValidateMessage As String, RowBuf As TrueDBGrid60.RowBuffer, ByVal RowBufRowIndex As Long, ByVal ObjectListIndex As Long) As Boolean
  Dim Band As FPCSBand
  Dim l As Long, lNewCC As Long, lNewMiles As Long
  
  On Error GoTo CheckForDuplicates_Err
  Call xSet("CheckForDuplicates")

  'get the CC and miles
  If ObjectListIndex = -1 Then
    'new line
    If Not IsNull(RowBuf.value(RowBufRowIndex, 2)) Then
      lNewCC = RowBuf.value(RowBufRowIndex, 2)
    End If
    If Not IsNull(RowBuf.value(RowBufRowIndex, 1)) Then
      lNewMiles = RowBuf.value(RowBufRowIndex, 1)
    End If
  Else
    Set Band = Bands(ObjectListIndex)
    If Not IsNull(RowBuf.value(RowBufRowIndex, 2)) Then
      lNewCC = RowBuf.value(RowBufRowIndex, 2)
    Else
      lNewCC = Band.GreaterThanCC
    End If
    If Not IsNull(RowBuf.value(RowBufRowIndex, 1)) Then
      lNewMiles = RowBuf.value(RowBufRowIndex, 1)
    Else
      lNewMiles = Band.GreaterThanMiles
    End If
    Set Band = Nothing
  End If

  For l = 1 To Bands.Count
    If l <> ObjectListIndex Then
      Set Band = Bands(l)
      If Not Band Is Nothing Then
        If lNewCC = Band.GreaterThanCC And lNewMiles = Band.GreaterThanMiles Then
          FirstColIndexInError = 1
          ValidateMessage = "This row has the same the same CC and miles as another band."
          'duplicate
          CheckForDuplicates = True
          Exit For
        End If
      End If
    End If
  Next
  
CheckForDuplicates_End:
  Set Band = Nothing
  Call xReturn("CheckForDuplicates")
  Exit Function

CheckForDuplicates_Err:
  Call ErrorMessage(ERR_ERROR, Err, "CheckForDuplicates", "Check For Duplicates", "Error checking for FPCS entries where the engine size and mileage are the same.")
  Resume CheckForDuplicates_End
  Resume
End Function


Public Sub Kill()
  Dim i As Long
  Dim Band As FPCSBand
  
  On Error GoTo Kill_Err
  Call xSet("Kill")

  For i = 1 To Bands.Count
    Set Band = Bands(i)
    Set Band = Nothing
  Next

Kill_End:
  Call xReturn("Kill")
  Exit Sub

Kill_Err:
  Call ErrorMessage(ERR_ERROR, Err, "Kill", "Kill", "Error killing a FPCS.")
  Resume Kill_End
End Sub

Private Sub Class_Terminate()
 'Debug.Print"FPCS - terminate"
End Sub
