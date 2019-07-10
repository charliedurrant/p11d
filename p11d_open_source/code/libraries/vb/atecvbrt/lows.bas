Attribute VB_Name = "lows"
Option Explicit

Public Function InCollectionEx(ByVal col As Object, ByVal vItem As Variant) As Boolean
  On Error GoTo InCollectionEx_err
  Call col.Item(vItem)
  InCollectionEx = True
  Exit Function
  
InCollectionEx_err:
  InCollectionEx = False
End Function

Public Sub GetAppSettings(AppPath As String, AppExeName As String, AppVersion As String, AppName As String, ByVal vbg As VB.Global)
  Dim p As Long
  Dim AppFile As String
      
  On Error GoTo GetAppSettings_err
  If Len(AppPath) = 0 Then
    If Not vbg Is Nothing Then AppPath = vbg.App.Path
  End If
  If vbg Is Nothing Then
    AppFile = GetModuleName(0, True)
    Dim fh As New FileHelper
    Call fh.SplitPath(GetModuleName(0, True), AppPath, AppExeName)
  Else
    AppPath = vbg.App.Path
    AppExeName = vbg.App.EXEName
  End If
  If Right$(AppPath, 1) = "\" Then AppPath = Left$(AppPath, Len(AppPath) - 1)
  p = InStr(1, AppExeName, ".EXE", vbTextCompare)
  If p > 1 Then AppExeName = Left$(AppExeName, p - 1)
      
  If vbg Is Nothing Then
    AppVersion = VersionQueryMap(AppFile, VQT_FILE_VERSION)
    AppName = AppExeName
  Else
    AppVersion = vbg.App.Major & "." & vbg.App.Minor & "." & vbg.App.Revision
    AppName = vbg.App.Title
  End If
  If (Len(AppVersion) = 0) Or (Len(AppName) = 0) Then Err.Raise ERR_INITIALISE, "GetAppSettings", "Unable to initialise standard library. No AppVersion or AppName found"
  Exit Sub
  
GetAppSettings_err:
  Err.Raise Err.Number, ErrorSourceEx(Err, "GetAppSettings"), Err.Description
End Sub
