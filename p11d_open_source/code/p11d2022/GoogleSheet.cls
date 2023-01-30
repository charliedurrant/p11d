VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GoogleSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_sheetJson As Object
Private m_iD As String
Private Const API_KEY As String = "AIzaSyB-fBNm14hANGp3tWKr--2Q4eGmRKat3JY"

Public AccessToken As String
Public Function Initialise(ByVal ID As String) As Boolean
  Dim ret As Boolean
  Dim csvData As String
  Dim Count As Long
  
On Error GoTo err_err:

  m_iD = ID
  AccessToken = ""
  Set m_sheetJson = SheetJson()
  Count = m_sheetJson.Item("sheets").Count()
  Initialise = True
  
err_end:
  Exit Function
err_err:
  m_iD = ""
  Exit Function
  
  
End Function
Private Function SheetIdentifier(Name As String) As String
  Dim i As Long, Count As Long
  Dim Title As String
  Dim ID As String
  Count = m_sheetJson.Item("sheets").Count()
  For i = 1 To Count
    'Print m_sheetJson.item("sheets").item(1).item("properties").item("sheetId")
    Title = m_sheetJson.Item("sheets").Item(i).Item("properties").Item("title")
    If Title = Name Then
      ID = m_sheetJson.Item("sheets").Item(i).Item("properties").Item("sheetId")
      SheetIdentifier = ID
      Exit Function
    End If
  Next
  Call Err.Raise(ERR_TESTING, "SheetId", "Failed to find the sheet id for the sheet named " & Name)
  
  
End Function
Private Function SheetJson() As Object
  Dim url As String
  Dim jsonString As String
  Dim ret As Object
  
On Error GoTo err_err

  url = "https://sheets.googleapis.com/v4/spreadsheets/" & m_iD & "?key=" & API_KEY
  jsonString = UrlContentsAsString(url)
  Set ret = JSON.parse(jsonString)
  
  
  Set SheetJson = ret
err_end:
  Exit Function
err_err:
  Call Err.Raise(ERR_ERROR, "SheetJson", Err.Description)
  
  
End Function
Public Property Get ID() As String
  ID = m_iD
End Property
Private Function SheetAsCsvUrl(ByVal sheetName As String)
  Dim sheetId As String
  
  sheetId = SheetIdentifier(sheetName)
  SheetAsCsvUrl = "https://docs.google.com/spreadsheets/u/1/d/" & Me.ID & "/export?format=csv&id=" & Me.ID & "&gid=" & sheetId & "&key=" & API_KEY
  'SheetAsCsvUrl = "https://docs.google.com/spreadsheets/d/" & Me.id & "/gviz/tq?tqx=out:csv&sheet=" & sheetName
End Function
Public Function SheetAsCsv(ByVal sheetName As String) As String
  Dim csvData As String
  
  If Not SheetAsCsvEx(sheetName, csvData) Then
    Call Err.Raise(ERR_TESTING, "SheetAsCsv", "The sheet does not exist or it is not publically available")
  End If
  SheetAsCsv = csvData
End Function
Public Sub SheetAsCsvToFile(ByVal sheetName As String, ByVal filePath)
  Dim csvData As String
  
  If Not SheetAsCsvEx(sheetName, csvData) Then
    Call Err.Raise(ERR_TESTING, "SheetAsCsv", "The sheet does not exist or it is not publically available")
  End If
  Call TextFileSave(filePath, csvData)
  
End Sub

Private Function SheetAsCsvEx(ByVal sheetName As String, ByRef csvData As String) As Boolean
  Dim url As String
  Dim urlContents As String
  Dim FileName As String
  
  SheetAsCsvEx = False
  url = SheetAsCsvUrl(sheetName)
  csvData = UrlContentsAsString(url)
    
  SheetAsCsvEx = True
  'csvData = urlContents
End Function
Public Function SheetAsRecordSet(ByVal sheetName As String) As ADODB.Recordset
  Dim tempDirectoryRoot As String
  Dim tempCsvFile As String
  Dim tempDirectory As String
  Dim csvData As String
  Dim rs As ADODB.Recordset
  Dim strConn As String
  Dim cn As ADODB.Connection
  
  Call GetTempDirectory(tempDirectoryRoot)
  
  tempDirectory = FullPath(tempDirectoryRoot) & GenerateGUID
  MkDirEx (tempDirectory)
  
  csvData = Me.SheetAsCsv(sheetName)
  tempCsvFile = FullPath(tempDirectory) & "test.csv"
  Call TextFileSave(tempCsvFile, csvData)
    
  strConn = "Driver={Microsoft Text Driver (*.txt; *.csv)};Dbq= " & tempDirectory & ";Extensions=csv;DefaultDir=" & tempDirectoryRoot & ";Extended Properties=TEXT"

  Set cn = New ADODB.Connection
  cn.CursorLocation = adUseClient
  cn.Open strConn

  Set rs = New ADODB.Recordset
  rs.Open "Select * from test.CSV", cn
  
  Set SheetAsRecordSet = rs
  
  'Call RemoveReadOnlyFolder(tempDirectory)
    
End Function
Public Function SheetAsParsedCsv(ByVal sheetName As String) As csvParser
  Dim csvData As String
  Dim csvParser As csvParser
  
  csvData = Me.SheetAsCsv(sheetName)
  Set csvParser = New csvParser
  Call csvParser.Init(csvData, True)
   
  Set SheetAsParsedCsv = csvParser
    
End Function
Public Function SheetAsParsedCsv2(ByVal sheetName As String, ByRef csvData As String) As csvParser
  Dim csvParser As csvParser
  
  csvData = Me.SheetAsCsv(sheetName)
  Set csvParser = New csvParser
  Call csvParser.Init(csvData, True)
   
  Set SheetAsParsedCsv2 = csvParser
    
End Function
Private Function GoogleError(ByVal data As String) As Boolean
  If InStr(1, data, "the file you have requested does not exist") > 0 Then
    GoogleError = True
    Exit Function
  End If
End Function

Private Sub AccessTokenCheck()
  If (Len(Me.AccessToken)) = 0 Then
    Call Err.Raise(ERR_TESTING, "AccessTokenCheck", "Acces token required")
  End If
End Sub
Public Sub UploadCsvToSheet(sheetName As String, csvData As String)
  Dim JSON As String
  Dim url As String
  
  Call AccessTokenCheck
  
  JSON = "{ 'requests': [ { 'updateCells': { 'range': { 'sheetId': 999999 }, 'fields': '*' } }, { 'pasteData': { 'coordinate': { 'sheetId': 999999, 'rowIndex': '0', 'columnIndex': '0' }, 'data': '%CSV_DATA%', 'type': 'PASTE_NORMAL', 'delimiter': ',' } } ] }"
  
  Dim xmlhttp As XMLHTTP60
    
  Set xmlhttp = New XMLHTTP60
  url = "https://sheets.googleapis.com/v4/spreadsheets/" & Me.ID & ":batchUpdate"
  
  Call xmlhttp.Open("POST", url, False)
  xmlhttp.setRequestHeader "Content-Type", "application/json"
  xmlhttp.setRequestHeader "User-Agent", "Firefox 3.6.4"
  xmlhttp.setRequestHeader "Authorization", "Bearer " & Me.AccessToken
  
  Dim sheetId As String
  sheetId = SheetIdentifier(sheetName)
  
  JSON = Replace$(JSON, 999999, sheetId)
  
  'csvData = """a"",""b"""
  
  csvData = Replace$(csvData, "\", "\\")
  csvData = Replace$(csvData, """", "\""")
  csvData = Replace$(csvData, vbCrLf, "\r\n")
  
  'csvData = "\""a\"",b\r\nc,d\""d"
  JSON = Replace$(JSON, "%CSV_DATA%", csvData)
  
  Call xmlhttp.send(JSON)
  
  Dim responseJson As String
  responseJson = xmlhttp.responseText
  
  If InStr(1, responseJson, "spreadsheetid", vbTextCompare) = 0 Then
    Call Err.Raise(ERR_ERROR, "UploadCsvToSheet", "Failed to upload a csv to the sheet " & sheetName)
  End If
  
  
End Sub







