Attribute VB_Name = "Document"
Option Explicit

Public Sub Reinitialise()
  Call SetCursor
  
  Call gProjects.kill
  Call InitialiseDocs
  Call ProjectsToScreen(frmMain.chkShowCategories.Value = vbChecked)
  Call ClearCursor
End Sub

Public Sub SearchAutoDoc(ByVal sToFind As String, ByVal SM As SEARCH_MODE)
  Dim c As Collection
  Dim fn As FunctionItem
  Dim s As String
  Dim li As ListItem
  Dim lv As ListView
  
  On Error GoTo SearchAutoDoc_ERR
  
  Set lv = frmMain.lvSearchResults
  If Len(sToFind) = 0 Then GoTo SearchAutoDoc_END
  'iterate the collection of functions
  
  
  lv.Enabled = False
  lv.ListItems.Clear
  
  Set c = gProjects.AllFunctions
  For Each fn In c
    Select Case SM
      Case SM_ALL
        s = fn.SearchText
      Case SM_DESCRIPTION
        s = fn.Description & " " & fn.DescriptionLong
      Case SM_NAME
        s = fn.Name
      Case SM_PARAMETERS
        s = fn.ParametersString
    End Select
    If InStr(1, s, sToFind, vbTextCompare) > 0 Then Set li = lv.ListItems.Add(, fn.Key, fn.Name, , IMG_SEARCH_RESULT)
    If fn.Selected Then
      fn.Selected = False
      frmMain.tvw.Nodes(fn.Key).Bold = False
    End If
  Next
  If lv.ListItems.Count > 0 Then Call ListViewSearchClick(lv.ListItems(1))
  
SearchAutoDoc_END:
  lv.Enabled = True
  Exit Sub
SearchAutoDoc_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "SearchAutoDoc", "Search Auto Doc", "Error in searching for " & sToFind & ".")
  Resume SearchAutoDoc_END
  Resume
End Sub
Public Sub SetupConst()
  Const MAX_CATEGORIES As Long = 20
  
  Set gProjects = New Projects
  ReDim CategoryMaps(1 To MAX_CATEGORIES)
  ReDim CategoryValues(1 To MAX_CATEGORIES)
  CategoryMaps(1) = "CLF"
  CategoryValues(1) = "Core Library"
  
  CategoryMaps(2) = "SF"
  CategoryValues(2) = "Static"
  
  CategoryMaps(3) = "EMF"
  CategoryValues(3) = "Errormessage"
  
  CategoryMaps(4) = "FNF"
  CategoryValues(4) = "File/Network"
  
  CategoryMaps(5) = "CF"
  CategoryValues(5) = "Clipboard"
  
  CategoryMaps(6) = "DTF"
  CategoryValues(6) = "Date/Time"
  
  CategoryMaps(7) = "FDF"
  CategoryValues(7) = "Form/Display"
  
  CategoryMaps(8) = "GCF"
  CategoryValues(8) = "General Conversion"
  
  CategoryMaps(9) = "IFF"
  CategoryValues(9) = "Ini files"
  
  CategoryMaps(10) = "DF"
  CategoryValues(10) = "Database"
  
  CategoryMaps(11) = "NF"
  CategoryValues(11) = "Numeric/Math"
  
  CategoryMaps(12) = "MF"
  CategoryValues(12) = "Miscellaneous"
  
  CategoryMaps(13) = "TCSMF"
  CategoryValues(13) = "TCS Menu"
  
  CategoryMaps(14) = "TCSPF"
  CategoryValues(14) = "TCS Password"
  
  CategoryMaps(15) = "SQF"
  CategoryValues(15) = "System"
  
  CategoryMaps(16) = "STF"
  CategoryValues(16) = "String functions"
  
  CategoryMaps(17) = "SORTF"
  CategoryValues(17) = "Sort functions"
  
  'CategoryMaps(4) = "FNF"
  'CategoryValues(4) = "File/Network functions"

End Sub

Public Sub InitialiseDocs()
  Dim cfgFile As String
  Dim irf As New TCSFileread
  Dim p As Long
  Dim nextfile As String, ClassList As String
    
  cfgFile = AppPath & "\" & "SYSTEM.CFG"
  If Not irf.OpenFile(cfgFile) Then Err.Raise ERR_INITDOCS, "InitialiseDocs", "Could not open configuration file " & cfgFile
  Call SetCursor
    Do While irf.GetLine(nextfile)
      ClassList = ""
      p = InStr(nextfile, ";")
      If p > 0 Then
        ClassList = Mid$(nextfile, p + 1)
        nextfile = Left$(nextfile, p - 1)
      End If
      If Not FileExists(nextfile) Then nextfile = AppPath & "\" & nextfile
      nextfile = GetCanonicalPathName(nextfile)
      Call gProjects.Add(nextfile, ClassList)
    Loop
  Call ClearCursor
End Sub


