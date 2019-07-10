Attribute VB_Name = "ScreenControl"
Option Explicit


Public Sub SetupScreen()
  On Error GoTo SetupScreen_ERR
    
  With frmMain.cboSearchWhat
    .AddItem ("All")
    .ItemData(frmMain.cboSearchWhat.ListCount - 1) = SM_ALL
    
    .AddItem ("Description")
    .ItemData(frmMain.cboSearchWhat.ListCount - 1) = SM_DESCRIPTION
    
    .AddItem ("Name")
    .ItemData(frmMain.cboSearchWhat.ListCount - 1) = SM_NAME
    
    .AddItem ("Parameters")
    .ItemData(frmMain.cboSearchWhat.ListCount - 1) = SM_PARAMETERS
    .ListIndex = 0
  End With
  
  
SetupScreen_END:
  Exit Sub
SetupScreen_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "SetupScreen", "Setup Screen", "Error setting up the screen.")
  Resume SetupScreen_END
  
End Sub

Public Sub ProjectsToScreen(ByVal ShowCategories As Boolean)
  Dim SelectNode As Node
  Dim n1 As Node, n2 As Node, n3 As Node, n4 As Node
  Dim s As String, CurrentCategory As String
  Dim i As Long, j As Long, k As Long
  Dim vp As Project, img As TREEVIEW_NODETYPE
  Dim cl As Class
  Dim fi As FunctionItem, errstring As String
  
  On Error GoTo FillScreen_ERR
  frmMain.tvw.Enabled = False
  frmMain.tvw.Nodes.Clear
  frmMain.lvSearchResults.ListItems.Clear
  frmMain.txtSearch.Text = ""
  frmOutPut.txtInfo.Text = ""
  Call gProjects.Sort
  
  For i = 1 To gProjects.Count
    Set vp = gProjects.Item(i)
    Call vp.Sort
    Set n1 = frmMain.tvw.Nodes.Add(, , vp.PathAndFile, vp.Name, IMG_PROJECT)
    n1.Tag = IMG_PROJECT
    n1.Expanded = True
    For j = 1 To vp.Count
      Set cl = vp.Item(j)
      If cl.GlobalNameSpace Then cl.SortByCategory = ShowCategories
      Call cl.Sort
      Set n2 = frmMain.tvw.Nodes.Add(n1, tvwChild, vp.PathAndFile & KEY_SEPARATOR & cl.PathAndFile, cl.Name, IMG_CLASS)
      n2.Tag = IMG_CLASS
      If SelectNode Is Nothing Then Set SelectNode = n2
      If cl.SortByCategory Then
        CurrentCategory = ""
        For k = 1 To cl.Count
          Set fi = cl.Item(k)
          If StrComp(fi.Category, CurrentCategory, vbTextCompare) <> 0 Then
            Set n3 = frmMain.tvw.Nodes.Add(n2, tvwChild, , fi.Category, IMG_CATEGORY)
            n3.Tag = IMG_CATEGORY
            CurrentCategory = fi.Category
          End If
          img = IMG_SUB
          If fi.IsFunction Then
            img = IMG_FUNCTION
          ElseIf fi.PropertyType <> PROPERTY_NONE Then
            img = IMG_PROPERTY
          End If
          Set n4 = frmMain.tvw.Nodes.Add(n3, tvwChild, vp.PathAndFile & KEY_SEPARATOR & cl.PathAndFile & KEY_SEPARATOR & fi.Name, fi.Name, img)
          fi.Key = n4.Key
          n4.Tag = IMG_FUNCTION
        Next k
      Else
        For k = 1 To cl.Count
          Set fi = cl.Item(k)
          If fi.IsFunction Then
            img = IMG_FUNCTION
          ElseIf fi.PropertyType <> PROPERTY_NONE Then
            img = IMG_PROPERTY
          End If
          Set n3 = frmMain.tvw.Nodes.Add(n2, tvwChild, vp.PathAndFile & KEY_SEPARATOR & cl.PathAndFile & KEY_SEPARATOR & fi.Name, fi.Name, img)
          n3.Tag = IMG_FUNCTION
          fi.Key = n3.Key
        Next k
      End If
    Next j
  Next i

FillScreen_END:
  frmMain.tvw.Enabled = True
  If Len(gLastFunctionKey) > 0 Then
    Set n1 = GetLastNode()
    If Not n1 Is Nothing Then Set SelectNode = n1
  End If
  If Not SelectNode Is Nothing Then
    Call NodeClick(SelectNode, False, False)
    frmMain.tvw.SetFocus
  End If
  Exit Sub
  
FillScreen_ERR:
  If Not fi Is Nothing Then errstring = vbCrLf & "Current Node: " & fi.Name
  Call ErrorMessage(ERR_ERROR, Err, "FillScreen", "Fill Screen", "Error placing file to screen." & errstring)
  Resume FillScreen_END
  Resume
End Sub

Private Function GetLastNode() As Node
  If Len(gLastFunctionKey) > 0 Then
    Set GetLastNode = frmMain.tvw.Nodes.Item(gLastFunctionKey)
  End If
End Function

Public Sub ListViewSearchClick(Item As ListItem)
  Call NodeClick(frmMain.tvw.Nodes(Item.Key), False)
End Sub

Public Sub NodeClick(Node As Node, ByVal bFromTree As Boolean, Optional ByVal SetBold As Boolean = True)
  Dim sItem As String, p0 As Long, p1 As Long
  Dim vp As Project, cl As Class, fi As FunctionItem
  Dim FullKey As String, TagType As TREEVIEW_NODETYPE
  Dim ifr As TCSFileread, sline As String, sScreen As String
    
  gLastFunctionKey = ""
  If Node Is Nothing Then Call Err.Raise(ERR_IS_NOTHING, "NodeClick", "NodeClick")
  FullKey = Node.Key
  TagType = Node.Tag
  If TagType = IMG_CATEGORY Then Exit Sub
  p0 = 1
  p1 = InStr(p0, FullKey, KEY_SEPARATOR)
  If p1 = 0 Then
    If TagType <> IMG_PROJECT Then Err.Raise ERR_NODECLICK, "NodeClick", "Expected node of type Project, found node with key " & FullKey
    sItem = Mid$(FullKey, p0)
  Else
    sItem = Mid$(FullKey, p0, p1 - p0)
  End If
  Set vp = gProjects.Item(sItem)
  If p1 > 0 Then
    p0 = p1 + Len(KEY_SEPARATOR)
    p1 = InStr(p0, FullKey, KEY_SEPARATOR)
    If p1 = 0 Then
      If TagType <> IMG_CLASS Then Err.Raise ERR_NODECLICK, "NodeClick", "Expected node of type Class, found node with key " & FullKey
      sItem = Mid$(FullKey, p0)
    Else
      sItem = Mid$(FullKey, p0, p1 - p0)
    End If
    Set cl = vp.Item(sItem)
  End If
  
  If p1 > 0 Then
    p0 = p1 + Len(KEY_SEPARATOR)
    p1 = InStr(p0, FullKey, KEY_SEPARATOR)
    If p1 = 0 Then
      If TagType <> IMG_FUNCTION Then Err.Raise ERR_NODECLICK, "NodeClick", "Expected node of type Function, found node with key " & FullKey
      sItem = Mid$(FullKey, p0)
    Else
      sItem = Mid$(FullKey, p0, p1 - p0)
    End If
    Set fi = cl.Item(sItem)
  End If
  sScreen = ""
  If Not vp Is Nothing Then
    sScreen = sScreen & vp.ScreenText
  End If
  If Not cl Is Nothing Then
    sScreen = sScreen & cl.ScreenText
  End If
  frmOutPut.txtInfo.Text = sScreen
  sScreen = ""
  If Not fi Is Nothing Then
    gLastFunctionKey = fi.Key
    Call fi.InitialScreenOutput(frmOutPut.txtInfo)
    If frmMain.chkViewCode.Value = vbChecked Then
      Set ifr = New TCSFileread
      If Not ifr.OpenFile(cl.PathAndFile) Then Err.Raise ERR_OPEN_FILE, "NodeClick", "Failed to open file " & cl.PathAndFile
      ifr.CurrentPos = fi.FileFunctionStartPos
      Do While ifr.GetLine(sline)
        sScreen = sScreen & sline & vbCrLf
        If (StrComp(LTrim$(sline), S_FUNCTION_END, vbTextCompare) = 0) Or _
           (StrComp(LTrim$(sline), S_SUB_END, vbTextCompare) = 0) Then Exit Do
      Loop
      Call fi.AddScreenOutput(vbCrLf & sScreen, vbBlack, False)
      'frmOutPut.txtInfo.Text = frmOutPut.txtInfo.Text & sScreen
    End If
    Call fi.OutputToScreen(frmOutPut.txtInfo)
    If Not bFromTree Then
      Set frmMain.tvw.SelectedItem = Node
      Node.Bold = SetBold
      Node.Selected = True
      Node.EnsureVisible
      fi.Selected = True
    End If
  End If
End Sub

