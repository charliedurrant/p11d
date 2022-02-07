VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{AF27A9B5-A3F4-11D2-8DB7-00C04FA9DD6F}#1.2#0"; "TCSPROG.OCX"
Begin VB.Form F_FindFiles 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Find Files"
   ClientHeight    =   4410
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8910
   ControlBox      =   0   'False
   Icon            =   "F_FindFiles.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4410
   ScaleWidth      =   8910
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstfolders 
      Height          =   255
      ItemData        =   "F_FindFiles.frx":030A
      Left            =   50
      List            =   "F_FindFiles.frx":030C
      TabIndex        =   10
      Top             =   3770
      Width           =   7524
   End
   Begin VB.CommandButton cmdAddRemove 
      Caption         =   "Add / Remove"
      Enabled         =   0   'False
      Height          =   252
      Left            =   7600
      TabIndex        =   9
      Top             =   3770
      Width           =   1260
   End
   Begin VB.CommandButton cmdRunStop 
      Caption         =   "&Run"
      Height          =   372
      Left            =   7600
      TabIndex        =   0
      Top             =   250
      Width           =   1250
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   372
      Left            =   7600
      TabIndex        =   1
      Top             =   660
      Width           =   1250
   End
   Begin VB.CommandButton cmdSaveResults 
      Caption         =   "&Save Results"
      Enabled         =   0   'False
      Height          =   372
      Left            =   7600
      TabIndex        =   2
      Top             =   1070
      Width           =   1250
   End
   Begin VB.CommandButton cmdLoadResults 
      Caption         =   "&Load Results"
      Height          =   372
      Left            =   7600
      TabIndex        =   3
      Top             =   1480
      Width           =   1250
   End
   Begin MSComctlLib.ListView lvwFilesFound 
      Height          =   3290
      Left            =   50
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   240
      Width           =   7524
      _ExtentX        =   13282
      _ExtentY        =   5794
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "Path"
         Text            =   "Path"
         Object.Width           =   7479
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Key             =   "Name"
         Text            =   "File Name"
         Object.Width           =   3739
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Employees"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "PAYE Ref."
         Object.Width           =   2540
      EndProperty
   End
   Begin TCSPROG.TCSProgressBar prg 
      Height          =   285
      Left            =   50
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   4080
      Width           =   8820
      _cx             =   15557
      _cy             =   508
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Min             =   0
      Max             =   1
      Value           =   0
      BarBackColor    =   12632256
      BarForeColor    =   8388608
      Appearance      =   1
      Style           =   0
      CaptionColor    =   0
      CaptionInvertColor=   16777215
      FillStyle       =   0
      FadeFromColor   =   0
      FadeToColor     =   16777215
      Caption         =   " "
      InnerCircle     =   0   'False
      Percentage      =   2
      Skew            =   0
      PictureOffsetTop=   0
      PictureOffsetLeft=   0
      Enabled         =   -1  'True
      Increment       =   1
      TextAlignment   =   2
   End
   Begin VB.Frame fraOptions 
      Caption         =   "Options"
      Enabled         =   0   'False
      Height          =   552
      Left            =   7600
      TabIndex        =   4
      Top             =   1890
      Width           =   1284
      Begin VB.CheckBox chkShowSubDirs 
         Caption         =   "S&ub-folders"
         Height          =   240
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   1092
      End
   End
   Begin VB.Label lblSaved 
      Height          =   255
      Left            =   1560
      TabIndex        =   12
      Top             =   45
      Width           =   6015
   End
   Begin VB.Label Label1 
      Caption         =   "Folders to search:"
      Height          =   255
      Left            =   90
      TabIndex        =   11
      Top             =   3550
      Width           =   1815
   End
   Begin VB.Label lblEmployers 
      Caption         =   "Employer files found"
      Height          =   255
      Left            =   90
      TabIndex        =   8
      Top             =   45
      Width           =   1455
   End
End
Attribute VB_Name = "F_FindFiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Implements IFindFile
Private m_bButtonStop As Boolean
Private m_dbfile As Database
Private ff As FindFiles

Private Sub lstFolders_Change()

End Sub

Private Sub chkShowSubDirs_Click()
  p11d32.FindFilesSearchSubDirs = ChkBoxToBool(chkShowSubDirs) 'JN
End Sub

Private Sub cmdAddRemove_Click()
  If F_AddRemoveFindFolders.Start() Then
    Call AddIniFilesToCombo
  End If
    Call AddIniFilesToCombo
End Sub
Private Sub AddIniFilesToCombo()
  Dim vPaths As Variant
  Dim i As Long
  
  lstfolders.Clear
  For i = 1 To GetDelimitedValues(vPaths, p11d32.FindFilesDirList, True, True, ";")
    lstfolders.AddItem (vPaths(i))
  Next
  
  lstfolders.Enabled = lstfolders.ListCount > 0
  
End Sub
Private Sub cmdClose_Click()
  Me.Hide
End Sub

Private Sub cmdLoadResults_Click()
  Call OpenFileList
End Sub
Private Sub cmdRunStop_Click()
  If m_bButtonStop Then
    cmdRunStop.DEFAULT = True
    ff.bSearchCancelled = True
    If lvwFilesFound.listitems.Count > 0 Then cmdSaveResults.Enabled = True
  Else
    cmdAddRemove.Enabled = False
    fraOptions.Enabled = False
    cmdRunStop.DEFAULT = False
    Call InitialiseSearch
  End If
End Sub
Private Sub InitialiseSearch()
  Dim i As Long
   
  On Error GoTo InitialiseSearch_Err
  
  lvwFilesFound.listitems.Clear
  lvwFilesFound.Refresh
 
  ff.lTotalNoOfFiles = 0
  If lstfolders.ListCount = 0 Then
    MsgBox "Select a folder to search", vbCritical + vbOKOnly, "No directory set"
  Else
    For i = 0 To lstfolders.ListCount - 1
      Call ChangeButton
      Call FinaliseSearch(ff.FindFiles(lstfolders.List(i), S_DATABASE_FILE_MASK, Me, p11d32.FindFilesSearchSubDirs), lstfolders.List(i))
    Next i
  End If
  cmdAddRemove.Enabled = True
  fraOptions.Enabled = True
  If F_FindFiles.lvwFilesFound.listitems.Count > 0 Then F_FindFiles.cmdSaveResults.Enabled = True

InitialiseSearch_End:
  Exit Sub
  
InitialiseSearch_Err:
  Call ErrorMessage(ERR_ERROR, Err, "InitialiseSearch", "Error Initialising Search", "Error Initialising Search")
End Sub
Private Sub FinaliseSearch(lFileCount As Long, sDirectory As String)
  
  On Error GoTo FinaliseSearch_Err
  
  Call ChangeButton
  ff.bSearchCancelled = False
    
  If lFileCount <> L_FOLDER_INVALID Then
    If lFileCount Then
      prg.Caption = ff.lTotalNoOfFiles & " files searched: " & lFileCount & " files found."
    Else
      prg.Caption = ff.lTotalNoOfFiles & " files searched: No files found."
      Call NoResults
      lvwFilesFound.HideColumnHeaders = False
      lvwFilesFound.TabStop = False
    End If
  End If
  
FinaliseSearch_End:
  Exit Sub
  
FinaliseSearch_Err:
  Call ErrorMessage(ERR_ERROR, Err, "FinaliseSearch", "Finalise Search", "Error Finalising Search")
  Resume FinaliseSearch_End
  
End Sub

Private Sub cmdSaveResults_Click()
  Call SaveFileList
End Sub

Private Sub Form_Load()
  Dim s As String
  Call p11d32.SavedFilesList(Ini_read)
  
  Set ff = New FindFiles
  cmdAddRemove.Enabled = True
  fraOptions.Enabled = True
  cmdRunStop.DEFAULT = True
  Call AddIniFilesToCombo
End Sub

Private Sub Form_Unload(Cancel As Integer)
  If m_bButtonStop Then Cancel = True
End Sub

Public Sub IFindFile_PreNotify(ByVal Path As String, ByVal Count As Long)
  If Count > 0 Then
    prg.Max = Count
    prg.Caption = Path
  End If
End Sub

Private Sub IFindFile_PostNotify()
  prg.value = 0
  End Sub
  Private Sub IFindFile_Notify(ByVal lTotalNoOfFiles As Long)
  Call prg.Step
  DoEvents
End Sub

Private Sub IFindFile_Process(ByVal sDirectory As String, ByVal sFIle As String, ByVal Index As Long, ByVal lEmployees As Long, ByVal sPAYEref As String)
  Dim FileItem As ListItem
  Dim ben As IBenefitClass
  Dim rs As Recordset
  Dim w As Long
  
  On Error GoTo IFindFile_Process_Err
  
  w = lvwFilesFound.width
  
  If Index = 1 Then
    lvwFilesFound.HideColumnHeaders = False
    lvwFilesFound.ColumnHeaders(1).width = w * 0.45
    lvwFilesFound.ColumnHeaders(2).width = w * 0.2
    lvwFilesFound.ColumnHeaders(3).width = w * 0.15
    lvwFilesFound.ColumnHeaders(4).width = w * 0.19
    lvwFilesFound.TabStop = True
  End If
  
  Set FileItem = lvwFilesFound.listitems.Add(Index, , sDirectory)
  FileItem.SubItems(1) = sFIle
  FileItem.SubItems(2) = lEmployees
  FileItem.SubItems(3) = sPAYEref
  
  If Not m_dbfile Is Nothing Then
    Set m_dbfile = Nothing
  End If
  
  If Right$(sDirectory, 1) = "\" Then
    Set m_dbfile = InitDB(p11d32.P11DWS, sDirectory & sFIle, "P11D Employer Database")
  Else
    Set m_dbfile = InitDB(p11d32.P11DWS, sDirectory & "\" & sFIle, "P11D Employer Database")
  End If
  
  Set rs = m_dbfile.OpenRecordset("T_Employer", dbOpenSnapshot, dbFailOnError)
  If rs.EOF And rs.BOF Then Err.Raise ERR_EMPLOYER_DB, "Initialize", "Unable to open the employer file as there are no records in the employer table of " & ben.value(sFIle)
  FileItem.SubItems(3) = "" & rs.Fields("PAYE").value
  
  Set rs = m_dbfile.OpenRecordset(sql.Queries(SELECT_EMPLOYEES_COUNT))
  If Not (rs.EOF And rs.BOF) Then
    FileItem.SubItems(2) = rs.Fields("Count").value
    If FileItem.SubItems(2) < 0 Then FileItem.SubItems(2) = 0
  Else
    FileItem.SubItems(2) = 0
  End If
  
 
IFindFile_Process_End:
  Exit Sub
  
IFindFile_Process_Err:
  Call ErrorMessage(ERR_ERROR, Err, "IFindFile_Process", "IFindFile_Process", "Error Processing Search Results")
  Resume IFindFile_Process_End
  Resume
End Sub

Private Sub ChangeButton()
  If m_bButtonStop Then
    cmdRunStop.Caption = "&Run"
    cmdClose.Enabled = True
    cmdLoadResults.Enabled = True
    cmdSaveResults.Enabled = False
    chkShowSubDirs.Enabled = True
    lstfolders.Enabled = True
    prg.Indicator = None
    If lvwFilesFound.listitems.Count > 0 Then cmdSaveResults.Enabled = True
  Else
    cmdRunStop.Caption = "&Stop"
    cmdRunStop.DEFAULT = False
    cmdClose.Enabled = False
    cmdLoadResults.Enabled = False
    cmdSaveResults.Enabled = False
    chkShowSubDirs.Enabled = False
    lstfolders.Enabled = False
    prg.Indicator = ValueOfMax
  End If
  
  m_bButtonStop = Not m_bButtonStop
  
    
End Sub

Private Sub SaveFileList()
    
  Dim FileItem As ListItem

  Dim sSaveAs As String
  Dim fs As FileSystemObject
  Dim File As TextStream
  Dim i As Long
    
  On Error GoTo SaveFileList_Err
  
  Set fs = New FileSystemObject
    
  sSaveAs = FileSaveAsDlg("Save Search Results", "File Searches(*.fss)|*.fss", p11d32.WorkingDirectory)
   
  If Len(sSaveAs) Then
    Set File = fs.CreateTextFile(sSaveAs, True)
   
    For i = 1 To lvwFilesFound.listitems.Count
      Set FileItem = lvwFilesFound.listitems(i)
      File.WriteLine FileItem.Text & "\" & ";" & FileItem.SubItems(1) & ";" & FileItem.SubItems(2) & ";" & FileItem.SubItems(3)
    Next
    
    File.Close
    prg.Caption = lvwFilesFound.listitems.Count & " filenames saved to " & sSaveAs
    lblSaved.Caption = "(saved to " & sSaveAs & ")"
    p11d32.FindFilesSavedFilesList = sSaveAs
    Call p11d32.SavedFilesList(Ini_Write)
  End If
  
SaveFileList_End:
  Exit Sub

SaveFileList_Err:
  Call ErrorMessage(ERR_ERROR, Err, "SaveFileAs", "Error Saving File", "Error Saving File")
  Resume SaveFileList_End
End Sub

Private Sub OpenFileList()
  Dim fr As TCSFileread
  Dim i As Long
  Dim FileItem As ListItem
  Dim lFileCount As Long
  Dim sOpen
  Dim sFileName, sPath As String
  Dim fs As FileSystemObject
  Dim File As TextStream
  Dim lEmployeesCount As Long
  Dim lEmployees As Long
  Dim sPAYEref As String
  Dim s As String
  Dim v As Variant
  Dim w As Long
  
  On Error GoTo OpenFileList_Err
        
  Set fs = New FileSystemObject
  w = lvwFilesFound.width
  sOpen = FileOpenDlg("Open Search Results", "File Searches(*.fss)|*.fss", p11d32.WorkingDirectory)
  
  If Len(sOpen) Then
      
    Set File = fs.OpenTextFile(sOpen)
    lvwFilesFound.listitems.Clear
    lvwFilesFound.Enabled = False
    lvwFilesFound.HideColumnHeaders = False
    lvwFilesFound.ColumnHeaders(1).width = w * 0.45
    lvwFilesFound.ColumnHeaders(2).width = w * 0.2
    lvwFilesFound.ColumnHeaders(3).width = w * 0.15
    lvwFilesFound.ColumnHeaders(4).width = w * 0.19
      
    Do Until File.AtEndOfStream
      i = i + 1
      s = File.ReadLine
      v = Split(s, ";")
      Set FileItem = lvwFilesFound.listitems.Add(, , v(0))
      Call FileItem.ListSubItems.Add(, , v(1))
      Call FileItem.ListSubItems.Add(, , v(2))
      Call FileItem.ListSubItems.Add(, , v(3))
    Loop
      
    lvwFilesFound.Enabled = True
    prg.Caption = sOpen & " (" & i & " files)"
    
  End If
  
OpenFileList_End:
  If lvwFilesFound.listitems.Count = 0 Then Call NoResults
  Exit Sub
  
OpenFileList_Err:
  Call ErrorMessage(ERR_ERROR, Err, "OpenFileList", "Error Opening File", "Error Opening File")
  prg.Caption = "Error Opening File"
  Resume OpenFileList_End
End Sub

Private Sub lvwFilesFound_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
  lvwFilesFound.Sorted = True
  If lvwFilesFound.SortKey = ColumnHeader.Index - 1 Then
    If lvwFilesFound.SortOrder = lvwAscending Then
      lvwFilesFound.SortOrder = lvwDescending
    Else
      lvwFilesFound.SortOrder = lvwAscending
    End If
  Else
    lvwFilesFound.SortKey = ColumnHeader.Index - 1
    lvwFilesFound.SortOrder = lvwAscending
  End If
End Sub
Private Sub ChangeDirectory(sCDir As String)

  On Error GoTo ChangeDirectory_Err
  p11d32.WorkingDirectory = FullPath(sCDir)
  Call ChDriveUNC(sCDir)
  Call ChDir(sCDir)
  Call ToolBarButton(TBR_REFRESH_EMPLOYERS, True)

ChangeDirectory_End:
  Call Me.Hide
  Exit Sub
  
ChangeDirectory_Err:
  Call ErrorMessage(ERR_ERROR, Err, "ChangeDirectory", "Error Changing Directory", "Error Changing Directory")

End Sub
Private Sub lvwFilesFound_DblClick()
  If FileExists(lvwFilesFound.SelectedItem.Text, True) Then
    Call ChangeDirectory(lvwFilesFound.SelectedItem.Text)
  Else
    Call ErrorMessage(ERR_ERROR, Err, "lvwFilesFound", "FindFiles", lvwFilesFound.SelectedItem.Text & " cannot be found")
  End If
End Sub

Private Sub NoResults()
  Dim w As Long
  w = lvwFilesFound.width
  lvwFilesFound.HideColumnHeaders = False
  lvwFilesFound.ColumnHeaders(1).width = w * 0.45
  lvwFilesFound.ColumnHeaders(2).width = w * 0.2
  lvwFilesFound.ColumnHeaders(3).width = w * 0.15
  lvwFilesFound.ColumnHeaders(4).width = w * 0.19
  cmdSaveResults.Enabled = False
  lvwFilesFound.TabStop = False
End Sub


