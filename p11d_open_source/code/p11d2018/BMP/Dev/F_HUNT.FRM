VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{AF27A9B5-A3F4-11D2-8DB7-00C04FA9DD6F}#1.2#0"; "TCSPROG.OCX"
Begin VB.Form F_Hunt 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Hunt"
   ClientHeight    =   5610
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9525
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5610
   ScaleMode       =   0  'User
   ScaleWidth      =   9525
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdLoadSearch 
      Caption         =   "Load"
      Height          =   375
      Left            =   7560
      TabIndex        =   12
      Top             =   720
      Width           =   1335
   End
   Begin VB.CommandButton cmdSaveSearch 
      Caption         =   "Save"
      Height          =   375
      Left            =   7560
      TabIndex        =   11
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton cmdAddServer 
      Caption         =   "Add Server"
      Height          =   375
      Left            =   5760
      TabIndex        =   10
      Top             =   1320
      Width           =   1335
   End
   Begin VB.TextBox txtAddServer 
      Height          =   285
      Left            =   1320
      TabIndex        =   9
      Text            =   "Text1"
      Top             =   1320
      Width           =   3855
   End
   Begin MSComctlLib.ListView lvwHunt 
      Height          =   3135
      Left            =   0
      TabIndex        =   7
      Top             =   1920
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   5530
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin TCSPROG.TCSProgressBar TCSProgressBar1 
      Height          =   375
      Left            =   0
      TabIndex        =   6
      Top             =   5160
      Width           =   9495
      _cx             =   16748
      _cy             =   661
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
      Max             =   100
      Value           =   50
      BarBackColor    =   -2147483633
      BarForeColor    =   6956042
      Appearance      =   1
      Style           =   0
      CaptionColor    =   0
      CaptionInvertColor=   16777215
      FillStyle       =   0
      FadeFromColor   =   0
      FadeToColor     =   16777215
      Caption         =   ""
      InnerCircle     =   0   'False
      Percentage      =   0
      Skew            =   0
      PictureOffsetTop=   0
      PictureOffsetLeft=   0
      Enabled         =   0   'False
      Increment       =   1
      TextAlignment   =   2
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   5760
      TabIndex        =   5
      Top             =   720
      Width           =   1335
   End
   Begin VB.CommandButton cmdSearchStop 
      Caption         =   "Search"
      Height          =   375
      Left            =   5760
      TabIndex        =   4
      Top             =   120
      Width           =   1335
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1320
      TabIndex        =   3
      Text            =   "Combo1"
      Top             =   720
      Width           =   3855
   End
   Begin VB.DriveListBox dlbHunt 
      Height          =   315
      Left            =   1320
      TabIndex        =   1
      Top             =   120
      Width           =   3855
   End
   Begin VB.Label lblAddServer 
      Caption         =   "Add server"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label lblLook 
      Caption         =   "Look in"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   720
      Width           =   735
   End
   Begin VB.Label lblMap 
      Caption         =   "Map to"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "F_Hunt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdClose_Click()
  Me.Hide
End Sub

Private Sub cmdLoadSearch_Click()
  Call OpenDatabaseList
End Sub

Private Sub cmdSearchStop_Click()
  If m_bButtonStop Then
    hd.bSearchCancelled = True
  Else
    Call InitialiseSearch
  End If
End Sub
Private Sub InitialiseHunt()
     
  On Error GoTo InitialiseHunt_Err
  
  lvwHuntDatabases.listitems.Clear
  lvwHuntDatabases.Refresh
  hd.lTotalNoOfDatabases = 0

  Call FinaliseHunt(hd.HuntDatabases(txtAddServer.Text, S_DIRECTORY_FILE_MASK, Me)) ', p11d32.HuntDatabases))
    
InitialiseHunt_End:
  Exit Sub
  
InitialiseHunt_Err:
  Call ErrorMessage(ERR_ERROR, Err, "InitialiseHunt", "Error Initialising Hunt", "Error Initialising Hunt")
End Sub
Private Sub FinaliseHunt(lDatabaseCount As Long)
  
  On Error GoTo FinaliseHunt_Err
  
  Call ChangeButton
  hd.bHuntCancelled = False
    
  If lDatabaseCount <> L_FOLDER_INVALID Then
    If lDatabaseCount Then
      prg.Caption = hd.lTotalNoOfDatabases & " Databases hunted: " & lDatabaseCount & " Databases found."
    Else
      prg.Caption = hd.lTotalNoOfDatabases & " Databases searched: No Databases found."
      Call NoResults
      lvwHuntDatabases.HideColumnHeaders = True
      lvwHuntDatabases.TabStop = False
      Call lvwHuntDatabases.listitems.Add(, , "No databases hunted. Please Try Another Server.")
     End If
  End If
FinaliseHunt_End:
  Exit Sub
  
FinaliseHunt_Err:
  Call ErrorMessage(ERR_ERROR, Err, "FinaliseHunt", "Finalise Hunt", "Error Finalising Hunt")
  Resume FinaliseHunt_End
  
End Sub
Private Sub OpenDatabaseList()
  Dim fr As TCSFileread
    
  Dim i As Long
  Dim DatabaseItem As ListItem
  Dim lDatabaseCount As Long
  Dim sOpen
  Dim sDatabaseName, sPath As String
  Dim DS As DatabaseSystemObject
  Dim File As TextStream
  
  On Error GoTo OpenDatabaseList_Err
        
  Set DS = New DatabaseSystemObject
  
  sOpen = DatabaseOpenDlg("Open Hunt Results", "Database Hunts(*.fss)|*.fss") ', p11d32.WorkingDrive)
  
  If Len(sOpen) Then
      
    Set File = DS.OpenTextFile(sOpen)
    
    lvwDatabasesHunted.listitems.Clear
    lvwDatabasesHunted.Enabled = False
        
    Do Until Database.AtEndOfStream
      i = i + 1
      Call SplitPath(File.Directory, sPath, sFileName)
      Call IHuntDatabase_Process(sPath, sDatabaseName, i)
    Loop
      
    lvwDatabasesHunted.Enabled = True
    prg.Caption = sOpen & " (" & i & " Databases)"
    
  End If
  
OpenDatabaseList_End:
  If lvwDatabasesHunted.listitems.Count = 0 Then Call NoResults
  Exit Sub
  
OpenDatabaseList_Err:
  Call ErrorMessage(ERR_ERROR, Err, "OpenDatabaseList", "Error Opening Database", "Error Opening Database")
  prg.Caption = "Error Opening Database"
  Resume OpenDatabaseList_End
End Sub

Private Sub ChangeButton()
  If m_bButtonStop Then
    cmdSearchStop.Caption = "Search"
    cmdClose.Enabled = True
    cmdLoadResults.Enabled = True
    cmdSaveResults.Enabled = True
    prg.Indicator = None
    If lvwDatabasesHunted.listitems.Count > 0 Then cmdSaveResults.Enabled = True
  Else
    cmdSearchStop.Caption = "Stop"
    cmdClose.Enabled = False
    cmdLoadResults.Enabled = False
    cmdSaveResults.Enabled = False
    prg.Indicator = ValueOfMax
  End If
  
  m_bButtonStop = Not m_bButtonStop
  txtStartAt.Enabled = Not txtAddServer.Enabled
    
End Sub
Private Sub cmdSaveSearch_Click()
  Call SaveDatabaseList
End Sub

Private Sub Form_Load()
  Set hd = New HuntDatabases
  txtStartAt = p11d32.WorkingDirectory
End Sub

Private Sub Form_Unload(Cancel As Integer)
  If m_bButtonStop Then Cancel = True
End Sub

Public Sub IHuntDatabase_PreNotify(ByVal Path As String, ByVal Count As Long)
  If Count > 0 Then
    prg.Max = Count
    prg.Caption = Path
  End If
End Sub

Private Sub IHuntDatabase_PostNotify()
  prg.value = 0
End Sub

Private Sub IHuntDatabase_Notify(ByVal lTotalNoOfDatabases As Long)
  Call prg.Step
  DoEvents
End Sub


Private Sub IHuntDatabase_Process(ByVal sDirectory As String, ByVal sDatabase As String, ByVal Index As Long)
  Dim DatabaseItem As ListItem
  
  On Error GoTo IHuntDatabase_Process_Err
  
  If Index = 1 Then
    lvwDatabasesHunted.HideColumnHeaders = False
    lvwDatabasesHunted.TabStop = True
  End If
  
  Set DatabaseItem = lvwDatabasesHunted.listitems.Add(Index, , sDirectory)
  DatabaseItem.SubItems(1) = sDatabase
    
IHuntDatabase_Process_End:
  Exit Sub
  
IHuntDatabase_Process_Err:
  Call ErrorMessage(ERR_ERROR, Err, "IHuntDatabase_Process", "IHuntDatabase_Process", "Error Processing Search Results")
  
End Sub

Private Sub SaveDatabaseList()
    
  Dim DatabaseItem As ListItem

  Dim sSaveAs As String
  Dim DS As DatabaseSystemObject
  Dim File As TextStream
  Dim i As Long
    
  On Error GoTo SaveDatabaseList_Err
  
  Set DS = New DatabaseSystemObject
    
  sSaveAs = DatabaseSaveAsDlg("Save Search Results", "Database Searches(*.fss)|*.fss", p11d32.WorkingDirectory)
   
  If Len(sSaveAs) Then
    Set Database = DS.CreateTextFile(sSaveAs, True)
   
    For i = 1 To lvwDatabasesHunted.listitems.Count
      Set DatabaseItem = lvwDatabasesHunted.listitems(i)
      Database.WriteLine (DatabaseItem.Text & DatabaseItem.SubItems(1))
    Next
    
    Database.Close
    prg.Caption = lvwDatabaseHunted.listitems.Count & " filenames saved to " & sSaveAs
    
  End If
  
SaveDatabaseList_End:
  Exit Sub

SaveDatabaseList_Err:
  Call ErrorMessage(ERR_ERROR, Err, "SaveDatabaseAs", "Error Saving Database", "Error Saving Database")
  Resume SaveDatabaseList_End
End Sub

Private Sub OpenDatabaseList()
  Dim fr As TCSFileread
    
  Dim i As Long
  Dim DatabaseItem As ListItem
  Dim lDatabaseCount As Long
  Dim sOpen
  Dim sDatabaseName, sPath As String
  Dim DS As DatabaseSystemObject
  Dim File As TextStream
  
  On Error GoTo OpenDatabaseList_Err
        
  Set DS = New DatabaseSystemObject
  
  sOpen = DatabaseOpenDlg("Open Search Results", "Database Searches(*.fss)|*.fss", p11d32.WorkingDirectory)
  
  If Len(sOpen) Then
      
    Set Database = DS.OpenTextFile(sOpen)
    
    lvwDatabasesHunted.listitems.Clear
    lvwDatabasesHunted.Enabled = False
        
    Do Until File.AtEndOfStream
      i = i + 1
      Call SplitPath(File.ReadLine, sPath, sDatabaseName)
      Call IHuntDatabase_Process(sPath, sDatabaseName, i)
    Loop
      
    lvwDatabasesHunted.Enabled = True
    prg.Caption = sOpen & " (" & i & " databases)"
    
  End If
  
OpenDatabaseList_End:
  If lvwDatabasesHunted.listitems.Count = 0 Then Call NoResults
  Exit Sub
  
OpenDatabaseList_Err:
  Call ErrorMessage(ERR_ERROR, Err, "OpenDatabaseList", "Error Opening Database", "Error Opening Database")
  prg.Caption = "Error Opening Database"
  Resume OpenDatabaseList_End
End Sub

Private Sub lvwDatabasesHunted_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
  lvwDatabasesHunted.Sorted = True
  If lvwDatabasesHunted.SortKey = ColumnHeader.Index - 1 Then
    If lvwDatabasesHunted.SortOrder = lvwAscending Then
      lvwDatabasesHunted.SortOrder = lvwDescending
    Else
      lvwDatabasesHunted.SortOrder = lvwAscending
    End If
  Else
    lvwDatabasesHunted.SortKey = ColumnHeader.Index - 1
    lvwDatabasesHunted.SortOrder = lvwAscending
  End If
End Sub
Private Sub ChangeDirectory(sCDir As String)

  On Error GoTo ChangeDirectory_Err
  
  Call ChDrive(sCDir)
  Call ChDir(sCDir)
  Call ToolBarButton(TBR_REFRESH_EMPLOYERS, True)

ChangeDirectory_End:
  Call Me.Hide
  Exit Sub
  
ChangeDirectory_Err:
  Call ErrorMessage(ERR_ERROR, Err, "ChangeDirectory", "Error Changing Directory", "Error Changing Directory")

End Sub


Private Sub lvwDatabasesHunted_DblClick()
  If DatabaseExists(lvwDatabasesHunted.SelectedItem.Text, True) Then
    Call ChangeDirectory(lvwDatabasesHunted.SelectedItem.Text)
  Else
    Call ErrorMessage(ERR_ERROR, Err, "lvwDatabasesHunted", "HuntDatabases", lvwDatabasesHunted.SelectedItem.Text & " cannot be found")
  End If
End Sub

Private Sub NoResults()
  lvwDatabasesHunted.HideColumnHeaders = True
  cmdSaveResults.Enabled = False
  lvwDatabasesHunted.TabStop = False
End Sub

