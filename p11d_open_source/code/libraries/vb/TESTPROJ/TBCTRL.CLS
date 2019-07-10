VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ToolBarControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private mTBItems As Collection
Private mTB As MSComctlLib.Toolbar
Private mNormalImages As ImageList
Private mHotImages As ImageList
Private mDisabledImages As ImageList

Public Enum TOOLBAR_IMAGES
  ICON_ADD = 200
  ICON_CANCEL = 210
  ICON_CONFIRM = 220
  ICON_COPY = 230
  ICON_CUT = 240
  ICON_PASTE = 250
  ICON_SUBTRACT = 260
End Enum

Private Sub Class_Initialize()
  Set mTBItems = New Collection
End Sub

Private Sub Class_Terminate()
  Set mTBItems = Nothing
End Sub

Public Sub ExecuteButton(ByVal Key As String)
  Dim tbi As ToolBarImage

  For Each tbi In mTBItems
    If StrComp(tbi.Key, Key) = 0 Then
      Call MsgBox("Execute " & Key)
    End If
  Next tbi
End Sub

Public Sub RefreshToolbar(Optional ByVal ButtonHeight As Long = 32)
  Dim tbi As ToolBarImage

  mTB.Buttons.Clear
  If Not mNormalImages Is Nothing Then mNormalImages.ListImages.Clear
  If Not mHotImages Is Nothing Then mHotImages.ListImages.Clear
  If Not mDisabledImages Is Nothing Then mDisabledImages.ListImages.Clear
  For Each tbi In mTBItems
    Call tbi.AddButtonImage(mTB, mNormalImages, mHotImages, mDisabledImages)
  Next tbi
  mTB.Style = tbrFlat
  mTB.ImageList = mNormalImages
  mTB.HotImageList = mHotImages
  mTB.DisabledImageList = mNormalImages ' DisabledImages
  mTB.Refresh
End Sub

Public Sub Initialise(tb As MSComctlLib.Toolbar, NormalImages As ImageList, HotImages As ImageList, DisabledImages As ImageList)
  Dim tbi As ToolBarImage
  
  On Error GoTo Initialise_err
  Set mTB = tb
  Set mNormalImages = NormalImages
  Set mHotImages = HotImages
  Set mDisabledImages = DisabledImages

  Set tbi = New ToolBarImage
  If tbi.SetToolbarImage("Add", ICON_ADD) Then Call mTBItems.Add(tbi, tbi.Name)
  Set tbi = New ToolBarImage
  If tbi.SetToolbarImage("Cancel", ICON_CANCEL) Then Call mTBItems.Add(tbi, tbi.Name)
  Set tbi = New ToolBarImage
  If tbi.SetToolbarImage("Confirm", ICON_CONFIRM) Then Call mTBItems.Add(tbi, tbi.Name)
  Set tbi = New ToolBarImage
  If tbi.SetToolbarImage("Copy", ICON_COPY) Then Call mTBItems.Add(tbi, tbi.Name)
  Set tbi = New ToolBarImage
  If tbi.SetToolbarImage("Cut", ICON_CUT) Then Call mTBItems.Add(tbi, tbi.Name)
  Set tbi = New ToolBarImage
  If tbi.SetToolbarImage("Paste", ICON_PASTE) Then Call mTBItems.Add(tbi, tbi.Name)
  Set tbi = New ToolBarImage
  If tbi.SetToolbarImage("Subtract", ICON_SUBTRACT) Then Call mTBItems.Add(tbi, tbi.Name)
      
Initialise_end:
  Exit Sub
  
Initialise_err:
  Call ErrorMessage(ERR_ERROR, Err, "Initialise", "Initialise Toolbar", "Unable to initialise toolbar")
  Resume Initialise_end
End Sub
