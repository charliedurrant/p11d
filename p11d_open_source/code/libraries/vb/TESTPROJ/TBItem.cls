VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ToolBarImage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private mName As String
Private mGroup As Long
Private mKey As String
Private mResIndex As Long
Private mImgNormal As IPictureDisp
Private mImgHot As IPictureDisp
Private mImgDisabled As IPictureDisp

Private Sub Class_Initialize()
  mResIndex = -1
End Sub

Private Sub Class_Terminate()
  On Error Resume Next
  Set mImgNormal = Nothing
  Set mImgHot = Nothing
  Set mImgDisabled = Nothing
End Sub

Public Property Get Name() As String
  Name = mName
End Property

Public Property Get Key() As String
  Key = mKey
End Property

Public Sub AddButtonImage(tb As Toolbar, NormalImages As ImageList, HotImages As ImageList, DisabledImages As ImageList)
  If mResIndex < 0 Then Err.Raise ERR_ADDBUTTONIMAGE, "AddButtonImage", "Unable add image as Item is not a valid toolbar image"
  If tb Is Nothing Then Err.Raise ERR_ADDBUTTONIMAGE, "AddButtonImage", "Unable add image as toolbar is not valid"
  If Not NormalImages Is Nothing Then Call NormalImages.ListImages.Add(, mKey, mImgNormal)
  If Not HotImages Is Nothing Then Call HotImages.ListImages.Add(, mKey, mImgHot)
  If Not DisabledImages Is Nothing Then Call DisabledImages.ListImages.Add(, mKey, mImgDisabled)
  Call tb.Buttons.Add(, mKey, mName, tbrDefault, mKey)
End Sub


'* Images Groups in resource file are offset 10 apart beginning at 200
'*       and in the order Normal, Hot, Disabled
Public Function SetToolbarImage(ByVal Name As String, ByVal ResourceIndex As TOOLBAR_IMAGES, Optional ByVal KeyName As Variant, Optional ByVal GroupID As Long = 0) As Boolean
  On Error GoTo SetToolbarImage_err
  mName = Name
  If IsMissing(KeyName) Then
    mKey = mName
  Else
    mKey = KeyName
  End If
  mGroup = GroupID
  mResIndex = ResourceIndex
  Set mImgNormal = LoadResPicture(mResIndex, vbResIcon)
  Set mImgHot = LoadResPicture(mResIndex + 1, vbResIcon)
  'Set mImgDisabled = LoadResPicture(mResIndex + 2, vbResIcon)
  SetToolbarImage = True
  
SetToolbarImage_end:
  Exit Function
  
SetToolbarImage_err:
  SetToolbarImage = False: mResIndex = -1
  Resume SetToolbarImage_end
End Function



