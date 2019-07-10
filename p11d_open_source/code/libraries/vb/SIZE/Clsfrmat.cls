VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFormResize"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mcolclsCtrlAttributes As Collection

Private m_InitOK As Boolean

Private mFrm As Form
Attribute mFrm.VB_VarHelpID = -1
Private MDIFrm As Form

Public Enum E_MINSIZE
  [_START_MINSIZE] = 0
  User = 0
  DESIGN
  VGA
  SVGA
  XGA
  [_END_MINSIZE] = XGA
End Enum

Private mMinimums([_START_MINSIZE] To [_END_MINSIZE]) As MIN_SIZES

Private mMinSizeType As E_MINSIZE
Private mMinWidth As Single
Private mMinHeight As Single

Private mDesignHeight As Single
Private mDesignWidth As Single





Private Sub Class_Initialize()
  If Not m_InitOK Then
    Set mcolclsCtrlAttributes = New Collection
    ParseItems(0).Name = "FREE"
    ParseItems(0).ParseItem = L_FREE
    ParseItems(1).Name = "LOCK"
    ParseItems(1).ParseItem = L_LOCK
    ParseItems(2).Name = "EQUALISE"
    ParseItems(2).ParseItem = L_EQUALISE
    ParseItems(3).Name = "EQUALISER"
    ParseItems(3).ParseItem = L_EQUALISERIGHT
    ParseItems(4).Name = "EQUALISEB"
    ParseItems(4).ParseItem = L_EQUALISEBOTTOM
    ParseItems(5).Name = "FREEEQUALISEBR"
    ParseItems(5).ParseItem = L_FREEEQUALISEBOTTOMRIGHT
    ParseItems(6).Name = "FREEEQUALISEB"
    ParseItems(6).ParseItem = L_FREEEQUALISEBOTTOM
    ParseItems(7).Name = "LOCKBR"
    ParseItems(7).ParseItem = L_LOCKBOTTOMRIGHT
    ParseItems(8).Name = "LOCKB"
    ParseItems(8).ParseItem = L_LOCKBOTTOM
    ParseItems(9).Name = "LOCKR"
    ParseItems(9).ParseItem = L_LOCKRIGHT
    ParseItems(10).Name = "MOVEONLY"
    ParseItems(10).ParseItem = L_MOVEONLY
    ParseItems(11).Name = "FONT"
    ParseItems(11).ParseItem = L_FONT
    ParseItems(12).Name = "FREELOCKR"
    ParseItems(12).ParseItem = L_FREELOCKRIGHT
    ParseItems(13).Name = "FREELOCKL"
    ParseItems(13).ParseItem = L_FREELOCKLEFT
    ParseItems(14).Name = "FREELOCKTBL"
    ParseItems(14).ParseItem = L_FREELOCKTOPBOTTOMLEFT
    ParseItems(15).Name = "FREELOCKTBR"
    ParseItems(15).ParseItem = L_FREELOCKTOPBOTTOMRIGHT
    ParseItems(16).Name = "FREELOCKTHL"
    ParseItems(16).ParseItem = L_FREELOCKTOPHEIGHTLEFT
    ParseItems(17).Name = "FREELOCKTHR"
    ParseItems(17).ParseItem = L_FREELOCKTOPHEIGHTRIGHT
    ParseItems(18).Name = "FREELOCKTB"
    ParseItems(18).ParseItem = L_FREELOCKTOPBOTTOM
    ParseItems(19).Name = "LOCKBEQUALISER"
    ParseItems(19).ParseItem = L_LOCKBOTTOMEQUALISERIGHT
    ParseItems(20).Name = "FREELOCKTR"
    ParseItems(20).ParseItem = L_FREELOCKTOPRIGHT
    ParseItems(21).Name = "FREELOCKBR"
    ParseItems(21).ParseItem = L_FREELOCKBOTTOMRIGHT
    ParseItems(22).Name = "FREELOCKT"
    ParseItems(22).ParseItem = L_FREELOCKTOP
    ParseItems(23).Name = "FREELOCKB"
    ParseItems(23).ParseItem = L_FREELOCKBOTTOM
    ParseItems(24).Name = "GRID"
    ParseItems(24).ParseItem = L_GRID
    ParseItems(25).Name = "CENTRE"
    ParseItems(25).ParseItem = L_CENTRE
    ParseItems(26).Name = "SCALEONLY"
    ParseItems(26).ParseItem = L_SCALEONLY
    ParseItems(27).Name = "BUDDY"
    ParseItems(27).ParseItem = L_BUDDY
    ParseItems(28).Name = "BUDDYL"
    ParseItems(28).ParseItem = L_BUDDYLEFT
    ParseItems(29).Name = "BUDDYR"
    ParseItems(29).ParseItem = L_BUDDYRIGHT
    ParseItems(30).Name = "BUDDYEQUALISEBR"
    ParseItems(30).ParseItem = L_BUDDYEQUALISEBOTTOMRIGHT
    m_InitOK = True
  End If
End Sub

' Produses a 1 based array of items
' parses a comma delimeted string into sItems array
Private Function ParseGeneral(Items() As TAG_PARSEITEM, Tag As String, Optional ByVal Delimeter As String = ",") As Long
  Dim ilen As Long, p0 As Long, p1 As Long, i As Long
  
  On Error GoTo ParseGeneral_err
  ilen = Len(Tag)
  If ilen = 0 Then GoTo ParseGeneral_end
  p0 = 1
  Do
    p1 = InStr(p0, Tag, Delimeter, vbBinaryCompare)
    i = i + 1
    ReDim Preserve Items(1 To i)
    If p1 <> 0 Then
      Items(i).Name = Mid$(Tag, p0, p1 - p0)
    Else
      Items(i).Name = Mid$(Tag, p0)
    End If
    Call ParseKey(Items(i))
    If p1 <> 0 Then p0 = p1 + 1
  Loop Until p1 = 0
  ParseGeneral = i
  
ParseGeneral_end:
  Exit Function
  
ParseGeneral_err:
  ParseGeneral = 0
  Resume ParseGeneral_end
End Function

' all valuse in twips
Private Sub InitDimensions(ByVal UserMinWidth As Single, ByVal UserMinHeight As Single)
  On Error GoTo InitDimensions_err
  
  mMinimums(E_MINSIZE.User).MinSize = E_MINSIZE.User
  mMinimums(E_MINSIZE.User).MinHeight = UserMinHeight
  mMinimums(E_MINSIZE.User).MinWidth = UserMinWidth
  
  mMinimums(E_MINSIZE.DESIGN).MinSize = E_MINSIZE.DESIGN
  mMinimums(E_MINSIZE.DESIGN).MinHeight = mDesignHeight
  mMinimums(E_MINSIZE.DESIGN).MinWidth = mDesignWidth
    
  mMinimums(E_MINSIZE.VGA).MinSize = E_MINSIZE.VGA
  mMinimums(E_MINSIZE.VGA).MinHeight = Screen.TwipsPerPixelY * 480
  mMinimums(E_MINSIZE.VGA).MinWidth = Screen.TwipsPerPixelX * 640
      
  mMinimums(E_MINSIZE.SVGA).MinSize = E_MINSIZE.SVGA
  mMinimums(E_MINSIZE.SVGA).MinHeight = Screen.TwipsPerPixelY * 600
  mMinimums(E_MINSIZE.SVGA).MinWidth = Screen.TwipsPerPixelX * 800
    
  mMinimums(E_MINSIZE.XGA).MinSize = E_MINSIZE.XGA
  mMinimums(E_MINSIZE.XGA).MinHeight = Screen.TwipsPerPixelY * 768
  mMinimums(E_MINSIZE.XGA).MinWidth = Screen.TwipsPerPixelX * 1024
    
  Call CheckScreenSize
 
InitDimensions_end:
  Exit Sub
  
InitDimensions_err:

  MsgBox Err.Description, vbInformation, "InitDimensions Error " & Err.Number
  Resume InitDimensions_end
End Sub

Private Function GetMinDimensions()
  Dim i As Long
  'ByVal MinSize As Long, MinWidth As Long, MinHeight As Long
  For i = [_START_MINSIZE] To [_END_MINSIZE]
    If mMinimums(i).MinSize = mMinSizeType Then
      mMinWidth = mMinimums(i).MinWidth
      mMinHeight = mMinimums(i).MinHeight
      Exit For
    End If
  Next i
End Function

Public Function InitResize(ByVal mFrmInit As Object, ByVal DesignHeight As Single, ByVal DesignWidth As Single, Optional ByVal DefaultMinType As E_MINSIZE = E_MINSIZE.VGA, Optional ByVal UserMinWidth As Single, Optional ByVal UserMinHeight As Single, Optional ByVal MDIParent As Object = Nothing) As Long
  Dim ctrl As Control
  Dim cCA As clsCtrlAttribute
  
On Error GoTo InitResize_ERR
  
  Set mFrm = mFrmInit
  If Not (MDIParent Is Nothing) Then Set MDIFrm = MDIParent
  mMinSizeType = DefaultMinType
  
  mDesignHeight = DesignHeight
  mDesignWidth = DesignWidth
  
  Call InitDimensions(UserMinWidth, UserMinHeight)
  Call GetMinDimensions
  Call ClearCollection
  
  For Each ctrl In mFrm.Controls
    Call GetCtrlAttr(ctrl)
  Next

  Call GetBuddies
  InitResize = True
  
InitResize_END:
  Set ctrl = Nothing
  Exit Function
InitResize_ERR:
  InitResize = False
  MsgBox Err.Description, vbInformation, "InitResize Error " & Err.Number
  Resume InitResize_END
End Function
Private Sub GetBuddies()
  Dim cCA As clsCtrlAttribute
  
On Error GoTo GetBuddies_ERR
  
  For Each cCA In mcolclsCtrlAttributes
    With cCA
      If Len(.sBuddy) Then
        If .ResizeAttributes And L_BUDDYLEFT Then
          Call GetBuddyDesignDistance(L_BUDDYLEFT, cCA)
        ElseIf .ResizeAttributes And L_BUDDYRIGHT Then
          Call GetBuddyDesignDistance(L_BUDDYRIGHT, cCA)
        End If
      End If
    End With
  Next
  
GetBuddies_END:
  Set cCA = Nothing
  Exit Sub
GetBuddies_ERR:
  MsgBox "GetBuddies " & Err.Description & " " & Err.Number
  Resume GetBuddies_END
End Sub

Private Sub GetParentWidthHeight(ctrlAttr As clsCtrlAttribute, ParentWidth As Single, ParentHeight As Single)

  On Error GoTo GetParentWidthHeight_ERR
  With ctrlAttr.cMe
    If ctrlAttr.bContainer Then
      ParentWidth = .Container.Width
      ParentHeight = .Container.Height
    Else
      ParentWidth = .Parent.Width
      ParentHeight = .Parent.Height
    End If
  
GetParentWidthHeight_END:
  End With
  Exit Sub
  
GetParentWidthHeight_ERR:
  MsgBox Err.Description, vbInformation, "GetParentWidthHeight Error " & Err.Number
  
  Resume GetParentWidthHeight_END
  Resume
End Sub

Private Function GetParentWidthHeightInit(ctrl As Control, ctrlAttr As clsCtrlAttribute, ParentWidth As Single, ParentHeight As Single)
  'This function gets the ctrls parent in the window collections height
  'and width, some controls ie data controls do not have
  'a container only a parent, unfortunately the data control does not
  'have a hwnd and thus can not use API GetParent which would get the container.
  
  'Thus if scaling is relative to a frame and not then form they willl not scale correctly
  
  'this function differs from getparentwidthheight as this is when initialising
  Dim lContainer As Long, lGettingDimensions As Long
  Dim sKey As String
  
  On Error GoTo GetParentWidthHeightInit_ERR
  
  With ctrl
    lContainer = .Container.Width
    'first referance the container, if error - property not available
    'then resume at parent
    'first find if has container property
NO_CONTAINER:
    If lContainer Then
      If Not (TypeOf .Container Is Form) Then
        Call GetObjectKey(.Container, sKey)
        ctrlAttr.sContainer = sKey
        ParentHeight = .Container.Height
        ParentWidth = .Container.Width
        ctrlAttr.bContainer = True
      Else
        ParentHeight = mDesignHeight
        ParentWidth = mDesignWidth
      End If
    Else
      ParentHeight = mDesignHeight
      ParentWidth = mDesignWidth
    End If
  End With
  
  GetParentWidthHeightInit = True
  
GetParentWidthHeightInit_END:
  Exit Function
GetParentWidthHeightInit_ERR:
  Select Case Err.Number
    Case ERR_NOPROPERTY 'obeject doesnt support method
      If lContainer Then
        GetParentWidthHeightInit = False
        Resume GetParentWidthHeightInit_END
      Else
        Resume NO_CONTAINER
      End If
    Case Else
      GetParentWidthHeightInit = False
      MsgBox Err.Description, vbInformation, "GetParentWidthHeightInit Error " & Err.Number
      Resume GetParentWidthHeightInit_END
  End Select
End Function
Private Sub SetAspectRatios(ctrlAttr As clsCtrlAttribute, sngParentWidth As Single, sngParentHeight As Single, sngCtrlLeft As Single, sngCtrlTop As Single, sngCtrlWidth As Single, sngCtrlHeight As Single)
  On Error GoTo SetAspectRatios_ERR
  
  ctrlAttr.sngXFactor = sngCtrlLeft / sngParentWidth
  ctrlAttr.sngYFactor = sngCtrlTop / sngParentHeight
  ctrlAttr.sngWidthFactor = sngCtrlWidth / sngParentWidth
  ctrlAttr.sngHeightFactor = sngCtrlHeight / sngParentHeight
  ctrlAttr.sngTopDistance = sngCtrlTop
  ctrlAttr.sngBottomDistance = sngParentHeight - ctrlAttr.sngTopDistance - sngCtrlHeight
  ctrlAttr.sngRightDistance = sngParentWidth - (sngCtrlWidth + sngCtrlLeft)
  ctrlAttr.sngLeftDistance = sngCtrlLeft
  ctrlAttr.sngDesignWidth = sngCtrlWidth
  ctrlAttr.sngDesignHeight = sngCtrlHeight
  ctrlAttr.sngFontHeightFactor = ctrlAttr.cMe.Font.Size / sngParentHeight
  ctrlAttr.sngFontWidthFactor = ctrlAttr.cMe.Font.Size / sngParentWidth
  ctrlAttr.sngCntrlOldWidth = sngCtrlWidth
  
SetAspectRatios_END:
  Exit Sub
SetAspectRatios_ERR:
  Select Case Err.Number
    Case ERR_NOARRAY, ERR_NORUNTIME, ERR_NOPROPERTY, ERR_NOTOP
      Resume Next
    Case Else
  End Select
  Resume Next
End Sub
Private Sub GetCtrlDimensions(c As Control, ctrlAttr As clsCtrlAttribute)
  Dim sngCtrlWidth As Single, sngCtrlHeight As Single, sngCtrlTop As Single, sngCtrlLeft As Single
  Dim sngParentHeight As Single, sngParentWidth As Single
  
On Error GoTo GetCtrlDimensions_ERR

    With c
      Call GetParentWidthHeightInit(c, ctrlAttr, sngParentWidth, sngParentHeight)
      
      sngCtrlWidth = .Width
      sngCtrlHeight = .Height
      sngCtrlTop = .Top
      
      If .Left < 0 Then
        sngCtrlLeft = .Left + SNG_TABLESS
      Else
        sngCtrlLeft = .Left
      End If
      Call SetAspectRatios(ctrlAttr, sngParentWidth, sngParentHeight, sngCtrlLeft, c.Top, c.Width, c.Height)
      
      
    End With
    
  
GetCtrlDimensions_END:
  Exit Sub
GetCtrlDimensions_ERR:
  Select Case Err.Number
    Case ERR_NOARRAY, ERR_NORUNTIME, ERR_NOPROPERTY, ERR_NOTOP
      Resume Next
    Case Else
  End Select
  Resume Next
End Sub
Private Sub GetCtrlAttr(ctrl As Control)
  Dim sName As String, sKey As String
  Dim ctrlAttr As New clsCtrlAttribute
  
On Error GoTo GetCtrlAttr_ERR
    
    With ctrl
      
      Set ctrlAttr.cMe = ctrl
      Call GetAttributes(ctrl.Tag, ctrlAttr)
      Call GetCtrlDimensions(ctrl, ctrlAttr)
      ctrlAttr.bCanMove = CanMove(ctrl)
      ctrlAttr.sName = .Name
      Call GetObjectKey(ctrl, sKey)
      
      ctrlAttr.sKey = sKey
    End With
    
    Add ctrlAttr
    
GetCtrlAttr_END:
  Set ctrlAttr = Nothing
  Exit Sub
GetCtrlAttr_ERR:
  Select Case Err.Number
      'not part of an array
      'objects property is not available at run time
      'object does not support this property
    Case ERR_NOARRAY, ERR_NORUNTIME, ERR_NOPROPERTY, ERR_NOTOP
      Resume Next
    Case Else
      
      MsgBox Err.Description, vbInformation, "GetCtrlAttr Error " & Err.Number
      Resume GetCtrlAttr_END
  End Select
End Sub
  
Public Sub ClearCollection()
  Do Until mcolclsCtrlAttributes.Count = 0
    mcolclsCtrlAttributes.Remove 1
  Loop
End Sub

Private Function NewEnum() As IUnknown
  Set NewEnum = mcolclsCtrlAttributes.[_NewEnum]
End Function

Private Sub GetAttributes(ByVal Tag As String, ctrlAttr As clsCtrlAttribute)
  Dim Items() As TAG_PARSEITEM, ICount As Long
  Dim i As Long, j As Long, l As TAG_PARSEITEMS
  Dim ResizeUserAttributes As TAG_PARSEITEMS
  
  On Error GoTo GetAttributes_ERR
  Tag = UCase$(Tag)
  ICount = ParseGeneral(Items, Tag)
  For i = 1 To ICount
    l = 0
    For j = [_PARSE_ITEMMIN] To [_PARSE_ITEMMAX]
      If StrComp(ParseItems(j).Name, Items(i).Name, vbTextCompare) = 0 Then
        l = l + ParseItems(j).ParseItem
        If ParseItems(j).ParseItem = L_BUDDY Then ctrlAttr.sBuddy = Items(i).Value
        Exit For
      End If
    Next j
    ResizeUserAttributes = ResizeUserAttributes + l
  Next i
  If ResizeUserAttributes = 0 Then ResizeUserAttributes = L_FREE
  ctrlAttr.ResizeAttributes = ResizeUserAttributes
  
GetAttributes_END:
  Exit Sub

GetAttributes_ERR:
  Resume GetAttributes_END
End Sub
Private Sub GetBuddyDesignDistance(lResizeType As Long, ctrlAttr As clsCtrlAttribute)
  Dim cBuddy As Control
  Dim cMe As Control
  Dim sngMeLeft As Single, sngMeRight As Single, sngBuddyLeft As Single, sngBuddyRight As Single
  
On Error GoTo GetBuddyDesignDistance_ERR
  
  Set cMe = ctrlAttr.cMe
  Set cBuddy = mcolclsCtrlAttributes(ctrlAttr.sBuddy).cMe
  
  With cMe
    sngMeLeft = .Left
    sngMeRight = .Width + sngMeLeft
  End With
  
  With cBuddy
    sngBuddyLeft = .Left
    sngBuddyRight = .Width + sngBuddyLeft
  End With
  
  Select Case lResizeType&
    Case L_BUDDYLEFT
      ctrlAttr.sngBuddyDesignDistance = sngBuddyLeft - sngMeRight
    Case L_BUDDYRIGHT
      ctrlAttr.sngBuddyDesignDistance = sngMeLeft - sngBuddyRight
  End Select
  
  
  
GetBuddyDesignDistance_END:
  Set cMe = Nothing
  Set cBuddy = Nothing
  Exit Sub
GetBuddyDesignDistance_ERR:
  MsgBox "GetBuddyDesignDistance " & Err.Description & ". You are possibly passing the wrong buddy name to " & ctrlAttr.sKey

  Resume GetBuddyDesignDistance_END
  
End Sub

' get equals pos and split string
Private Sub ParseKey(tp As TAG_PARSEITEM)
  Dim ilen As Long, p As Long
  
  ilen = Len(tp.Name)
  p = InStr(1, tp.Name, "=")
  If p > 0 Then
    tp.Value = Right(tp.Name, ilen - p)
    tp.Name = Left(tp.Name, p - 1)
  End If
End Sub

Private Sub Add(ctrlAttr As clsCtrlAttribute)
  mcolclsCtrlAttributes.Add ctrlAttr, ctrlAttr.sKey
End Sub


Public Function Resize() As Long
  Dim ctrlAttr As clsCtrlAttribute
  
  On Error GoTo Resize_ERR
  'see if less than minimum height as set by user
  
  If (Not mFrm.Visible) And (mFrm.WindowState <> vbMaximized) Then GoTo Resize_END
  If mFrm.WindowState = vbMinimized Then GoTo Resize_END
  ' Take account of on the fly screen res changes
  Call CheckScreenSize
  If mFrm.MDIChild Then
    If MDIFrm Is Nothing Then Call Err.Raise(ERR_RESIZE, "Resize", "The form being resized has its MDIChild property set to true." & vbCrLf & "Please pass the MDIParent as a parameter in InitResize. ")
    If MDIFrm.WindowState = vbMinimized Then GoTo Resize_END
    If mFrm.WindowState = vbMaximized Then
      If MDIFrm.WindowState <> vbMaximized Then
        If MDIFrm.Height < mMinHeight Then MDIFrm.Height = mMinHeight
        If MDIFrm.Width < mMinWidth Then MDIFrm.Width = mMinWidth
      End If
    End If
  Else
    If mFrm.Height < mMinHeight Then mFrm.Height = mMinHeight
    If mFrm.Width < mMinWidth Then mFrm.Width = mMinWidth
  End If
  
  For Each ctrlAttr In mcolclsCtrlAttributes
    Call Process(ctrlAttr)
  Next
  
  For Each ctrlAttr In mcolclsCtrlAttributes
    ctrlAttr.Resized = False
  Next
  
Resize_END:
  Set ctrlAttr = Nothing
  Exit Function
  
Resize_ERR:
  Resize = False
  MsgBox Err.Description, vbInformation, "Resize Error " & Err.Number
  Resume Next
End Function

Public Sub ReDoAspectRatios(sControlKey$)
  Dim cCA As clsCtrlAttribute
  Dim sngParentWidth As Single, sngParentHeight As Single
  
On Error GoTo ReDoAspectRatios_ERR
  
  Set cCA = mcolclsCtrlAttributes(sControlKey$)
  With cCA
     Call GetParentWidthHeight(cCA, sngParentWidth, sngParentHeight)
     Call SetAspectRatios(cCA, sngParentWidth, sngParentHeight, .cMe.Left, .cMe.Top, .cMe.Width, .cMe.Height)
    '.sngXFactor = .cMe.Left / ParentWidth
    '.sngYFactor = .cMe.Top / ParentHeight
    '.sngWidthFactor = .cMe.Width / ParentWidth
    
    '.sngHeightFactor = .cMe.Height / ParentHeight
    '.sngLeftDistance = .cMe.Left
    '.sngRightDistance = ParentWidth - (.cMe.Left + .cMe.Width)
    '.sngTopDistance = .cMe.Top
    
    '.sngBottomDistance = ParentHeight - (.sngTopDistance + .cMe.Height)
ReDoAspectRatios_END:
  End With
  Set cCA = Nothing
  Exit Sub
  
ReDoAspectRatios_ERR:
  MsgBox Err.Description, vbInformation, "ReDoAspectRatios Error " & Err.Number
  Resume ReDoAspectRatios_END
End Sub

Private Sub Process(ctrlAttr As clsCtrlAttribute)
  Dim ResizeAttributes&
  Dim ctrlContainer As clsCtrlAttribute
  Dim sBuddysBuddy$
  
  On Error GoTo Process_Err
  
  With ctrlAttr
    If Not .Resized Then
      If L_BUDDY And .ResizeAttributes Then
        'if it has a buddy then must process it
        sBuddysBuddy = mcolclsCtrlAttributes(.sBuddy).sBuddy
        'check for circular buddy chains
        If StrComp(sBuddysBuddy, ctrlAttr.sName, vbTextCompare) = 0 Then
          MsgBox (ctrlAttr.sName & " and " & .sBuddy & " form a circular buddy chain.")
          GoTo Process_End
        End If
        Call Process(mcolclsCtrlAttributes(.sBuddy))
        
      End If
      If .bContainer Then
        Set ctrlContainer = mcolclsCtrlAttributes(.sContainer)
        If ctrlContainer.Resized = 0 Then
          Call Process(ctrlContainer)
        End If
      End If
      ResizeAttributes& = .ResizeAttributes
      Call ProcessBitFields(ResizeAttributes, ctrlAttr)
      .Resized = True
    End If
  End With
  

  
Process_End:
  Set ctrlContainer = Nothing
  Exit Sub
Process_Err:

  Select Case Err.Number
    'invalid procedure or argument call
    Case 5
      MsgBox Err.Description & ". It is possible the buddy controls name given in tag of " & ctrlAttr.sName & " is incorrect.", vbInformation, "GetCtrlAttr Error " & Err.Number
    Case Else
      MsgBox Err.Description, vbInformation, "GetCtrlAttr Error " & Err.Number
  End Select
  Resume Process_End
  Resume
End Sub

'apf cd control set code should use UDT and Move/Size methods if possible

Public Function IsAdded&(c As Object)
  Dim cCA As New clsCtrlAttribute
  Dim sKey$
  Dim prvC As Control
  
  On Error GoTo IsAdded_ERR
  
  Set prvC = Nothing
  With prvC
    Call GetObjectKey(prvC, sKey)
    cCA.sKey = sKey
    Call Add(cCA)
    mcolclsCtrlAttributes.Remove (sKey)
    IsAdded = False
  End With
  
  
IsAdded_END:
  Set prvC = Nothing
  Set cCA = Nothing
  Exit Function
IsAdded_ERR:
  IsAdded = True
  Select Case Err.Number
    'object not part of an array
    Case ERR_NOARRAY
      sKey = c.Name
    Case Else
      'do not set isadded to false
      Resume IsAdded_END
  End Select
End Function

Public Sub AddCtrl(c As Object)
  Dim sName$
  Dim prvC As Control
On Error GoTo AddCtrl_ERR

  'first test if there is such control
  Set prvC = c
  sName = prvC.Name
  Call GetCtrlAttr(prvC)
  
AddCtrl_END:
  Set prvC = Nothing
  Exit Sub
AddCtrl_ERR:
  MsgBox "The control passed to AddCtrl is not available", vbInformation, "AddCtrl Error " & Err.Number
  Resume AddCtrl_END
End Sub
 
Public Sub RemoveCtrl(c As Object)
  Dim sKey$
  Dim prvC As Control
  
On Error GoTo RemoveCtrl_ERR
  Set prvC = c
  Call GetObjectKey(prvC, sKey)
  mcolclsCtrlAttributes.Remove (sKey)
  Set prvC = Nothing
  
RemoveCtrl_END:
  Exit Sub
RemoveCtrl_ERR:
  MsgBox "The control passed to RemoveCtrl is not part of the resize collection", vbInformation, "RemoveCtrl Error " & Err.Number
  Resume RemoveCtrl_END
End Sub

Private Sub GetObjectKey(o As Object, sKey As String)
  
On Error GoTo GetObjectKey_ERR
  
  With o
    sKey = .Name & "(" & CStr(.Index) & ")"
  End With
  
GetObjectKey_END:
  Exit Sub
GetObjectKey_ERR:
  Select Case Err.Number
    'object not part of an array
    Case ERR_NOARRAY, ERR_NOPROPERTY
      sKey = o.Name
      Resume Next
    Case Else
      MsgBox "Function GetObjectKey, " & Err.Description, "Error " & Err.Number
      Resume GetObjectKey_END
  End Select
End Sub

Private Sub ProcessBitFields(ResizeAttributes&, ctrlAttr As clsCtrlAttribute, Optional lIncreaseInSize&)
  Dim lField As Long
  Dim l As Long, ParentWidth As Single, ParentHeight As Single
    
  On Error GoTo ProcessBitFields_ERR
  
  Call GetParentWidthHeight(ctrlAttr, ParentWidth, ParentHeight)
  'the reason for calling this twice (FONT) is that the first pass
  'accounts for a decrease in form size and the second below an increase
  'If this was not done the re-sized control would prevent
  'the resizing of the font. The second pass is simply in the
  'for next loop below
  
  'ie an decrease in form size
  If L_FONT And ResizeAttributes Then Call ResizeType(L_FONT, ctrlAttr, ParentWidth, ParentHeight, True)
  
  For l = [_PARSE_ITEMMIN] To ([_PARSE_ITEMMAX] - 1)
    lField = 2 ^ l
    If lField And ResizeAttributes Then
      'accounts for increase in form size as element lFontDecrease is false
      If ctrlAttr.bCanMove Then
        Call ResizeType(lField, ctrlAttr, ParentWidth, ParentHeight)
      Else
        Call ResizeTypeNoMoveAvailable(lField, ctrlAttr, ParentWidth, ParentHeight)
      End If
    End If
  Next l
  
  ctrlAttr.Resized = True

ProcessBitFields_END:
  Exit Sub
  
ProcessBitFields_ERR:
  MsgBox Err.Description, vbInformation, "ProcessBitFields Error " & Err.Number
  Resume ProcessBitFields_END
End Sub
  

Private Sub Class_Terminate()
  Set mcolclsCtrlAttributes = Nothing
  Set mFrm = Nothing
End Sub

Private Sub CheckScreenSize()
  Static ScrY As Single
  Static ScrX As Single
  Dim i As Long
  
  If (ScrX <> Screen.Width) Or (ScrY <> Screen.Height) Then
    ScrX = Screen.Width
    ScrY = Screen.Height
    For i = [_START_MINSIZE] To [_END_MINSIZE]
      If mMinimums(i).MinHeight > ScrY Then mMinimums(i).MinHeight = ScrY
      If mMinimums(i).MinWidth > ScrX Then mMinimums(i).MinWidth = ScrX
    Next i
  End If
End Sub

Private Sub ResizeType(ByVal ResizeAttribute As Long, ctrlAttr As clsCtrlAttribute, ByVal ParentWidth As Single, ByVal ParentHeight As Single, Optional ByVal FontDecrease As Long)
  Dim lFontSizeHeight As Long, lFontSizeWidth As Long, lNewFontSize As Long
  Dim sngColFactor As Single, l As Long
  Dim c As Control, InTabControl  As Boolean
  Dim SZNew As SIZE_SNG
  
  'if change here then change in ResizeTypeNoMoveAvaialable below
  
On Error GoTo ResizeType_ERR
  
  With ctrlAttr.cMe
    SZNew.Height = .Height
    SZNew.Width = .Width
    SZNew.Left = .Left
    SZNew.Top = .Top
    
    If SZNew.Left < 0 Then
      SZNew.Left = SZNew.Left + SNG_TABLESS
      InTabControl = True
    End If
    
    Select Case ResizeAttribute
      Case L_BUDDYRIGHT
        Set c = mcolclsCtrlAttributes(ctrlAttr.sBuddy).cMe
        
        SZNew.Left = c.Left + c.Width + ctrlAttr.sngBuddyDesignDistance
        SZNew.Top = c.Top
      Case L_BUDDYLEFT
        Set c = mcolclsCtrlAttributes(ctrlAttr.sBuddy).cMe
        
        SZNew.Left = c.Left - SZNew.Width - ctrlAttr.sngBuddyDesignDistance
        SZNew.Top = c.Top
      Case L_LOCK
      Case L_FREE
        'free to move and change size proportionally
        SZNew.Left = ctrlAttr.sngXFactor * ParentWidth
        SZNew.Top = ctrlAttr.sngYFactor * ParentHeight
        SZNew.Width = ctrlAttr.sngWidthFactor * ParentWidth
        SZNew.Height = ctrlAttr.sngHeightFactor * ParentHeight
      Case L_EQUALISE
        'resizes to keep same distances from form edges
        
        SZNew.Width = ParentWidth - ctrlAttr.sngRightDistance - ctrlAttr.sngLeftDistance
        SZNew.Height = ParentHeight - ctrlAttr.sngBottomDistance - ctrlAttr.sngTopDistance
      Case L_EQUALISERIGHT
        'resizes width keeping top,left same but width changes
        'width is kept same distance from right side as design time
        SZNew.Width = ParentWidth - ctrlAttr.sngRightDistance - ctrlAttr.sngLeftDistance
      Case L_EQUALISEBOTTOM
        'keep height the same
        SZNew.Height = ParentHeight - ctrlAttr.sngBottomDistance - ctrlAttr.sngTopDistance
      Case L_LOCKBOTTOMRIGHT
        SZNew.Left = ParentWidth - ctrlAttr.sngRightDistance - ctrlAttr.sngDesignWidth
        SZNew.Top = ParentHeight - ctrlAttr.sngBottomDistance - ctrlAttr.sngDesignHeight
      Case L_LOCKBOTTOM
        SZNew.Top = ParentHeight - ctrlAttr.sngBottomDistance - ctrlAttr.sngDesignHeight
        
      Case L_LOCKRIGHT
        SZNew.Left = ParentWidth - ctrlAttr.sngRightDistance - ctrlAttr.sngDesignWidth
        
      Case L_FREEEQUALISEBOTTOM
        SZNew.Left = ctrlAttr.sngXFactor * ParentWidth
        SZNew.Top = ctrlAttr.sngYFactor * .Parent.ScaleHeight
        SZNew.Width = ctrlAttr.sngWidthFactor * ParentWidth
        SZNew.Height = ParentHeight - .Top - ctrlAttr.sngBottomDistance
      Case L_FREEEQUALISEBOTTOMRIGHT
        SZNew.Left = ctrlAttr.sngXFactor * ParentWidth
        SZNew.Top = ctrlAttr.sngYFactor * ParentHeight
        SZNew.Height = .ParentHeight - (.Top + ctrlAttr.sngBottomDistance)
      Case L_MOVEONLY
        SZNew.Left = ctrlAttr.sngXFactor * ParentWidth
        SZNew.Top = ctrlAttr.sngYFactor * ParentHeight
      Case L_FONT
        lFontSizeWidth = ParentWidth * ctrlAttr.sngFontWidthFactor
        lFontSizeHeight = ParentHeight * ctrlAttr.sngFontHeightFactor
        'chose the lower one as if reduce width then reduce to this
        'not the unreduced height factor
        If lFontSizeHeight < lFontSizeWidth Then
          lNewFontSize = lFontSizeHeight
        Else
          lNewFontSize = lFontSizeWidth
        End If
        
        If lNewFontSize > .Font.Size Then
          If FontDecrease = 0 Then .Font.Size = lNewFontSize
          'increase in size with second pass from processbitfields
        Else
          If FontDecrease <> 0 Then .Font.Size = lNewFontSize
          'decrease in size with first pass
        End If
      Case L_FREELOCKRIGHT
        SZNew.Top = ctrlAttr.sngYFactor * ParentHeight
        SZNew.Width = ctrlAttr.sngWidthFactor * ParentWidth
        SZNew.Height = ctrlAttr.sngHeightFactor * ParentHeight
        SZNew.Left = ParentWidth - ctrlAttr.sngRightDistance - .Width
      Case L_FREELOCKLEFT
        SZNew.Left = ctrlAttr.sngLeftDistance
        SZNew.Top = ctrlAttr.sngYFactor * ParentHeight
        SZNew.Width = ctrlAttr.sngWidthFactor * ParentWidth
        SZNew.Height = ctrlAttr.sngHeightFactor * ParentHeight
      Case L_FREELOCKTOPBOTTOMLEFT
        SZNew.Width = ctrlAttr.sngWidthFactor * ParentWidth
        SZNew.Height = ParentHeight - ctrlAttr.sngBottomDistance - ctrlAttr.sngTopDistance
      Case L_FREELOCKTOPBOTTOMRIGHT
        SZNew.Height = ParentHeight - ctrlAttr.sngBottomDistance - ctrlAttr.sngTopDistance
        SZNew.Width = ctrlAttr.sngWidthFactor * ParentWidth
        SZNew.Left = ParentWidth - ctrlAttr.sngRightDistance - SZNew.Width
      Case L_FREELOCKTOPHEIGHTLEFT
        SZNew.Width = ctrlAttr.sngWidthFactor * ParentWidth
        
      Case L_FREELOCKTOPHEIGHTRIGHT
        SZNew.Width = ctrlAttr.sngWidthFactor * ParentWidth
        SZNew.Left = ParentWidth - ctrlAttr.sngRightDistance - SZNew.Width
        
      Case L_FREELOCKTOPBOTTOM
        SZNew.Left = ctrlAttr.sngXFactor * ParentWidth
        
        SZNew.Width = ctrlAttr.sngWidthFactor * ParentWidth
        SZNew.Height = ParentHeight - ctrlAttr.sngBottomDistance - ctrlAttr.sngTopDistance
        
      Case L_LOCKBOTTOMEQUALISERIGHT
        SZNew.Top = ParentHeight - ctrlAttr.sngBottomDistance - ctrlAttr.sngDesignHeight
        SZNew.Width = ParentWidth - ctrlAttr.sngRightDistance - ctrlAttr.sngLeftDistance
      Case L_FREELOCKTOPRIGHT
        
        SZNew.Height = ctrlAttr.sngHeightFactor * ParentHeight
        SZNew.Width = ctrlAttr.sngWidthFactor * ParentWidth
        SZNew.Left = ParentWidth - ctrlAttr.sngRightDistance - SZNew.Width
        
      Case L_FREELOCKBOTTOMRIGHT
        SZNew.Height = ctrlAttr.sngHeightFactor * ParentHeight
        SZNew.Top = ParentHeight - ctrlAttr.sngBottomDistance - .Height
        SZNew.Width = ctrlAttr.sngWidthFactor * ParentWidth
        SZNew.Left = ParentWidth - ctrlAttr.sngRightDistance - .Width
        
      Case L_FREELOCKTOP
        SZNew.Left = ctrlAttr.sngXFactor * ParentWidth
        
      Case L_SCALEONLY
        SZNew.Width = ctrlAttr.sngWidthFactor * ParentWidth
        SZNew.Height = ctrlAttr.sngHeightFactor * ParentHeight
        
      Case L_FREELOCKBOTTOM
        SZNew.Left = ctrlAttr.sngXFactor * ParentWidth
        SZNew.Top = ParentHeight - ctrlAttr.sngBottomDistance - .Height
        
      Case L_GRID
        sngColFactor = SZNew.Width / ctrlAttr.sngCntrlOldWidth
        If TypeOf ctrlAttr.cMe Is MSFlexGrid Then
          For l = 0 To .Cols - 1
            .ColWidth(l) = .ColWidth(l) * sngColFactor
          Next l
        End If
        ctrlAttr.sngCntrlOldWidth = SZNew.Width
      Case L_CENTRE
        'assume control is resized at this point
        SZNew.Left = (ParentWidth \ 2) - (SZNew.Width \ 2)
      Case L_BUDDYEQUALISEBOTTOMRIGHT
        SZNew.Width = ParentWidth - .Left - ctrlAttr.sngRightDistance
        SZNew.Height = ParentHeight - .Top - ctrlAttr.sngBottomDistance
      Case Else
        'Resize proportionally  -  default
        SZNew.Width = ctrlAttr.sngWidthFactor * ParentWidth
        SZNew.Height = ctrlAttr.sngHeightFactor * ParentHeight
    End Select
    
    If InTabControl Then
      SZNew.Left = SZNew.Left - SNG_TABLESS
    End If
    
    If ResizeAttribute <> L_FONT Then
      If ctrlAttr.bCanMove Then
        Call .Move(SZNew.Left, SZNew.Top, SZNew.Width, SZNew.Height)
      Else
        .Left = SZNew.Left
        .Top = SZNew.Top
        .Width = SZNew.Width
        .Height = SZNew.Height
      End If
    End If
  End With
ResizeType_END:
  Exit Sub
ResizeType_ERR:
  Resume Next
End Sub
Private Sub ResizeTypeNoMoveAvailable(ByVal ResizeAttribute As Long, ctrlAttr As clsCtrlAttribute, ByVal ParentWidth As Single, ByVal ParentHeight As Single, Optional ByVal FontDecrease As Long)
  Dim lFontSizeHeight As Long, lFontSizeWidth As Long, lNewFontSize As Long
  Dim sngColFactor As Single, l As Long
  Dim c As Control, InTabControl  As Boolean
  
  
  'if change here then chage in ResizeType above
On Error GoTo ResizeType_ERR
  If ctrlAttr.cMe.Left < 0 Then
    ctrlAttr.cMe.Left = ctrlAttr.cMe.Left + SNG_TABLESS
    InTabControl = True
  End If
  With ctrlAttr.cMe
    Select Case ResizeAttribute
      Case L_BUDDYRIGHT
        Set c = mcolclsCtrlAttributes(ctrlAttr.sBuddy).cMe
        .Top = c.Top
        .Left = c.Left + c.Width + ctrlAttr.sngBuddyDesignDistance
      Case L_BUDDYLEFT
        Set c = mcolclsCtrlAttributes(ctrlAttr.sBuddy).cMe
        .Top = c.Top
        .Left = c.Left - .Width - ctrlAttr.sngBuddyDesignDistance
      Case L_LOCK
      Case L_FREE
        'free to move and change size proportionally
        .Left = ctrlAttr.sngXFactor * ParentWidth
        .Top = ctrlAttr.sngYFactor * ParentHeight
        .Width = ctrlAttr.sngWidthFactor * ParentWidth
        .Height = ctrlAttr.sngHeightFactor * ParentHeight
      Case L_EQUALISE
        'resizes to keep same distances from form edges
        .Width = ParentWidth - ctrlAttr.sngRightDistance - ctrlAttr.sngLeftDistance
        .Height = ParentHeight - ctrlAttr.sngBottomDistance - ctrlAttr.sngTopDistance
      Case L_EQUALISERIGHT
        'resizes width keeping top,left same but width changes
        'width is kept same distance from right side as design time
        .Width = ParentWidth - ctrlAttr.sngRightDistance - ctrlAttr.sngLeftDistance
      Case L_EQUALISEBOTTOM
        'keep height the same
        .Height = ParentHeight - ctrlAttr.sngBottomDistance - ctrlAttr.sngTopDistance
      Case L_LOCKBOTTOMRIGHT
        .Left = ParentWidth - ctrlAttr.sngRightDistance - ctrlAttr.sngDesignWidth
        .Top = ParentHeight - ctrlAttr.sngBottomDistance - ctrlAttr.sngDesignHeight
      Case L_LOCKBOTTOM
        .Top = ParentHeight - ctrlAttr.sngBottomDistance - ctrlAttr.sngDesignHeight
      Case L_LOCKRIGHT
        .Left = ParentWidth - ctrlAttr.sngRightDistance - ctrlAttr.sngDesignWidth
      Case L_FREEEQUALISEBOTTOM
        .Left = ctrlAttr.sngXFactor * ParentWidth
        .Top = ctrlAttr.sngYFactor * .Parent.ScaleHeight
        .Width = ctrlAttr.sngWidthFactor * ParentWidth
        .Height = ParentHeight - .Top - ctrlAttr.sngBottomDistance
      Case L_FREEEQUALISEBOTTOMRIGHT
        .Left = ctrlAttr.sngXFactor * ParentWidth
        .Top = ctrlAttr.sngYFactor * ParentHeight
        .Height = ParentHeight - (.Top + ctrlAttr.sngBottomDistance)

      Case L_MOVEONLY
        .Left = ctrlAttr.sngXFactor * ParentWidth
        .Top = ctrlAttr.sngYFactor * ParentHeight
      Case L_FONT
        lFontSizeWidth = ParentWidth * ctrlAttr.sngFontWidthFactor
        lFontSizeHeight = ParentHeight * ctrlAttr.sngFontHeightFactor
        'chose the lower one as if reduce width then reduce to this
        'not the unreduced height factor
        If lFontSizeHeight < lFontSizeWidth Then
          lNewFontSize = lFontSizeHeight
        Else
          lNewFontSize = lFontSizeWidth
        End If
        
        If lNewFontSize > .Font.Size Then
          If FontDecrease = 0 Then .Font.Size = lNewFontSize
          'increase in size with second pass from processbitfields
        Else
          If FontDecrease <> 0 Then .Font.Size = lNewFontSize
          'decrease in size with first pass
        End If
        
      Case L_FREELOCKRIGHT
        .Top = ctrlAttr.sngYFactor * ParentHeight
        .Width = ctrlAttr.sngWidthFactor * ParentWidth
        .Height = ctrlAttr.sngHeightFactor * ParentHeight
        .Left = ParentWidth - ctrlAttr.sngRightDistance - .Width
      Case L_FREELOCKLEFT
        .Top = ctrlAttr.sngYFactor * ParentHeight
        .Width = ctrlAttr.sngWidthFactor * ParentWidth
        .Height = ctrlAttr.sngHeightFactor * ParentHeight
        .Left = ctrlAttr.sngLeftDistance
      Case L_FREELOCKTOPBOTTOMLEFT
        .Width = ctrlAttr.sngWidthFactor * ParentWidth
        .Height = ParentHeight - ctrlAttr.sngBottomDistance - ctrlAttr.sngTopDistance
      Case L_FREELOCKTOPBOTTOMRIGHT
        .Height = ParentHeight - ctrlAttr.sngBottomDistance - ctrlAttr.sngTopDistance
        .Width = ctrlAttr.sngWidthFactor * ParentWidth
        .Left = ParentWidth - ctrlAttr.sngRightDistance - .Width
      Case L_FREELOCKTOPHEIGHTLEFT
        .Width = ctrlAttr.sngWidthFactor * ParentWidth
      Case L_FREELOCKTOPHEIGHTRIGHT
        .Width = ctrlAttr.sngWidthFactor * ParentWidth
        .Left = ParentWidth - ctrlAttr.sngRightDistance - .Width
      Case L_FREELOCKTOPBOTTOM
        .Height = ParentHeight - ctrlAttr.sngBottomDistance - ctrlAttr.sngTopDistance
        .Left = ctrlAttr.sngXFactor * ParentWidth
        .Width = ctrlAttr.sngWidthFactor * ParentWidth
      Case L_LOCKBOTTOMEQUALISERIGHT
        .Top = ParentHeight - ctrlAttr.sngBottomDistance - ctrlAttr.sngDesignHeight
        .Width = ParentWidth - ctrlAttr.sngRightDistance - ctrlAttr.sngLeftDistance
      Case L_FREELOCKTOPRIGHT
        .Height = ctrlAttr.sngHeightFactor * ParentHeight
        .Width = ctrlAttr.sngWidthFactor * ParentWidth
        .Left = ParentWidth - ctrlAttr.sngRightDistance - .Width
      Case L_FREELOCKBOTTOMRIGHT
        .Height = ctrlAttr.sngHeightFactor * ParentHeight
        .Top = ParentHeight - ctrlAttr.sngBottomDistance - .Height
        .Width = ctrlAttr.sngWidthFactor * ParentWidth
        .Left = ParentWidth - ctrlAttr.sngRightDistance - .Width
      Case L_FREELOCKTOP
        .Left = ctrlAttr.sngXFactor * ParentWidth
      Case L_SCALEONLY
        .Width = ctrlAttr.sngWidthFactor * ParentWidth
        .Height = ctrlAttr.sngHeightFactor * ParentHeight
      Case L_FREELOCKBOTTOM
        .Top = ParentHeight - ctrlAttr.sngBottomDistance - .Height
        .Left = ctrlAttr.sngXFactor * ParentWidth
      Case L_GRID
        sngColFactor = .Width / ctrlAttr.sngCntrlOldWidth
        If TypeOf ctrlAttr.cMe Is MSFlexGrid Then
          For l = 0 To .Cols - 1
            .ColWidth(l) = .ColWidth(l) * sngColFactor
          Next l
'        ElseIf TypeOf ctrlAttr.cMe Is DBGrid Then
'         For l = 0 To .Columns.Count - 1
'            .Columns(l).Width = .Columns(l).Width * sngColFactor
'          Next
        End If
        ctrlAttr.sngCntrlOldWidth = .Width
      Case L_CENTRE
        'assume control is resized at this point
        .Left = (ParentWidth \ 2) - (.Width \ 2)
      Case L_BUDDYEQUALISEBOTTOMRIGHT
        .Height = ParentHeight - .Top - ctrlAttr.sngBottomDistance
        .Width = ParentWidth - .Left - ctrlAttr.sngRightDistance
      Case Else
        'Resize proportionally  -  default
        .Width = ctrlAttr.sngWidthFactor * ParentWidth
        .Height = ctrlAttr.sngHeightFactor * ParentHeight
    End Select
  End With
  If InTabControl Then
    ctrlAttr.cMe.Left = ctrlAttr.cMe.Left - SNG_TABLESS
  End If
  
ResizeType_END:
  Set c = Nothing
  Exit Sub
ResizeType_ERR:
  Resume Next
End Sub

Private Function CanMove(c As Control)
  On Error GoTo CanMove_ERR
  Call c.Move(c.Left, c.Top, c.Width, c.Height)
  CanMove = True
CanMove_ERR:
End Function
