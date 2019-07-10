Attribute VB_Name = "TrVChkBox"
'Option Explicit
'
'Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
'Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
'Private Declare Function SendMessageAny Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Any, lParam As Any) As Long
'Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
'Private Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
'Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
'
'Private Const IIL_UNCHECKED As Long = 1
'Private Const IIL_CHECKED As Long = 2
'Private Const GWL_STYLE As Long = -16
'
'Private Const TVS_CHECKBOXES  As Long = &H100  '>= IE3
'Private Const TV_FIRST      As Long = &H1100
'Private Const TVM_GETITEM   As Long = (TV_FIRST + 12)
'Private Const TVM_SETITEM   As Long = (TV_FIRST + 13)
'Private Const TVM_HITTEST   As Long = (TV_FIRST + 17)
'
'Private Const TVIF_TEXT     As Long = &H1
'Private Const TVIF_STATE    As Long = &H8
'Private Const TVIF_HANDLE   As Long = &H10
'Private Const TVIS_STATEIMAGEMASK  As Long = &HF000
'
'Private Const TVHT_ONITEMSTATEICON  As Long = &H40
'Private Const TVM_GETNEXTITEM  As Long = (TV_FIRST + 10)
'Private Const MAX_ITEM As Long = 256
'
'Private Type POINTAPI
'  x As Long
'  y As Long
'End Type
'
'Private Type TVITEM
'  Mask As Long
'  hItem As Long
'  state As Long
'  stateMask As Long
'  pszText As String   'Long   'pointer
'  cchTextMax As Long
'  iImage As Long
'  iSelectedImage As Long
'  cChildren As Long
'  lParam As Long
'End Type
'
'Private Type TVHITTESTINFO
'  pt As POINTAPI
'  flags As Long
'  hItem As Long
'End Type
'
'Private Enum TVGN_Flags
'    TVGN_ROOT = &H0
'    TVGN_NEXT = &H1
'    TVGN_PREVIOUS = &H2
'    TVGN_PARENT = &H3
'    TVGN_CHILD = &H4
'    TVGN_FIRSTVISIBLE = &H5
'    TVGN_NEXTVISIBLE = &H6
'    TVGN_PREVIOUSVISIBLE = &H7
'    TVGN_DROPHILITE = &H8
'    TVGN_CARET = &H9
'  #If (WIN32_IE >= &H400) Then
'    TVGN_LASTVISIBLE = &HA
'  #End If
'End Enum
'
'Private Function IsTVItemChecked(ByVal hwndTV As Long, ByVal hItem As Long) As Boolean
'   Dim tvi As TVITEM
'
'   With tvi
'      .Mask = TVIF_STATE
'      .hItem = hItem
'      .stateMask = TVIS_STATEIMAGEMASK
'   End With
'   Call TreeView_GetItem(hwndTV, tvi)
'   IsTVItemChecked = (tvi.state And INDEXTOSTATEIMAGEMASK(IIL_CHECKED))
'End Function
'
'Private Function IsTVClickinCheckBox(ByVal hwndTV As Long, ByVal x As Long, ByVal y As Long) As Boolean
'   Dim tvhti As TVHITTESTINFO
'
'   tvhti.pt.x = x
'   tvhti.pt.y = y
'   IsTVClickinCheckBox = False
'   If TreeView_HitTest(hwndTV, tvhti) Then   'returns an hItem also
'     IsTVClickinCheckBox = (tvhti.flags And TVHT_ONITEMSTATEICON)
'   End If
'End Function
'
'  '---------------------------------------------------
'  'Determines if the current state image of the
'  'item under the specified point (if any) is
'  'set to the checked checkbox image index.
'  '
'  'hwndTV - treeview window handle
'  'x, y   - treeview co-ordinates in which
'  '         to retrieve the item from
'  '
'  'Returns True if the item's state image is
'  'set to the checked checkbox index, or False
'  'otherwise.
'  '---------------------------------------------------
'Private Function IsTVItemCheckedFromClick(ByVal hwndTV As Long, ByVal x As Long, ByVal y As Long) As Boolean
'   Dim tvhti As TVHITTESTINFO
'   Dim fChecked As Boolean
'
'   tvhti.pt.x = x
'   tvhti.pt.y = y
'
'   If TreeView_HitTest(hwndTV, tvhti) Then   'returns an hItem also
'
'      fChecked = IsTVItemChecked(hwndTV, tvhti.hItem)
'
'     'Since we retrieved the item's handle from
'     'a treeview co-ordinate as a result of a
'     'NodeClick event (or MouseUp event, both are
'     'invoked from a NM_CLICK notification), if
'     'this co-ordinate is within the area of the
'     'item's state icon, then the item's checkbox
'     'image is *in the process* of being toggled,
'     'but *not yet* toggled. So we'll toggle the
'     'return value reflecting the soon-to-be-set
'     'state value.
'      If (tvhti.flags And TVHT_ONITEMSTATEICON) Then fChecked = Not fChecked
'
'      IsTVItemCheckedFromClick = fChecked
'
'   End If
'
'End Function
'
'
'Public Function SetTVItemCheckImage(ByVal hwndTV As Long, ByVal hItem As Long, ByVal CheckItem As Boolean) As Boolean
'   Dim tvi As TVITEM
'
'   With tvi
'      .Mask = TVIF_HANDLE Or TVIF_STATE
'      .hItem = hItem
'      .stateMask = TVIS_STATEIMAGEMASK
'      .state = INDEXTOSTATEIMAGEMASK(Abs(CheckItem) + 1)
'   End With
'   SetTVItemCheckImage = TreeView_SetItem(hwndTV, tvi)
'End Function
'
'Private Function GetTVItemText(ByVal hwndTV As Long, ByVal hItem As Long, Optional ByVal cbItem As Long = MAX_ITEM) As String
'
'   Dim tvi As TVITEM
'   With tvi
'      .Mask = TVIF_TEXT
'      .hItem = hItem
'      .pszText = String$(cbItem, 0)
'      .cchTextMax = cbItem
'   End With
'   If TreeView_GetItem(hwndTV, tvi) Then
'      GetTVItemText = GetStrFromBufferA(tvi.pszText)
'   End If
'End Function
'
'Private Function GetStrFromBufferA(Item As String) As String
'  If InStr(Item, vbNullChar) Then
'    GetStrFromBufferA = Left$(Item, InStr(Item, vbNullChar) - 1)
'  Else
'    GetStrFromBufferA = Item
'  End If
'End Function
'
'Public Function GetTVItemFromNode(ByVal hwndTV As Long, nod As Node) As Long
'   Dim nodeCur As Node
'   Dim asNodes() As String
'   Dim nNodes As Integer
'   Dim i As Integer
'   Dim hitemParent As Long
'   Dim hItem As Long
'
'   Set nodeCur = nod
'   Do While (nodeCur Is Nothing) = False
'      nNodes = nNodes + 1
'      ReDim Preserve asNodes(nNodes)
'      asNodes(nNodes) = nodeCur.Text
'      Set nodeCur = nodeCur.Parent
'   Loop
'   hitemParent = TreeView_GetRoot(hwndTV)
'   If hitemParent Then
'      Do While nNodes
'         hItem = FindTVItemFromText(hwndTV, hitemParent, asNodes(nNodes))
'         If hItem Then
'            hitemParent = TreeView_GetChild(hwndTV, hItem)
'         Else
'            Exit Function
'         End If
'         nNodes = nNodes - 1
'      Loop
'      GetTVItemFromNode = hItem
'   End If
'End Function
'
'Private Function FindTVItemFromText(ByVal hwndTV As Long, ByVal hitemChild As Long, ByVal sItem As String) As Long
'   If Len(sItem) = 0 Then Exit Function
'
'   Do While hitemChild
'      If GetTVItemText(hwndTV, hitemChild, MAX_ITEM) = sItem Then
'         FindTVItemFromText = hitemChild
'         Exit Function
'      End If
'      hitemChild = TreeView_GetNextSibling(hwndTV, hitemChild)
'  Loop
'End Function
'
'Private Function TreeView_HitTest(hWnd As Long, lpht As TVHITTESTINFO) As Long
'   TreeView_HitTest = SendMessageAny(hWnd, TVM_HITTEST, 0&, lpht)
'End Function
'
'Private Function TreeView_GetItem(hWnd As Long, pitem As TVITEM) As Boolean
'   TreeView_GetItem = SendMessageAny(hWnd, TVM_GETITEM, 0&, pitem)
'End Function
'
'Private Function TreeView_SetItem(hWnd As Long, pitem As TVITEM) As Boolean
'   TreeView_SetItem = SendMessageAny(hWnd, TVM_SETITEM, 0&, pitem)
'End Function
'
'Private Function INDEXTOSTATEIMAGEMASK(iState As Long) As Long
'   INDEXTOSTATEIMAGEMASK = iState * (2 ^ 12)
'End Function
'
'Private Function TreeView_GetNextItem(hWnd As Long, hItem As Long, flag As Long) As Long
'   TreeView_GetNextItem = SendMessageAny(hWnd, TVM_GETNEXTITEM, flag, ByVal hItem)
'End Function
'
'Private Function TreeView_GetChild(hWnd As Long, hItem As Long) As Long
'   TreeView_GetChild = TreeView_GetNextItem(hWnd, hItem, TVGN_CHILD)
'End Function
'
'Private Function TreeView_GetNextSibling(hWnd As Long, hItem As Long) As Long
'   TreeView_GetNextSibling = TreeView_GetNextItem(hWnd, hItem, TVGN_NEXT)
'End Function
'
'Private Function TreeView_GetRoot(hWnd As Long) As Long
'   TreeView_GetRoot = TreeView_GetNextItem(hWnd, 0, TVGN_ROOT)
'End Function
'
'Public Function SetTVStyle(ByVal hWnd As Long) As Boolean
'  Dim dwStyle As Long
'  dwStyle = GetWindowLong(hWnd, GWL_STYLE)
'  If dwStyle Then
'    SetTVStyle = CBool(SetWindowLong(hWnd, GWL_STYLE, dwStyle Or TVS_CHECKBOXES))
'  End If
'End Function
'
'Public Sub NodeClickEx(hItem As Long, IsSelectedItemAlreadyChecked As Boolean, WasNodeClickWithinCheckBox As Boolean, Node As MSComCtlLib.Node, ByVal hWnd As Long)
'  Dim pt As POINTAPI
'
'  Call GetCursorPos(pt)
'  Call ScreenToClient(hWnd, pt)
'  hItem = GetTVItemFromNode(hWnd, Node)
'  IsSelectedItemAlreadyChecked = IsTVItemChecked(hWnd, hItem)
'  WasNodeClickWithinCheckBox = IsTVClickinCheckBox(hWnd, pt.x, pt.y)
'End Sub
'
