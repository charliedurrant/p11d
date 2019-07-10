Attribute VB_Name = "Module1"
Option Explicit
'
'Brad Martinez, http://www.mvps.org/ccrp/
'

'These are the indices of the treeview
'checkbox state images when the treeview
'TVS_CHECKBOXES style bit is set.
Public Const IIL_UNCHECKED As Long = 1
Public Const IIL_CHECKED As Long = 2

Public Const GWL_STYLE = (-16)

Declare Function GetWindowLong Lib "user32" _
     Alias "GetWindowLongA" _
    (ByVal hwnd As Long, _
     ByVal nIndex As Long) As Long
     
Declare Function SetWindowLong Lib "user32" _
     Alias "SetWindowLongA" _
    (ByVal hwnd As Long, _
     ByVal nIndex As Long, _
     ByVal dwNewLong As Long) As Long

Declare Function SendMessageAny Lib "user32" _
    Alias "SendMessageA" _
    (ByVal hwnd As Long, _
     ByVal wMsg As Long, _
     ByVal wParam As Any, _
     lParam As Any) As Long

Public Type POINTAPI   'pt
  x As Long
  y As Long
End Type

Declare Function GetCursorPos Lib "user32" _
     (lpPoint As POINTAPI) As Long
     
Declare Function ScreenToClient Lib "user32" _
    (ByVal hwnd As Long, _
     lpPoint As POINTAPI) As Long
     
Declare Function GetAsyncKeyState Lib "user32" _
    (ByVal vKey As Long) As Integer

'--------------------------------------------
'treeview definitions defined in Commctrl.h at:
'http://premium.microsoft.com/msdn/library/sdkdoc/c67_4c8m.htm

'style
Public Const TVS_CHECKBOXES  As Long = &H100  '>= IE3

'messages
Public Const TV_FIRST      As Long = &H1100
Public Const TVM_GETITEM   As Long = (TV_FIRST + 12)
Public Const TVM_SETITEM   As Long = (TV_FIRST + 13)
Public Const TVM_HITTEST   As Long = (TV_FIRST + 17)

Public Type TVITEM   'was TV_ITEM
  mask As Long
  hItem As Long
  state As Long
  stateMask As Long
  pszText As String   'Long   'pointer
  cchTextMax As Long
  iImage As Long
  iSelectedImage As Long
  cChildren As Long
  lParam As Long
End Type

'TVITEM.mask flags
Public Const TVIF_TEXT     As Long = &H1
Public Const TVIF_STATE    As Long = &H8
Public Const TVIF_HANDLE   As Long = &H10

'TVITEM.state bit value
Public Const TVIS_STATEIMAGEMASK  As Long = &HF000

Public Type TVHITTESTINFO   'was TV_HITTESTINFO
  pt As POINTAPI
  flags As Long
  hItem As Long
End Type

'TVHITTESTINFO.flags value
Public Const TVHT_ONITEMSTATEICON  As Long = &H40

'User-defined as the maximum treeview item
'text length. If an items text exceeds this
'value when calling GetTVItemText there could
'be problems...
Public Const MAX_ITEM = 256
Public Const TVM_GETNEXTITEM  As Long = (TV_FIRST + 10)

'TVM_GETNEXTITEM wParam values
Public Enum TVGN_Flags
    TVGN_ROOT = &H0
    TVGN_NEXT = &H1
    TVGN_PREVIOUS = &H2
    TVGN_PARENT = &H3
    TVGN_CHILD = &H4
    TVGN_FIRSTVISIBLE = &H5
    TVGN_NEXTVISIBLE = &H6
    TVGN_PREVIOUSVISIBLE = &H7
    TVGN_DROPHILITE = &H8
    TVGN_CARET = &H9
#If (WIN32_IE >= &H400) Then
    TVGN_LASTVISIBLE = &HA
#End If
End Enum


Public Function IsTVItemChecked(hwndTV As Long, _
                                hItem As Long) As Boolean
                                
  '---------------------------------------------------
  'Determines if the current state image of the
  'specified treeview item is set to the checked
  'checkbox image index.
  '
  'hwndTV   - treeview window handle
  'hItem    - item's handle whose checkbox state is to be to returned
  '
  'Returns True if the item's state image is
  'set to the checked checkbox index, returns
  'False otherwise.
  '---------------------------------------------------
                                
   Dim tvi As TVITEM

  'Initialize the struct and get the item's state value.
   With tvi
      .mask = TVIF_STATE
      .hItem = hItem
      .stateMask = TVIS_STATEIMAGEMASK
   End With
   
   Call TreeView_GetItem(hwndTV, tvi)

  'We have to test to see if the treeview
  'checked state image *is* set since the logical
  'And test on the unchecked image (1) will
  'evaluate to True when either checkbox image
  'is set.
   IsTVItemChecked = (tvi.state And INDEXTOSTATEIMAGEMASK(IIL_CHECKED))
  
End Function


Public Function IsTVItemCheckedFromClick(hwndTV As Long, _
                                         x As Long, _
                                         y As Long) As Boolean
                                         
  '---------------------------------------------------
  'Determines if the current state image of the
  'item under the specified point (if any) is
  'set to the checked checkbox image index.
  '
  'hwndTV - treeview window handle
  'x, y   - treeview co-ordinates in which
  '         to retrieve the item from
  '
  'Returns True if the item's state image is
  'set to the checked checkbox index, or False
  'otherwise.
  '---------------------------------------------------

   Dim tvhti As TVHITTESTINFO
   Dim fChecked As Boolean
    
   tvhti.pt.x = x
   tvhti.pt.y = y
  
   If TreeView_HitTest(hwndTV, tvhti) Then   'returns an hItem also
    
      fChecked = IsTVItemChecked(hwndTV, tvhti.hItem)
    
     'Since we retrieved the item's handle from
     'a treeview co-ordinate as a result of a
     'NodeClick event (or MouseUp event, both are
     'invoked from a NM_CLICK notification), if
     'this co-ordinate is within the area of the
     'item's state icon, then the item's checkbox
     'image is *in the process* of being toggled,
     'but *not yet* toggled. So we'll toggle the
     'return value reflecting the soon-to-be-set
     'state value.
      If (tvhti.flags And TVHT_ONITEMSTATEICON) Then fChecked = Not fChecked
    
      IsTVItemCheckedFromClick = fChecked
  
   End If
  
End Function


Public Function SetTVItemCheckImage(hwndTV As Long, _
                                    hItem As Long, _
                                    fCheck As Boolean) As Boolean
                                    
  '---------------------------------------------------
  'Set the specified checkbox state for the
  'specified item. Returns True if successful,
  'returns False otherwise.
  '
  'hwndTV   - treeview window handle
  'hItem    - item's handle whose checkbox state is to be to set
  'fCheck   - If True, sets the checkbox state to the checked image,
  '           if False, sets the unchecked image.
  '---------------------------------------------------

   Dim tvi As TVITEM
  
   With tvi
      .mask = TVIF_HANDLE Or TVIF_STATE
      .hItem = hItem
      .stateMask = TVIS_STATEIMAGEMASK
  
     'As the values for the check constants are 1 for
     'unchecked (IIL_UNCHECKED) and 2 for checked
     '(IIL_CHECKED), fCheck (which is either True or
     'False) can be used directly to toggle the
     'INDEXTOIMAGESTATE parameter.
     '
     'This is accomplished by using the ABS() of
     'fCheck (turning True (-1) and False (0)
     'into 1 and 0 respectively.) Now, by adding 1,
     'the value toggles between 2 and 1 respectively,
     'exactly the same as using the IIL_ constants.
     '
     'Therefore, the single line of code below is
     'equivalent to an If..Then statement of:
     '
     'If fCheck Then
     '      tvi.state = INDEXTOSTATEIMAGEMASK(IIL_CHECKED)
     'Else: tvi.state = INDEXTOSTATEIMAGEMASK(IIL_UNCHECKED)
     'End If
     '
     'See the comments section for code that more
     'clearly demonstrates using ABS() to achieve this.
     
      .state = INDEXTOSTATEIMAGEMASK(Abs(fCheck) + 1)
   
   End With
     
   SetTVItemCheckImage = TreeView_SetItem(hwndTV, tvi)
  
End Function


Public Function GetTVItemText(hwndTV As Long, _
                              hItem As Long, _
                              Optional cbItem As Long = MAX_ITEM) As String
  
  '---------------------------------------------
  'Returns the text of the specified treeview
  'item if successful, returns an empty string
  'otherwise.
  '
  'hwndTV   - treeview window handle
  'hItem    - item's handle whose text is to be to returned
  'cbItem   - length of the specified item's text.
  '---------------------------------------------
   
   Dim tvi As TVITEM
   
   With tvi
      .mask = TVIF_TEXT
      .hItem = hItem
      .pszText = String$(cbItem, 0)
      .cchTextMax = cbItem
   End With
   
   If TreeView_GetItem(hwndTV, tvi) Then
      GetTVItemText = GetStrFromBufferA(tvi.pszText)
   End If
  
End Function


Public Function GetStrFromBufferA(item As String) As String

  'Returns the string before first null char
  'encountered (if any) from an ANSII string.

  If InStr(item, vbNullChar) Then
    GetStrFromBufferA = Left$(item, InStr(item, vbNullChar) - 1)
  
  Else
    
     'If item had no null char, the Left$ function
     'above would return a zero length string ("").
      GetStrFromBufferA = item
  
  End If

End Function


Public Function GetTVItemFromNode(hwndTV As Long, _
                                  nod As Node) As Long
   
   
  'If successful, returns the treeview item
  'handle represented by the specified Node,
  'returns 0 otherwise.
   
   Dim nodeCur As Node
   Dim asNodes() As String
   Dim nNodes As Integer
   Dim i As Integer
   Dim hitemParent As Long
   Dim hItem As Long
   
   Set nodeCur = nod
   
  'Cache the node and all of it's parent
  'node's text in the array
   Do While (nodeCur Is Nothing) = False
      nNodes = nNodes + 1
      
      ReDim Preserve asNodes(nNodes)
      asNodes(nNodes) = nodeCur.Text
      
      Set nodeCur = nodeCur.Parent
   Loop

  'Get the hItem of the first root in the
  'treeview, it will be the first parent
   hitemParent = TreeView_GetRoot(hwndTV)
   
   If hitemParent Then
    
     'Walk through the cached node text from
     'the root to the specified node (backwards
     'through the array)
     
      Do While nNodes
      
        'Get the hItem of the current node
         hItem = FindTVItemFromText(hwndTV, _
                                    hitemParent, _
                                    asNodes(nNodes))
         
         If hItem Then
         
           'Make the the current parent's first
           'child item the new parent
            hitemParent = TreeView_GetChild(hwndTV, hItem)
            
         Else: Exit Function
         End If
      
         nNodes = nNodes - 1
    
      Loop
  
      GetTVItemFromNode = hItem
    
   End If
  
End Function


Public Function FindTVItemFromText(hwndTV As Long, _
                                   ByVal hitemChild As Long, _
                                   sItem As String) As Long
                                   
  '---------------------------------------------
  'Returns the first encountered item handle
  'whose text label matches the specified text.
  '*Is case sensitive*.
  '
  'hwndTV      - treeview window handle
  'hitemChild  - first sibling item's handle in which to search
  'sItem       - specified item's text we're looking for
  '
  'If the text represented by sItem is found, it's
  'hItem is returned, otherwise 0 is returned.
  '---------------------------------------------
                                  
  'Can't find the hItem of an item with no text...
   If Len(sItem) = 0 Then Exit Function
    
   Do While hitemChild
    
    'If the current sibling item label
    'matches our target text, we're done.
    
      If GetTVItemText(hwndTV, hitemChild, MAX_ITEM) = sItem Then
         FindTVItemFromText = hitemChild
         Exit Function
      End If
    
     'Keep going while we have subsequent
     'sibling items
      hitemChild = TreeView_GetNextSibling(hwndTV, hitemChild)
  
  Loop
  
End Function


Public Function TreeView_HitTest(hwnd As Long, _
                                 lpht As TVHITTESTINFO) As Long
  
  'Determines the location of the specified point
  'relative to the client area of a treeview control.
  'Returns the handle to the tree-view item that
  'occupies the specified point or NULL if no item
  'occupies the point.

   TreeView_HitTest = SendMessageAny(hwnd, TVM_HITTEST, 0&, lpht)
   
End Function


Public Function TreeView_GetItem(hwnd As Long, pitem As TVITEM) As Boolean
   
  'Retrieves some or all of a tree-view
  'item's attributes. Returns TRUE if
  'successful or FALSE otherwise.

   TreeView_GetItem = SendMessageAny(hwnd, TVM_GETITEM, 0&, pitem)
   
End Function


Public Function TreeView_SetItem(hwnd As Long, pitem As TVITEM) As Boolean

  'Sets some or all of a tree-view item's
  'attributes. Old docs say returns zero if
  'successful or - 1 otherwise.
  'New docs say returns TRUE if successful,
  'or FALSE otherwise!
   
   TreeView_SetItem = SendMessageAny(hwnd, TVM_SETITEM, 0&, pitem)
  
End Function


Public Function INDEXTOSTATEIMAGEMASK(iState As Long) As Long

  'Prepares the index of a state image so that a
  'treeview control or listview control can use the
  'index to retrieve the state image for an item.
  'Returns the one-based index of the state image
  'shifted left twelve bits. A common control
  'utility macro.
  'This macro is defined in commctrl.h as:
  '#define INDEXTOSTATEIMAGEMASK(i) ((i) << 12)
  
   INDEXTOSTATEIMAGEMASK = iState * (2 ^ 12)

End Function


Public Function TreeView_GetNextItem(hwnd As Long, _
                                     hItem As Long, _
                                     flag As Long) As Long

  'Retrieves the tree-view item that bears the
  'specified relationship to a specified item.
  'Returns the handle to the item if successful
  'or 0 otherwise.

   TreeView_GetNextItem = SendMessageAny(hwnd, _
                                      TVM_GETNEXTITEM, _
                                      flag, _
                                      ByVal hItem)
   
End Function


Public Function TreeView_GetChild(hwnd As Long, hItem As Long) As Long
  
  'Retrieves the first child item. The hitem
  'parameter must be NULL. Returns the handle
  'to the item if successful or 0 otherwise.
      
   TreeView_GetChild = TreeView_GetNextItem(hwnd, hItem, TVGN_CHILD)
  
End Function


Public Function TreeView_GetNextSibling(hwnd As Long, _
                                        hItem As Long) As Long

   'Retrieves the next sibling item.
   'Returns the handle to the item if
   'successful or 0 otherwise.

   TreeView_GetNextSibling = TreeView_GetNextItem(hwnd, hItem, TVGN_NEXT)
   
End Function


Public Function TreeView_GetRoot(hwnd As Long) As Long

  'Retrieves the topmost or very first item
  'of the tree-view control. Returns the handle
  'to the item if successful or 0 otherwise.
   
   TreeView_GetRoot = TreeView_GetNextItem(hwnd, 0, TVGN_ROOT)
  
End Function


