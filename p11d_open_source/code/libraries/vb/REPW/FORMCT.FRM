VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.2#0"; "COMCTL32.OCX"
Begin VB.Form Formct 
   Caption         =   "Formct"
   ClientHeight    =   5115
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   5115
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdUnCheckChildren 
      Caption         =   "Uncheck Children"
      Height          =   390
      Left            =   2250
      TabIndex        =   4
      Top             =   3825
      Width           =   1515
   End
   Begin VB.CommandButton cmdCheckChildren 
      Caption         =   "Check Children"
      Height          =   390
      Left            =   2250
      TabIndex        =   3
      Top             =   3300
      Width           =   1515
   End
   Begin VB.OptionButton optMouseUp 
      Caption         =   "Mouse Up"
      Height          =   540
      Left            =   450
      TabIndex        =   2
      Top             =   3750
      Width           =   1665
   End
   Begin VB.OptionButton optNodeClick 
      Caption         =   "Node Click"
      Height          =   390
      Left            =   450
      TabIndex        =   1
      Top             =   3300
      Width           =   1740
   End
   Begin ComctlLib.TreeView TreeView1 
      Height          =   3015
      Left            =   300
      TabIndex        =   0
      Top             =   75
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   5318
      _Version        =   327682
      Style           =   7
      Appearance      =   1
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   315
      Left            =   1050
      TabIndex        =   5
      Top             =   4500
      Width           =   2265
   End
End
Attribute VB_Name = "Formct"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit ''Brad Martinez, http://www.mvps.org/ccrp/'
Private hwndTV As Long

Private Sub Form_Load()
  
  Dim Node1 As Node
  Dim Node2 As Node
  Dim Node3 As Node
  Dim i As Integer
  Dim j As Integer
  Dim k As Integer
  'Set some treeview properties, and
  'fill up the treeview with two root
  'items, each having a child node with
  'with four children...
  With TreeView1
    'For convenience...
    .HideSelection = False
    .LabelEdit = tvwManual
    hwndTV = .hwnd
    'Use the API to set the checkbox style
    Call SetTVStyle(hwndTV)
    For i = 1 To 2
      Set Node1 = .Nodes.Add(, tvwLast, , "Root" & i)
      For j = 1 To 1
        Set Node2 = .Nodes.Add(Node1.Index, tvwChild, , "Root" & i & "Child" & j)
        For k = 1 To 4
          Set Node3 = .Nodes.Add(Node2.Index, tvwChild, , "Grandchild" & k)
        Next k
        Node2.Expanded = True
      Next j
      Node1.Expanded = True
    Next i
  End With
  
  'Let Label1 reflect the first root Node's text
  'and check state. This is subsequently done
  'in the NodeClick event, but here on loading,
  'a NodeClick is not generated when the treeview
  'first appears, even though the first root
  'is selected). So we manually call the routine.
  
  Call DisplaySelectedNodeState(TreeView1.Nodes(1), False)
End Sub

Private Function SetTVStyle(hwndTV As Long) As Boolean
  Dim dwStyle As Long
  dwStyle = GetWindowLong(hwndTV, GWL_STYLE)
  'Set the treeview checkbox style. Note that
  'this style is applied across the entire
  'treeview - you can not have some items
  'allowing checks while others don't. If this
  'functionality is needed, you must use your
  'own state images to mimic the checkboxes.
  'This is not covered in this code example.
  
  If dwStyle Then
    SetTVStyle = CBool(SetWindowLong(hwndTV, _
    GWL_STYLE, _
    dwStyle Or TVS_CHECKBOXES))
  End If
End Function

Private Sub DisplaySelectedNodeState(sNode As String, fIsChecked As Boolean)
  If fIsChecked Then
    Label1 = sNode & " is checked"
  Else
    Label1 = sNode & " is unchecked"
  End If
End Sub

Private Sub TreeView1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
  'The only way to determine if a Node is checked or
  'not is by sending the treeview a TVM_GETITEM
  'message and retrieving the Node's checked state.
  'In order to send this message, we must have the
  'treeview item handle of the Node. The easiest
  'way to obtain the Node's item handle is by sending
  'the treeview a TVM_HITTEST message with the
  'co-ordinates of the mouse when the left button  'is released.
  
  Dim nodeSel As Node
  Dim fChecked As Boolean
  'Only perform this if the 'use MouseUp'  'option was selected.
  
  If optMouseUp Then
    If Button = vbLeftButton Then
      'Get the left-clicked node
      Set nodeSel = TreeView1.HitTest(x, y)
      If (nodeSel Is Nothing) = False Then
        fChecked = IsTVItemCheckedFromClick(hwndTV, x / Screen.TwipsPerPixelX, y / Screen.TwipsPerPixelY)
        Call DisplaySelectedNodeState(nodeSel.Text, fChecked)
      End If   '(nodeSel Is Nothing) = False
    End If   'Button = vbLeftButton
  End If   'optMouseUp
End Sub

Private Sub TreeView1_NodeClick(ByVal Node As ComctlLib.Node)
  'We could do things a bit differently here and
  'obtain the treeview item handle of the Node with
  'the GetTVItemFromNode call, and pass the handle
  'directly to the IsTVItemChecked proc. But since
  'we don't determine if this event was invoked due
  'to a checkbox click (toggling the Node's checkbox
  'state), IsTVItemChecked will not return an accurate
  'value.
  Dim fChecked As Boolean
  Dim hItem As Long
  Dim pt As POINTAPI
  'Only perform this if the 'use node'
  'option was selected.
  If optNodeClick Then
    'If this event was invoked from a left
    'mouse button click (if the left mouse
    'button was depressed when the NM_CLICK
    'message was received by the treeview
    'parent window).
    If GetAsyncKeyState(vbKeyLButton) Then
      'Get the current cursor pos in screen
      'coords, convert it to treeview coords,
      'and get the check state.
      Call GetCursorPos(pt)
      Call ScreenToClient(hwndTV, pt)
      fChecked = IsTVItemCheckedFromClick(hwndTV, pt.x, pt.y)
      Call DisplaySelectedNodeState(Node.Text, fChecked)
    End If   'GetAsyncKeyState
  End If   'optNodeClick
End Sub

Private Sub cmdCheckChildren_Click()
  'check the children of the selected item
  Call SetCheckStateOfChildren(TreeView1.SelectedItem, True)
End Sub

Private Sub cmdUncheckChildren_Click()
  'uncheck the children of the selected item
  Call SetCheckStateOfChildren(TreeView1.SelectedItem, False)
End Sub

Private Sub SetCheckStateOfChildren(nodeParent As Node, fCheck As Boolean)
  Dim nodeChild As Node
  Dim hItem As Long
  
  Set nodeChild = nodeParent.Child
  Do While (nodeChild Is Nothing) = False
    'obtain the item handle of the node
    hItem = GetTVItemFromNode(hwndTV, nodeChild)
    'if a valid handle, set the checked state
    If hItem Then Call SetTVItemCheckImage(hwndTV, hItem, fCheck)
    'if the node has child nodes itself,
    'recursively call this routine to set
    'the state of those as well
    If (nodeChild.Child Is Nothing) = False Then
      Call SetCheckStateOfChildren(nodeChild, fCheck)
    End If
    Set nodeChild = nodeChild.Next
  Loop
End Sub
