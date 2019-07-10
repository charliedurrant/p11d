VERSION 5.00
Begin VB.Form frmPopUp 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Items:"
   ClientHeight    =   3780
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3180
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3780
   ScaleWidth      =   3180
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdUserItem 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2610
      Picture         =   "frmPopUp.frx":0000
      TabIndex        =   6
      ToolTipText     =   "Add"
      Top             =   2880
      Width           =   375
   End
   Begin VB.TextBox txtUserItem 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      ToolTipText     =   "Enter a new item"
      Top             =   2880
      Width           =   2535
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2040
      TabIndex        =   3
      Top             =   3360
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   990
      TabIndex        =   2
      Top             =   3360
      Width           =   975
   End
   Begin VB.ListBox lstItems 
      Height          =   2010
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   0
      ToolTipText     =   "Select an item"
      Top             =   495
      Width           =   2895
   End
   Begin VB.Label lblUserItem 
      BackColor       =   &H00C0C0C0&
      Caption         =   "New item:"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   2640
      Width           =   2055
   End
   Begin VB.Label lblItems 
      Caption         =   "lblList"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   2175
   End
End
Attribute VB_Name = "frmPopUp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private lOK&
Private vT As ValText

Friend Property Set VTRef(o As ValText)
  Set vT = o
End Property
Friend Property Get OKPressed&()
  OKPressed = lOK
End Property

Private Sub cmdCancel_Click()
  lOK = False
  Me.Hide
End Sub

Private Sub cmdOK_Click()
  lOK = True
  Me.Hide
End Sub



Private Sub cmdUserItem_Click()
  Call lAddNewItem
End Sub

Private Sub lblList_Click()

End Sub

Private Sub lstItems_DblClick()
  If vT.TypeOfData <> VT_NOTINLIST Then
    cmdOK.Value = True
  End If
End Sub

Private Sub lstItems_KeyDown(KeyCode As Integer, Shift As Integer)
  'if return key pressed
  With lstItems
    If .ListCount Then
      Select Case KeyCode
        Case 13
          cmdOK.Value = True
        Case vbKeyDelete
          If vT.AllowUserDeleteItems Then
            vT.ListItems(.List(.ListIndex)).lUserDeleted = True
            .RemoveItem (.ListIndex)
          End If
      End Select
    End If
  End With
End Sub

Private Sub txtUserItem_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then
    lAddNewItem
  End If
End Sub
Private Function lAddNewItem&()
  Dim s$
  Dim lLen&, lMaxLen&
  
  s = txtUserItem.Text
  
  lLen = Len(s)
  With txtUserItem
    If Len(s) Then
      If vT.AddListItem(s, , True) Then
        lstItems.AddItem (s)
        .Text = ""
      Else
        .SelStart = 0
        .SelLength = lLen
      End If
    End If
  End With
  
lAddNewItem_END:
  Exit Function
lAddNewItem_ERR:
  Resume Next
End Function
