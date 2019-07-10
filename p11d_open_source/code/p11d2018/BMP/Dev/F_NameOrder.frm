VERSION 5.00
Begin VB.Form F_NameOrder 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Set Name Order"
   ClientHeight    =   1155
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3150
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1155
   ScaleWidth      =   3150
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   900
      TabIndex        =   3
      Top             =   720
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2070
      TabIndex        =   2
      Top             =   720
      Width           =   1050
   End
   Begin VB.ComboBox cbo 
      Height          =   315
      Left            =   45
      TabIndex        =   0
      Top             =   270
      Width           =   3120
   End
   Begin VB.Label lbl 
      Caption         =   "Select the order of name parts required"
      Height          =   240
      Left            =   45
      TabIndex        =   1
      Top             =   45
      Width           =   3075
   End
End
Attribute VB_Name = "F_NameOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
  Me.Hide
End Sub

Private Sub cmdOK_Click()
  Call NameOrderToScreen
  Me.Hide
End Sub
Private Sub NameOrderToScreen()
  Dim ibf As IBenefitForm2
  Dim lv As ListView
  Dim li As ListItem
  Dim ben As IBenefitClass
  Dim employees As ObjectList
  Dim i As Long
  
  On Error GoTo NameOrderToScreen_ERR
  
  Call xSet("NameOrderToScreen")
  
  On Error GoTo NameOrderToScreen_ERR
  
  Call SetCursor
  If cbo.ItemData(cbo.ListIndex) = p11d32.NameOrder Then GoTo NameOrderToScreen_END
  
  p11d32.NameOrder = cbo.ItemData(cbo.ListIndex)
  
  Set ibf = F_Employees
  Set lv = ibf.lv
  Set employees = p11d32.CurrentEmployer.employees
  Call PrgStartCaption(employees.Count, "Updating names", "Employee", ValueOfMax)
  For Each li In lv.listitems
    Set ben = employees(li.Tag)
    Call ibf.UpdateBenefitListViewItem(li, ben)
    Call PrgStep
  Next
  
  If CurrentForm Is F_Employees And Not p11d32.CurrentEmployeeIsNothing Then Call MDIMain.NavigateBarUpdate(p11d32.CurrentEmployer.CurrentEmployee)
  
NameOrderToScreen_END:
  Call PrgStopCaption
  Call ClearCursor
  Call xReturn("NameOrderToScreen")
  Exit Sub
NameOrderToScreen_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "NameOrderToScreen", "Name Order To Screen", "Error placing the name order to the screen.")
  Resume NameOrderToScreen_END

End Sub
Private Sub Form_Load()
  Call SettingsToScreen
End Sub
Private Sub SettingsToScreen()
  Dim s As String
  Dim i As NAME_ORDERS
  On Error GoTo SettingsToScreen_Err
  
  Call xSet("SettingsToScreen")
  
  For i = NO_FIRST_ITEM To NO_LAST_ITEM
    Select Case i
      Case NO_SURNAME_TITLE_FN_INITIALS
        s = "Surname, Title, First Name, Initials"
      Case NO_SURNAME_INITIALS
        s = "Surname, Initials"
      Case NO_TITLE_FN_INITIALS_SURNAME
        s = "Title, First Name, Initials, Surname"
      Case NO_FN_INITIALS_SURNAME
        s = "First Name, Initials, Surname"
      Case NO_FN_SURNAME
        s = "First Name, Surname"
      Case NO_SURNAME_FN_INITIALS
        s = "Surname, First Name, Initials"
      Case NO_SURNAME_FN
        s = "Surname, First Name"
      Case NO_TITLE_INITIALS_SURNAME
        s = "Title, Initials, Surname"
      Case NO_INITIALS_SURNAME
        s = "Initials, Surname"
      Case Else
        Call ECASE("Invalid Name Order in settings to screen.")
    End Select
    cbo.AddItem (s)
    cbo.ItemData(i) = i
  Next
  
  If cbo.ListCount = 0 Then Call ECASE("No Name orders in settings to screen.")
  If cbo.ListCount < (p11d32.NameOrder + 1) Then
    p11d32.NameOrder = 0
    cbo.ListIndex = 0
  Else
    cbo.ListIndex = p11d32.NameOrder
  End If
  
SettingsToScreen_End:
  Call xReturn("SettingsToScreen")
  Exit Sub
SettingsToScreen_Err:
  Call ErrorMessage(ERR_ERROR, Err, "SettingsToScreen", "Settings To Screen", "Error placing the name order settings to the screen.")
  Resume SettingsToScreen_End
End Sub
