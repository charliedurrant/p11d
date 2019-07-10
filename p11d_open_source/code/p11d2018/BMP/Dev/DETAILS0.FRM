VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.2#0"; "COMCTL32.OCX"
Object = "{4582CA9E-1A45-11D2-8D2F-00C04FA9DD6F}#1.0#0"; "ATC2VTEXT.OCX"
Begin VB.Form F_Details0 
   Caption         =   "Personal Details"
   ClientHeight    =   5520
   ClientLeft      =   1080
   ClientTop       =   2505
   ClientWidth     =   8505
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   Icon            =   "Details0.frx":0000
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5520
   ScaleWidth      =   8505
   WindowState     =   2  'Maximized
   Begin ComctlLib.ListView LB 
      Height          =   2565
      Left            =   75
      TabIndex        =   0
      Tag             =   "free,font"
      Top             =   105
      Width           =   8265
      _ExtentX        =   14579
      _ExtentY        =   4524
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   7
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Reference"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   2
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "NI Number"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   3
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Status"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   4
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Group1"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   5
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Group2"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   6
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Group3"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Frame P_Details 
      Height          =   2775
      Left            =   75
      TabIndex        =   14
      Top             =   2700
      Width           =   8265
      Begin VB.CommandButton B_ChangePNum 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   4800
         TabIndex        =   3
         Tag             =   "FREE"
         Top             =   630
         Width           =   315
      End
      Begin atc2valtext.ValText TB_Data 
         Height          =   285
         Index           =   6
         Left            =   1800
         TabIndex        =   5
         Tag             =   "FREE,FONT"
         Top             =   1305
         Width           =   975
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "Details0.frx":030A
         Text            =   ""
         TypeOfData      =   3
      End
      Begin atc2valtext.ValText TB_Data 
         Height          =   585
         Index           =   7
         Left            =   120
         TabIndex        =   11
         Tag             =   "FREE,FONT"
         Top             =   2040
         Width           =   7815
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "Details0.frx":0326
         Text            =   ""
         TypeOfData      =   3
      End
      Begin VB.ComboBox CB_Status 
         DataField       =   "Status"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6360
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Tag             =   "FREE,FONT"
         Top             =   240
         Width           =   1545
      End
      Begin atc2valtext.ValText TB_Data 
         Height          =   285
         Index           =   2
         Left            =   6360
         TabIndex        =   8
         Tag             =   "FREE,FONT"
         Top             =   600
         Width           =   1545
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "Details0.frx":0342
         Text            =   ""
         TypeOfData      =   3
      End
      Begin atc2valtext.ValText TB_Date 
         Height          =   285
         Index           =   1
         Left            =   6375
         TabIndex        =   9
         Tag             =   "FREE,FONT"
         Top             =   960
         Width           =   1545
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "Details0.frx":035E
         Text            =   ""
         TypeOfData      =   2
         Maximum         =   "5/4/1999"
         Minimum         =   "6/4/1998"
      End
      Begin atc2valtext.ValText TB_Date 
         Height          =   285
         Index           =   2
         Left            =   6360
         TabIndex        =   10
         Tag             =   "FREE,FONT"
         Top             =   1320
         Width           =   1545
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "Details0.frx":037A
         Text            =   ""
         TypeOfData      =   2
         Maximum         =   "5/4/1999"
         Minimum         =   "6/4/1998"
      End
      Begin atc2valtext.ValText TB_Data 
         Height          =   285
         Index           =   8
         Left            =   4173
         TabIndex        =   6
         Tag             =   "FREE,FONT"
         Top             =   1305
         Width           =   975
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "Details0.frx":0396
         Text            =   ""
         TypeOfData      =   3
      End
      Begin atc2valtext.ValText TB_Data 
         Height          =   285
         Index           =   3
         Left            =   1800
         TabIndex        =   4
         Tag             =   "FREE,FONT"
         Top             =   945
         Width           =   3348
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "Details0.frx":03B2
         Text            =   ""
         TypeOfData      =   3
      End
      Begin atc2valtext.ValText TB_Data 
         Height          =   285
         Index           =   1
         Left            =   1800
         TabIndex        =   2
         Tag             =   "FREE,FONT"
         Top             =   600
         Width           =   2985
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Locked          =   -1  'True
         MouseIcon       =   "Details0.frx":03CE
         Text            =   ""
         TypeOfData      =   3
      End
      Begin atc2valtext.ValText TB_Data 
         Height          =   285
         Index           =   0
         Left            =   1800
         TabIndex        =   1
         Tag             =   "FREE,FONT"
         Top             =   270
         Width           =   3348
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "Details0.frx":03EA
         Text            =   ""
         TypeOfData      =   3
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Comments"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   192
         Index           =   2
         Left            =   60
         TabIndex        =   21
         Tag             =   "FREE,FONT"
         Top             =   1740
         Width           =   1512
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Group Code 3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   1
         Left            =   2940
         TabIndex        =   16
         Tag             =   "FREE,FONT"
         Top             =   1395
         Width           =   1155
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   60
         TabIndex        =   12
         Tag             =   "FREE,FONT"
         Top             =   360
         Width           =   1515
      End
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Personnel Reference"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   60
         TabIndex        =   13
         Tag             =   "FREE,FONT"
         Top             =   675
         Width           =   1755
      End
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "NI Number"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   5280
         TabIndex        =   18
         Tag             =   "FREE,FONT"
         Top             =   675
         Width           =   1095
      End
      Begin VB.Label Label10 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Date started"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   5280
         TabIndex        =   19
         Tag             =   "FREE,FONT"
         Top             =   1035
         Width           =   1110
      End
      Begin VB.Label Label11 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Date leaving"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   5280
         TabIndex        =   20
         Tag             =   "FREE,FONT"
         Top             =   1395
         Width           =   1095
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Group Code 1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   60
         TabIndex        =   22
         Tag             =   "FREE,FONT"
         Top             =   1035
         Width           =   1635
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Group Code 2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   0
         Left            =   60
         TabIndex        =   15
         Tag             =   "FREE,FONT"
         Top             =   1395
         Width           =   1755
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Status"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   5280
         TabIndex        =   17
         Tag             =   "FREE,FONT"
         Top             =   360
         Width           =   1095
      End
   End
End
Attribute VB_Name = "F_Details0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Implements IBenefitForm
Implements IFrmGeneral

Private mclsResize As New clsFormResize
Private Const L_DES_HEIGHT = 5925
Private Const L_DES_WIDTH = 8505
Private m_InvalidVt As atc2valtext.ValText


Private Function IFrmGeneral_CheckChanged(C As Control, ByVal UpdateCurrentListItem As Boolean) As Boolean

End Function

Public Property Get IFrmGeneral_InvalidVT() As atc2valtext.ValText
  Set IFrmGeneral_InvalidVT = m_InvalidVt
End Property

Public Property Set IFrmGeneral_InvalidVT(NewValue As atc2valtext.ValText)
  Set m_InvalidVt = NewValue
End Property


Private Sub CB_Status_Lostfocus()
  Call CheckChanged(CB_Status)
End Sub

Private Sub TB_Date_LostFocus(Index As Integer)
  Call CheckChanged(TB_Date(Index))
End Sub


Private Sub Form_Load()
  If Not (mclsResize.InitResize(Me, L_DES_HEIGHT, L_DES_WIDTH, DESIGN, , , MDIMain)) Then ', DESIGN)) Then
    Err.Raise ERR_Application
  End If
  CB_Status.Clear
  CB_Status.AddItem S_STAFF
  CB_Status.AddItem S_DIRECTOR
End Sub


Private Sub IBenefitForm_AddBenefit()
  Dim q As New SQLQUERIES
  Dim lst As ListItem
  Dim rs As Recordset
  Dim ee As clsEmployee
  Dim frm As Form
  Dim I As Long
  On Error GoTo AddBenefit_Err
  Call xSet("AddBenefit")
  
  Set frm = F_EeNew
'Addbenefit_retry:
  'Do While True
    
AddEmployee_ReTest:
    frm.Show 1
    
    'Create new employee and fill in details...
    
    If frm.m_OK = False Then
      If F_Details.LB.ListItems.count = 0 Then F_Details.Enabled = False
      GoTo AddBenefit_End
    End If
    
    Set ee = New clsEmployee
    With ee
      Set .Parent = CurrentEmployer
      .Name = frm.TxtBx(0).Text
      .PersonelNo = frm.TxtBx(1).Text
      .NINumber = frm.TxtBx(2).Text
      .Payeref = frm.TxtBx(3).Text
      Set rs = CurrentEmployer.db.OpenRecordset(q.Queries(SELECT_EMPLOYEES), dbOpenDynaset)
      rs.AddNew
      rs.Fields(S_EMPLOYEE_NUM_FIELD) = .PersonelNo
      rs.Fields("NAME") = .Name
      rs.Fields("NI") = .NINumber
      rs.Update
      Set rs = CurrentEmployer.db.OpenRecordset(q.Queries(SELECT_EMPLOYEE, ee.PersonelNo), dbOpenSnapshot)
      Call .ReadDB(rs)
      Set rs = Nothing
      .Dirty = True
      .WriteDB
      
      I = Employees.Add(ee)
      Set lst = Me.LB.ListItems.Add(, , .Name)
      lst.Tag = I
      lst.SubItems(1) = .PersonelNo
      lst.SubItems(2) = .NINumber
      lst.SubItems(3) = IIf(.Status, "Director", "Staff")
      lst.SubItems(4) = .Group1
      lst.SubItems(5) = .Group2
      lst.SubItems(6) = .Group3
      
      Set LB.SelectedItem = lst
      F_Details.Enabled = True
      Call EmployeeDetails(I)
      Call MDIMain.SetDelete
      Set lst = Nothing
      Call frm.ClearFields
    End With
    Set ee = Nothing
    
AddBenefit_End:
  Unload frm
  Set frm = Nothing
  MDIMain.Enabled = True
  Set ee = Nothing
  Set rs = Nothing
  Call DBEngine.Idle(dbFreeLocks)
  Call xReturn("AddBenefit")
  Exit Sub
AddBenefit_Err:
  Select Case Err.Number
    Case 3022
      Call ErrorMessage(ERR_ERROR, Err, "Duplicate Personnel reference", "ERR_ADDEMPOYEE", "The personnel reference you are trying to add already exists in the database." & vbCrLf & "Please use an alternative reference")
      rs.CancelUpdate
      Resume AddEmployee_ReTest
    Case 3315
      Call ErrorMessage(ERR_ERROR, Err, "Empty field", "ERR_ADDEMPOYEE", "You must complete all the fields.")
      rs.CancelUpdate
      Resume AddEmployee_ReTest
    Case Else
      Call ErrorMessage(ERR_ERROR, Err, "AddBenefit", "ERR_ADDBENEFIT", "Error in AddBenefit function, called from the form " & Me.Name & ".")
      Resume AddBenefit_End
    End Select
    Resume
End Sub
Private Sub IBenefitForm_ClearFields()
    TB_Data(0).Text = ""
    TB_Data(1).Text = ""
    TB_Data(2).Text = ""
    TB_Data(3).Text = ""
    TB_Data(6).Text = ""
    TB_Data(7).Text = ""
    CB_Status = S_STAFF
    TB_Data(8).Text = ""
    TB_Date(1).Text = ""
    TB_Date(2).Text = ""
End Sub

Private Function IBenefitForm_Remove(I As Long) As Boolean
  On Error GoTo KillBenefit_Err
  Call xSet("KillBenefit")
  
  Call DeleteEmployee  'GetEmployee(i))
  Call IBenefitForm_ClearFields
  
KillBenefit_End:
  Call xReturn("KillBenefit")
  Exit Function
KillBenefit_Err:
  Call ErrorMessage(ERR_ERROR, Err, "KillBenefit", "ERR_UNDEFINED", "Undefined error.")
  Resume KillBenefit_End
End Function

Private Function IBenefitForm_ListBenefits(Optional ByVal Index As Long = 0&) As Boolean

End Function

Private Sub LB_DblClick()
  
  Call BenefitToolBar(1, GetEmployeeIndexFromSelectedEmployee)
End Sub

Private Sub LB_ItemClick(ByVal Item As ComctlLib.ListItem)

  Static InClick As Boolean
  Static Ignore As Boolean
  Dim lItem As ListItem, ret As Boolean
  Dim I As Long, id As Long
  
  Debug.Print "ItemClick"
  If Not InClick Then
    InClick = True: I = -1
    If Item.Selected Then I = Item.Tag
    If Ignore Then
      Ignore = False
      ret = True
    Else
      ret = Not EmployeeDetails(I)
      If ret Then Ignore = True
    End If
    If ret Then
      Item.Selected = False
      Set Me.LB.SelectedItem = Nothing
      If Not CurrentEmployee Is Nothing Then
      'apf
        id = GetObjectIndex(Employees, CurrentEmployee)
        If id <> -1 Then
          For I = 1 To Me.LB.ListItems.count
            Set lItem = Me.LB.ListItems(I)
            If lItem.Tag = id Then
              Set Me.LB.SelectedItem = Me.LB.ListItems(I)
              
              Exit For
            End If
          Next I
        End If
      End If
    End If
    InClick = False
  End If
End Sub

Private Sub LB_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then 'Return key
    Call LB_DblClick
  End If
End Sub

Private Sub TB_data_Lostfocus(Index As Integer)
 Call CheckChanged(TB_Data(Index))
End Sub

Private Function CheckChanged(ctl As Control) As Boolean
  Dim I As Long, vDate As Variant
  On Error GoTo CheckChanged_Err
  Call xSet("CheckChanged")
  
  If CurrentEmployee Is Nothing Then GoTo CheckChanged_End
  Select Case ctl.Name
    Case "TB_Data"
      Select Case ctl.Index
        Case 0
          I = StrComp(ctl.Text, CurrentEmployee.Name, vbBinaryCompare)
          If I <> 0 Then CurrentEmployee.Name = ctl.Text
        Case 1
          'Ignore the P_NUM field
        Case 2
          I = StrComp(ctl.Text, CurrentEmployee.NINumber, vbBinaryCompare)
          If I <> 0 Then CurrentEmployee.NINumber = ctl.Text
        Case 3
          I = StrComp(ctl.Text, CurrentEmployee.Group1, vbBinaryCompare)
          If I <> 0 Then CurrentEmployee.Group1 = ctl.Text
        Case 6
          I = StrComp(ctl.Text, CurrentEmployee.Group2, vbBinaryCompare)
          If I <> 0 Then CurrentEmployee.Group2 = ctl.Text
        Case 7
          I = StrComp(ctl.Text, CurrentEmployee.Comments, vbBinaryCompare)
          If I <> 0 Then CurrentEmployee.Comments = ctl.Text
        Case 8
          I = StrComp(ctl.Text, CurrentEmployee.Group3, vbBinaryCompare)
          If I <> 0 Then CurrentEmployee.Group3 = ctl.Text
        Case Else
          ECASE "Unknown control"
      End Select
    Case "TB_Date"
      Select Case ctl.Index
        Case 1
          vDate = TryConvertDateDMY(ctl.Text, UNDATED)
          I = vDate <> CurrentEmployee.Joined
          CurrentEmployee.Joined = vDate
        Case 2
          vDate = TryConvertDateDMY(ctl.Text, UNDATED)
          I = vDate <> CurrentEmployee.Left
          If I <> 0 Then CurrentEmployee.Left = vDate
      End Select
    Case "CB_Status"
      I = StrComp(ctl.Text, IIf(CurrentEmployee.Status, S_DIRECTOR, S_STAFF), vbBinaryCompare)
      If I <> 0 Then CurrentEmployee.Status = IIf(ctl.Text = S_DIRECTOR, True, False)
    Case Else
  End Select
  
  If I <> 0 Then CurrentEmployee.InvalidFields = InvalidFields(Me)
  If CurrentEmployee.InvalidFields > 0 Then
    Call MDIMain.sts.SetStatus(0, "", S_NOEMPSAVE)
    Call MDIMain.SetUndo
    CurrentEmployee.Dirty = False
  ElseIf I <> 0 Then
    Call MDIMain.sts.SetStatus(0, "", "")
    Call MDIMain.SetConfirmUndo
    With LB.SelectedItem
      .Text = CurrentEmployee.Name
      .SubItems(1) = CurrentEmployee.PersonelNo
      .SubItems(2) = CurrentEmployee.NINumber
      .SubItems(3) = IIf(CurrentEmployee.Status, "Staff", "Director")
      .SubItems(4) = CurrentEmployee.Group1
      .SubItems(5) = CurrentEmployee.Group2
      .SubItems(6) = CurrentEmployee.Group3
    End With
    CurrentEmployee.Dirty = True
  ElseIf CurrentEmployee.Dirty Then
    Call MDIMain.sts.SetStatus(0, "", "")
    Call MDIMain.SetConfirmUndo
  End If

CheckChanged_End:
  Call xReturn("CheckChanged")
  Exit Function

CheckChanged_Err:
  Call ErrorMessage(ERR_ERROR, Err, "CheckChanged", "ERR_CHECKCHANGED", "Undefined error.")
  Resume CheckChanged_End
  Resume
End Function


Private Sub Form_Resize()
  mclsResize.Resize
  Call ColumnWidths(Me.LB, L_NAME_COL, L_REFERENCE_COL, L_NINUMBER_COL&, L_STATUS_COL&, L_GROUP1_COL&, L_GROUP2_COL&, L_GROUP3_COL&)
End Sub

Private Sub LB_ColumnClick(ByVal ColumnHeader As ComctlLib.ColumnHeader)
  Me.LB.SortKey = ColumnHeader.Index - 1
  LB.SelectedItem.EnsureVisible
End Sub


