VERSION 5.00
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmGridTest 
   Caption         =   "Form1"
   ClientHeight    =   6945
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10020
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6945
   ScaleWidth      =   10020
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdExtra 
      Caption         =   "Refresh"
      Height          =   375
      Left            =   5520
      TabIndex        =   2
      Top             =   6240
      Width           =   1575
   End
   Begin VB.CommandButton cmdPopulate 
      Caption         =   "Populate"
      Height          =   375
      Left            =   7320
      TabIndex        =   1
      Top             =   6240
      Width           =   1575
   End
   Begin MSAdodcLib.Adodc griddc 
      Height          =   375
      Left            =   240
      Top             =   6000
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin TrueOleDBGrid60.TDBGrid Grid 
      Bindings        =   "frmGridTest.frx":0000
      Height          =   5535
      Left            =   120
      OleObjectBlob   =   "frmGridTest.frx":0015
      TabIndex        =   0
      Top             =   120
      Width           =   8775
   End
End
Attribute VB_Name = "frmGridTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private cn As adodb.Connection
Private rs As adodb.Recordset
Private m_NewRowBMK As Variant
Private m_TrackAddNew As Boolean

Private Sub RefreshCurrentRow()
  Dim col As Integer
  Dim cbmk As Variant
  
  On Error Resume Next
  'col = Me.Grid.col
  cbmk = Me.Grid.Bookmark
  Me.Grid.Bookmark = m_NewRowBMK
  Call Me.griddc.Recordset.Resync(adAffectCurrent, adResyncAllValues)
  Me.Grid.Bookmark = cbmk
  'Me.Grid.col = col
  Call Me.Grid.SetFocus
End Sub

Private Sub Grid_OnAddNew()
  m_NewRowBMK = Me.griddc.Recordset.Bookmark
  Debug.Print "OAN: " & Me.Grid.Bookmark & " " & Me.griddc.Recordset.Bookmark & " Editmode: " & Me.griddc.Recordset.EditMode
End Sub

Private Sub Grid_PostEvent(ByVal MsgId As Integer)
  If MsgId = 1 Then Call RefreshCurrentRow
End Sub

Private Sub cmdExtra_Click()
  Dim col As Integer, row As Integer
  Dim vbmk As Variant
  Dim t0 As Long, t1 As Long
  
  t0 = GetTicks
  col = Me.Grid.col
  Call Me.griddc.Recordset.Resync(adAffectCurrent, adResyncAllValues)
  row = Me.Grid.row
  Me.Grid.col = col
  Call Me.Grid.SetFocus
  t1 = GetTicks
  Debug.Print "BMK: " & vbmk & " row: " & row & " col: " & col & " time: " & (t1 - t0) / 1000
End Sub

Private Sub cmdPopulate_Click()
  Dim dsn As String
  Dim rs0  As adodb.Recordset
  If Me.griddc.Recordset Is Nothing Then
    Set cn = ADOConnect(ADOAccess4ConnectString(AppPath & "\" & "test.mdb"))
    
    Set rs = New adodb.Recordset
    rs.CursorLocation = adUseClient
    rs.Open "SELECT * from Contacts order by company", cn, adOpenStatic, adLockOptimistic
    
    
    Set Me.griddc.Recordset = rs
    Call Me.Grid.ReBind
  End If
End Sub

Private Sub Form_Resize()
  Const TOP_BORDER As Single = 50
  Const BOTTON_BORDER As Single = (TOP_BORDER * 2) + 1300

  Me.Grid.Width = Me.ScaleWidth - (2 * Me.Grid.Left)
  Me.Grid.Height = Me.ScaleHeight - (2 * Me.Grid.Left) - BOTTON_BORDER
  Me.cmdPopulate.Top = Me.ScaleHeight - Me.cmdPopulate.Height - (BOTTON_BORDER / 2)
  Me.cmdPopulate.Left = Me.ScaleWidth - TOP_BORDER - Me.cmdPopulate.Width
  Me.cmdExtra.Top = Me.cmdPopulate.Top
  Me.cmdExtra.Left = Me.cmdPopulate.Left - TOP_BORDER - Me.cmdExtra.Width
  
  Me.griddc.Left = TOP_BORDER
  Me.griddc.Top = Me.cmdPopulate.Top
End Sub

Private Sub Grid_AfterInsert()
  'm_NewRowBMK = Me.Grid.Bookmark
  'Debug.Print "AI: " & Me.Grid.Bookmark & " " & Me.griddc.Recordset.Bookmark
  Grid.PostMsg 1
End Sub

Private Sub Grid_AfterUpdate()
  'Debug.Print "AU: " & Me.Grid.Bookmark & " " & Me.griddc.Recordset.Bookmark
End Sub

Private Sub Grid_BeforeInsert(Cancel As Integer)
  'Debug.Print "BI: " & Me.Grid.Bookmark & " " & Me.griddc.Recordset.Bookmark
End Sub

Private Sub Grid_BeforeUpdate(Cancel As Integer)
  m_TrackAddNew = False
  'm_NewRowBMK = Me.griddc.Recordset.Bookmark
  'Debug.Print "BU: " & Me.Grid.Bookmark & " " & Me.griddc.Recordset.Bookmark
End Sub


