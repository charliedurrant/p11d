VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.UserControl PayeOnlineXMLVlaidatorControl 
   ClientHeight    =   9105
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9585
   ScaleHeight     =   9105
   ScaleWidth      =   9585
   Begin MSComctlLib.ListView listViewErrors 
      Height          =   2055
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   3625
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin RichTextLib.RichTextBox rt 
      Height          =   5535
      Left            =   120
      TabIndex        =   0
      Top             =   3000
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   9763
      _Version        =   393217
      Enabled         =   -1  'True
      HideSelection   =   0   'False
      ScrollBars      =   3
      TextRTF         =   $"PayeOnlineXMLVlaidatorControl.ctx":0000
   End
   Begin VB.Label lblCursorPos 
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   8640
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   $"PayeOnlineXMLVlaidatorControl.ctx":0082
      Height          =   735
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   9255
   End
End
Attribute VB_Name = "PayeOnlineXMLVlaidatorControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private m_Validator As PayeOnlineXMLValidator
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SendMessageStr Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Private Declare Function SendMessageBynum& Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)

'declare some constants for sendmessage
Private Const EM_LINESCROLL = &HB6 'needed in my version
Private Const EM_SCROLL As Long = &HB5 'needed in both versions
Private Const EM_GETLINECOUNT As Long = &HBA 'needed in both versions
Private Const EM_GETFIRSTVISIBLELINE = &HCE 'used to re-sync lines, returns topmost visible line #
'list box top constants
Private Const LB_GETTOPINDEX As Long = &H18E
Private Const LB_SETTOPINDEX As Long = &H197
Private Const EM_LINEINDEX As Long = &HBB
Private Const EM_LINEFROMCHAR As Long = &HC9
Public Sub Init(validator As PayeOnlineXMLValidator, xml As String)
  rt.TextRTF = xml
  Set m_Validator = validator
  Call ValidatorErrorsToList(validator)
End Sub
Private Sub ValidatorErrorsToList(validator As PayeOnlineXMLValidator)
  Dim error As PayeOnlineXmlValidationError
  Dim li As ListItem
  Dim ch As ColumnHeader
  
  Set ch = listViewErrors.ColumnHeaders.Add(, , "Description")
  ch.width = 5000
  Call listViewErrors.ColumnHeaders.Add(, , "Line")
  Call listViewErrors.ColumnHeaders.Add(, , "Column")
  
  For Each error In validator.Errors
    Set li = listViewErrors.listitems.Add(, , Replace$(Trim$(error.Description), vbCrLf, "-"))
    Call li.ListSubItems.Add(, , error.LineNumber)
    Call li.ListSubItems.Add(, , error.ColumnNumber)
  Next
  
  If IsRunningInIDE() Then
    Set li = listViewErrors.listitems.Add(, , "IDE Test error")
    Call li.ListSubItems.Add(, , 1)
    Call li.ListSubItems.Add(, , 1)
  
  End If
  
    
  If (listViewErrors.listitems.Count > 0) Then
    Call listViewErrors_ItemClick(listViewErrors.listitems(1))
    listViewErrors.listitems(1).Selected = True
    
  End If
  
End Sub

Private Sub listViewErrors_ItemClick(ByVal Item As MSComctlLib.ListItem)
  Dim line As Long
  Dim column As Long
  Dim dif As Long, firstvis As Long
  Dim charindex As Long, charindex2 As Long
  
  On Error GoTo err_Err
  
  If (Len(Item.SubItems(1)) = 0) Then GoTo err_End
  
  line = CLng(Item.SubItems(1))
  line = line - 1
  column = CLng(Item.SubItems(2))
  column = column - 1
  Call SendMessage(rt.hwnd, EM_LINESCROLL, 0, 10)
  
  firstvis = SendMessageBynum(rt.hwnd, EM_GETFIRSTVISIBLELINE, 0, 0) 'get the topmost visible line #
  dif = (line) - firstvis 'calculate the change needed
  SendMessageBynum rt.hwnd, EM_LINESCROLL, 0, dif 'scroll there
  
  
  charindex = SendMessage(rt.hwnd, EM_LINEINDEX, ByVal line, ByVal CLng(0))
  
  rt.SelStart = charindex
  charindex2 = SendMessage(rt.hwnd, EM_LINEINDEX, ByVal line + 1, ByVal CLng(0))
  rt.SelLength = charindex2 - charindex
  rt.SelColor = vbRed
  
  rt.SelLength = column + 1
  rt.SelUnderline = True
  rt.SelLength = 0
  rt.SelStart = charindex + (column)
  
  
err_End:
  Exit Sub
err_Err:
  Call ErrorMessage(ERR_ERROR, Err, "ItemClick", "Item Click", "Failed to goto the error")
  Resume err_End:
  
End Sub

Private Sub rt_SelChange()
  Dim s As String
  Dim line As Long
  Dim charindex As Long
  
  line = SendMessageBynum(rt.hwnd, EM_LINEFROMCHAR, rt.SelStart + rt.SelLength, 0)  'get the topmost visible line #
  charindex = SendMessage(rt.hwnd, EM_LINEINDEX, ByVal line, ByVal CLng(0))
  line = line + 1
  
  charindex = charindex - 1
  charindex = (rt.SelStart + rt.SelLength) - charindex
  
  lblCursorPos.Caption = "Line: " & line & ", column: " & charindex
End Sub

