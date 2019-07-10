VERSION 5.00
Object = "{8D988532-0F0C-460C-B00E-7B5637E97680}#1.0#0"; "ATC2VTEXT.OCX"
Begin VB.Form F_PayeOnlineP46Options 
   Caption         =   "P46(Car) Options"
   ClientHeight    =   3150
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   3630
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3150
   ScaleWidth      =   3630
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   420
      Left            =   2450
      TabIndex        =   1
      Top             =   2640
      Width           =   1050
   End
   Begin VB.Frame fraP46Options 
      Caption         =   "P46(Car) options"
      Height          =   2490
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3390
      Begin VB.PictureBox pctFrame 
         BorderStyle     =   0  'None
         Height          =   2220
         Left            =   45
         ScaleHeight     =   2220
         ScaleWidth      =   3300
         TabIndex        =   2
         Top             =   180
         Width           =   3300
         Begin VB.OptionButton optQuarter 
            Caption         =   "Range"
            Height          =   330
            Index           =   4
            Left            =   135
            TabIndex        =   7
            Top             =   1215
            Width           =   2760
         End
         Begin VB.OptionButton optQuarter 
            Caption         =   "Quarter 4"
            Height          =   285
            Index           =   3
            Left            =   135
            TabIndex        =   6
            Top             =   855
            Width           =   3120
         End
         Begin VB.OptionButton optQuarter 
            Caption         =   "Quarter 3"
            Height          =   330
            Index           =   2
            Left            =   135
            TabIndex        =   5
            Top             =   540
            Width           =   3120
         End
         Begin VB.OptionButton optQuarter 
            Caption         =   "Quarter 2"
            Height          =   330
            Index           =   1
            Left            =   135
            TabIndex        =   4
            Top             =   270
            Width           =   3120
         End
         Begin VB.OptionButton optQuarter 
            Caption         =   "Quarter 1"
            Height          =   330
            Index           =   0
            Left            =   135
            TabIndex        =   3
            Top             =   0
            Width           =   3120
         End
         Begin atc2valtext.ValText txtP46Date 
            Height          =   285
            Index           =   0
            Left            =   1080
            TabIndex        =   8
            Top             =   1575
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   503
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MouseIcon       =   "F_PAYEOnlineP46Options.frx":0000
            Text            =   ""
            TypeOfData      =   2
            Maximum         =   "05/04/1999"
            Minimum         =   "06/04/1998"
            AutoSelect      =   0
         End
         Begin atc2valtext.ValText txtP46Date 
            Height          =   285
            Index           =   1
            Left            =   1080
            TabIndex        =   9
            Top             =   1890
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   503
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MouseIcon       =   "F_PAYEOnlineP46Options.frx":001C
            Text            =   ""
            TypeOfData      =   2
            Maximum         =   "05/04/1999"
            Minimum         =   "06/04/1998"
            AutoSelect      =   0
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00808080&
            X1              =   90
            X2              =   3060
            Y1              =   1170
            Y2              =   1170
         End
         Begin VB.Label lblDateTo 
            Caption         =   "Date to"
            Height          =   240
            Left            =   180
            TabIndex        =   11
            Top             =   1935
            Width           =   645
         End
         Begin VB.Label lblDateFrom 
            Caption         =   "Date from"
            Height          =   195
            Left            =   180
            TabIndex        =   10
            Top             =   1620
            Width           =   690
         End
      End
   End
End
Attribute VB_Name = "F_PayeOnlineP46Options"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOK_Click()
  Unload Me
End Sub
Public Function SettingsToScreen() As Boolean
  
  Dim i As Long
  
  On Error GoTo SettingsToScreen_Err
  Call xSet("SettingsToScreen")

  Call SetDefaultVTDate(txtP46Date(0))
  Call SetDefaultVTDate(txtP46Date(1))
  
  Call txtP46Date_LostFocus(-1)
  Call optQuarter_Click(-1)
  
  
SettingsToScreen_End:
  Call xReturn("SettingsToScreen")
  Exit Function

SettingsToScreen_Err:
  Call ErrorMessage(ERR_ERROR, Err, "SettingsToScreen", "Settings To Screen", "Error placing the print options to the screen.")
  Resume SettingsToScreen_End
  Resume
End Function
Private Sub Form_Load()
  Call SettingsToScreen
End Sub



Private Sub optQuarter_Click(Index As Integer)
  Dim i As Long, lCurrentQuarterIndex As Long
  Dim dQuarterEnd As Date, dQuarterStart As Date, dNow As Date
  Dim bSetValue As Boolean
  dNow = Now
  
  Select Case Index
    Case -1
      For i = 0 To 4
        If i < 4 Then
          
          Call p11d32.Rates.GetP46QuarterStartEnd(dQuarterStart, dQuarterEnd, i + 1)
          If DateInRange(dNow, dQuarterStart, dQuarterEnd) Then lCurrentQuarterIndex = i
          optQuarter(i).Caption = "Quarter " & CStr(i + 1) & " (" & DateValReadToScreen(dQuarterStart) & " - " & DateValReadToScreen(dQuarterEnd) & ")"
        End If
        If p11d32.PAYEonline.P46Range = i Then
          optQuarter(i) = True
          bSetValue = True
        End If
      Next
      If Not bSetValue Then
        optQuarter(lCurrentQuarterIndex) = True
        p11d32.PAYEonline.P46UserDateFrom = DateValReadToScreen(p11d32.PAYEonline.P46DateFrom)
        p11d32.PAYEonline.P46UserDateTo = DateValReadToScreen(p11d32.PAYEonline.P46DateTo)
      End If
    Case Is < 4
      Call p11d32.Rates.GetP46QuarterStartEnd(dQuarterStart, dQuarterEnd, Index + 1)
      p11d32.PAYEonline.P46DateFrom = dQuarterStart
      p11d32.PAYEonline.P46DateTo = dQuarterEnd
      txtP46Date(0).Enabled = False
      txtP46Date(1).Enabled = False
      txtP46Date(0).Validate = False
      txtP46Date(1).Validate = False
      p11d32.PAYEonline.P46Range = Index
    Case 4
      txtP46Date(0).Enabled = True
      txtP46Date(1).Enabled = True
      txtP46Date(0).Validate = True
      txtP46Date(1).Validate = True
      txtP46Date(0).AllowEmpty = False
      txtP46Date(1).AllowEmpty = False
      p11d32.PAYEonline.P46Range = Index
      Call SetRangeOnSelectRange
  End Select
  
End Sub
Private Sub SetRangeOnSelectRange()
  On Error GoTo err_Err
  
  Call txtP46Date_LostFocus(0)
  Call txtP46Date_LostFocus(1)

err_Err:
 Exit Sub
End Sub

Public Function CheckP46Date() As Boolean
  Dim i As Long
  On Error GoTo CheckP46Date_Err
  Call xSet("CheckP46Date")

  CheckP46Date = True

  If fraP46Options.Enabled And p11d32.PAYEonline.P46Range = P46_USERRANGE Then
    'check the validity of the two txtP46DAte
    For i = 0 To 1
      If txtP46Date(i).FieldInvalid Then
        Call ErrorMessage(ERR_ERROR, Err, "CheckP46Date", "Check P46 Date", "The P46(Car) date range is invalid.")
        CheckP46Date = False
        txtP46Date(i).SetFocus
        Exit For
      End If
    Next
  End If


CheckP46Date_End:
  Call xReturn("CheckP46Date")
  Exit Function

CheckP46Date_Err:
  Call ErrorMessage(ERR_ERROR, Err, "CheckP46Date", "Check P46 Date", "Error checking the user input P46 date range.")
  Resume CheckP46Date_End
End Function

Private Sub txtP46Date_LostFocus(Index As Integer)
  Select Case Index
    Case 0
      p11d32.PAYEonline.P46UserDateFrom = txtP46Date(0).Text
      p11d32.PAYEonline.P46DateFrom = TryConvertDateDMY(txtP46Date(Index).Text, UNDATED)
    Case 1
      p11d32.PAYEonline.P46UserDateTo = txtP46Date(1).Text
      p11d32.PAYEonline.P46DateTo = TryConvertDateDMY(txtP46Date(Index).Text, UNDATED)
    Case -1
      txtP46Date(0).Text = p11d32.PAYEonline.P46UserDateFrom
      txtP46Date(1).Text = p11d32.PAYEonline.P46UserDateTo
  End Select
End Sub

