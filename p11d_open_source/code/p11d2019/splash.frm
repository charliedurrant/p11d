VERSION 5.00
Object = "{AF27A9B5-A3F4-11D2-8DB7-00C04FA9DD6F}#1.2#0"; "TCSPROG.OCX"
Object = "{89056D22-ECDA-4A64-B90B-25EBB3AE8DB8}#1.0#0"; "atc2hook.ocx"
Begin VB.Form frmSplash 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6120
   ClientLeft      =   225
   ClientTop       =   1380
   ClientWidth     =   7560
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   FontTransparent =   0   'False
   Icon            =   "splash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "splash.frx":000C
   ScaleHeight     =   6120
   ScaleWidth      =   7560
   StartUpPosition =   2  'CenterScreen
   Begin TCSPROG.TCSProgressBar prgStartup 
      Height          =   255
      Left            =   480
      TabIndex        =   4
      Top             =   4560
      Visible         =   0   'False
      Width           =   6495
      _cx             =   11456
      _cy             =   450
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Min             =   0
      Max             =   100
      Value           =   50
      BarBackColor    =   16777215
      BarForeColor    =   10040064
      Appearance      =   2
      Style           =   1
      CaptionColor    =   0
      CaptionInvertColor=   16777215
      FillStyle       =   0
      FadeFromColor   =   0
      FadeToColor     =   16777215
      Caption         =   ""
      InnerCircle     =   0   'False
      Percentage      =   2
      Skew            =   0
      PictureOffsetTop=   0
      PictureOffsetLeft=   0
      Enabled         =   -1  'True
      Increment       =   1
      TextAlignment   =   2
   End
   Begin atc2hook.HOOK HOOK 
      Left            =   6720
      Top             =   5040
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin VB.Label lblMessage 
      BackColor       =   &H00A481AA&
      BackStyle       =   0  'Transparent
      Caption         =   "lblMessage"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   330
      Left            =   480
      TabIndex        =   3
      Top             =   3240
      Width           =   4335
   End
   Begin VB.Label lblVersion 
      BackColor       =   &H00A481AA&
      BackStyle       =   0  'Transparent
      Caption         =   "lblVersion"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   510
      Left            =   480
      TabIndex        =   2
      Top             =   2400
      Width           =   4335
   End
   Begin VB.Label lblProduct 
      BackColor       =   &H00A481AA&
      BackStyle       =   0  'Transparent
      Caption         =   "lblProduct"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   510
      Left            =   480
      TabIndex        =   1
      Top             =   1800
      Width           =   4335
   End
   Begin VB.Label lblDescription 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "lblProduct Description"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   480
      TabIndex        =   0
      Top             =   2880
      Width           =   3960
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1
Private Const HWND_TOP = 0
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2

Private Const WM_QUERYNEWPALETTE = &H30F
Private Const WM_PALETTECHANGED = &H311
Private Const WM_PALETTEISCHANGING = &H310

Private Declare Function SelectPalette Lib "gdi32" (ByVal hdc As Long, ByVal hPalette As Long, ByVal bForceBackground As Long) As Long
Private Declare Function RealizePalette Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function InvalidateRect Lib "user32" (ByVal hWnd As Long, ByVal lpRect As Long, ByVal bErase As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hdc As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private m_hPal As Long

Private Sub Form_Load()
  On Error Resume Next
    
  lblMessage.Caption = ""
  lblProduct.Caption = UCASE$(App.Title)
  lblDescription.Caption = App.Comments
  lblVersion.Caption = "Version " & GetVersionString(True)

  m_hPal = CreateAAHPal
  prgStartup.hPal = m_hPal
  HOOK.hWnd = Me.hWnd
  HOOK.Messages(WM_PALETTECHANGED) = True
  HOOK.Messages(WM_QUERYNEWPALETTE) = True
  Call SetWindowPos(Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE)
End Sub

Public Property Let Message(ByVal Msg As String)
  On Error Resume Next
  Me.lblMessage = Msg
End Property

Public Property Let HideProgressBar(ByVal NewValue As Boolean)
  On Error Resume Next
  Me.prgStartup.Visible = NewValue
End Property

Public Sub InitProgressBar()
  On Error Resume Next
  Me.prgStartup.Visible = True
  Me.prgStartup.Min = 0
  Me.prgStartup.Max = 10
  Me.prgStartup.value = 1
End Sub

Public Sub IncrementProgressBar(Optional ByVal Finish As Boolean)
  On Error Resume Next
  If Finish Then
    Me.prgStartup.value = 10
  ElseIf Me.prgStartup.value < 10 Then
    Me.prgStartup.value = Me.prgStartup.value + 1
  End If
End Sub

Private Sub Form_Terminate()
  If m_hPal <> 0 Then Call DeleteObject(m_hPal)
End Sub

Private Sub HOOK_WndProc(Discard As Boolean, MsgReturn As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long)
  Dim hdc As Long
  Dim hOldPal As Long
  Dim iRet As Long
  Dim newCol As Long
  
  Select Case Msg
    Case WM_PALETTECHANGED
      If wParam = Me.hWnd Then
        MsgReturn = 0
        Discard = True
        Exit Sub
      Else
        GoTo QUERY_PAL
      End If
      
    Case WM_QUERYNEWPALETTE
QUERY_PAL:
      hdc = GetDC(hWnd)
      hOldPal = SelectPalette(hdc, m_hPal, 0)
      newCol = RealizePalette(hdc)
      iRet = (newCol <> 0)
      If (iRet) Then
        Call InvalidateRect(hWnd, 0, 1)
      End If
      Call SelectPalette(hdc, hOldPal, 1)
      Call RealizePalette(hdc)
      Call ReleaseDC(hWnd, hdc)
      If iRet Then MsgReturn = 1
      Discard = True
  End Select
End Sub

