VERSION 5.00
Begin VB.UserControl FWCtrl 
   ClientHeight    =   4425
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4980
   BeginProperty Font 
      Name            =   "Courier"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   4425
   ScaleWidth      =   4980
   Begin atc2ADOIM.FWSlider FWSlider1 
      Height          =   3492
      Index           =   0
      Left            =   120
      TabIndex        =   5
      Top             =   480
      Visible         =   0   'False
      Width           =   132
      _ExtentX        =   238
      _ExtentY        =   6165
   End
   Begin VB.PictureBox Picture3 
      Enabled         =   0   'False
      Height          =   2292
      Left            =   360
      ScaleHeight     =   2235
      ScaleWidth      =   4035
      TabIndex        =   3
      Top             =   1080
      Width           =   4092
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   2412
         Left            =   -120
         ScaleHeight     =   2415
         ScaleWidth      =   4215
         TabIndex        =   4
         Top             =   -120
         Width           =   4212
      End
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   612
      Left            =   360
      ScaleHeight     =   615
      ScaleWidth      =   4095
      TabIndex        =   2
      Top             =   480
      Width           =   4092
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   2292
      Left            =   4440
      TabIndex        =   1
      Top             =   1080
      Width           =   252
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   252
      Left            =   360
      TabIndex        =   0
      Top             =   3360
      Width           =   4095
   End
End
Attribute VB_Name = "FWCtrl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private zz As Long
Private w As Long
Private tpp As Long
Private h As Long
Private CharBreak As String
Private CharWidth As Long
Private CharDispWidth As Long
Private CharFirst As Long
Private LineFirst As Long
Private LineLength As Long
Private LineDispLength As Long
Private LineText() As String
Private FontName As String
Private FontSize As Long
Private Pointers As Boolean

Private mOriginalWidth As Long
Private mOriginalHeight As Long
Private mNoRefresh As Boolean

Public Property Let UsePointers(UseP As String)
  Pointers = UseP
End Property

Public Property Get UsePointers() As String
  UsePointers = Pointers
End Property

Public Property Let FWFontName(FName As String)
  FontName = FName
End Property

Public Property Get FWFontName() As String
  FWFontName = FontName
End Property

Public Property Let FWFontSize(FSize As String)
  FontSize = FSize
End Property

Public Property Get FWFontSize() As String
  FWFontSize = FontSize
End Property

Public Property Let BreakString(BS As String)
  CharBreak = BS
End Property

Public Property Get BreakString() As String
  BreakString = CharBreak
End Property

Public Property Let OriginalWidth(Value As Long)
  mOriginalWidth = Value
End Property

Public Property Get OriginalWidth() As Long
  OriginalWidth = mOriginalWidth
End Property

Public Property Let OriginalHeight(Value As Long)
  mOriginalHeight = Value
End Property

Public Property Get OriginalHeight() As Long
  OriginalHeight = mOriginalHeight
End Property

Public Property Let LinesCopied(LL As Long)
  LineLength = LL
End Property

Public Property Get LinesCopied() As Long
  LinesCopied = LineLength
End Property

Public Property Let CharsWide(CW As Long)
  CharWidth = CW
End Property

Public Property Get CharsWide() As Long
  CharsWide = CharWidth
End Property


Public Property Get Lines() As Variant
  
  Dim i As Long, ii As Long, NF As Long
  Dim Start() As Long

  Mid$(CharBreak, Len(CharBreak), 1) = "!"

  NF = 0
  i = 0
  Do
    ii = i
    i = InStr(ii + 1, CharBreak, "!")
    If i = 0 Then Exit Do
    NF = NF + 1
    ReDim Preserve Start(2, NF)
    Start(1, NF) = ii + 1
    Start(2, NF) = i - ii
  Loop

  Lines = Start
  
End Property

Public Property Let LinesIn(LI As Variant)
  Dim i As Long, p As Long
  
  ReDim LineText(1 To UBound(LI))
  For i = 1 To UBound(LI)
    LineText(i) = LI(i)
    p = 1
    Do
      p = InStr(p, LineText(i), vbTab)
      If p > 0 Then
        Mid$(LineText(i), p) = Chr$(127)
        p = p + 1
      End If
    Loop Until p = 0
  Next i
End Property

Private Sub HScroll1_Change()
  CharFirst = HScroll1.Value
  Call FillPictureBox
End Sub

Private Sub FWSlider1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
  
  x = x + FWSlider1(Index).Left

  Call UserControl_MouseDown(Button, Shift, x, y)

End Sub

Private Sub FWSlider1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

  x = x + FWSlider1(Index).Left

  Call UserControl_MouseMove(Button, Shift, x, y)

End Sub

Private Sub FWSlider1_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
  
  x = x + FWSlider1(Index).Left

  Call UserControl_MouseUp(Button, Shift, x, y)

End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  
  Dim z As Long
   
  z = CInt((x - Picture3.Left - 2 * tpp) / w)
  
  If z > 0 And z <= CharDispWidth Then
    If Button = 1 And Pointers = True Then
      FWSlider1(z).Visible = True
      Mid$(CharBreak, CharFirst + z - 1) = "!"
      zz = z
    End If
    If Button = 2 And Pointers = True Then
      FWSlider1(z).Visible = False
      Mid$(CharBreak, CharFirst + z - 1) = "."
      zz = -999
    End If
    Call FillPictureBox2
  End If

End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  
  Dim z As Long
   
  z = CInt((x - Picture3.Left - 2 * tpp) / w)
  
  If z > 0 And z <= CharDispWidth And zz <> -999 Then
    If Button = 1 Then
      FWSlider1(zz).Left = Picture3.Left + (z - 1) * w + 2 * tpp
    End If
  End If
  
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
  
  Dim z As Long
   
  z = CInt((x - Picture3.Left - 2 * tpp) / w)
  
  If z < 1 Then z = 1
  If z > CharDispWidth Then z = CharDispWidth

  If Button = 1 And Pointers = True Then
    If zz <> -999 And zz <> z Then
      FWSlider1(zz).Visible = False
      Mid$(CharBreak, CharFirst + zz - 1) = "."
      FWSlider1(zz).Left = Picture3.Left + (zz - 1) * w + 2 * tpp
      FWSlider1(z).Visible = True
      Mid$(CharBreak, CharFirst + z - 1) = "!"
      Call FillPictureBox2
    End If
  End If

End Sub

Private Sub UserControl_Resize()

  If Not mNoRefresh Then Call Refresh
  
End Sub

Private Sub UserControl_Show()

  Call Refresh

End Sub

Public Sub ClearBreaks()

  CharBreak = String$(CharWidth, ".")
  Call Refresh

End Sub

Public Sub RestoreOriginalSize()
  mNoRefresh = True
  UserControl.Width = mOriginalWidth
  UserControl.Height = mOriginalHeight
  mNoRefresh = False
End Sub

Public Sub Refresh()
  Dim i As Long
  Dim H0 As Long, W0 As Long
  
  On Error GoTo Refresh_Err
  Call xSet("Refresh")
  
  If Trim$(FontName) = "" Or FontSize <= 0 Then
    FontName = "Courier"
    FontSize = 10
  End If
  
  With UserControl
    
    H0 = .Height
    W0 = .Width
  
    .Font.Name = FontName
    .Font.Size = FontSize
    w = .TextWidth("W")
  
    If w <> .TextWidth("l") Then
      FontName = "Courier"
      FontSize = 10
      .Font.Name = FontName
      .Font.Size = FontSize
      w = .TextWidth("W")
    End If
  
    h = .TextHeight("X")
    tpp = Screen.TwipsPerPixelX

    'CharDispWidth = Picture1.Width / w
    CharDispWidth = (.Width - 4 * tpp - VScroll1.Width) / w
    'LineDispLength = Picture1.Height / h
    LineDispLength = (.Height - 4 * tpp - HScroll1.Height) / h - 2
        
  End With
        
  If CharWidth <= CharDispWidth Then
    HScroll1.Visible = False
    CharDispWidth = CharWidth
  Else
    HScroll1.Visible = True
  End If
  
  If LineLength <= LineDispLength Then
    VScroll1.Visible = False
    LineDispLength = LineLength
  Else
    VScroll1.Visible = True
  End If
        
  With Picture1
    .Font.Name = FontName
    .Font.Size = FontSize
    .Width = CharDispWidth * w
    .Height = LineDispLength * h
    .Top = 0
    .Left = 0
  End With
  With Picture2
    .Font.Name = FontName
    .Font.Size = FontSize
    .BackColor = UserControl.BackColor
    .Width = Picture1.Width + 2 * tpp
    .Height = 2 * h
    .Left = 0 + 2 * tpp
    .Top = 0
  End With
  With Picture3
    .Left = 0
    .Width = Picture1.Width + 4 * tpp
    .Height = Picture1.Height + 4 * tpp
    .Top = Picture2.Top + Picture2.Height
  End With
  
  '  Create Pointers
  If Pointers = True Then
    For i = FWSlider1.Count To CharDispWidth
      Load FWSlider1(FWSlider1.Count)
      With FWSlider1(FWSlider1.Count - 1)
        .Left = Picture3.Left + (i - 1) * w + 2 * tpp
        .Top = Picture2.Top + 0.75 * Picture2.Height
        .Height = Picture3.Height + 0.25 * Picture2.Height
        .Visible = False
        .ZOrder
      End With
    Next i
  End If
      
  With HScroll1
    .Height = h
    .Left = Picture3.Left
    .Width = Picture3.Width
    .Top = Picture3.Top + Picture3.Height
    .Min = 1
    .Max = Min((CharWidth - CharDispWidth + 1), 32000)
    .LargeChange = 10
  End With
  
  With VScroll1
    .Width = h
    .Height = Picture3.Height
    .Left = Picture3.Left + Picture3.Width
    .Top = Picture3.Top
    .Min = 1
    .Max = LineLength - LineDispLength + 1
    If Int(LineLength / 2) > 0 Then
      .LargeChange = Int(LineLength / 2)
    Else
      .LargeChange = 1
    End If
    
  End With
  
  If CharWidth = 0 Or LineLength = 0 Then
    UserControl.Height = H0
    UserControl.Width = W0
  Else
    UserControl.Height = Picture3.Height + Picture2.Height + HScroll1.Height
    UserControl.Width = Picture3.Width + VScroll1.Width
  End If
      
  CharFirst = 1
  LineFirst = 1
  
  zz = -999
  
  If Len(CharBreak) <> CharWidth Then
    CharBreak = String$(CharWidth, ".")
  End If
  
  Call FillPictureBox
  
Refresh_End:
  Call xReturn("Refresh")
  Exit Sub

Refresh_Err:
  Call ErrorMessage(ERR_ERROR, Err, "FWCtrl.Refresh", "ERR_FWCTRL_REFRESH", "An error occurred whilst refreshing the Fixed Width Control." & vbCrLf & "(CharWidth = " & Trim$(CStr(CharWidth)))
  Resume Refresh_End
End Sub

'Public Sub Refresh()
'
'  Dim i As Long
'  Dim H0 As Long, W0 As Long
'
'  If Trim$(FontName) = "" Or FontSize <= 0 Then
'    FontName = "Courier"
'    FontSize = 10
'  End If
'
'  With UserControl
'
'    H0 = .Height
'    W0 = .Width
'
'    .Font.Name = FontName
'    .Font.Size = FontSize
'    w = .TextWidth("W")
'
'    If w <> .TextWidth("l") Then
'      FontName = "Courier"
'      FontSize = 10
'      .Font.Name = FontName
'      .Font.Size = FontSize
'      w = .TextWidth("W")
'    End If
'
'    h = .TextHeight("X")
'    tpp = Screen.TwipsPerPixelX
'
'    'CharDispWidth = Picture1.Width / w
'    CharDispWidth = (.Width - 4 * tpp - VScroll1.Width) / w
'    'LineDispLength = Picture1.Height / h
'    LineDispLength = (.Height - 4 * tpp - HScroll1.Height) / h - 2
'
'  End With
'
'  If CharWidth <= CharDispWidth Then
'    HScroll1.Visible = False
'    CharDispWidth = CharWidth
'  Else
'    HScroll1.Visible = True
'  End If
'
'  If LineLength <= LineDispLength Then
'    VScroll1.Visible = False
'    LineDispLength = LineLength
'  Else
'    VScroll1.Visible = True
'  End If
'
'  With Picture1
'    .Font.Name = FontName
'    .Font.Size = FontSize
'    .Width = CharDispWidth * w
'    .Height = LineDispLength * h
'    .Top = 0
'    .Left = 0
'  End With
'  With Picture2
'    .Font.Name = FontName
'    .Font.Size = FontSize
'    .BackColor = UserControl.BackColor
'    .Width = Picture1.Width + 2 * tpp
'    .Height = 2 * h
'    .Left = 0 + 2 * tpp
'    .Top = 0
'  End With
'  With Picture3
'    .Left = 0
'    .Width = Picture1.Width + 4 * tpp
'    .Height = Picture1.Height + 4 * tpp
'    .Top = Picture2.Top + Picture2.Height
'  End With
'
'  '  Create Pointers
'  If Pointers = True Then
'    For i = FWSlider1.Count To CharDispWidth
'      Load FWSlider1(FWSlider1.Count)
'      With FWSlider1(FWSlider1.Count - 1)
'        .Left = Picture3.Left + (i - 1) * w + 2 * tpp
'        .Top = Picture2.Top + 0.75 * Picture2.Height
'        .Height = Picture3.Height + 0.25 * Picture2.Height
'        .Visible = False
'        .ZOrder
'      End With
'    Next i
'  End If
'
'  With HScroll1
'    .Height = h
'    .Left = Picture3.Left
'    .Width = Picture3.Width
'    .Top = Picture3.Top + Picture3.Height
'    .Min = 1
'    .Max = Min((CharWidth - CharDispWidth + 1), 32000)
'    .LargeChange = 10
'  End With
'
'  With VScroll1
'    .Width = h
'    .Height = Picture3.Height
'    .Left = Picture3.Left + Picture3.Width
'    .Top = Picture3.Top
'    .Min = 1
'    .Max = LineLength - LineDispLength + 1
'    If Int(LineLength / 2) > 0 Then
'      .LargeChange = Int(LineLength / 2)
'    Else
'      .LargeChange = 1
'    End If
'
'  End With
'
'  If CharWidth = 0 Or LineLength = 0 Then
'    UserControl.Height = H0
'    UserControl.Width = W0
'  Else
'    UserControl.Height = Picture3.Height + Picture2.Height + HScroll1.Height
'    UserControl.Width = Picture3.Width + VScroll1.Width
'  End If
'
'  CharFirst = 1
'  LineFirst = 1
'
'  zz = -999
'
'  If Len(CharBreak) <> CharWidth Then
'    CharBreak = String$(CharWidth, ".")
'  End If
'
'  Call FillPictureBox
'
'End Sub

Private Sub UserControl_Terminate()
    Dim i As Long
    For i = FWSlider1.Count - 1 To 1 Step -1
        Unload FWSlider1(i)
    Next i
End Sub

Private Sub VScroll1_Change()

  LineFirst = VScroll1.Value

  Call FillPictureBox

End Sub

Private Sub FillPictureBox()
  
  Dim i As Long
  
  Picture1.Cls

  For i = 1 To LineDispLength
    Picture1.Print Mid$(LineText(LineFirst + i - 1), CharFirst, CharDispWidth)
  Next i

  Call FillPictureBox2
  
End Sub

Private Sub FillPictureBox2()
  
  Dim i As Long, j As Long, ii As Long
  Dim x As String, jj As String
  Dim hh As Single
  
  x = String$(CharDispWidth + 10, " ")
  
  For i = 1 To CharDispWidth
    j = CharFirst + i - 1
    jj = Trim$(Str$(j))
    If j / 10 = Int(j / 10) Or j = 1 Then
      ii = i - Fix((Len(jj) - 1) / 2)
      If ii > 0 Then
        Mid$(x, ii, Len(jj) + 1) = jj & " "
        If ii + Len(jj) > CharDispWidth Then
          Mid$(x, CharDispWidth - Len(jj) + 1, Len(jj) + 1) = jj & " "
        End If
      Else
        Mid$(x, 1, Len(jj) + 1) = jj & " "
      End If
    End If
  Next i
  
  Picture2.Cls
  Picture2.Print x
  Picture2.Line (0, 1.5 * h)-(Picture2.Width, 1.5 * h), vbBlack
  
  For i = 1 To CharDispWidth
    j = CharFirst + i - 1
    hh = 0.2
    If j / 5 = Int(j / 5) Then hh = 0.4
    If j / 10 = Int(j / 10) Or j = 1 Then hh = 0.6
    Picture2.Line (i * w, 1.5 * h)-(i * w, 1.5 * h - h * hh), vbBlack
  Next i
  
  If Pointers = True Then
    For i = 1 To CharDispWidth
      If Mid$(CharBreak, CharFirst + i - 1, 1) = "!" Then
        FWSlider1(i).Visible = True
      Else
        FWSlider1(i).Visible = False
      End If
    Next i
  End If
  
End Sub

