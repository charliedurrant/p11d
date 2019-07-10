VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FontStyle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Name As String
Private mSize As Single
Public bold As Boolean
Public Italic As Boolean
Public UnderLine As Boolean
Public StrikeThrough As Boolean
Public Align As ALIGNMENT_TYPE
Public FontType As FONT_TYPE
Private mFontHeight As Single

Private Sub Class_Initialize()
  mFontHeight = -1
End Sub

Private Function GetTextHeight() As Single
  Dim cTarget As REPORT_TARGET, org As FontStyle
  Const str As String = "Ay"
     
  cTarget = ReportControl.rTarget
  If Not ReportControl.fStyle Is Me Then
    Set org = ReportControl.fStyle
    Set ReportControl.fStyle = Me
    ReportControl.fStyle.FontType = VALID_FONT_TYPE
  End If
  If IsPrinterAvail(False) Then
    ReportControl.rTarget = RPT_PRINTER
    Call SetFont
    GetTextHeight = Printer.ScaleX(Printer.TextHeight(str), Printer.ScaleMode, vbTwips)
  Else
    ReportControl.rTarget = RPT_CONFIG
    Call SetFont
    GetTextHeight = ReportControl.Preview.ScaleX(ReportControl.Preview.TextHeight(str), ReportControl.Preview.ScaleMode, vbTwips)
  End If
  If Not org Is Nothing Then
    Set ReportControl.fStyle = org
    ReportControl.fStyle.FontType = VALID_FONT_TYPE
    Call SetFont
  End If
  ReportControl.rTarget = cTarget
End Function
Public Property Let FontHeightInternal(ByVal NewValue As Single)
  mFontHeight = NewValue
End Property
Public Property Get FontHeight() As Single
  If mFontHeight < 0 Then
    If ReportControl.UseExactFontHeight Then
      mFontHeight = GetTextHeight()
    Else
      mFontHeight = FontHeights(mSize)
    End If
  End If
  FontHeight = mFontHeight
End Property

Public Property Get Size() As Single
  Size = mSize
End Property

Public Property Let Size(ByVal NewValue As Single)
  If NewValue <> mSize Then mFontHeight = -1
  mSize = NewValue
End Property
