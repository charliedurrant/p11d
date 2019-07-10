VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ParseColorTextBoxFontHeight"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'cad new class
Implements IParseItem

Private Property Get IParseItem_ActionStr() As String
  IParseItem_ActionStr = "COLORTBFONTHEIGHT"
End Property

Private Property Get IParseItem_ActionParamCount() As Long
  IParseItem_ActionParamCount = 5
End Property
'filled box with string L, R, C
Private Function IParseItem_OnAction(psett As IParseSettings, objCurItem As Variant, pre As String, ByVal paramcount As Long, params() As String) As PARSESTATUS
  Dim ForeColorOld As Long, FillColorOld As Long
  Dim XWidth As Single, TWidth As Single, x0 As Single
  Dim YHeight  As Single, THeight  As Single, y0 As Single
  Dim sngWidth As Single, sngHeight As Single
  Dim sText As String
  Dim iBackColor As Long, iForeColor As Long
  Dim sAlign As String
  Dim at As ALIGNMENT_TYPE
  Dim sngCurrentY As Single, sngNextY As Single
  
  Call xSet("COLORTBFONTHEIGHT")

  sngWidth = CSng(params(1))
  sText = params(2)
  iForeColor = params(3)
  iBackColor = params(4)
  sAlign = params(5) 'L,R,C'
  
  '1 = width %
  '2 = text
  '3 = forecolor
  '4 = backcolor
  'alignment
  
  sAlign = LCase$(sAlign)
  Select Case sAlign
    Case "l"
      at = ALIGN_LEFT
    Case "c"
      at = ALIGN_CENTER
    Case Else
      at = ALIGN_LEFT
  End Select
  
  sngHeight = (ReportControl.fStyle.LineHeight / ReportControl.PageHeight) * 100!
  If IsExportHTML Then
    Call HTMLBox(sText, GetHexHTMLLong(iBackColor), GetHexHTMLLong(iForeColor), sngHeight, sngWidth, at, True)
  Else
    ForeColorOld = ProcessForeColorRGB(iForeColor)
    FillColorOld = ProcessFillColorRGB(iBackColor)
    
    Call DisplayTextBox(sText, sngWidth, sngHeight, True, at)
    
    Call ProcessForeColorRGB(ForeColorOld)
    Call ProcessFillColorRGB(FillColorOld)
  End If
  Call xReturn("COLORTBFONTHEIGHT")
End Function





