VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ParseArial6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements IParseItem

Private Property Get IParseItem_ActionStr() As String
  IParseItem_ActionStr = "Arial6"
End Property

Private Property Get IParseItem_ActionParamCount() As Long
  IParseItem_ActionParamCount = 0
End Property


Private Function IParseItem_OnAction(psett As TCSPARSE.IParseSettings, objCurItem As Variant, pre As String, params() As String) As TCSPARSE.PARSESTATUS
  
  
  Call xSet("setstylei")

  If rpt(iReport).curstylenum = i Then GoTo endsetstyle

  If (i < 0) Or (i >= numstyles) Then
    Call ErrorMessage(ERR_ERROR, "SetStylei", "ERR_STYLE", "Selected style:" & i & " only " & numstyles & " are defined")
    GoTo endsetstyle
  End If

  If rpt(iReport).rTarget = P_PRINTER Then
    If Styles(i).FontSize < 8 Then
      Printer.FontSize = Styles(i).FontSize
    End If
    Printer.FontName = GetFontMap(Styles(i).Font)
    Printer.FontSize = Styles(i).FontSize
    Printer.FontBold = Styles(i).bold
    Printer.FontUnderline = Styles(i).underline
    Printer.FontItalic = Styles(i).italic
'    Printer.FontTransparent = True
  ElseIf (rpt(iReport).rTarget = P_PREVIEWPAGE) Or (rpt(iReport).rTarget = P_PReview) Or (rpt(iReport).rTarget = P_CONFIG) Then
    If (Styles(i).FontSize * (mZoom / 100)) < 8 Then
      Target.FontSize = Styles(i).FontSize * (mZoom / 100)
    End If
    Target.FontName = GetFontMap(Styles(i).Font)
    Target.FontSize = Styles(i).FontSize * (mZoom / 100)
    Target.FontBold = Styles(i).bold
    Target.FontUnderline = Styles(i).underline
    Target.FontItalic = Styles(i).italic
  End If
  rpt(iReport).curstyle = Styles(i).Name
  rpt(iReport).curstylenum = i

endsetstyle:
  Call xReturn("SetStylei")
  
    
    
End Sub
