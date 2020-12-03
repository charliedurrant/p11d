Attribute VB_Name = "HTML"
Option Explicit
Private mReportTitle As String
Private HelpString As String

Public Function IsExportHTML() As Boolean
  IsExportHTML = (ReportControl.rTarget = EXPORT_HTML_IE) Or (ReportControl.rTarget = EXPORT_HTML_NETSCAPE) Or (ReportControl.rTarget = EXPORT_HTML_INTEXP5)  'km
End Function

Public Sub ConstructHTMLHeader(ByVal ReportTitle As String)
  Dim s As String
  
  ReportControl.HTML.ReportPages = 0
  mReportTitle = ReportTitle
  Call SetHelpPage
  Call ReportControl.HTML.HTMLString.Append("<HTML>")
  Call ReportControl.HTML.HTMLString.Append(vbCrLf)
  Call ReportControl.HTML.HTMLString.Append(vbTab)
  Call ReportControl.HTML.HTMLString.Append("<HEAD>")
  Call ReportControl.HTML.HTMLString.Append(vbCrLf)
  Call ReportControl.HTML.HTMLString.Append(vbTab)
  Call ReportControl.HTML.HTMLString.Append(vbTab)
  Call ReportControl.HTML.HTMLString.Append("<TITLE>")
  Call ReportControl.HTML.HTMLString.Append(vbCrLf)
  Call ReportControl.HTML.HTMLString.Append(vbTab)
  Call ReportControl.HTML.HTMLString.Append(vbTab)
  Call ReportControl.HTML.HTMLString.Append(vbTab)
  Call ReportControl.HTML.HTMLString.Append(ReportTitle)
  Call ReportControl.HTML.HTMLString.Append(vbCrLf)
  Call ReportControl.HTML.HTMLString.Append(vbTab)
  Call ReportControl.HTML.HTMLString.Append(vbTab)
  Call ReportControl.HTML.HTMLString.Append("</TITLE>")
  Call ReportControl.HTML.HTMLString.Append(vbCrLf)
  Call SetStyleBlock
  Call SetScriptBlock
  Call ReportControl.HTML.HTMLString.Append("</HEAD>")
  Call ReportControl.HTML.HTMLString.Append(vbCrLf)
  Call ReportControl.HTML.HTMLString.Append(vbTab)
  Call ReportControl.HTML.HTMLString.Append("<BODY")
  Call ReportControl.HTML.HTMLString.Append(vbTab)
  Call ReportControl.HTML.HTMLString.Append(" leftmargin=0 bottommargin=0 rightmargin=0 topmargin=0")
  If ReportControl.rTarget = EXPORT_HTML_IE Then
    Call ReportControl.HTML.HTMLString.Append(" onload='SetScreen()' onresize='SetScreen()' scroll=NO ")
  Else
    If ReportControl.rTarget = EXPORT_HTML_NETSCAPE Then  'km
      Call ReportControl.HTML.HTMLString.Append(" onresize='Resize()' ")
    End If
  End If
  
  Call ReportControl.HTML.HTMLString.Append(">")
  Call ReportControl.HTML.HTMLString.Append(vbCrLf)
  If ReportControl.rTarget = EXPORT_HTML_IE Then
    Call ReportControl.HTML.HTMLString.Append("<DIV ID='WholeThing' style='display: none;'>")
    Call ReportControl.HTML.HTMLString.Append(vbCrLf)
    Call SetNavDiv
    Call ReportControl.HTML.HTMLString.Append("<DIV ID='Holder' style='HEIGHT=100%; overflow: scroll '>")
    Call OpenPageDiv(1, True)
    ReportControl.HTML.CurrentY = NAV_HEADER
  Else
    ReportControl.HTML.CurrentY = 0
  End If
End Sub

Public Sub ConstructHTMLFooter()
  Dim s As String
  
  If ReportControl.HTML.OpenDiv Then
    If (ReportControl.rTarget = EXPORT_HTML_IE) Or (ReportControl.rTarget = EXPORT_HTML_INTEXP5) Then  'km
      Call ReportControl.HTML.HTMLString.Append("'></DIV>")
    Else
      Call ReportControl.HTML.HTMLString.Append("'></LAYER>")
    End If
  End If
  Call ReportControl.HTML.HTMLString.Append(vbTab)
  If (ReportControl.rTarget = EXPORT_HTML_IE) Or (ReportControl.rTarget = EXPORT_HTML_INTEXP5) Then 'km
    Call ReportControl.HTML.HTMLString.Append("</DIV>")
    Call ReportControl.HTML.HTMLString.Append("</DIV>")
  End If
  Call ReportControl.HTML.HTMLString.Append("</BODY>")
  Call ReportControl.HTML.HTMLString.Append(vbCrLf)
  Call ReportControl.HTML.HTMLString.Append("</HTML>")
  Call ReportControl.HTML.HTMLString.Append(vbCrLf)
  Call SetMaxPage
End Sub

Public Sub SetOpenDiv()
  If Not ReportControl.HTML.OpenDiv Then
    If (ReportControl.rTarget = EXPORT_HTML_IE) Or (ReportControl.rTarget = EXPORT_HTML_INTEXP5) Then 'km
      Call ReportControl.HTML.HTMLString.Append("<DIV STYLE=""")
    Else
      Call ReportControl.HTML.HTMLString.Append("<LAYER STYLE=""")
    End If
    ReportControl.HTML.OpenDiv = True
    ReportControl.HTML.CloseDiv = False
    ReportControl.HTML.HTMLFontSet = False
    ReportControl.HTML.XSetHTML = False
    ReportControl.HTML.YSetHTML = False
    ReportControl.HTML.Position = False
  End If
End Sub

Public Sub SetFontHTML()
  Dim TDecoration As Boolean
  
  If ReportControl.HTML.HTMLFontSet Then Exit Sub
  ReportControl.HTML.HTMLFontSet = True
  Call SetOpenDiv
  With ReportControl.fStyle
    Call ReportControl.HTML.HTMLString.Append("font-family: ")
    Call ReportControl.HTML.HTMLString.Append(.Name)
    Call ReportControl.HTML.HTMLString.Append("; ")
    Call ReportControl.HTML.HTMLString.Append("font-size: ")
    Call ReportControl.HTML.HTMLString.Append(.Size)
    Call ReportControl.HTML.HTMLString.Append("pt; ")
    If .bold Then
      Call ReportControl.HTML.HTMLString.Append("font-weight: bold; ")
    End If
    If .Italic Then
      Call ReportControl.HTML.HTMLString.Append("font-style: italic; ")
    End If
    If .UnderLine Then
      If Not TDecoration Then
        Call ReportControl.HTML.HTMLString.Append("text-decoration: ")
        TDecoration = True
      End If
      Call ReportControl.HTML.HTMLString.Append("underline ")
    End If
    If .StrikeThrough Then
      If Not TDecoration Then
        Call ReportControl.HTML.HTMLString.Append("text-decoration: ")
        TDecoration = True
      End If
      Call ReportControl.HTML.HTMLString.Append("line-through ")
    End If
    If TDecoration Then Call ReportControl.HTML.HTMLString.Append("; ")
  End With
End Sub

Public Sub CloseOpenDiv()
  If ReportControl.HTML.OpenDiv Then
    If (ReportControl.rTarget = EXPORT_HTML_IE) Or (ReportControl.rTarget = EXPORT_HTML_INTEXP5) Then 'km
      Call ReportControl.HTML.HTMLString.Append("""")
    End If
    Call ReportControl.HTML.HTMLString.Append(">")
    ReportControl.HTML.OpenDiv = False
    ReportControl.HTML.CloseDiv = True
  End If
End Sub

Public Sub MakeLeftAdjustment()

  'KM - the idea is to stretch the html page out as much as possible for printing
  '     purposes in IE5 mode
  Dim largeAdjNeg As Double
  Dim smallAdjNeg As Double
  Dim largeAdjPos As Double
  Dim smallAdjPos As Double
  
  Select Case (ReportControl.HTML.CurrentX / ReportControl.PageWidth)
    Case Is < 0.1
      'adjust object to left by 3%
      largeAdjNeg = (ReportControl.HTML.CurrentX / ReportControl.PageWidth) - 0.02 '- 0.03
      If largeAdjNeg <= 0 Then
        Call ReportControl.HTML.HTMLString.Append(0)
      Else
        Call ReportControl.HTML.HTMLString.Append(largeAdjNeg * 100)
      End If
    Case Is < 0.5
      'adjust object to left by 1%
      smallAdjNeg = (ReportControl.HTML.CurrentX / ReportControl.PageWidth) - 0.01 ' - 0.01
      Call ReportControl.HTML.HTMLString.Append(smallAdjNeg * 100)
    Case Is > 0.9
      'adjust object to right by 3%
      largeAdjPos = (ReportControl.HTML.CurrentX / ReportControl.PageWidth) + 0.02 '+ 0.03
      Call ReportControl.HTML.HTMLString.Append(largeAdjPos * 100)
    Case Else
      'adjust object to right by 1%
      smallAdjPos = (ReportControl.HTML.CurrentX / ReportControl.PageWidth) + 0.01  ' +0.01
      Call ReportControl.HTML.HTMLString.Append(smallAdjPos * 100)
  End Select
  
End Sub

Public Sub SetXHTML()
  If Not ReportControl.HTML.XSetHTML Then
    If (ReportControl.rTarget = EXPORT_HTML_IE) Or (ReportControl.rTarget = EXPORT_HTML_INTEXP5) Then 'km
      Call SetPositionHTML
      Call ReportControl.HTML.HTMLString.Append("LEFT: ")
      If (ReportControl.rTarget = EXPORT_HTML_INTEXP5) Then
        Call MakeLeftAdjustment
      Else
        Call ReportControl.HTML.HTMLString.Append((ReportControl.HTML.CurrentX / ReportControl.PageWidth) * 100)
      End If
      Call ReportControl.HTML.HTMLString.Append("%; ")
    Else
      Call ReportControl.HTML.HTMLString.Append(" LEFT= ")
      Call ReportControl.HTML.HTMLString.Append(RoundN((ReportControl.HTML.CurrentX / ReportControl.PageWidth) * 100, 0))
      Call ReportControl.HTML.HTMLString.Append("% ")
    End If
    ReportControl.HTML.XSetHTML = True
  End If
End Sub

Public Sub SetYHTML()
  If Not ReportControl.HTML.YSetHTML Then
    If (ReportControl.rTarget = EXPORT_HTML_IE) Or (ReportControl.rTarget = EXPORT_HTML_INTEXP5) Then 'km
      Call SetPositionHTML
      Call ReportControl.HTML.HTMLString.Append("TOP: ")
      Call ReportControl.HTML.HTMLString.Append(ReportControl.HTML.CurrentY / Screen.TwipsPerPixelY)
      Call ReportControl.HTML.HTMLString.Append("px; ")
    Else
      Call ReportControl.HTML.HTMLString.Append("TOP= ")
      Call ReportControl.HTML.HTMLString.Append(ReportControl.HTML.CurrentY / Screen.TwipsPerPixelY)
      Call ReportControl.HTML.HTMLString.Append("px ")
    End If
    ReportControl.HTML.YSetHTML = True
  End If
End Sub

Public Sub SetWHTML(Width As Single)
  Call SetPositionHTML
  If (ReportControl.rTarget = EXPORT_HTML_IE) Or (ReportControl.rTarget = EXPORT_HTML_INTEXP5) Then 'km
    Call ReportControl.HTML.HTMLString.Append("WIDTH: ")
    Call ReportControl.HTML.HTMLString.Append((Width / ReportControl.PageWidth) * 100)
    Call ReportControl.HTML.HTMLString.Append("%; ")
  Else
    Call ReportControl.HTML.HTMLString.Append("WIDTH= ")
    Call ReportControl.HTML.HTMLString.Append((Width / ReportControl.PageWidth) * 100)
    Call ReportControl.HTML.HTMLString.Append("% ")
  End If
End Sub

Public Sub SetHHTML(Height As Single)
  Call SetPositionHTML
  If (ReportControl.rTarget = EXPORT_HTML_IE) Or (ReportControl.rTarget = EXPORT_HTML_INTEXP5) Then 'km
    Call ReportControl.HTML.HTMLString.Append("HEIGHT: ")
    Call ReportControl.HTML.HTMLString.Append(Height / Screen.TwipsPerPixelY)
    Call ReportControl.HTML.HTMLString.Append("px; ")
  Else
    Call ReportControl.HTML.HTMLString.Append("HEIGHT= ")
    Call ReportControl.HTML.HTMLString.Append(Height / Screen.TwipsPerPixelY)
    Call ReportControl.HTML.HTMLString.Append("px ")
  End If
End Sub

Public Sub CloseDiv()
  If ReportControl.HTML.CloseDiv Then
    ReportControl.HTML.HTMLFontSet = False
    ReportControl.HTML.OpenDiv = False
    ReportControl.HTML.XSetHTML = False
    ReportControl.HTML.YSetHTML = False
    ReportControl.HTML.CloseDiv = False
    ReportControl.HTML.Position = False
    If (ReportControl.rTarget = EXPORT_HTML_IE) Or (ReportControl.rTarget = EXPORT_HTML_INTEXP5) Then 'km
      Call ReportControl.HTML.HTMLString.Append("</DIV>")
    Else
      Call ReportControl.HTML.HTMLString.Append("</LAYER>")
    End If
    Call ReportControl.HTML.HTMLString.Append(vbCrLf)
  End If
End Sub

Public Sub SetNewHTMLPage()
  If ReportControl.rTarget = EXPORT_HTML_NETSCAPE Then
    Call ReportControl.HTML.HTMLString.Append("""")
  End If
  Call CloseOpenDiv
  Call CloseDiv
  If (ReportControl.rTarget = EXPORT_HTML_IE) Or (ReportControl.rTarget = EXPORT_HTML_INTEXP5) Then 'km
    If Not (ReportControl.rTarget = EXPORT_HTML_INTEXP5) Then ReportControl.HTML.CurrentY = NAV_HEADER  'km
    Call ReportControl.HTML.HTMLString.Append("</DIV>")
  Else
    ReportControl.HTML.CurrentY = ReportControl.HTML.CurrentY + (3 * ReportControl.fStyle.FontHeight)
  End If
  Call OpenPageDiv(ReportControl.CurPage + 1, False)
  ReportControl.HTML.ReportPages = ReportControl.HTML.ReportPages + 1
End Sub

Public Sub HTMLBox(Text As String, BColor As String, FColor As String, Height As Single, Width As Single, Align As ALIGNMENT_TYPE, FillBox As Boolean, Optional ByVal Absolute As Boolean = False)
  Dim THeight As Single
  Dim TWidth As Single
  Dim PHeight As Single
  
  Call PushCoord(PUSH_BOTH)
  Call CloseDiv
  Call SetOpenDiv
  If Not Absolute Then
    Width = GetPercent(Width, ReportControl.PageWidth)
    Height = GetPercent(Height, ReportControl.PageHeight)
  End If
  THeight = GetTextHeight(Text)
  If FillBox Then
    Call ReportControl.HTML.HTMLString.Append("BACKGROUND-COLOR:")
    Call ReportControl.HTML.HTMLString.Append(BColor)
    Call ReportControl.HTML.HTMLString.Append("; ")
  Else
    If (ReportControl.rTarget = EXPORT_HTML_IE) Or (ReportControl.rTarget = EXPORT_HTML_INTEXP5) Then  'km
      Call ReportControl.HTML.HTMLString.Append("BORDER-BOTTOM: ")
      Call ReportControl.HTML.HTMLString.Append(FColor)
      Call ReportControl.HTML.HTMLString.Append(" 1px solid; BORDER-LEFT: ")
      Call ReportControl.HTML.HTMLString.Append(FColor)
      Call ReportControl.HTML.HTMLString.Append(" 1px solid; BORDER-RIGHT: ")
      Call ReportControl.HTML.HTMLString.Append(FColor)
      Call ReportControl.HTML.HTMLString.Append(" 1px solid; BORDER-TOP: ")
      Call ReportControl.HTML.HTMLString.Append(FColor)
      Call ReportControl.HTML.HTMLString.Append(" 1px solid; ")
    End If
  End If
  If ReportControl.rTarget = EXPORT_HTML_NETSCAPE Then
    Call ReportControl.HTML.HTMLString.Append("""")
  End If
  Call SetXHTML
  Call SetYHTML
  Call SetWHTML(Width)
  Call SetHHTML(Height)
  Call CloseOpenDiv
  If Len(Text) > 0 Then
    Call ReportControl.HTML.HTMLString.Append("<TABLE WIDTH=100% HEIGHT=100% border=0 cellpadding=0 cellspacing=0")
    If ReportControl.rTarget = EXPORT_HTML_NETSCAPE And FillBox Then
      Call ReportControl.HTML.HTMLString.Append(" STYLE=""BACKGROUND-COLOR:")
      Call ReportControl.HTML.HTMLString.Append(BColor)
      Call ReportControl.HTML.HTMLString.Append("; """)
    End If
    Call ReportControl.HTML.HTMLString.Append(">")
    Call ReportControl.HTML.HTMLString.Append(vbCrLf)
    Call ReportControl.HTML.HTMLString.Append("<TR>") ' WIDTH=100% HEIGHT=100%
    Call ReportControl.HTML.HTMLString.Append(vbCrLf)
    Call ReportControl.HTML.HTMLString.Append("<TD VALIGN=""middle"" style=""")
    Call ReportControl.HTML.HTMLString.Append("text-align: ")
    If Align = ALIGN_LEFT Then
      Call ReportControl.HTML.HTMLString.Append("left")
    ElseIf Align = ALIGN_CENTER Then
      Call ReportControl.HTML.HTMLString.Append("center")
    ElseIf Align = ALIGN_RIGHT Then
      'PC Fix for netscape printing
      If ReportControl.rTarget = EXPORT_HTML_NETSCAPE Then
        Call ReportControl.HTML.HTMLString.Append("left")
      Else
        Call ReportControl.HTML.HTMLString.Append("right")
      End If
    End If
    Call ReportControl.HTML.HTMLString.Append("; ")
    Call ReportControl.HTML.HTMLString.Append("COLOR: ")
    Call ReportControl.HTML.HTMLString.Append(FColor)
    Call ReportControl.HTML.HTMLString.Append("; ")
    ReportControl.HTML.OpenDiv = True
    Call SetFontHTML
    ReportControl.HTML.OpenDiv = False
    Call ReportControl.HTML.HTMLString.Append(""">")
    Call ReportControl.HTML.HTMLString.Append(vbCrLf)
    If ReportControl.rTarget = EXPORT_HTML_NETSCAPE Then
      If Not FillBox Then
        Call ReportControl.HTML.HTMLString.Append("<SPAN STYLE=""")
        Call ReportControl.HTML.HTMLString.Append("border-width: 1px; ")
        Call ReportControl.HTML.HTMLString.Append("border-style: solid; ")
        Call ReportControl.HTML.HTMLString.Append("border-color: ")
        Call ReportControl.HTML.HTMLString.Append(FColor)
        Call ReportControl.HTML.HTMLString.Append("; ")
        Call ReportControl.HTML.HTMLString.Append("width: ")
        Call ReportControl.HTML.HTMLString.Append(Width)
        Call ReportControl.HTML.HTMLString.Append("px; ")
        Call ReportControl.HTML.HTMLString.Append("height: ")
        Call ReportControl.HTML.HTMLString.Append(Height)
        Call ReportControl.HTML.HTMLString.Append("px; ")
        Call ReportControl.HTML.HTMLString.Append(""">")
      End If
      Call ReportControl.HTML.HTMLString.Append("<SPAN STYLE=""COLOR: ")
      Call ReportControl.HTML.HTMLString.Append(FColor)
      Call ReportControl.HTML.HTMLString.Append("; ")
      Call ReportControl.HTML.HTMLString.Append("text-align: ")
      If Align = ALIGN_LEFT Then
        Call ReportControl.HTML.HTMLString.Append("left")
      ElseIf Align = ALIGN_CENTER Then
        Call ReportControl.HTML.HTMLString.Append("center")
      ElseIf Align = ALIGN_RIGHT Then
        'PC Fix for netscape printing
        Call ReportControl.HTML.HTMLString.Append("left")
      End If
      Call ReportControl.HTML.HTMLString.Append("; ")
      ReportControl.HTML.OpenDiv = True
      ReportControl.HTML.HTMLFontSet = False
      Call SetFontHTML
      ReportControl.HTML.OpenDiv = False
      Call ReportControl.HTML.HTMLString.Append(""">")
      Call ReportControl.HTML.HTMLString.Append(vbCrLf)
    End If
    Call SetHTMLSpaces(Text)
    Call ReportControl.HTML.HTMLString.Append(Text)
    Call ReportControl.HTML.HTMLString.Append(vbCrLf)
    If ReportControl.rTarget = EXPORT_HTML_NETSCAPE Then
      If Not FillBox Then
        Call ReportControl.HTML.HTMLString.Append("</SPAN>")
      End If
      Call ReportControl.HTML.HTMLString.Append("</SPAN>")
      Call ReportControl.HTML.HTMLString.Append(vbCrLf)
    End If
    Call ReportControl.HTML.HTMLString.Append("</TD>")
    Call ReportControl.HTML.HTMLString.Append("</TR>")
    Call ReportControl.HTML.HTMLString.Append("</TABLE>")
  Else
    If ReportControl.rTarget = EXPORT_HTML_NETSCAPE Then
      Call ReportControl.HTML.HTMLString.Append("<TABLE WIDTH=100% HEIGHT=100% border=0 cellpadding=0 cellspacing=0")
      Call ReportControl.HTML.HTMLString.Append(" STYLE=""BACKGROUND-COLOR:")
      Call ReportControl.HTML.HTMLString.Append(BColor)
      Call ReportControl.HTML.HTMLString.Append("; """)
      Call ReportControl.HTML.HTMLString.Append(">")
      Call ReportControl.HTML.HTMLString.Append(vbCrLf)
      Call ReportControl.HTML.HTMLString.Append("<TR>")
      Call ReportControl.HTML.HTMLString.Append(vbCrLf)
      Call ReportControl.HTML.HTMLString.Append("<TD>")
      Call ReportControl.HTML.HTMLString.Append(vbCrLf)
      If Not FillBox Then
        Call ReportControl.HTML.HTMLString.Append("<SPAN STYLE=""")
        Call ReportControl.HTML.HTMLString.Append("border-width: 1px; ")
        Call ReportControl.HTML.HTMLString.Append("border-style: solid; ")
        Call ReportControl.HTML.HTMLString.Append("border-color: ")
        Call ReportControl.HTML.HTMLString.Append(FColor)
        Call ReportControl.HTML.HTMLString.Append("; ")
        Call ReportControl.HTML.HTMLString.Append("width: ")
        Call ReportControl.HTML.HTMLString.Append(Width)
        Call ReportControl.HTML.HTMLString.Append("px; ")
        Call ReportControl.HTML.HTMLString.Append("height: ")
        Call ReportControl.HTML.HTMLString.Append(Height)
        Call ReportControl.HTML.HTMLString.Append("px; ")
        Call ReportControl.HTML.HTMLString.Append(""">")
      End If
      Call ReportControl.HTML.HTMLString.Append("&nbsp")
      Call ReportControl.HTML.HTMLString.Append(vbCrLf)
      If Not FillBox Then
        Call ReportControl.HTML.HTMLString.Append("</SPAN>")
      End If
      Call ReportControl.HTML.HTMLString.Append("</TD>")
      Call ReportControl.HTML.HTMLString.Append(vbCrLf)
      Call ReportControl.HTML.HTMLString.Append("</TR>")
      Call ReportControl.HTML.HTMLString.Append(vbCrLf)
      Call ReportControl.HTML.HTMLString.Append("</TABLE>")
      Call ReportControl.HTML.HTMLString.Append(vbCrLf)
    End If
  End If
  Call CloseDiv
  Call PopCoord
End Sub

Public Sub SetPositionHTML()
  If (ReportControl.rTarget = EXPORT_HTML_IE) Or (ReportControl.rTarget = EXPORT_HTML_INTEXP5) Then 'km
    If Not ReportControl.HTML.Position Then
      Call ReportControl.HTML.HTMLString.Append("POSITION: absolute; ")
      ReportControl.HTML.Position = True
    End If
  End If
End Sub

Public Sub SetNavDiv()
  If (ReportControl.rTarget = EXPORT_HTML_IE) Or (ReportControl.rTarget = EXPORT_HTML_INTEXP5) Then 'km
    Call ReportControl.HTML.HTMLString.Append(LoadResString(104))
  End If
End Sub

Public Sub OpenPageDiv(Page As Long, Visible As Boolean)
  If (ReportControl.rTarget = EXPORT_HTML_IE) Then
    Call ReportControl.HTML.HTMLString.Append("<DIV ID='Page")
    Call ReportControl.HTML.HTMLString.Append(Page)
    Call ReportControl.HTML.HTMLString.Append("' style='display: ")
    If Visible Then
      Call ReportControl.HTML.HTMLString.Append("block")
    Else
      Call ReportControl.HTML.HTMLString.Append("none")
    End If
    Call ReportControl.HTML.HTMLString.Append("'>")
  End If
End Sub

Public Sub SetScriptBlock()
  Dim Script As String
  Dim pos As Long
  Dim TempString As String
    
  On Error GoTo SetScriptBlock_err
  If (ReportControl.rTarget = EXPORT_HTML_IE) Or (ReportControl.rTarget = EXPORT_HTML_INTEXP5) Then 'km
    Script = LoadResString(102)
    pos = InStr(1, Script, "[Help]", vbTextCompare)
    TempString = Left$(Script, pos - 1)
    TempString = TempString & HelpString
    TempString = TempString & Mid$(Script, pos + 6, Len(Script) - pos)
    Call ReportControl.HTML.HTMLString.Append(TempString)
  Else
    ReportControl.HTML.HTMLString.Append "<SCRIPT>function Resize(){ if ( isnavigator461() ) { window.location.href=window.location.href; }}"
    ReportControl.HTML.HTMLString.Append "function isnavigator461() { var an = navigator.appName; var ver = 0; if (an == ""Netscape"") { ver = parseFloat(navigator.appVersion); return (ver >= 4.61);  } return false; }"
    ReportControl.HTML.HTMLString.Append "</SCRIPT>"
  End If
  Exit Sub
  
SetScriptBlock_err:
  Err.Raise Err.Number, ErrorSource(Err, "SetScriptBlock"), Err.Description
End Sub

Public Function GetRGB(ByVal RGBval As Long, ByRef R As Integer, ByRef G As Integer, ByRef B As Integer) As Boolean
  R = RGBval \ 256 ^ (0) And 255
  G = RGBval \ 256 ^ (1) And 255
  B = RGBval \ 256 ^ (2) And 255
  GetRGB = True
End Function

Public Function GetHexHTML(ByRef R As Integer, ByRef G As Integer, ByRef B As Integer) As String
  GetHexHTML = "#" & Hex(B) + Hex(G) + Hex(R)
End Function

Public Sub SetHTMLColor(rgb As Long)
  Dim R As Integer
  Dim G As Integer
  Dim B As Integer
  
  Call GetRGB(rgb, R, G, B)
  ReportControl.HTML.FillColor = GetHexHTML(R, G, B)
End Sub

Public Sub SetHTMLSpaces(ByRef Text As String)
  Dim pos As Long
  Dim Spaces As Long
  Dim tmp As String
  Dim Wid As Single
  
  For pos = 1 To Len(Text)
    If Mid$(Text, pos, 1) = " " Then
      Spaces = Spaces + 1
    Else
      Exit For
    End If
  Next pos
  If Spaces > 0 Then
    Wid = GetTextWidth(" ")
    For pos = 1 To Spaces
      tmp = tmp & "&nbsp "
      'ReportControl.HTML.CurrentX = ReportControl.HTML.CurrentX + Wid
    Next pos
    If pos < Len(Text) Then
      tmp = tmp & Mid$(Text, pos, Len(Text) - pos + 1)
      'tmp = Mid$(Text, pos, Len(Text) - pos + 1)
    End If
    Text = tmp
  End If
End Sub

Public Sub SetHTMLLine(ByVal PercentWidth As Single, ByVal DoubleLine As Boolean, ByVal Absolute As Boolean)
  Dim x0 As Single, x1 As Single, y0 As Single
  
  Call PushCoord
  x0 = StackTopX
  If Absolute Then
    If PercentWidth >= 100 Then x0 = 1 'apf
    If PercentWidth <= -100 Then x0 = ReportControl.PageWidth
    x1 = GetRelativePercent(x0, PercentWidth, ReportControl.PageWidth)
    Call HTMLLine(Min(x1, x0), Abs(x1 - x0))
    If DoubleLine Then
      ReportControl.HTML.CurrentY = ReportControl.HTML.CurrentY + LINESPACE
      Call HTMLLine(Min(x1, x0), Abs(x1 - x0))
    End If
  Else
  
  End If
  Call PopCoord
End Sub

Private Sub HTMLLine(Left As Single, Width As Single)
  Dim FontHeight As Single
  Dim FontHeightPt As Single
  
  If (ReportControl.rTarget = EXPORT_HTML_IE) Or (ReportControl.rTarget = EXPORT_HTML_INTEXP5) Then 'km
    Call SetOpenDiv
    Call ReportControl.HTML.HTMLString.Append("BORDER-TOP: ")
    Call ReportControl.HTML.HTMLString.Append("#000000 ")
    Call ReportControl.HTML.HTMLString.Append(" 1px solid; ")
    ReportControl.HTML.CurrentX = Left
    If ReportControl.rTarget = EXPORT_HTML_NETSCAPE Then
      Call ReportControl.HTML.HTMLString.Append("""")
    End If
    
    Call SetXHTML
    Call SetYHTML
    Call SetWHTML(Width)
    Call SetHHTML(Screen.TwipsPerPixelY)
    Call CloseOpenDiv
    Call CloseDiv
  Else
'    Call SetOpenDiv
    'Call ReportControl.HTML.HTMLString.Append("border-width: 1px; ")
    'Call ReportControl.HTML.HTMLString.Append("border-style: solid; ")
    'Call ReportControl.HTML.HTMLString.Append("border-color: ")
'    Call ReportControl.HTML.HTMLString.Append(" COLOR: #000000")
'    ReportControl.HTML.CurrentX = Left
'    Call ReportControl.HTML.HTMLString.Append("""")
'    Call SetXHTML
'    Call SetYHTML
'    Call SetWHTML(Width)
'    Call SetHHTML(Screen.TwipsPerPixelY)
'    Call CloseOpenDiv
'    Call ReportControl.HTML.HTMLString.Append("<TABLE WIDTH=100% HEIGHT=100% border=0 cellpadding=0 cellspacing=0")
'    Call ReportControl.HTML.HTMLString.Append(" STYLE=""BACKGROUND-COLOR:")
'    Call ReportControl.HTML.HTMLString.Append(BColor)
'    Call ReportControl.HTML.HTMLString.Append("; "">")
    
'    Call CloseDiv
  End If
End Sub

Private Sub SetStyleBlock()
  If (ReportControl.rTarget = EXPORT_HTML_IE) Or (ReportControl.rTarget = EXPORT_HTML_INTEXP5) Then 'km
    Call ReportControl.HTML.HTMLString.Append(LoadResString(101))
  End If
End Sub

Public Sub SetHelpPage()
  Dim TempString As String
  Dim pos As Long
  Dim Pos2 As Long
  Dim m_File As Long
  
  If (ReportControl.rTarget = EXPORT_HTML_IE) Or (ReportControl.rTarget = EXPORT_HTML_INTEXP5) Then 'km
    HelpString = LoadResString(103)
    Pos2 = InStr(1, HelpString, "[Application Name]", vbTextCompare)
    TempString = Left$(HelpString, Pos2 - 1)
    TempString = TempString & AppName
    pos = InStr(1, HelpString, "[ORIENTATION]", vbTextCompare)
    TempString = TempString & Mid$(HelpString, Pos2 + 18, pos - Pos2 - 18)
    If Len(ReportControl.HTML.OrientationString) = 0 Then
      If ReportControl.Orientation = PORTRAIT Then
        TempString = TempString & "Portrait"
      Else
        TempString = TempString & "Landscape"
      End If
    Else
      TempString = TempString & ReportControl.HTML.OrientationString
    End If
    Pos2 = InStr(1, HelpString, "[TOP]", vbTextCompare)
    TempString = TempString & Mid$(HelpString, pos + 13, Pos2 - pos - 13)
    TempString = TempString & ReportControl.HTML.TopString
    pos = InStr(1, HelpString, "[BOTTOM]", vbTextCompare)
    TempString = TempString & Mid$(HelpString, Pos2 + 5, pos - Pos2 - 5)
    TempString = TempString & ReportControl.HTML.BottomString
    Pos2 = InStr(1, HelpString, "[LEFT]", vbTextCompare)
    TempString = TempString & Mid$(HelpString, pos + 8, Pos2 - pos - 8)
    TempString = TempString & ReportControl.HTML.LeftString
    pos = InStr(1, HelpString, "[RIGHT]", vbTextCompare)
    TempString = TempString & Mid$(HelpString, Pos2 + 6, pos - Pos2 - 6)
    TempString = TempString & ReportControl.HTML.RightString
    Pos2 = InStr(1, HelpString, "[Contact Details]", vbTextCompare)
    TempString = TempString & Mid$(HelpString, pos + 7, Pos2 - pos - 7)
    If Len(ReportControl.HTML.ContactString) > 0 Then
      TempString = TempString & ReportControl.HTML.ContactString
    Else
      TempString = TempString & GetStatic("Contact")
    End If
    TempString = TempString & Mid$(HelpString, Pos2 + 17, Len(HelpString) - Pos2)
    HelpString = TempString
  End If
End Sub

Public Sub SetMaxPage()
  Const MAX_PAGE As String = "[MaxPage]" ' denotes tag to replace
  
  Call InsertTag(ReportControl.HTML.HTMLString, MAX_PAGE, Left$(CStr(ReportControl.HTML.ReportPages) & ";" & Space$(Len(MAX_PAGE)), Len(MAX_PAGE)))
End Sub

Private Sub InsertTag(ByVal qsHTML As QString, ByVal tag As String, ByVal value As String)
  Dim s As String, p As Long
  
  s = qsHTML
  p = InStr(1, s, tag, vbTextCompare)
  If p > 0 Then
    Mid$(s, p, Len(tag)) = value
    qsHTML = s
  End If
End Sub

