VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ParseFont"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements IParseItem

Private Property Get IParseItem_ActionParamCount() As Long
  IParseItem_ActionParamCount = 8
End Property

Private Property Get IParseItem_ActionStr() As String
  IParseItem_ActionStr = "FONT"
End Property

Private Function IParseItem_OnAction(psett As IParseSettings, objCurItem As Variant, pre As String, ByVal paramcount As Long, params() As String) As PARSESTATUS
  Dim rcls As ReportDetails
  Dim rFld As ReportField
  Dim OnAction As PARSESTATUS
  Dim FDetails As FontDetails
  
  Set FDetails = Nothing
  
  Select Case UCase$(params(1))
    Case "REPHEADERFONT"
      Set rcls = psett.ParseParameter(2)
      Set FDetails = rcls.RepHeaderFontL
      Set FDetails = rcls.RepHeaderFontC
      Set FDetails = rcls.RepHeaderFontR
      OnAction = PARSESTATUS_INVALID
    Case "PAGEHEADERFONT"
      Set rcls = psett.ParseParameter(2)
      Set FDetails = rcls.PageHeaderFontL
      Set FDetails = rcls.PageHeaderFontC
      Set FDetails = rcls.PageHeaderFontR
      OnAction = PARSESTATUS_INVALID
    Case "PAGEFOOTERFONT"
      Set rcls = psett.ParseParameter(2)
      Set FDetails = rcls.PageFooterFontL
      Set FDetails = rcls.PageFooterFontC
      Set FDetails = rcls.PageFooterFontR
      OnAction = PARSESTATUS_INVALID
    Case "REPHEADERFONTL"
      Set rcls = psett.ParseParameter(2)
      Set FDetails = rcls.RepHeaderFontL
      OnAction = PARSESTATUS_INVALID
    Case "REPHEADERFONTC"
      Set rcls = psett.ParseParameter(2)
      Set FDetails = rcls.RepHeaderFontC
      OnAction = PARSESTATUS_INVALID
    Case "REPHEADERFONTR"
      Set rcls = psett.ParseParameter(2)
      Set FDetails = rcls.RepHeaderFontR
      OnAction = PARSESTATUS_INVALID
    Case "PAGEHEADERFONTL"
      Set rcls = psett.ParseParameter(2)
      Set FDetails = rcls.PageHeaderFontL
      OnAction = PARSESTATUS_INVALID
    Case "PAGEHEADERFONTC"
      Set rcls = psett.ParseParameter(2)
      Set FDetails = rcls.PageHeaderFontC
      OnAction = PARSESTATUS_INVALID
    Case "PAGEHEADERFONTR"
      Set rcls = psett.ParseParameter(2)
      Set FDetails = rcls.PageHeaderFontR
      OnAction = PARSESTATUS_INVALID
    Case "PAGEFOOTERFONTL"
      Set rcls = psett.ParseParameter(2)
      Set FDetails = rcls.PageFooterFontL
      OnAction = PARSESTATUS_INVALID
    Case "PAGEFOOTERFONTC"
      Set rcls = psett.ParseParameter(2)
      Set FDetails = rcls.PageFooterFontC
      OnAction = PARSESTATUS_INVALID
    Case "PAGEFOOTERFONTR"
      Set rcls = psett.ParseParameter(2)
      Set FDetails = rcls.PageFooterFontR
      OnAction = PARSESTATUS_INVALID
    Case "GLOBALDATAFONT"
      Set rcls = psett.ParseParameter(2)
      Set FDetails = rcls.DataFont
      OnAction = PARSESTATUS_INVALID
    Case "HEADINGFONT"
      Set rFld = objCurItem
      Set FDetails = rFld.HeadingFont
      OnAction = PARSESTATUS_OK
    Case "DATAFONT"
      Set rFld = objCurItem
      Set FDetails = rFld.DataFont
      OnAction = PARSESTATUS_OK
    Case Else
  End Select
    
  If Not FDetails Is Nothing Then
    FDetails.Name = (params(2))
    FDetails.Size = CLng(params(3))
    FDetails.ForeColor = CLng(params(4))
    FDetails.Bold = CBoolean(params(5))
    FDetails.Italic = CBoolean(params(6))
    FDetails.Strikethru = CBoolean(params(7))
    FDetails.Underline = CBoolean(params(8))
    IParseItem_OnAction = OnAction
  End If
  
End Function


