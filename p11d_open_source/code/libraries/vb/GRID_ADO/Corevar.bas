Attribute VB_Name = "Const"
Option Explicit

Public Enum TCSAUTO_ERRORS
  ERR_NOSYNCH = TCSADOAUTO_ERROR + 1
  ERR_REPORTACTIVE
  ERR_AUTOPARSE
  ERR_PASTEROW
  ERR_SHOWGRID
  ERR_INITGRID
  ERR_UPDATEROW
  ERR_SETGRID
  ERR_GETVALUE
  ERR_BEFOREUPDATE
  ERR_MOVECOLUMNS
  ERR_AUDITSET
  ERR_GETCALCVALUE
  ERR_GRID_BEFOREUPDATE
  ERR_INITAUTODATA
  ERR_CALCDERIVED
  ERR_PASTE
  ERR_NOREMOVEFORMAT
  ERR_NOFILTER
  ERR_SQLTOOLONG
  ERR_INVALIDGROUP
  ERR_GRIDVALIDATE
  ERR_GRIDVALIDATEIGNORE
  ERR_INVALID_CELLVALUE
  ERR_INVALID_CELLVALUES
  ERR_GRIDDATES
  ERR_GRIDCALCFAIL
  ERR_FINDRECORD
  ERR_GRIDFORMAT
  ERR_AUTOCOL
  ERR_INVALIDSQL
  ERR_GETCOMBORS
  ERR_INVALID_VAR_TYPE
  ERR_NUMBER_OF_COLUMNS
  ERR_FILTER
End Enum

Public Const GRID_MINCOLWIDTH As Long = 100
Public Const one_cm As Single = 1440 / 2.54
Public Const MINCOLWIDTH As Single = one_cm
Public Const MINEMPTYCOLWIDTH As Single = one_cm / 4
Public Const MININTERCOLSPACE As Single = one_cm / 4

Public Const MINLEFTMARGIN As Single = one_cm / 2
Public Const MINRIGHTMARGIN As Single = one_cm / 2
Public Const ADJRIGHTALIGN As Single = one_cm / 4
Public Const ADJSPACING As Single = one_cm / 5

Public Const PREVIEWROWS = 25

' Resource bitmaps
Public Const TICK_BMP As Long = 1000
Public Const CROSS_BMP As Long = 1001
Public Const CROSS_BLANK_BMP As Long = 1002
Public Const BUTTON_BASE_BMP As Long = 1500

' Current Auto Names
Public AutoParser As Parser
Public AutoClipHandle As Long
Public AutoNames As StringList
Public AutoCount As Long
Public AutoControlRegistered As Boolean
Public FormatRemove As Boolean

' SQL parameters
Public Const BEGIN_PARAM As String = "<%"
Public Const END_PARAM As String = "%>"


Public Sub SetUpParser()
  Set AutoParser.ParseSettings = New AutoParseSettings
  Call AutoParser.AddParseItem(New ParseAlignment)
  Call AutoParser.AddParseItem(New ParseBoolean)
  Call AutoParser.AddParseItem(New ParseButton)
  Call AutoParser.AddParseItem(New ParseCBoolean)
  Call AutoParser.AddParseItem(New ParseCaption)
  Call AutoParser.AddParseItem(New ParseCaptionFormat)
  Call AutoParser.AddParseItem(New ParseCollapseLike)
  Call AutoParser.AddParseItem(New ParseDataFormat)
  Call AutoParser.AddParseItem(New ParseDateFormat)
  Call AutoParser.AddParseItem(New ParseDefault)
  Call AutoParser.AddParseItem(New ParseFormat)
  Call AutoParser.AddParseItem(New ParseDrop)
  Call AutoParser.AddParseItem(New ParseDropCombo)
  Call AutoParser.AddParseItem(New ParseDropList)
  Call AutoParser.AddParseItem(New ParseDropQuery)
  Call AutoParser.AddParseItem(New ParseDropQueryCombo)
  Call AutoParser.AddParseItem(New ParseFixedL)
  Call AutoParser.AddParseItem(New ParseFixedR)
  Call AutoParser.AddParseItem(New ParseGridCaption)
  Call AutoParser.AddParseItem(New ParseGroup)
  Call AutoParser.AddParseItem(New ParseHide)
  Call AutoParser.AddParseItem(New ParseMean)
  Call AutoParser.AddParseItem(New ParseMaxValue)
  Call AutoParser.AddParseItem(New ParseMaxDropItems)
  Call AutoParser.AddParseItem(New ParseMinWidth)
  Call AutoParser.AddParseItem(New ParseMinValue)
  Call AutoParser.AddParseItem(New ParseMinXOffset)
  Call AutoParser.AddParseItem(New ParseNewRecord)
  Call AutoParser.AddParseItem(New ParseNoAddNew)
  Call AutoParser.AddParseItem(New ParseNoCalc)
  Call AutoParser.AddParseItem(New ParseNoCopy)
  Call AutoParser.AddParseItem(New ParseNoEdit)
  Call AutoParser.AddParseItem(New ParseNoPrint)
  Call AutoParser.AddParseItem(New ParseNoSquash)
  Call AutoParser.AddParseItem(New ParseOnAddNew)
  Call AutoParser.AddParseItem(New ParseOnChangeEvent)
  Call AutoParser.AddParseItem(New ParseOnNull)
  Call AutoParser.AddParseItem(New ParseOnUpdate)
  Call AutoParser.AddParseItem(New ParsePrintCaption)
  Call AutoParser.AddParseItem(New ParseQuery)
  Call AutoParser.AddParseItem(New ParseRealColumn)
  Call AutoParser.AddParseItem(New ParseReportFormat)
  Call AutoParser.AddParseItem(New ParseSplit)
  Call AutoParser.AddParseItem(New ParseSum)
  Call AutoParser.AddParseItem(New ParseTitle)
  Call AutoParser.AddParseItem(New ParseToolTip)
  Call AutoParser.AddParseItem(New ParseUnboundColumn)
  Call AutoParser.AddParseItem(New ParseWidth)
  Call AutoParser.AddParseItem(New ParseWrap)
  'FIX CAD1
  Call AutoParser.AddParseItem(New ParseBackColor)
  Call AutoParser.AddParseItem(New ParseForeColor)
  AutoParser.ParseTokensOnly = True
End Sub

