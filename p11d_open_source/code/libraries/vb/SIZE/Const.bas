Attribute VB_Name = "Const"
Option Explicit

Public Const SNG_TABLESS As Single = 75000

Public Type SIZE_SNG
  Left As Single
  Top As Single
  Width As Single
  Height As Single
End Type


Public Type TAG_PARSEITEM
  Name As String
  Value As String
  ParseItem As TAG_PARSEITEMS
End Type

Public Enum TAG_PARSEITEMS
  [_PARSE_ITEMMIN] = 0
  L_BUDDYRIGHT = 2 ^ 0
  L_BUDDYLEFT = 2 ^ 1
  L_FREE = 2 ^ 2
  L_LOCK = 2 ^ 3
  L_EQUALISE = 2 ^ 4
  L_EQUALISEBOTTOM = 2 ^ 5
  L_EQUALISERIGHT = 2 ^ 6
  L_LOCKRIGHT = 2 ^ 7
  L_LOCKBOTTOM = 2 ^ 8
  L_LOCKBOTTOMRIGHT = 2 ^ 9
  L_FREEEQUALISEBOTTOMRIGHT = 2 ^ 10
  L_FREEEQUALISEBOTTOM = 2 ^ 11
  L_MOVEONLY = 2 ^ 12
  L_FREELOCKRIGHT = 2 ^ 13
  L_FREELOCKLEFT = 2 ^ 14
  L_FREELOCKTOPBOTTOMLEFT = 2 ^ 15
  L_FREELOCKTOPBOTTOMRIGHT = 2 ^ 16
  L_FREELOCKTOPBOTTOM = 2 ^ 17
  L_LOCKBOTTOMEQUALISERIGHT = 2 ^ 18
  L_FREELOCKTOPRIGHT = 2 ^ 19
  L_FREELOCKBOTTOMRIGHT = 2 ^ 20
  L_FREELOCKTOP = 2 ^ 21
  L_FREELOCKBOTTOM = 2 ^ 22
  L_SCALEONLY = 2 ^ 23
  L_BUDDY = 2 ^ 24
  L_FREELOCKTOPHEIGHTLEFT = 2 ^ 25
  L_FREELOCKTOPHEIGHTRIGHT = 2 ^ 26
  L_BUDDYEQUALISEBOTTOMRIGHT = 2 ^ 27
  L_GRID = 2 ^ 28
  L_CENTRE = 2 ^ 29
  L_FONT = 2 ^ 30
  
  [_PARSE_ITEMMAX] = 31
End Enum

Public ParseItems([_PARSE_ITEMMIN] To [_PARSE_ITEMMAX]) As TAG_PARSEITEM

Public Type MIN_SIZES
  MinSize As Long
  MinWidth As Single
  MinHeight As Single
End Type

'Top' property cannot be read at run time
Public Enum RESIZE_ERRORS
  ERR_NOTOP = 393
  ERR_NOARRAY = 343
  ERR_NORUNTIME = 382
  ERR_NOPROPERTY = 438
  ERR_RESIZE = TCSSIZE_ERROR
End Enum

