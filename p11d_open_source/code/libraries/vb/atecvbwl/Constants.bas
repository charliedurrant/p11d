Attribute VB_Name = "Constants"
Option Explicit

Public Enum ATECVBWL_ERRORS
  ERR_INCLUDE_STYLE
  ERR_INCLUDE_SCRIPT
  ERR_NAVIGATOR
  ERR_FILETOSTRING
  ERR_URLERROR
  ERR_HTMLATTR
  ERR_MULTI_PART_FORM
  ERR_FIELD_REQUIRED
  ERR_TREEVIEW
  'cadxx #RD
  ERR_DATA_ACCESS
  'cadxx #RD
  ERR_PAGE_MARKUP
End Enum


'misc html constants
Public Const S_QUOT As String = """"
Public Const S_QUOTQUOT As String = """"""

Public Const CDATA_START As String = "<![CDATA["
Public Const CDATA_END As String = "]]>"

' Note: Both ReplaceXMLMetacharacters & XMLText are dependent on this value
Public Const S_INVALID_XML_CHARS As String = S_QUOT & "<>&'"

'cadxx #RD
Public Type ELEMENT_TO_SEARCH_FOR
  Name As String
  Position As Long
  Search As String
  SearchPos As Long
  LenSearch As Long
  FindInnerHTML As Boolean
End Type
'
Public Const L_SPACE As Long = 32
Public Const L_SINGLE_QUOTE As Long = 39
Public Const L_DOUBLE_QUOTE As Long = 34
Public Const L_EQUALS As Long = 61
Public Const L_RESIZE_INCREASE As Long = 10
'#RD added in reference to microsoft scripting runtime

Public Const QS_INCREMENT As Long = 32768  ' 2^ 15
Public Const COMPONENT_NAME As String = "atc2vbwl"

Public Const S_NAV_ATTRIB_FIELD_CHILDREN As String = "children"
Public Const S_NAV_ATTRIB_FIELD_OPEN As String = "open"
Public Const S_NAV_ATTRIB_FIELD_SELECTED As String = "selected"
Public Const S_NAV_ATTRIB_FIELD_IMAGE_CLOSED As String = "image_open"
Public Const S_NAV_ATTRIB_FIELD_IMAGE_OPEN As String = "image_closed"
Public Const S_NAV_ATTRIB_FIELD_IMAGE_LEAF As String = "image_leaf"


Public Const S_NAV_NODE_FIELD_TOOLTIP As String = "tooltip"
Public Const S_NAV_NODE_FIELD_NAME As String = "name"


