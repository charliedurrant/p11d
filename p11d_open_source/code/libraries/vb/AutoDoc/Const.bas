Attribute VB_Name = "Const"
Option Explicit
Public gbForceExit As Boolean
Public gbAllowAppExit As Boolean
Public gbMDILoaded As Boolean

Public Enum ApplicationErrors
  ERR_CORESETUP = TCSCLIENT_ERROR
  ERR_FILE_NOT_EXIST
  ERR_INVALID_CLASS_FILE
  ERR_NOT_VBP
  ERR_FILE_OPEN_EXCLUSIVE
  ERR_INVALID_VBP_CLASS
  ERR_FILE_WRONG_TYPE
  ERR_INVALIDNAME
  ERR_IS_NOTHING
  ERR_INITDOCS
  ERR_NODECLICK
  ERR_NOVARNAME
  ERR_OPEN_FILE
  ERR_PARSEFUNCTION
  'ERR_NEXTERROR etc...
End Enum
' END --- TEMPLATE CODE

Public Const S_AUTODOCSTART As String = "'##"
Public Const S_AUTODOC_CLASSDESCRIPTION As String = S_AUTODOCSTART & "CD "
Public Const S_AUTODOC_DESCRIPTION As String = S_AUTODOCSTART & "D "
Public Const S_AUTODOC_CATEGORY As String = S_AUTODOCSTART & "C "
Public Const S_AUTODOC_LONG_DESCRIPTION As String = S_AUTODOCSTART & "LD "
Public Const S_AUTODOC_RETURN_VALUE As String = S_AUTODOCSTART & "RV "
Public Const S_AUTODOC_VARNAME As String = S_AUTODOCSTART & "V "
Public Const S_AUTODOC_STUB As String = S_AUTODOCSTART & "CCORE "

Public Const S_FUNTION_SEARCH_BARE_SUB  As String = "Sub"
Public Const S_FUNTION_SEARCH_BARE_FUNCTION  As String = "Function"
Public Const S_FUNTION_SEARCH_PUBLIC_FUNCTION  As String = "Public Function"
Public Const S_FUNTION_SEARCH_PUBLIC_SUB  As String = "Public Sub"
Public Const S_FUNTION_PUBLIC_PROPERTY As String = "Public Property"
Public Const S_FUNTION_PUBLIC_ANY As String = "Public"
Public Const S_FUNCTION_END As String = "End Function"
Public Const S_SUB_END As String = "End Sub"
Public Const S_PROPERTY_GET As String = "Get"
Public Const S_PROPERTY_LET As String = "Let"
Public Const S_PROPERTY_SET As String = "Set"

Public Const S_INI_SECTION_FILE As String = "File"

Public Const KEY_SEPARATOR As String = ":?:"

Public Const L_GAP As Long = 10

Public Enum TREEVIEW_NODETYPE
  IMG_PROJECT = 1
  IMG_CLASS
  IMG_FUNCTION
  IMG_SUB
  IMG_PROPERTY
  IMG_CATEGORY
  IMG_SEARCH_RESULT
End Enum

Public Enum SEARCH_MODE
  SM_ALL
  SM_DESCRIPTION
  SM_NAME
  SM_PARAMETERS
End Enum

Public Enum CLASS_PARSEMODE
  CP_SEARCH_ANY
  CP_SEARCH_FUNCTION
  CP_PROCESS_AUTODOC
End Enum

Public Enum PROPERTY_TYPE
  PROPERTY_NONE
  PROPERTY_GET = 1
  PROPERTY_LET = 2
  PROPERTY_SET = 4
End Enum

Public gProjects As Projects
Public CategoryMaps() As String
Public CategoryValues() As String
Public gLastFunctionKey As String

Public Function InList(ClassList As String, ByVal ClassName As String) As Boolean
  Dim p As Long, ch As String
  
  p = InStr(1, ClassList, ClassName, vbTextCompare)
  If p > 0 Then
    If p > 1 Then
      ch = Mid$(ClassList, p - 1)
      If (ch <> ";") And (ch <> " ") Then Exit Function
    End If
    InList = True
  End If
End Function
