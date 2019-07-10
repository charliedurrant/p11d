Attribute VB_Name = "Funcs"
Option Explicit

Public Type RECENT_FILE
  CanonicalPathAndFile As String
  NonCanonicalPathAndFile As String
  Enabled As Boolean
  Valid As Boolean
End Type

Public Const S_INI_SECTION_RECENT_FILES As String = "RECENT_FILES"
Public Const S_INI_ENTRY_RECENT_FILE As String = "FILE"
Public Const L_MAX_CHARS_SUB As Long = 10
Public Const S_MENU_PREFIX = "rfilemenu_"

Public Function vbMenuName(ByVal Index As Long) As String
  vbMenuName = S_MENU_PREFIX & CStr(Index)
End Function

Public Sub LoadvbMenuItems(ByVal mnuRecentFile As VBMenu, ByVal mnuItemRecentFile As VBMenuItem, ByVal MaxNoOfFiles As Long)
  Dim i As Long
  Dim mnu As VBMenuItem
  
  On Error GoTo LoadvbMenuItems_ERR
  Call xSet("LoadvbMenuItems")
  ' unload all current recent file menus
  For i = mnuRecentFile.Count To 1
    Set mnu = mnuRecentFile.Item(i)
    If InStr(1, mnu.Name, S_MENU_PREFIX, vbBinaryCompare) = 1 Then
      Call mnuRecentFile.Remove(i)
    End If
  Next i
  
  ' add in the new indices
  For i = 1 To MaxNoOfFiles
    Set mnu = mnuRecentFile.Add(vbMenuName(i), "", mnuItemRecentFile.ParentName)
    mnu.Visible = False
  Next i

LoadvbMenuItems_END:
  Call xReturn("LoadvbMenuItems")
  Exit Sub
  
LoadvbMenuItems_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "LoadvbMenuItems", "Load vb Menu Objects", "Error loading menu objects")
  Resume LoadvbMenuItems_END
End Sub

Public Sub LoadMenuObjects(mnuRecentFile As Object, ByVal MaxNoOfFiles As Long)
  Dim i As Long
  Dim m As Menu
  
  On Error GoTo LoadMenuObjects_ERR
  Call xSet("LoadMenuObjects")
  For i = 1 To mnuRecentFile.UBound
    Call Unload(mnuRecentFile(i))
  Next i
  For i = 1 To MaxNoOfFiles
    Call Load(mnuRecentFile(i))
  Next i

LoadMenuObjects_END:
  Call xReturn("LoadMenuObjects")
  Exit Sub
  
LoadMenuObjects_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "LoadMenuObjects", "Load Menu Objects", "Error loading menu objects")
  Resume LoadMenuObjects_END
End Sub

Public Function GetMenuCaption(ByVal sPathAndFile As String, ByVal MaxMenuCaptionStringLength As Long) As String
  
  On Error GoTo GetMenuCaption_ERR
  Call xSet("GetMenuCaption")
   
  If Len(sPathAndFile) > MaxMenuCaptionStringLength Then
    If Left$(sPathAndFile, 2) = "\\" Then
      GetMenuCaption = Canonical(sPathAndFile, MaxMenuCaptionStringLength)
    Else
      GetMenuCaption = NonCanonical(sPathAndFile, MaxMenuCaptionStringLength)
    End If
  Else
  GetMenuCaption = sPathAndFile
  End If
      
GetMenuCaption_END:
  Call xReturn("GetMenuCaption")
  Exit Function
  
GetMenuCaption_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "GetMenuCaption", "Get Menu Caption", "Error getting menu caption.")
  Resume GetMenuCaption_END
End Function

Public Function NonCanonical(ByVal sPathAndFile As String, ByVal MaxMenuCaptionStringLength) As String
  Dim dir As String
  Dim file As String
  Dim ext As String

  On Error GoTo NonCanonical_ERR
  
  Call xSet("NonCanonical")

  Call SplitPath(sPathAndFile, dir, file, ext)
  If Len(file) > L_MAX_CHARS_SUB Then
    file = Left$(file, L_MAX_CHARS_SUB)
  End If
  dir = Left$(sPathAndFile, MaxMenuCaptionStringLength - Len(file) - Len(ext) - Len("...\"))
  NonCanonical = dir & "...\" & file & ext

NonCanonical_END:
  Call xReturn("NonCanonical")
  Exit Function
NonCanonical_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "NonCanonical", "Get Non-Canonical", "Error getting non-canonical menu caption.")
  Resume NonCanonical_END
End Function

Public Function Canonical(ByVal sPathAndFile As String, ByVal MaxMenuCaptionStringLength) As String
  
  On Error GoTo Canonical_ERR
  Call xSet("Canonical")
  
  sPathAndFile = Right$(sPathAndFile, MaxMenuCaptionStringLength - Len("...\"))
  sPathAndFile = Right$(sPathAndFile, Len(sPathAndFile) - InStr(1, sPathAndFile, "\", vbBinaryCompare))
  Canonical = "...\" & sPathAndFile

Canonical_END:
  Call xReturn("Canonical")
  Exit Function
Canonical_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "Canonical", "Get Canonical", "Error getting canonical menu caption.")
  Resume Canonical_END
End Function

