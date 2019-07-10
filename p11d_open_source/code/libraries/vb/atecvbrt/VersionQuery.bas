Attribute VB_Name = "VersionHelper"
Option Explicit

Public Function VersionQueryMap(ByVal sPathAndFile As String, Optional ByVal VQT As VER_QUERY_TYPE = VQT_FILE_VERSION) As String
  Dim sProperty As String

  Select Case VQT
    Case VQT_PRODUCT_VERSION
      sProperty = "ProductVersion"
    Case VQT_PRODUCT_NAME
      sProperty = "ProductName"
    Case VQT_COMPANY_NAME
      sProperty = "CompanyName"
    Case VQT_FILE_DESCRIPTION
      sProperty = "FileDescription"
    Case VQT_FILE_VERSION
      sProperty = "FileVersion"
    Case VQT_INTERNAL_NAME
      sProperty = "InternalName"
    Case VQT_LEGAL_COPYRIGHT
      sProperty = "LegalCopyright"
    Case VQT_ORIGINAL_FILE_NAME
      sProperty = "OriginalFilename"
    Case VQT_COMMENTS
      sProperty = "OriginalFilename"
    Case VQT_LEGAL_TRADEMARKS
      sProperty = "LegalTrademarks"
    Case VQT_PRIVATE_BUILD
      sProperty = "PrivateBuild"
    Case VQT_SPECIAL_BUILD
      sProperty = "SpecialBuild"
    Case Else
      Err.Raise ERR_VERQUERY, "VersionQueryMap", "Unknown Verquery type [" & VQT & "]"
  End Select
  VersionQueryMap = VersionQuery(sPathAndFile, sProperty)
End Function

