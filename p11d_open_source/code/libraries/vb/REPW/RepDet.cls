VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ReportDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public ReportName As String

Public IncludeRepHeader As Boolean
Public IncludePageHeader As Boolean
Public IncludePageFooter As Boolean

Public RepHeaderL As String
Public RepHeaderC As String
Public RepHeaderR As String

Public PageHeaderL As String
Public PageHeaderC As String
Public PageHeaderR As String

Public PageFooterL As String
Public PageFooterC As String
Public PageFooterR As String

Public Orientation As REPORT_ORIENTATION
Public ExportType As String

Public RepHeaderFontL As FontDetails
Public RepHeaderFontC As FontDetails
Public RepHeaderFontR As FontDetails
Public PageHeaderFontL As FontDetails
Public PageHeaderFontC As FontDetails
Public PageHeaderFontR As FontDetails
Public PageFooterFontL As FontDetails
Public PageFooterFontC As FontDetails
Public PageFooterFontR As FontDetails

Public DataFont As FontDetails
Public HeadingFont As FontDetails

Public ReportPages As Long
Public AutoWidth As Boolean
Public TrimHeadings As Boolean
Public PreviewLines As Long

Public GroupHeaderDelimiter As String
Public GroupHeaderSeparator As String
Public PrintRecordCount As Boolean
Public ReportHeaderOnAllPages As Boolean
Public FitToPage As Boolean
Public SummaryReport As Boolean
Public IncludeHeaders As Boolean
Public AlignHeaders As Boolean
Public HideGroupHeaderTypes As Boolean
Public LeftJoinLeaves As Boolean

'RK AR Reporter properties
#If AbacusReporter Then
  Public ARFileGroupSelection As String
  Public ARFileGroups As String
  Private m_ARAbacusProductType As Long


  Public Property Get ARAbacusProductType() As Long
    'Default to current Abacus Product Type e.g. for template
    If m_ARAbacusProductType = 0 Then
      ARAbacusProductType = g_AbacusReporter.AbacusDataProvider.ProductType
    Else
      ARAbacusProductType = m_ARAbacusProductType
    End If
  End Property

  Public Property Let ARAbacusProductType(NewVal As Long)
    m_ARAbacusProductType = NewVal
  End Property

#End If

Private Sub Class_Initialize()
  Set RepHeaderFontL = New FontDetails
  Set RepHeaderFontC = New FontDetails
  Set RepHeaderFontR = New FontDetails
  Set PageHeaderFontL = New FontDetails
  Set PageHeaderFontC = New FontDetails
  Set PageHeaderFontR = New FontDetails
  Set PageFooterFontL = New FontDetails
  Set PageFooterFontC = New FontDetails
  Set PageFooterFontR = New FontDetails
  
  Set DataFont = New FontDetails
  Set HeadingFont = New FontDetails
  
  Call Reinitialise
End Sub

Public Function SaveDetails() As String
  Dim s As String
  
  'If IncludeRepHeader Then
    s = s & "(IncludeRepHeader)" & vbCrLf
    s = s & "(RepHeaderL=" & AddEscapeChars(RepHeaderL) & ")" & vbCrLf
    s = s & "(RepHeaderC=" & AddEscapeChars(RepHeaderC) & ")" & vbCrLf
    s = s & "(RepHeaderR=" & AddEscapeChars(RepHeaderR) & ")" & vbCrLf
    s = s & RepHeaderFontL.SaveString("RepHeaderFontL") & vbCrLf
    s = s & RepHeaderFontC.SaveString("RepHeaderFontC") & vbCrLf
    s = s & RepHeaderFontR.SaveString("RepHeaderFontR") & vbCrLf
  'End If
  'If IncludePageHeader Then
    s = s & "(IncludePageHeader)" & vbCrLf
    s = s & "(PageHeaderL=" & AddEscapeChars(PageHeaderL) & ")" & vbCrLf
    s = s & "(PageHeaderC=" & AddEscapeChars(PageHeaderC) & ")" & vbCrLf
    s = s & "(PageHeaderR=" & AddEscapeChars(PageHeaderR) & ")" & vbCrLf
    s = s & PageHeaderFontL.SaveString("PageHeaderFontL") & vbCrLf
    s = s & PageHeaderFontC.SaveString("PageHeaderFontC") & vbCrLf
    s = s & PageHeaderFontR.SaveString("PageHeaderFontR") & vbCrLf
  'End If
  'If IncludePageFooter Then
    s = s & "(IncludePageFooter)" & vbCrLf
    s = s & "(PageFooterL=" & AddEscapeChars(PageFooterL) & ")" & vbCrLf
    s = s & "(PageFooterC=" & AddEscapeChars(PageFooterC) & ")" & vbCrLf
    s = s & "(PageFooterR=" & AddEscapeChars(PageFooterR) & ")" & vbCrLf
    s = s & PageFooterFontL.SaveString("PageFooterFontL") & vbCrLf
    s = s & PageFooterFontC.SaveString("PageFooterFontC") & vbCrLf
    s = s & PageFooterFontR.SaveString("PageFooterFontR") & vbCrLf
  'End If
  s = s & "(Orientation=" & Orientation & ")" & vbCrLf
  s = s & DataFont.SaveString("GlobalDataFont") & vbCrLf
  If AutoWidth Then s = s & "(AutoWidth)" & vbCrLf
  If TrimHeadings Then s = s & "(TrimHeadings)" & vbCrLf
  
  s = s & "(GroupHeaderSeparator=" & AddEscapeChars(GroupHeaderDelimiter) & ")" & vbCrLf
  s = s & "(GroupHeaderFieldSuffix=" & AddEscapeChars(GroupHeaderSeparator) & ")" & vbCrLf
  If PrintRecordCount Then s = s & "(PrintRecordCount)" & vbCrLf
  If ReportHeaderOnAllPages Then s = s & "(ReportHeaderOnAllPages)" & vbCrLf
  If FitToPage Then s = s & "(FitToPage)" & vbCrLf
  If SummaryReport Then s = s & "(CollapseAllGroups)" & vbCrLf
  If IncludeHeaders Then s = s & "(IncludeColHeaders)" & vbCrLf
  If AlignHeaders Then s = s & "(AlignColHeaders)" & vbCrLf
  If HideGroupHeaderTypes Then s = s & "(HideGroupHeaderTypes)" & vbCrLf
  If LeftJoinLeaves Then s = s & "(LeftJoinLeaves)" & vbCrLf
  
  'RK Addition of AR Properties
  #If AbacusReporter Then
    s = s & "(FileGroupSelection=" & ARFileGroupSelection & ")" & vbCrLf
    s = s & "(FileGroups=" & ARFileGroups & ")" & vbCrLf
    s = s & "(AbacusProductType=" & g_AbacusReporter.AbacusDataProvider.ProductType & ")" & vbCrLf
  #End If
  SaveDetails = s
End Function

Public Function AddIntoAuto(ac As AutoClass, RepWiz As ReportWizard) As Boolean
  Dim s As String
  
  On Error GoTo AddIntoAuto_Err
  Call xSet("AddIntoAuto")
    
  ac.ReportPreviewRows = PreviewLines
  ac.TrimColHeaders = TrimHeadings
  ac.ReportFormat = DataFont.FontString
  
  If IncludeRepHeader Then
    s = ""
    If Len(RepHeaderL) > 0 Then s = s & RepHeaderFontL.FontString & "{X=0}{LEFT}" & RepHeaderL
    If Len(RepHeaderC) > 0 Then s = s & RepHeaderFontC.FontString & "{CENTERPAGE}" & RepHeaderC
    If Len(RepHeaderR) > 0 Then s = s & RepHeaderFontR.FontString & "{X=100}{RIGHT}" & RepHeaderR
    s = RepWiz.ReplaceCtrlChars(s)
    ac.ReportHeader = s
  Else
    ac.ReportHeader = ""
  End If
  If IncludePageHeader Then
    s = ""
    If Len(PageHeaderL) > 0 Then s = s & PageHeaderFontL.FontString & "{X=0}" & PageHeaderL
    If Len(PageHeaderC) > 0 Then s = s & PageHeaderFontC.FontString & "{CENTERPAGE}" & PageHeaderC
    If Len(PageHeaderR) > 0 Then s = s & PageHeaderFontR.FontString & "{X=100}{RIGHT}" & PageHeaderR
    s = RepWiz.ReplaceCtrlChars(s)
    ac.PageHeader = s
  Else
    ac.PageHeader = ""
  End If
  If IncludePageFooter Then
    s = ""
    If Len(PageFooterL) > 0 Then s = s & PageFooterFontL.FontString & "{X=0}" & PageFooterL
    If Len(PageFooterC) > 0 Then s = s & PageFooterFontC.FontString & "{CENTERPAGE}" & PageFooterC
    If Len(PageFooterR) > 0 Then s = s & PageFooterFontR.FontString & "{X=100}{RIGHT}" & PageFooterR
    s = RepWiz.ReplaceCtrlChars(s)
    ac.PageFooter = s
  Else
    ac.PageFooter = ""
  End If
         
  ac.PrintRecordCount = PrintRecordCount
  ac.GroupHeaderSeparator = GroupHeaderDelimiter
  ac.GroupHeaderFieldSuffix = GroupHeaderSeparator
  If HideGroupHeaderTypes Then ac.GroupHeaderFieldSuffix = ""
  ac.ReportHeaderOnAllPages = ReportHeaderOnAllPages
  ac.FitToPage = FitToPage
  ac.CollapseAllGroups = SummaryReport
  ac.IncludeColHeaders = IncludeHeaders
  ac.AlignColHeaders = AlignHeaders

AddIntoAuto_End:
  Call xReturn("AddIntoAuto")
  Exit Function

AddIntoAuto_Err:
  Call ErrorMessage(ERR_ERROR, Err, "AddIntoAuto", "ERR_UNDEFINED", "Error adding report details into reporter.")
  Resume AddIntoAuto_End
  Resume
End Function

Private Sub ModifyHeaderFooter(Location As TCSREPWIZ_HEADERSFOOTERS, Modification As String)
  Dim TextType As Long, LocateWithin As Long, InsertionType As Long
    
  TextType = Location And (HEADERSFOOTERS_REPORTHEADER + HEADERSFOOTERS_PAGEHEADER + HEADERSFOOTERS_PAGEFOOTER)
  LocateWithin = Location And (HEADERSFOOTERS_LEFT + HEADERSFOOTERS_CENTER + HEADERSFOOTERS_RIGHT)
  InsertionType = Location And (HEADERSFOOTERS_BEFORETEXT + HEADERSFOOTERS_AFTERTEXT)
  
  Select Case TextType
    Case HEADERSFOOTERS_REPORTHEADER
      Select Case LocateWithin
        Case HEADERSFOOTERS_LEFT
          If InsertionType = HEADERSFOOTERS_BEFORETEXT Then
            RepHeaderL = Modification & RepHeaderL
          Else
            RepHeaderL = RepHeaderL & Modification
          End If
        Case HEADERSFOOTERS_CENTER
          If InsertionType = HEADERSFOOTERS_BEFORETEXT Then
            RepHeaderC = Modification & RepHeaderC
          Else
            RepHeaderC = RepHeaderC & Modification
          End If
        Case HEADERSFOOTERS_RIGHT
          If InsertionType = HEADERSFOOTERS_BEFORETEXT Then
            RepHeaderR = Modification & RepHeaderR
          Else
            RepHeaderR = RepHeaderR & Modification
          End If
      End Select
    Case HEADERSFOOTERS_PAGEHEADER
      Select Case LocateWithin
        Case HEADERSFOOTERS_LEFT
          If InsertionType = HEADERSFOOTERS_BEFORETEXT Then
            PageHeaderL = Modification & PageHeaderL
          Else
            PageHeaderL = PageHeaderL & Modification
          End If
        Case HEADERSFOOTERS_CENTER
          If InsertionType = HEADERSFOOTERS_BEFORETEXT Then
            PageHeaderC = Modification & PageHeaderC
          Else
            PageHeaderC = PageHeaderC & Modification
          End If
        Case HEADERSFOOTERS_RIGHT
          If InsertionType = HEADERSFOOTERS_BEFORETEXT Then
            PageHeaderR = Modification & PageHeaderR
          Else
            PageHeaderR = PageHeaderR & Modification
          End If
      End Select
    Case HEADERSFOOTERS_PAGEFOOTER
      Select Case LocateWithin
        Case HEADERSFOOTERS_LEFT
          If InsertionType = HEADERSFOOTERS_BEFORETEXT Then
            PageFooterL = Modification & PageFooterL
          Else
            PageFooterL = PageFooterL & Modification
          End If
        Case HEADERSFOOTERS_CENTER
          If InsertionType = HEADERSFOOTERS_BEFORETEXT Then
            PageFooterC = Modification & PageFooterC
          Else
            PageFooterC = PageFooterC & Modification
          End If
        Case HEADERSFOOTERS_RIGHT
          If InsertionType = HEADERSFOOTERS_BEFORETEXT Then
            PageFooterR = Modification & PageFooterR
          Else
            PageFooterR = PageFooterR & Modification
          End If
      End Select
  End Select

End Sub


Public Sub Reinitialise()
  ReportName = "TCS Report Wizard" 'apf mps check
  RepHeaderL = ""
  RepHeaderC = ""
  RepHeaderR = ""
  PageHeaderL = ""
  PageHeaderC = ""
  PageHeaderR = ""
  
  PageFooterL = "" '"{x=2}{Times=8}{DATE} {TIME}  " & GetStatic("ApplicationName") & " Version " & GetStatic("Version")
  PageFooterC = ""
  PageFooterR = "" '"{Times=8,N}{X=95}Page {PAGE}"
  
  IncludeRepHeader = False
  IncludePageHeader = False
  IncludePageFooter = False
  
  Orientation = PORTRAIT
  AutoWidth = True
  TrimHeadings = False
  PreviewLines = 50

  GroupHeaderDelimiter = ";  "
  GroupHeaderSeparator = " = "
  PrintRecordCount = False
  ReportHeaderOnAllPages = False
  FitToPage = False
  SummaryReport = False
  IncludeHeaders = True
  AlignHeaders = False
  HideGroupHeaderTypes = False
  LeftJoinLeaves = False

  RepHeaderFontL.Reinitialise
  RepHeaderFontC.Reinitialise
  RepHeaderFontR.Reinitialise
  PageHeaderFontL.Reinitialise
  PageHeaderFontC.Reinitialise
  PageHeaderFontR.Reinitialise
  PageFooterFontL.Reinitialise
  PageFooterFontC.Reinitialise
  PageFooterFontR.Reinitialise
  DataFont.Reinitialise
  HeadingFont.Reinitialise
  
End Sub
