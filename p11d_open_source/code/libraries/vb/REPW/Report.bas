Attribute VB_Name = "Report"
Option Explicit
Private Const FIELD_SEP As String = " , "

Public Function InitReportColumns(ac As AutoClass, ReportFields As Variant, ByVal MaxCol As Long, RepDetails As ReportDetails) As Boolean
  Dim rFld As ReportField, aCol As AutoCol
  Dim i As Long, j As Long, s As String
  
  On Error GoTo InitReportColumns_Err
  Call xSet("InitReportColumns")
  Call ac.ClearColumns
  Call ac.ClearPreview
  For i = 1 To MaxCol
    Set rFld = ReportFields(i)
    Set aCol = ac.Add(rFld.KeyString, rFld.DataType)
    Call rFld.SetAutoCol(aCol, RepDetails.DataFont.Name, RepDetails.DataFont.Size, RepDetails.HideGroupHeaderTypes)
  Next i
  InitReportColumns = True

InitReportColumns_End:
  Call xReturn("InitReportColumns")
  Exit Function

InitReportColumns_Err:
  Call ErrorMessage(ERR_ERROR, Err, "InitReportColumns", "ERR_UNDEFINED", "Failed to initialise report columns.")
  Resume InitReportColumns_End
End Function


' do output
' 1 to MaxCol
Public Function PrepareOutputLine(ac As AutoClass, rep As Reporter, ReportFields() As ReportField, PrintLine As Variant, ByVal MaxCol As Long, ByVal MaxCriteria As Long, ByVal rw As ReportWizard) As Boolean
  Dim rFld As ReportField, aCol As AutoCol
  Dim i As Long, j As Long, MeetsCriteria As Boolean
  Dim sNotify As String
  
  #If DEBUGVER Then
    Dim oString As String
    oString = isNullEx(PrintLine(1))
    For i = 2 To MaxCol
      oString = oString & FIELD_SEP & isNullEx(PrintLine(i))
    Next i
    Call OutputDebug("PrepareOutputline", oString)
  #End If
    
    
    
  If OutputLineMeetsCriteria(MaxCol, MaxCriteria, ReportFields, PrintLine) Then
    If (Not rw.ReportInterface.Notify Is Nothing) And rw.NotifyLineMeetsCriteria Then
      sNotify = "OUTPUT_LINE:"
      For i = 1 To MaxCol
         Set rFld = ReportFields(i)
         sNotify = sNotify & """" & rFld.Name & """;""" & PrintLine(i) & """"
         If (i < MaxCol) Then
           sNotify = sNotify & ","
         End If
      Next
      Call rw.ReportInterface.Notify.Notify(-1, -1, sNotify)
    End If
    
    Call ModifyPrintLine(MaxCol, PrintLine, ReportFields)
    Call ac.PreviewAutoLine(PrintLine)
    PrepareOutputLine = True
  End If
End Function
        

Private Function ModifyPrintLine(MaxCol As Long, PrintLine As Variant, ReportFields() As ReportField) As Boolean
  Dim i As Long, rFld As ReportField

  On Error GoTo ModifyPrintLine_Err
  Call xSet("ModifyPrintLine")

  For i = 1 To MaxCol
    Set rFld = ReportFields(i)
    Select Case rFld.DataType
      Case TYPE_BOOL
        If CBoolean(PrintLine(i)) Then
          PrintLine(i) = rFld.BooleanTrue
        Else
          PrintLine(i) = rFld.BooleanFalse
        End If
      Case Else
    End Select
    If Len(rFld.Prefix) > 0 Then PrintLine(i) = rFld.Prefix & PrintLine(i)
    If Len(rFld.Suffix) > 0 Then PrintLine(i) = PrintLine(i) & rFld.Suffix
  Next i
  
ModifyPrintLine_End:
  Call xReturn("ModifyPrintLine")
  Exit Function

ModifyPrintLine_Err:
  Call ErrorMessage(ERR_MODIFYPRINTLINE, Err, "ModifyPrintLine", "ERR_MODIFYPRINTLINE", "The field value '" & PrintLine(i) & "' could not be modified as required." & vbCrLf & "Field Key: " & rFld.KeyString & vbCrLf & "Field Data Type: " & DataTypeName(rFld.DataType))
  Resume ModifyPrintLine_End
End Function

Public Sub FieldWidthLimits(rFields As Collection, ByVal SelectedOrder As Long, TotalWidth As Double, WidthToRight As Double)
  Dim rFld As ReportField
  
  TotalWidth = 0
  WidthToRight = 0
  For Each rFld In rFields
    If (Not rFld.Hide) And Not (rFld.Group And rFld.GroupHeader) Then
      TotalWidth = TotalWidth + rFld.Width
      If rFld.Order > SelectedOrder Then
        WidthToRight = WidthToRight + rFld.Width
      End If
    End If
  Next rFld
End Sub

Public Sub SetEqualFieldWidths(rFields As Collection)
  Dim NumDisp As Long, rFld As ReportField
  Dim CurrWidth As Double, TotalWidth As Double
  
  NumDisp = 0
  For Each rFld In rFields
    If (Not rFld.Hide) And Not (rFld.Group And rFld.GroupHeader) Then
      NumDisp = NumDisp + 1
    End If
  Next rFld
  TotalWidth = 0
  For Each rFld In rFields
    If NumDisp > 0 Then
      CurrWidth = PercentDP(Val(100 / NumDisp), 1)
      rFld.Width = CurrWidth
      TotalWidth = TotalWidth + CurrWidth
    End If
  Next rFld
  If TotalWidth > 100 Then
    For Each rFld In rFields
      CurrWidth = Min(TotalWidth - 100, rFld.Width)
      TotalWidth = TotalWidth - CurrWidth
      rFld.Width = rFld.Width - CurrWidth
      If TotalWidth <= 100 Then Exit For
    Next rFld
  End If
End Sub

