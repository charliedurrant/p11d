Attribute VB_Name = "lows"
Option Explicit

Public Sub SwitchForm(iFrm As IImportForm, Dest As IMPORT_GOTOFORM)
  Dim frm As Form
  
  iFrm.ImpWiz.GotoForm = Dest
  Set frm = iFrm
  frm.Visible = False
End Sub

Public Sub OmitLines(txtOmitH As TextBox, txtOmitF As TextBox, TakeFieldNamesFrom As Long, MaxLines As Long)
  Dim oH As Long, oF As Long
  On Error GoTo OmitLines_Err
  Call xSet("OmitLines")

  
  oH = CLngEx(txtOmitH.Text, 0)
  oF = CLngEx(txtOmitF.Text, 0)
  
  'If oH < 0 Then oH = 0
  If oH < TakeFieldNamesFrom Then oH = TakeFieldNamesFrom
  If oH > (MaxLines - oF - 1) Then oH = (MaxLines - oF - 1)

  If oF < 0 Then oF = 0
  If oF > (MaxLines - oH - 1) Then oF = (MaxLines - oH - 1)
  
  txtOmitH.Text = CStr(oH)
  txtOmitF.Text = CStr(oF)
  
OmitLines_End:
  Call xReturn("OmitLines")
  Exit Sub

OmitLines_Err:
  Call ErrorMessage(ERR_ERROR, Err, "OmitLines", "Omit lines from Header and Footer", "Unable to calculate Header/Footer lines.")
  Resume OmitLines_End
End Sub

Public Function CopyLines(FileFrom As String, ToArray() As String, ByVal FirstLine As Long, ByVal LinesToCopy As Long) As Long
  Dim i As Long
  Dim rf As New TCSFileread
  Dim buffer As String
  
  On Error GoTo CopyLines_Err
  Call xSet("CopyLines")
  If Not rf.OpenFile(FileFrom) Then Err.Raise ERR_IMPORT, "CopyLines", "Unable to open file " & FileFrom
  If FirstLine > rf.LineCount Then FirstLine = 1
  If (FirstLine + LinesToCopy) > rf.LineCount Then LinesToCopy = rf.LineCount - FirstLine + 1
  ReDim ToArray(1 To LinesToCopy)
  For i = 1 To FirstLine - 1
    Call rf.GetLine(buffer)
  Next i
  For i = 1 To LinesToCopy
    Call rf.GetLine(buffer)
    ToArray(i) = buffer
  Next i
  CopyLines = FirstLine
  
CopyLines_End:
  Call xReturn("CopyLines")
  Exit Function

CopyLines_Err:
  CopyLines = 0
  Call ErrorMessage(ERR_ERROR, Err, "CopyLines", "Copy Lines to array", "Error copying lines.")
  Resume CopyLines_End
End Function

   
Public Sub FixUpCopyFields(NewCol As FieldSpecs)
  Dim ispec0 As ImportSpec, ispec1 As ImportSpec
  Dim i As Long, j As Long
  
  For j = 1 To NewCol.Count
    Set ispec0 = NewCol(j)
    If ispec0.CopyFieldKey > 0 Then
      For i = 1 To NewCol.Count
        Set ispec1 = NewCol.Item(i)
        If ispec0.CopyFieldKey = ispec1.FieldKey Then
          Set ispec0.CopyField = ispec1
          ispec0.vartype = ispec1.vartype
          GoTo nextcolumn
        End If
      Next i
      Call Err.Raise(ERR_FIXUPCOLUMNS, "FixUpCopyFields", "Could not find Field Key: " & CStr(ispec0.CopyFieldKey) & " in the Import Specification Columns. Unable to use Copyfield")
    End If
nextcolumn:
  Next j
End Sub

Public Function addcriteriaequal(ispec As ImportSpec) As String
   addcriteriaequal = "([" & ispec.DestField & "] = "
   If ispec.vartype = TYPE_BOOL Then
     addcriteriaequal = addcriteriaequal & Format$(ispec.Value, "True/False")
   ElseIf (ispec.vartype = TYPE_DOUBLE) Or (ispec.vartype = TYPE_LONG) Then
     'CAD P11D numsql
     addcriteriaequal = addcriteriaequal & NumSQL(ispec.Value)
   ElseIf ispec.vartype = TYPE_DATE Then
     addcriteriaequal = addcriteriaequal & DateSQL(ispec.Value)
   ElseIf ispec.vartype = TYPE_STR Then
     'cad P11d, str sql
     addcriteriaequal = addcriteriaequal & StrSQL(ispec.Value)
   Else
     Call ECASE("Import AddCriteria - Unknown datatype")
   End If
   addcriteriaequal = addcriteriaequal & ")"
End Function

Public Function FieldsLinkCount(fSpecs As FieldSpecs, ByVal IncludeHidden As Boolean) As Long
  Dim vCount As Long, hCount As Long
  Dim i As Long
  
  vCount = 0: hCount = 0
  For i = 1 To fSpecs.Count
    If Len(fSpecs(i).DestField) > 0 Then
      If fSpecs(i).Hide Then
        hCount = hCount + 1
      Else
        vCount = vCount + 1
      End If
    End If
  Next i
  FieldsLinkCount = vCount
  If IncludeHidden Then FieldsLinkCount = FieldsLinkCount + hCount
End Function
  
Public Function IsValidValue(ByVal v As Variant) As Boolean
  IsValidValue = Not (IsEmpty(v) Or IsNull(v))
End Function
  

Public Function CLngEx(v As Variant, ByVal Default As Long) As Long
  On Error GoTo CLngEx_err
  CLngEx = CLng(v)
CLngEx_end:
  Exit Function
CLngEx_err:
  CLngEx = Default
  Resume CLngEx_end
End Function
