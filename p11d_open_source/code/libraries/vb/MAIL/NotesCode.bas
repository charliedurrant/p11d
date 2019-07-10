Attribute VB_Name = "NotesCode"
Option Explicit
Private Sub BodyToNote(msgBody As String, ByVal ne As NoteEdit, ByVal lNotesBodyFontSize As Long)
  Dim p0 As Long, p1 As Long
  Dim lLen As Long
  Dim sLine As String
  
  On Error GoTo BodyToNote_err
  lLen = Len(msgBody)
  If lLen = 0 Then GoTo BodyToNote_end
  p0 = 1
  p1 = InStr(1, msgBody, vbCrLf, vbBinaryCompare)
  Do While p1 > 0
    sLine = Mid$(msgBody, p0, p1 - p0)
    Call ne.AppendRichText("Body", sLine, lNotesBodyFontSize)
    p0 = p1 + 2
    If p0 > lLen Then GoTo BodyToNote_end
    p1 = InStr(p0, msgBody, vbCrLf, vbBinaryCompare)
    sLine = ""
  Loop
  
BodyToNote_end:
  Exit Sub
BodyToNote_err:
  Call Err.Raise(ERR_BODY_TO_NOTE, "BodyToNote", "Error in BodyToNote")
End Sub

Public Sub NotesMailSendEx(ByVal dbMail As NotesDB, ByVal msgSubject As String, ByVal msgBody As String, ToList As Variant, CCList As Variant, Attachments As MailAttachments, ByVal lNotesBodyFontSize As Long)
  Dim Attachment As MailAttachment
  Dim ne As NoteEdit
  Dim v As Variant
  
  On Error GoTo NotesMailSendEx_err
  Set ne = dbMail.CreateMail()
  If VariantArrayToStringArray(v, ToList) > 0 Then Call ne.SetField("SendTo", v)
  If VariantArrayToStringArray(v, CCList) > 0 Then Call ne.SetField("CopyTo", v)
  If VariantArrayToStringArray(v, ToList, CCList) > 0 Then Call ne.SetField("Recipients", v)
  
  Call ne.SetField("Subject", msgSubject)
  Call ne.SetField("PostedDate", FormatDateTime(Now, vbGeneralDate))
  Call BodyToNote(msgBody, ne, lNotesBodyFontSize)
  Call ne.SetField("LastName", dbMail.GetUserName)
  Call ne.SetField("Form", "Memo")
  Call ne.SetField("From", dbMail.GetUserName)
  
  If Not Attachments Is Nothing Then
    For Each Attachment In Attachments
      Call ne.AttachFile(Attachment.FileName, Attachment.DisplayName)
    Next
  End If
  Call ne.SendMail
  Call ne.CloseNote(True)
  
NotesMailSendEx_end:
  Set ne = Nothing
  Exit Sub
NotesMailSendEx_err:
  Call ErrorMessage(ERR_ERROR, Err, ErrorSource(Err, "NotesMailSendEx"), "Notes Mail Sendmail (extended)", "Unable to complete send mail (Notes)")
  Resume NotesMailSendEx_end
  Resume
End Sub

