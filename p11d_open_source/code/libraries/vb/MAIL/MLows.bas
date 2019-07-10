Attribute VB_Name = "MAPI_Lows"
Option Explicit

Private Declare Function MAPIFindNext Lib "MAPI32.DLL" Alias "BMAPIFindNext" (ByVal SessionID As Long, ByVal UIParam As Long, MsgType As String, SeedMsgID As String, ByVal Flags As Long, ByVal Reserved As Long, MsgID As String) As Long
Private Declare Function BMAPIReadMail Lib "MAPI32.DLL" (msgPtr As Long, nRecipients As Long, nFiles As Long, ByVal SessionID As Long, ByVal UIParam As Long, MsgID As String, ByVal Flags As Long, ByVal Reserved As Long) As Long
'Private Declare Function BMAPIGetReadMail Lib "MAPI32.DLL" (ByVal msgPtr As Long, ByVal Message As Long, ByVal Recip As Long, ByVal File As Long, ByVal Originator As Long) As Long

Private Declare Function BMAPIGetReadMail Lib "MAPI32.DLL" (ByVal lMsg&, Message As MapiMessage, Recip() As MapiRecip, File() As MAPIfile, Originator As MapiRecip) As Long

Private Declare Function MAPIDeleteMail Lib "MAPI32.DLL" (ByVal SessionID As Long, ByVal UIParam As Long, ByVal MsgID As String, ByVal Flags As Long, ByVal Reserved As Long) As Long


'Declare Function BMAPIGetReadMail Lib "MAPI.DLL" (ByVal lMsg&, Message As MapiMessage, Recip As MapiRecip, File As MAPIfile, Originator As MapiRecip) As Long
'Private Declare Function MAPIReadMail Lib "MAPI32.DLL" (ByVal SessionID As Long, ByVal UIParam As Long, MsgID As String, ByVal Flags As Long, ByVal Reserved As Long, lpMapiMessage As Long) As Long
Private Declare Function MAPIFreeBuffer Lib "MAPI32.DLL" (ByVal lpMapiMessage As Long) As Long

 
Type MapiMessage
  Reserved As Long
  Subject As String
  NoteText As String
  MessageType As String
  DateReceived As String
  ConversationID As String
  Flags As Long
  RecipCount As Long
  FileCount As Long
End Type
 
Type MapiRecip
  Reserved As Long
  RecipClass As Long
  Name As String
  Address As String
  EIDSize As Long
  EntryID As String
End Type
 
Type MAPIfile
  Reserved As Long
  Flags As Long
  Position As Long
  PathName As String
  FileName As String
  FileType As String
End Type

Type MessageHeader
  FromName As String
  FromAddress As String
  DateReceived As String
  Subject As String
  Body As String
  MsgRead As Boolean
  Sent As Boolean
End Type


'* MAPIFindNext() flags *
Private Const MAPI_UNREAD_ONLY As Long = &H20
Private Const MAPI_GUARANTEE_FIFO As Long = &H100
Private Const MAPI_LONG_MSGID As Long = &H4000

'* MAPIReadMail() flags *
Private Const MAPI_ENVELOPE_ONLY As Long = &H40
Private Const MAPI_PEEK As Long = &H80
Private Const MAPI_BODY_AS_FILE As Long = &H200
Private Const MAPI_SUPPRESS_ATTACH As Long = &H800

'* MAPI GetMail flags
Private Const MAPI_UNREAD As Long = &H1
Private Const MAPI_RECEIPT_REQUESTED As Long = &H2
Private Const MAPI_SENT As Long = &H4


' Return values
Private Const SUCCESS_SUCCESS = 0
Private Const MAPI_E_INSUFFICIENT_MEMORY = 5
Private Const MAPI_E_INVALID_MESSAGE = 17
Private Const MAPI_E_INVALID_SESSION = 19
Private Const MAPI_E_NO_MESSAGES = 16

Private Const MSGID_LENGTH As Long = 512
 
Public Function GetNextMessage(ByVal SessionID As Long, MsgID As String) As Boolean
  Dim retval As Long
  Dim iSeedID As String, iMsgID As String
  
fetch_next:
  iSeedID = MsgID & String$(MSGID_LENGTH, vbNullChar)
  iMsgID = String$(MSGID_LENGTH, vbNullChar)
  retval = MAPIFindNext(SessionID, 0, vbNullString, iSeedID, MAPI_LONG_MSGID, 0, iMsgID)
  MsgID = iMsgID
  If retval = MAPI_E_INVALID_MESSAGE Then GoTo fetch_next
  If retval = MAPI_E_INVALID_SESSION Then Err.Raise ERR_MAPI_INVSESSION, "GetNextMessage", "GetNextMessage failed due to invalid session MAPI error code=" & retval
  If Not ((retval = MAPI_E_NO_MESSAGES) Or (retval = SUCCESS_SUCCESS)) Then Err.Raise ERR_MAPI_FETCHNEXT, "GetNextMessage", "GetNextMessage failed with MAPI error code=" & retval
  GetNextMessage = (retval = SUCCESS_SUCCESS)
End Function

Public Function ReadMessage(msg As MessageHeader, ByVal SessionID As Long, ByVal MsgID As String, ByVal HeaderOnly As Boolean, Atts As MailAttachments) As Boolean
  Dim retval As Long, Flags As Long, i As Long
  Dim Info As Long, nFiles As Long, nRecips As Long
  Dim Recips() As MapiRecip, Files() As MAPIfile
  Dim Originator As MapiRecip, Message As MapiMessage
  
  retval = SUCCESS_SUCCESS
  If HeaderOnly Then Flags = MAPI_ENVELOPE_ONLY + MAPI_PEEK
  retval = BMAPIReadMail(Info, nRecips, nFiles, SessionID, 0, MsgID, Flags, 0)
  If retval = SUCCESS_SUCCESS Then
    'Message is now read into the handles array, We have to redim the arrays and read the stuff in
    If nRecips = 0 Then
      ReDim Recips(0 To 0) As MapiRecip
    Else
      ReDim Recips(0 To nRecips - 1) As MapiRecip
    End If
    If nFiles = 0 Then
      ReDim Files(0 To 0) As MAPIfile
    Else
      ReDim Files(0 To nFiles - 1) As MAPIfile
    End If

    retval = BMAPIGetReadMail(Info, Message, Recips(), Files(), Originator)
    If retval = SUCCESS_SUCCESS Then
      msg.FromName = Originator.Name
      msg.FromAddress = Originator.Address
      msg.Subject = Message.Subject
      msg.Body = Message.NoteText
      msg.DateReceived = Message.DateReceived
      msg.MsgRead = Not ((Message.Flags And MAPI_UNREAD) = MAPI_UNREAD)
      msg.Sent = (Message.Flags And MAPI_SENT) = MAPI_SENT
      If Not Atts Is Nothing Then
        For i = 0 To nFiles - 1
          Call Atts.Add("Attachment: " & CStr(i), Files(i).PathName, MAIL_ATT_DATA)
        Next i
      End If
    End If
  End If
  ReadMessage = (retval = SUCCESS_SUCCESS)
End Function
Private Function IsNothingVariant(v) As Boolean
  On Error GoTo IsNothingVariant_ERR
  
  If v Is Nothing Then IsNothingVariant = True
  
  Exit Function
IsNothingVariant_ERR:
End Function
Public Function VariantArrayToStringArray(vContainingStringArray As Variant, ParamArray vArray()) As Long
  Dim i As Long
  Dim j As Long
  Dim k As Long
  Dim s() As String
  Dim v As Variant
  
  On Error GoTo VariantArrayToStringArray_ERR

  If (Not IsEmpty(vArray)) And IsArray(vArray) Then
    For i = LBound(vArray) To UBound(vArray)
      If Not IsNothingVariant(vArray(i)) Then
        v = vArray(i)
        If (Not IsEmpty(v)) And IsArray(v) Then j = j + ((UBound(v) - LBound(v)) + 1)
      End If
    Next
  End If
  
  If j > 0 Then
    ReDim s(0 To j - 1)
    k = 0
    For i = LBound(vArray) To UBound(vArray)
      If Not IsNothingVariant(vArray(i)) Then
        v = vArray(i)
        If (Not IsEmpty(v)) And IsArray(v) Then
          For j = LBound(v) To UBound(v)
            k = k + 1
            s(k - 1) = v(j)
          Next
        End If
      End If
    Next
  End If
  
  vContainingStringArray = s()
  VariantArrayToStringArray = k
  
  
VariantArrayToStringArray_END:
  Exit Function
VariantArrayToStringArray_ERR:
  Call Err.Raise(ERR_VARIANT_ARRAY_TO_STRING_ARRAY, "VariantArrayToStringArray", Err.Description)
  Resume
End Function
Public Function DeleteMessage(ByVal SessionID As Long, ByVal MsgID As String) As Boolean
  Dim retval As Long
  retval = MAPIDeleteMail(SessionID, 0, MsgID, 0, 0)
  DeleteMessage = (retval = SUCCESS_SUCCESS)
End Function
