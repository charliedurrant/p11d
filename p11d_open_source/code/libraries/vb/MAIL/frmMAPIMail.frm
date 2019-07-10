VERSION 5.00
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "MSMAPI32.OCX"
Begin VB.Form frmMAPIMail 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "MAPI Mail"
   ClientHeight    =   3960
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6045
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3960
   ScaleWidth      =   6045
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   375
      Left            =   4680
      TabIndex        =   4
      Top             =   3480
      Width           =   1215
   End
   Begin VB.CheckBox chkDebug 
      Caption         =   "Debug Mode"
      Height          =   255
      Left            =   1440
      TabIndex        =   3
      Top             =   3480
      Width           =   1335
   End
   Begin VB.TextBox txtMsg 
      Height          =   2205
      Left            =   1440
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   1050
      Width           =   4455
   End
   Begin VB.TextBox txtSubject 
      Height          =   285
      Left            =   1440
      TabIndex        =   1
      Top             =   570
      Width           =   4455
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "&Send Mail"
      Default         =   -1  'True
      Height          =   375
      Left            =   3240
      TabIndex        =   5
      Top             =   3480
      Width           =   1215
   End
   Begin VB.TextBox txtTo 
      Height          =   285
      Left            =   1440
      TabIndex        =   0
      Top             =   90
      Width           =   4455
   End
   Begin MSMAPI.MAPIMessages MMessages 
      Left            =   120
      Top             =   2760
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      AddressEditFieldCount=   1
      AddressModifiable=   0   'False
      AddressResolveUI=   0   'False
      FetchSorted     =   0   'False
      FetchUnreadOnly =   0   'False
   End
   Begin MSMAPI.MAPISession MSession 
      Left            =   120
      Top             =   2040
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DownloadMail    =   0   'False
      LogonUI         =   -1  'True
      NewSession      =   0   'False
   End
   Begin VB.Label lblMsg 
      Caption         =   "Message:"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label lblSubject 
      Caption         =   "Subject:"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label lblTo 
      Caption         =   "To:"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "frmMAPIMail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mSessionID As Long
Private mDebug As Boolean
Public mUseAPI As Boolean

Public Property Get MAILDebug() As Boolean
  MAILDebug = mDebug
End Property

Public Property Let MAILDebug(ByVal NewValue As Boolean)
  mDebug = NewValue
  If NewValue Then
    chkDebug = vbChecked
  Else
    chkDebug = vbUnchecked
  End If
End Property

Private Sub MAPISignOn(MapiSess As MAPISession)
  
  On Error GoTo MAPISignOn_err
  If mSessionID = -1 Then
    MapiSess.DownLoadMail = True
    'MapiSess.Name
    'MapiSess.Password
    MapiSess.LogonUI = True
    MapiSess.SignOn
    mSessionID = MapiSess.SessionID
  End If
  
MAPISignOn_end:
  Exit Sub

MAPISignOn_err:
  mSessionID = -1
  Call ErrorMessage(ERR_ERROR, Err, "MAPISignOn", "MAPI initial session initialise", "Unable to initialise MAPI session.")
  Resume MAPISignOn_end
End Sub

Public Sub MAPISignOff()
  On Error GoTo MAPISignOff_Err
  Call xSet("MAPISignOff")
  
  If mSessionID <> -1 Then
    Call MSession.SignOff
    mSessionID = -1
  End If

MAPISignOff_End:
  Call xReturn("MAPISignOff")
  Exit Sub

MAPISignOff_Err:
  Call ErrorMessage(ERR_ERROR, Err, "MAPISignOff", "MAPI session close", "Error closing MAPI session.")
  Resume MAPISignOff_End
End Sub

Private Sub chkDebug_Click()
  mDebug = (chkDebug = vbChecked)
End Sub

Private Sub cmdClose_Click()
  Unload Me
End Sub

Private Sub cmdSend_Click()
  Dim v As Variant
  If (Len(txtTo) > 0) And ((Len(txtSubject) > 0) Or (Len(txtMsg) > 0)) Then
    Call MAPISignOn(MSession)
    ReDim v(0 To 0) As Variant
    v(0) = txtTo
    Call MAILSendEx(txtSubject, txtMsg, v, Empty, Nothing)
  End If
End Sub

Private Sub Form_Initialize()
  mSessionID = -1
End Sub

Private Sub Form_Load()
  txtTo = GetStatic("MAPISendTo")
  txtSubject = GetStatic("MAPISubject")
End Sub

Private Sub Form_Unload(Cancel As Integer)
  If Len(txtTo) > 0 Then Call AddStatic("MAPISendTo", , txtTo)
  If Len(txtSubject) > 0 Then Call AddStatic("MAPISubject", , txtSubject)
  Call MAPISignOff
End Sub

Public Sub MAILSendEx(ByVal msgSubject As String, ByVal msgBody As String, ToList As Variant, CCList As Variant, Attachments As MailAttachments)
  Dim i As Long, rIndex As Long, lAttatchmentIndex As Long
  Dim Attachment As MailAttachment, cChars As Long
  
  On Error GoTo MAILSendEx_err
  Call MAPISignOn(MSession)
  If mSessionID <> -1 Then
    MMessages.SessionID = mSessionID
    MMessages.AddressResolveUI = False
    Call MMessages.Compose
    rIndex = 0
    If (Not IsEmpty(ToList)) And IsArray(ToList) Then
      For i = LBound(ToList) To UBound(ToList)
        MMessages.RecipIndex = rIndex
        MMessages.RecipType = mapToList
        MMessages.RecipDisplayName = ToList(i)
        rIndex = rIndex + 1
      Next i
    End If
    If (Not IsEmpty(CCList)) And IsArray(CCList) Then
      For i = LBound(CCList) To UBound(CCList)
        MMessages.RecipIndex = rIndex
        MMessages.RecipType = mapToList
        MMessages.RecipDisplayName = CCList(i)
        rIndex = rIndex + 1
      Next i
    End If
    MMessages.msgSubject = msgSubject
    MMessages.MsgNoteText = msgBody
    cChars = Len(msgBody)
    If Not Attachments Is Nothing Then
      MMessages.MsgNoteText = MMessages.MsgNoteText & Space$(Attachments.Count)
      lAttatchmentIndex = 0
      For Each Attachment In Attachments
        MMessages.AttachmentIndex = lAttatchmentIndex
        MMessages.AttachmentPathName = Attachment.FileName
        MMessages.AttachmentType = Attachment.MAILType
        MMessages.AttachmentPosition = cChars + lAttatchmentIndex
        lAttatchmentIndex = lAttatchmentIndex + 1
      Next
    End If
    Call MMessages.Send(mDebug)
  End If
  
MAILSendEx_end:
  Exit Sub

MAILSendEx_err:
  Call ErrorMessage(ERR_ERROR, Err, "MAILSendEx", "MAIL Sendmail (extended)", "Unable to complete send mail")
  Resume MAILSendEx_end
  Resume
End Sub

Public Sub ForEachMessage(msg As IMailMessage, ByVal UnReadOnly As Boolean, ByVal FromFilter As String, ByVal SubjectFilter As String, NotifyIF As IBaseNotify, ByVal CutOffDate As Date)
  Dim Atts As MailAttachments
  Dim msgAction As MessageAction
  Dim inFetch As Boolean
  Dim msgDate As Date
  Dim MsgID As String, msgOK As Boolean, LastValidMsgID As String
  Dim sFrom As String, sSubject As String
  Dim readMsg As Boolean
  Dim nMsg As Long, i As Long, j As Long, errno As Long
  Dim msgHeader As MessageHeader
  
  On Error GoTo ForEachMessage_Err
  Call xSet("ForEachMessage")
  inFetch = False
  Call SetCursor
  Call MAPISignOn(MSession)
  DoEvents
  If mSessionID <> -1 Then
    If Not mUseAPI Then
      MMessages.SessionID = mSessionID
      'MMessages.FetchMsgType = "IPM.NOTE"
      MMessages.FetchSorted = True
      MMessages.FetchUnreadOnly = False
      MMessages.Fetch
      nMsg = (MMessages.MsgCount - 1)
    Else
      MsgID = ""
      nMsg = -1
    End If
    i = 0
    Do While (i <= nMsg) Or mUseAPI
      inFetch = True
      If mUseAPI Then
        LastValidMsgID = MsgID
        If Not GetNextMessage(mSessionID, MsgID) Then Exit Do
      End If
      If Not NotifyIF Is Nothing Then Call NotifyIF.Notify(i, nMsg, "Retrieving message")
      If mUseAPI Then
        If Not ReadMessage(msgHeader, mSessionID, MsgID, True, Nothing) Then GoTo skip_message
        If UnReadOnly And msgHeader.MsgRead Then GoTo skip_message
        sFrom = msgHeader.FromName
        sSubject = msgHeader.Subject
        msgDate = CDateEx(msgHeader.DateReceived, Now)
      Else
        MMessages.MsgIndex = i
        If UnReadOnly And MMessages.MsgRead Then GoTo skip_message
        sFrom = MMessages.MsgOrigDisplayName
        sSubject = MMessages.msgSubject
        msgDate = CDateEx(MMessages.MsgDateReceived, Now)
      End If
      readMsg = True
      If Len(FromFilter) > 0 Then
        readMsg = (sFrom Like FromFilter)
      End If
      If Len(SubjectFilter) > 0 Then
        readMsg = (sSubject Like SubjectFilter)
      End If
      
      If CutOffDate <> UNDATED Then
        If msgDate < CutOffDate Then GoTo skip_message
      End If
      If readMsg Then
        Set Atts = New MailAttachments
        If mUseAPI Then
          If Not ReadMessage(msgHeader, mSessionID, MsgID, False, Atts) Then GoTo skip_message
          msgAction = msg.Message(sFrom, sSubject, msgHeader.Body, Atts)
          If msgAction = MESSAGE_DELETE Then
            If Not DeleteMessage(mSessionID, MsgID) Then Err.Raise ERR_DELETEFAILED, "ForEachMessage", "Unable to delete message." & vbCrLf & "From: " & sFrom & vbCrLf & "Subject: " & sSubject
            MsgID = LastValidMsgID
          End If
        Else
          For j = 0 To (MMessages.AttachmentCount - 1)
            MMessages.AttachmentIndex = j
            Call Atts.Add("Attachment: " & CStr(j), MMessages.AttachmentPathName, MMessages.AttachmentType)
          Next j
          msgAction = msg.Message(sFrom, sSubject, MMessages.MsgNoteText, Atts)
          If msgAction = MESSAGE_DELETE Then
            Call MMessages.Delete(mapMessageDelete)
            ' messages reindexed on delete
            nMsg = nMsg - 1
            i = i - 1
          End If
        End If
        Call Atts.RemoveAll
      End If
skip_message:
      i = i + 1
    Loop
    inFetch = False
  End If
  
ForEachMessage_End:
  Set Atts = Nothing
  Call ClearCursor
  Call xReturn("ForEachMessage")
  Exit Sub

ForEachMessage_Err:
  errno = Err.Number
  Call ErrorMessage(ERR_ERROR + ERR_ALLOWIGNORE, Err, "ForEachMessage", "Mail Fetch mail messages", "Error fetching mail messages")
  If errno = ERR_MAPI_INVSESSION Then Resume ForEachMessage_End
  If inFetch Then Resume skip_message
  Resume ForEachMessage_End
End Sub

Private Function CDateEx(ByVal DateString As String, ByVal DefaultDate As Date) As Date
  On Error GoTo CDateEx_err
  CDateEx = CDate(DateString)
  
CDateEx_end:
  Exit Function
  
CDateEx_err:
  CDateEx = DefaultDate
  Resume CDateEx_end
End Function
