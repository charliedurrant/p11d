VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.ocx"
Begin VB.Form F_Testing 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Testing"
   ClientHeight    =   9240
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9330
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9240
   ScaleWidth      =   9330
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSWinsockLib.Winsock winsock 
      Left            =   240
      Top             =   720
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin RichTextLib.RichTextBox richTextResults 
      Height          =   7695
      Left            =   120
      TabIndex        =   4
      Top             =   1440
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   13573
      _Version        =   393217
      TextRTF         =   $"F_Testing.frx":0000
   End
   Begin VB.TextBox txtGoogleSheetId 
      Height          =   375
      Left            =   1560
      TabIndex        =   2
      Text            =   "1h4p4kJraA_WiVl5kUiQvW2kXxaXvvdQ7CInDHSTYxaI"
      Top             =   240
      Width           =   7575
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   7920
      TabIndex        =   1
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Run"
      Default         =   -1  'True
      Height          =   375
      Left            =   6480
      TabIndex        =   0
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Note the google sheet is is per year and some of the import specs in the sheet need the year changing"
      Height          =   495
      Left            =   840
      TabIndex        =   7
      Top             =   720
      Width           =   4095
   End
   Begin VB.Label lblMessage 
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   5040
      TabIndex        =   6
      Top             =   720
      Width           =   3975
   End
   Begin VB.Label Label2 
      Caption         =   "Results"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Google sheet Id"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   360
      Width           =   1215
   End
End
Attribute VB_Name = "F_Testing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_googleSheet As GoogleSheet
Private Const GOOGLE_CLIENT_ID As String = "154107348377-juocv2nkknla74td3h06934el9je9dru.apps.googleusercontent.com"
Private Const GOOGLE_CLIENT_SEC As String = "AQAAANCMnd8BFdERjHoAwE/Cl+sBAAAAvMtiR1IXlUGG0iwZwm1/cAQAAAAsAAAAYQBiAGEAdABlAGMAIABlAG4AYwByAHkAcAB0AGUAZAAgAGQAYQB0AGEAAAAQZgAAAAEAACAAAAC2F04u+PSxeVhyeM1rk12jtxiYz2kPK1QoBKvUpZlQGwAAAAAOgAAAAAIAACAAAAB0+3zFx42mNhhgTnRWAmHpoOTOTDpjyiC9eCoja2L26VAAAADllFjkC78SluoceFZToeiO4pNPlnFQ5p6lj/TyUYX9U92jFwowKOJWF+52rHzDlb++XRWrCQ9jlVluZpGDz8UhCypFlVLY9RDNu/9O6fEv3kAAAABWE08GpEbdoFryHIfPQ3CwLOYROG1W0oNqCBHuwuNNvyjDm7n2KRx41VwGVBaSnEizHqp/REYkouXb09vkz2tL"
Private Const GOOGLE_REDIRECT_URL As String = "http://localhost:5001"

Private Sub cmdCancel_Click()
  Call Unload(Me)
End Sub

Private Sub cmdOk_Click()
  Dim results As String
  Dim googleSheetId As String
  Dim gSheet As GoogleSheet
  
  
  On Error GoTo err_err
  
  googleSheetId = Trim$(txtGoogleSheetId.Text)
  
  If Len(googleSheetId) = 0 Then
    Call Err.Raise(ERR_TESTING, "Testing", "Enter google sheet Guid (the sheet needs to be public access, see url for Guid)")
  End If
  
  Set m_googleSheet = New GoogleSheet
  If Not m_googleSheet.Initialise(googleSheetId) Then
    Call Err.Raise(ERR_ERROR, "Testing", "A sheet with Id '" & googleSheetId & "' either does not exist or is not public access")
  End If
                  
  p11d32.Testing.LastGoogleTestSheetId = googleSheetId
  
  If (p11d32.Testing.GoggleAccessTokenExpired) Then
    Dim url As String
    url = "https://accounts.google.com/o/oauth2/v2/auth?client_id=" & GOOGLE_CLIENT_ID & "&scope=https://www.googleapis.com/auth/spreadsheets&response_type=code&redirect_uri=" & GOOGLE_REDIRECT_URL
    Call ShellExecute(Me.hwnd, "open", url, 0, 0, 1)
    lblMessage.caption = "A new browser has started, plaase login to google"
    GoTo err_end
  Else
    Call Run
  End If
  
  
  
err_end:
  Exit Sub
err_err:
Call ErrorMessage(ERR_ERROR, Err, "cmdOk_Click", "Test via google sheet", Err.Description)
  Resume err_end
  Resume
End Sub

Private Sub Reply(fail As Boolean)
  Dim Message As String
  Dim data As String
  
  Dim page As String
  page = ""
  page = page & "<html>"
    page = page & "<head>"
    page = page & "</head>"
    page = page & "<body>"
      If fail Then
        page = page & "Failed, try again in the application"
      Else
        page = page & "Success, return to application"
      End If
      
      page = page & "<script>"
      If (Not fail) Then
        page = page & "window.close();"
      End If
      page = page & "</script>"
    
    page = page & "</body>"
  page = page & "</html>"
  
  data = "HTTP/1.1 200 OK" & vbCrLf & _
          "Connection: Close" & vbCrLf & _
          "Content-Length: " & CStr(Len(page)) & vbCrLf & _
          vbCrLf & page
  winsock.SendData (data)
End Sub
Private Sub ProcessMessage(Message As String)
  Dim Buff As String
  Dim data As String
  Dim failed As Boolean
  Dim p0 As Long, p1 As Long
  Const MATCH As String = "GET /?code="
  Const MATCH_REFERRER As String = "Referer: http://localhost:5001"
  lblMessage.caption = ""
  
  p0 = InStr(1, Message, MATCH)
  
  
  If p0 = 0 Then
    p0 = InStr(1, Message, "Referer: " & GOOGLE_REDIRECT_URL)
    If (p0 <> 0) Then
      'ignore message
      Exit Sub
    End If
    failed = True
  End If
  
  Call Reply(failed)
  
  If Not failed Then
    Dim authCode As String
    p1 = InStr(p0, Message, "&")
    p0 = p0 + Len(MATCH)
    authCode = Mid(Message, p0, p1 - p0)
    Dim url As String
    Dim xmlhttp As XMLHTTP60
    
    Set xmlhttp = New XMLHTTP60
    
    url = "https://accounts.google.com/o/oauth2/token"
    
    Call xmlhttp.Open("POST", url, False)
    xmlhttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    xmlhttp.setRequestHeader "User-Agent", "Firefox 3.6.4"
    
    Dim secretDecrypted As String
    
    secretDecrypted = DecryptLMSecret(GOOGLE_CLIENT_SEC)
    
    data = "client_id=" & GOOGLE_CLIENT_ID & "&client_secret=" & secretDecrypted & "&redirect_uri=" & GOOGLE_REDIRECT_URL
    data = data & "&grant_type=authorization_code"
    data = data & "&code=" & authCode
    
    Call xmlhttp.send(data)
    
    Dim jsonString As String
    Dim jsonObj As Object
    jsonString = xmlhttp.responseText
    Set jsonObj = JSON.parse(jsonString)
    
    p11d32.Testing.GoogleAccessToken = jsonObj("access_token")
    Dim expiresIn As Long
    expiresIn = CLng(jsonObj("expires_in"))
    p11d32.Testing.GoogleAccessTokenExpiresAt = DateAdd("s", expiresIn - 500, Now)
    Call Run
  Else
    lblMessage.caption = "Please try again"
  End If
  

End Sub

Private Sub Run()
  Dim results As String
  Dim failed As Boolean
  
On Error GoTo err_err

  m_googleSheet.AccessToken = p11d32.Testing.GoogleAccessToken
  
  Call SetCursor(vbHourglass)
  Call p11d32.Testing.Init(m_googleSheet)
  
  failed = p11d32.Testing.Run(results)
  richTextResults.Text = results
  ClearAllCursors
  
  Call Me.SetFocus
  If failed Then
    Call MsgBox("Tests failed")
  Else
    Call MsgBox("Tests succeeded")
  End If
    
err_end:
  ClearAllCursors
  Exit Sub
err_err:
  ClearAllCursors
  Call ErrorMessage(ERR_ERROR, Err, "mnuTestingGoogleSheet_Click", "Test via google sheet", Err.Description)
  Resume err_end
  Resume
End Sub
Private Sub Form_Load()
  txtGoogleSheetId.Text = p11d32.Testing.LastGoogleTestSheetId
  With winsock
    .LocalPort = 5001                      'set the port to listen on
    .Listen                                'start listening
  End With 'Winsock1
End Sub
Private Sub Form_Unload(Cancel As Integer)
  winsock.Close
End Sub

Private Sub winsock_ConnectionRequest(ByVal requestID As Long)
  With winsock
    If .State <> sckClosed Then .Close     'close the port when not closed (you could also use another winsock control to accept the connection)
    .Accept requestID                      'accept the connection request
  End With 'Winsock1
End Sub
Private Sub winsock_Close()
  winsock.Close
  winsock.Listen
  
End Sub

Private Sub winsock_Error(ByVal Number As Integer, Description As String, ByVal sCode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
  Call MsgBox(Description, vbCritical, "Error " & CStr(Number))
End Sub
Private Sub winsock_DataArrival(ByVal bytesTotal As Long)
  Dim strData As String
  winsock.GetData strData                 'get the data
  Call ProcessMessage(strData)                      'process the data
End Sub



