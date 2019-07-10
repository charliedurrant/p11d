VERSION 5.00
Object = "{AF27A9B5-A3F4-11D2-8DB7-00C04FA9DD6F}#1.2#0"; "TCSPROG.OCX"
Begin VB.Form frmCancel 
   Caption         =   "frmCancel"
   ClientHeight    =   8670
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4455
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   8670
   ScaleWidth      =   4455
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picAVI 
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   120
      ScaleHeight     =   735
      ScaleWidth      =   4095
      TabIndex        =   11
      Top             =   1080
      Width           =   4095
   End
   Begin VB.Frame fraTimings 
      BorderStyle     =   0  'None
      Caption         =   "fraTimings"
      Height          =   5055
      Left            =   0
      TabIndex        =   3
      Top             =   2760
      Width           =   4335
      Begin VB.Frame fraFileTimings 
         Caption         =   "File timings"
         Height          =   3675
         Left            =   120
         TabIndex        =   12
         Top             =   600
         Width           =   4215
         Begin VB.Label lblCloseLast 
            Caption         =   "lblCloseLast"
            Height          =   255
            Left            =   1320
            TabIndex        =   40
            Top             =   2280
            Width           =   1155
         End
         Begin VB.Label lblOpenLast 
            Caption         =   "lblOpenLast"
            Height          =   255
            Left            =   1320
            TabIndex        =   39
            Top             =   600
            Width           =   1155
         End
         Begin VB.Label lblExtractLast 
            Caption         =   "lblExtractLast"
            Height          =   255
            Left            =   1320
            TabIndex        =   38
            Top             =   960
            Width           =   1155
         End
         Begin VB.Label lblSaveLast 
            Caption         =   "lblSaveLast"
            Height          =   255
            Left            =   1320
            TabIndex        =   37
            Top             =   2640
            Width           =   1155
         End
         Begin VB.Label lblOpen 
            Caption         =   "Abacus Open"
            Height          =   255
            Left            =   120
            TabIndex        =   36
            Tag             =   "-1"
            Top             =   600
            Width           =   1095
         End
         Begin VB.Label lblExtract 
            Caption         =   "Extract"
            Height          =   255
            Left            =   120
            TabIndex        =   35
            Tag             =   "-1"
            Top             =   960
            Width           =   495
         End
         Begin VB.Label lblClose 
            Caption         =   "Abacus Close"
            Height          =   255
            Left            =   120
            TabIndex        =   34
            Tag             =   "-1"
            Top             =   2280
            Width           =   975
         End
         Begin VB.Label lblSave 
            Caption         =   "Save"
            Height          =   255
            Left            =   120
            TabIndex        =   33
            Tag             =   "-1"
            Top             =   2640
            Width           =   495
         End
         Begin VB.Label lblSaveAll 
            Caption         =   "lblSaveAll"
            Height          =   255
            Left            =   2640
            TabIndex        =   32
            Top             =   2640
            Width           =   1155
         End
         Begin VB.Label lblExtractAll 
            Caption         =   "lblExtractAll"
            Height          =   255
            Left            =   2640
            TabIndex        =   31
            Top             =   960
            Width           =   1155
         End
         Begin VB.Label lblOpenAll 
            Caption         =   "lblOpenAll"
            Height          =   255
            Left            =   2640
            TabIndex        =   30
            Top             =   600
            Width           =   1155
         End
         Begin VB.Label lblCloseAll 
            Caption         =   "lblCloseAll"
            Height          =   255
            Left            =   2640
            TabIndex        =   29
            Top             =   2280
            Width           =   1155
         End
         Begin VB.Label lblTotalAll 
            Caption         =   "lblTotalAll"
            Height          =   255
            Left            =   2640
            TabIndex        =   28
            Top             =   3000
            Width           =   1155
         End
         Begin VB.Label lblTotal 
            Caption         =   "Total:"
            Height          =   255
            Left            =   120
            TabIndex        =   27
            Tag             =   "-1"
            Top             =   3000
            Width           =   495
         End
         Begin VB.Label lblTotalLast 
            Caption         =   "lblTotalLast"
            Height          =   255
            Left            =   1320
            TabIndex        =   26
            Top             =   3000
            Width           =   1155
         End
         Begin VB.Label Label1 
            Caption         =   "Last file"
            Height          =   255
            Left            =   1320
            TabIndex        =   25
            Top             =   240
            Width           =   1155
         End
         Begin VB.Label Label2 
            Caption         =   "All files"
            Height          =   255
            Left            =   2640
            TabIndex        =   24
            Top             =   240
            Width           =   1155
         End
         Begin VB.Label lblEvaluateAll 
            Caption         =   "lblEvaluateAll"
            Height          =   255
            Left            =   2640
            TabIndex        =   23
            Top             =   1320
            Width           =   1155
         End
         Begin VB.Label Label4 
            Caption         =   "Evaluate"
            Height          =   255
            Left            =   120
            TabIndex        =   22
            Tag             =   "-1"
            Top             =   1320
            Width           =   1455
         End
         Begin VB.Label lblEvaluateLast 
            Caption         =   "lblEvaluateLast"
            Height          =   255
            Left            =   1320
            TabIndex        =   21
            Top             =   1320
            Width           =   1155
         End
         Begin VB.Label lblExecuteAll 
            Caption         =   "lblExecuteAll"
            Height          =   255
            Left            =   2640
            TabIndex        =   20
            Top             =   1560
            Width           =   1155
         End
         Begin VB.Label Label7 
            Caption         =   "Execute"
            Height          =   255
            Left            =   120
            TabIndex        =   19
            Tag             =   "-1"
            Top             =   1560
            Width           =   1095
         End
         Begin VB.Label lblExecuteLast 
            Caption         =   "lblExecuteLast"
            Height          =   255
            Left            =   1320
            TabIndex        =   18
            Top             =   1560
            Width           =   1155
         End
         Begin VB.Label lblSaveCursorAll 
            Caption         =   "lblSaveCursorAll"
            Height          =   255
            Left            =   2640
            TabIndex        =   17
            Top             =   1800
            Width           =   1155
         End
         Begin VB.Label Label13 
            Caption         =   "SaveCursor"
            Height          =   255
            Left            =   120
            TabIndex        =   16
            Tag             =   "-1"
            Top             =   1800
            Width           =   1095
         End
         Begin VB.Label lblSaveCursorLast 
            Caption         =   "lblSaveCursorLast"
            Height          =   255
            Left            =   1320
            TabIndex        =   15
            Top             =   1800
            Width           =   1155
         End
         Begin VB.Label lblFileCount 
            Caption         =   "lblFileCount"
            Height          =   255
            Left            =   240
            TabIndex        =   14
            Top             =   3360
            Width           =   1815
         End
         Begin VB.Label lblExceptionCount 
            Caption         =   "lblExceptionCount"
            Height          =   255
            Left            =   2160
            TabIndex        =   13
            Top             =   3360
            Width           =   1815
         End
      End
      Begin VB.Frame fraUDM 
         Caption         =   "UDM timings"
         Height          =   615
         Left            =   120
         TabIndex        =   7
         Top             =   4320
         Width           =   4215
         Begin VB.Label lblLineLast 
            Caption         =   "lblLineLast"
            Height          =   255
            Left            =   1320
            TabIndex        =   9
            Top             =   240
            Width           =   1155
         End
         Begin VB.Label lblLineAll 
            Caption         =   "lblLineAll"
            Height          =   255
            Left            =   2640
            TabIndex        =   8
            Top             =   240
            Width           =   1155
         End
      End
      Begin VB.Frame fraFileGroupTimings 
         Caption         =   "File group timings"
         Height          =   555
         Left            =   120
         TabIndex        =   4
         Top             =   0
         Width           =   4215
         Begin VB.Label lblFileGroupAll 
            Caption         =   "lblFileGroupAll"
            Height          =   255
            Left            =   2640
            TabIndex        =   6
            Top             =   240
            Width           =   1155
         End
         Begin VB.Label lblFileGroupLast 
            Caption         =   "lblFileGroupLast"
            Height          =   255
            Left            =   1320
            TabIndex        =   5
            Top             =   240
            Width           =   1155
         End
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3120
      TabIndex        =   0
      Top             =   8145
      Width           =   1215
   End
   Begin TCSPROG.TCSProgressBar PBar 
      Height          =   285
      Left            =   120
      TabIndex        =   10
      Top             =   1920
      Width           =   4185
      _cx             =   7382
      _cy             =   503
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Min             =   0
      Max             =   100
      Value           =   50
      BarBackColor    =   -2147483633
      BarForeColor    =   12937777
      Appearance      =   1
      Style           =   0
      CaptionColor    =   0
      CaptionInvertColor=   16777215
      FillStyle       =   0
      FadeFromColor   =   0
      FadeToColor     =   16777215
      Caption         =   ""
      InnerCircle     =   0   'False
      Percentage      =   0
      Skew            =   0
      PictureOffsetTop=   0
      PictureOffsetLeft=   0
      Enabled         =   -1  'True
      Increment       =   1
      TextAlignment   =   2
   End
   Begin TCSPROG.TCSProgressBar PBarStep 
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Top             =   2280
      Visible         =   0   'False
      Width           =   4185
      _cx             =   7382
      _cy             =   503
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Min             =   0
      Max             =   100
      Value           =   50
      BarBackColor    =   -2147483633
      BarForeColor    =   12937777
      Appearance      =   1
      Style           =   0
      CaptionColor    =   0
      CaptionInvertColor=   16777215
      FillStyle       =   0
      FadeFromColor   =   0
      FadeToColor     =   16777215
      Caption         =   ""
      InnerCircle     =   0   'False
      Percentage      =   0
      Skew            =   0
      PictureOffsetTop=   0
      PictureOffsetLeft=   0
      Enabled         =   0   'False
      Increment       =   1
      TextAlignment   =   2
   End
   Begin VB.Label lblStatus 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4215
   End
End
Attribute VB_Name = "frmCancel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public rw As ReportWizard
Private m_AVI As AVI
Private m_Loaded As Boolean
Private m_Started As Boolean
Private m_Cancelled As Boolean
Public Excepted As Boolean
Private Sub cmdCancel_Click()
  m_Cancelled = True
  Call StopAVI
  rw.Cancel = True
End Sub
Private Sub DisplayControls()
  On Error GoTo DisplayControls_Err
  Call xSet("DisplayControls")
  
  'Set form size without timings frame
  'RK TODO: Scale heights/top more sensibly
  
  fraTimings.Visible = False
  'PBarStep.Height +
  Me.Height = lblStatus.Height + cmdCancel.Height + PBar.Height + picAVI.Height + (L_MARGIN_LENGTH * 6)
  
  lblStatus.top = L_MARGIN_LENGTH
    
  'Locate Cancel Button
  'PBarStep.Height +
  cmdCancel.top = lblStatus.Height + PBar.Height + picAVI.Height + (L_MARGIN_LENGTH * 3)
  
  #If AbacusReporter Then
    'Override settings
    Me.Icon = g_AbacusReporter.AbacusReporterForm.Icon
    If g_AbacusReporter.DisplayTimings Then
      'Set form size with timings frame
      '+ PBarStep.Height
      Me.Height = lblStatus.Height + fraTimings.Height + cmdCancel.Height + PBar.Height + picAVI.Height + (L_MARGIN_LENGTH * 7)
      
      'Locate Timing frame
      '+ PBarStep.Height
      fraTimings.Visible = True
      fraTimings.top = lblStatus.Height + PBar.Height + picAVI.Height + L_MARGIN_LENGTH * 3
      
      'Locate Cancel Button
      '+ PBarStep.Height
      cmdCancel.top = lblStatus.Height + fraTimings.Height + PBar.Height + picAVI.Height + (L_MARGIN_LENGTH * 4)
    
      'Clear Labels
      Call StartTimings
    End If
  #End If
  
  
DisplayControls_End:
  Call xReturn("DisplayControls")
  Exit Sub
    
DisplayControls_Err:
  Call Err.Raise(Err.Number, ErrorSource(Err, "DisplayControls"), Err.Description)
  Resume DisplayControls_End
  Resume
End Sub

Private Sub Form_Activate()
  On Error GoTo Err_Err
  
  Call rw.PrepareReportSub
  Call StopAVI
  Call Me.Hide
  
Err_End:
  Exit Sub
Err_Err:
  Call ErrorMessagePush(Err)
  Me.Excepted = True
  StopAVI
  Me.Hide
End Sub


Private Sub Form_Load()
  If Not m_Loaded Then
    m_Loaded = True
    Me.Caption = rw.Title & "Status"
    lblStatus.Caption = "Loading..."
    Call DisplayAVI
    Call DisplayControls
  End If
End Sub

Private Sub StopAVI()
    m_AVI.StopPlay
    DoEvents
End Sub

Private Sub DisplayAVI()
  Dim sExt As String
  
  On Error GoTo DisplayAVI_Err
  Call xSet("DisplayAVI")
    Set m_AVI = New AVI
    
    m_AVI.AutoPlay = True
    m_AVI.Centre = True
    m_AVI.Transparent = True
     
    #If AbacusReporter Then
      m_AVI.ResourceFileName = FullPath(App.Path) & App.EXEName & ".exe"
    #Else
      m_AVI.ResourceFileName = FullPath(App.Path) & App.EXEName & ".dll"
    #End If
    m_AVI.ResourceId = 101
    Set m_AVI.Owner = picAVI

DisplayAVI_End:
  Exit Sub
DisplayAVI_Err:
  Call Err.Raise(Err.Number, ErrorSource(Err, "DisplayAVI"), Err.Description)
  Resume DisplayAVI_End
  Resume
End Sub
#If AbacusReporter Then
  
  Public Sub StartTimings()
    'Make form visible and clear captions
    Call xSet("StartTimings")
    On Error GoTo StartTimings_Err
    
    If g_AbacusReporter.DisplayTimings Then
      Call SetFileTimings(0, 0, 0, 0, True)
      Call SetFileGroupTimings(0, True)
      Call SetReportLineTimings(0, True)
      Call SetAbacusTimings(0, 0, 0, True)
    End If
  
StartTimings_End:
    Call xReturn("StartTimings")
    Exit Sub
    
StartTimings_Err:
    Call ErrorMessage(ERR_ERROR, Err, ErrorSource(Err, "StartTimings"), "Start Report Wizard", "Error starting the report wizard.")
    Resume StartTimings_End
    Resume
  End Sub
  
 Public Sub SetFileTimings(ByVal tFileOpen As Double, ByVal tFileExtract As Double, ByVal tFileClose As Double, ByVal tFileSave As Double, Optional ByVal bReset As Boolean)
    Static tCumulativeFileOpen As Double 'RK TODO: Make these larger!
    Static tCumulativeFileExtract As Double
    Static tCumulativeFileClose As Double
    Static tCumulativeFileSave As Double
    Static lFileCount As Long
    
    Call xSet("SetFileTimings")
    On Error GoTo SetFileTimings_Err

    If g_AbacusReporter.DisplayTimings Then
      If bReset Then
        tCumulativeFileOpen = 0
        tCumulativeFileExtract = 0
        tCumulativeFileClose = 0
        tCumulativeFileSave = 0
        tFileOpen = 0
        tFileExtract = 0
        tFileClose = 0
        tFileSave = 0
        lblFileCount.Caption = "File count: " & 0
        lblExceptionCount.Caption = "Exception count: " & 0
      Else
        lFileCount = lFileCount + 1
        lblFileCount.Caption = "Count: " & lFileCount
        lblExceptionCount.Caption = "Exception count: " & g_AbacusReporter.Session.ExceptionCount
      End If

      tCumulativeFileOpen = tCumulativeFileOpen + tFileOpen
      tCumulativeFileExtract = tCumulativeFileExtract + tFileExtract
      tCumulativeFileClose = tCumulativeFileClose + tFileClose
      tCumulativeFileSave = tCumulativeFileSave + tFileSave

      lblOpenLast.Caption = xStrPad(tFileOpen / 1000, " ", 20)
      lblExtractLast.Caption = xStrPad(tFileExtract / 1000, " ", 20)
      lblCloseLast.Caption = xStrPad(tFileClose / 1000, " ", 20)
      lblSaveLast.Caption = xStrPad(tFileSave / 1000, " ", 20)

      lblOpenAll.Caption = tCumulativeFileOpen / 1000
      lblExtractAll.Caption = tCumulativeFileExtract / 1000
      lblCloseAll.Caption = tCumulativeFileClose / 1000
      lblSaveAll.Caption = tCumulativeFileSave / 1000
          
      lblTotalLast = (tFileOpen + tFileExtract + tFileClose + tFileSave) / 1000
      lblTotalAll = (tCumulativeFileOpen + tCumulativeFileExtract + tCumulativeFileClose + tCumulativeFileSave) / 1000
    
    End If
  
SetFileTimings_End:
    Call xReturn("SetFileTimings")
    Exit Sub
    
SetFileTimings_Err:
    Call ErrorMessage(ERR_ERROR, Err, ErrorSource(Err, "SetFileTimings"), "Start Report Wizard", "Error starting the report wizard.")
    Resume SetFileTimings_End
    Resume
  End Sub
  
  Public Sub SetAbacusTimings(ByVal tEvaluate As Double, ByVal tExecute As Double, ByVal tSaveCursor As Double, Optional ByVal bReset As Boolean)
    Static tCumulativeEvaluate As Double
    Static tCumulativeExecute As Double
    Static tCumulativeSaveCursor As Double

    Call xSet("SetAbacusTimings")
    On Error GoTo SetAbacusTimings_Err

    If g_AbacusReporter.DisplayTimings Then
      If bReset Then
        tCumulativeEvaluate = 0
        tCumulativeExecute = 0
        tCumulativeSaveCursor = 0
        tEvaluate = 0
        tExecute = 0
        tSaveCursor = 0
      Else
        tEvaluate = Max(tEvaluate, 0)
        tExecute = Max(tExecute, 0)
        tSaveCursor = Max(tSaveCursor, 0)
      End If

      tCumulativeEvaluate = tCumulativeEvaluate + tEvaluate
      tCumulativeExecute = tCumulativeExecute + tExecute
      tCumulativeSaveCursor = tCumulativeSaveCursor + tSaveCursor

      lblEvaluateLast.Caption = xStrPad(tEvaluate / 1000, " ", 20)
      lblExecuteLast.Caption = xStrPad(tExecute / 1000, " ", 20)
      lblSaveCursorLast.Caption = xStrPad(tSaveCursor / 1000, " ", 20)

      lblEvaluateAll.Caption = tCumulativeEvaluate / 1000
      lblExecuteAll.Caption = tCumulativeExecute / 1000
      lblSaveCursorAll.Caption = tCumulativeSaveCursor / 1000
    End If
  
SetAbacusTimings_End:
    Call xReturn("SetAbacusTimings")
    Exit Sub
    
SetAbacusTimings_Err:
    Call ErrorMessage(ERR_ERROR, Err, ErrorSource(Err, "SetAbacusTimings"), "Start Report Wizard", "Error starting the report wizard.")
    Resume SetAbacusTimings_End
    Resume
  End Sub
  
  Friend Sub SetFileGroupTimings(ByVal tIndividual As Double, Optional ByVal bReset As Boolean)
    Static tCumulative As Double
    
    Call xSet("SetFileGroupTimings")
    On Error GoTo SetFileGroupTimings_Err

    If g_AbacusReporter.DisplayTimings Then
      If bReset Then
        tCumulative = 0
        tIndividual = 0
      End If

      tCumulative = tCumulative + tIndividual

      lblFileGroupLast.Caption = xStrPad("Last: " & tIndividual / 1000, " ", 20)
      lblFileGroupAll.Caption = "All: " & tCumulative / 1000
    End If
  
SetFileGroupTimings_End:
    Call xReturn("SetFileGroupTimings")
    Exit Sub
    
SetFileGroupTimings_Err:
    Call ErrorMessage(ERR_ERROR, Err, ErrorSource(Err, "SetFileGroupTimings"), "Start Report Wizard", "Error starting the report wizard.")
    Resume SetFileGroupTimings_End
    Resume
  End Sub

  Friend Sub SetReportLineTimings(ByVal tIndividual As Double, Optional ByVal bReset As Boolean)
    Static tCumulative As Double
    
    Call xSet("SetReportLineTimings")
    On Error GoTo SetReportLineTimings_Err

    If g_AbacusReporter.DisplayTimings Then
      If bReset Then
        tCumulative = 0
        tIndividual = 0
      End If

      tCumulative = tCumulative + tIndividual

      lblLineLast.Caption = xStrPad("Last: " & tIndividual / 1000, " ", 20)
      lblLineAll.Caption = "All: " & tCumulative / 1000
    End If
  
SetReportLineTimings_End:
    Call xReturn("SetReportLineTimings")
    Exit Sub
    
SetReportLineTimings_Err:
    Call ErrorMessage(ERR_ERROR, Err, ErrorSource(Err, "SetReportLineTimings"), "Start Report Wizard", "Error starting the report wizard.")
    Resume SetReportLineTimings_End
    Resume
  End Sub

#End If

