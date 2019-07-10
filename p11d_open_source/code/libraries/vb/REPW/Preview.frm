VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmPreview 
   BackColor       =   &H8000000C&
   Caption         =   "Print Preview"
   ClientHeight    =   7275
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   11820
   Icon            =   "Preview.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7275
   ScaleWidth      =   11820
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin MSComctlLib.Toolbar tbrPreview 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   11820
      _ExtentX        =   20849
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   2
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Toggle View"
            Object.Tag             =   "Toggle"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Print Report"
            Object.Tag             =   "PrintReport"
            ImageIndex      =   9
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin VB.PictureBox picPage 
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   7920
         ScaleHeight     =   495
         ScaleWidth      =   1935
         TabIndex        =   6
         Top             =   0
         Width           =   1935
         Begin VB.Label lblPage 
            Caption         =   "Label1"
            Height          =   255
            Left            =   240
            TabIndex        =   7
            Top             =   120
            Width           =   2175
         End
      End
   End
   Begin VB.CommandButton cmdTurn 
      Height          =   600
      Left            =   11520
      Picture         =   "Preview.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6390
      Width           =   600
   End
   Begin VB.HScrollBar hscrPage 
      Height          =   240
      Left            =   10440
      Max             =   10
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   6750
      Width           =   1215
   End
   Begin VB.VScrollBar vscrPaper 
      Height          =   6690
      Left            =   9840
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   480
      Width           =   255
   End
   Begin VB.HScrollBar hscrPaper 
      Height          =   240
      Left            =   45
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   6750
      Width           =   10440
   End
   Begin VB.PictureBox picPaper 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   5100
      Left            =   840
      MouseIcon       =   "Preview.frx":0614
      ScaleHeight     =   5100
      ScaleWidth      =   4890
      TabIndex        =   0
      Top             =   1200
      Width           =   4890
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   120
      Top             =   480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   13
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Preview.frx":091E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Preview.frx":0E60
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Preview.frx":13A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Preview.frx":18E4
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Preview.frx":1E26
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Preview.frx":2368
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Preview.frx":28AA
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Preview.frx":2DEC
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Preview.frx":332E
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Preview.frx":3870
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Preview.frx":3DB2
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Preview.frx":42F4
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Preview.frx":4836
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Shape shpMargin 
      BackStyle       =   1  'Opaque
      Height          =   5640
      Left            =   750
      Top             =   720
      Width           =   5190
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuExport 
         Caption         =   "&Export..."
      End
      Begin VB.Menu mnuPrint 
         Caption         =   "&Print"
      End
      Begin VB.Menu mnuClose 
         Caption         =   "&Close"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuZoom 
         Caption         =   "&Zoom"
         Begin VB.Menu mnuZoomLevel 
            Caption         =   "25%"
            Index           =   0
         End
         Begin VB.Menu mnuZoomLevel 
            Caption         =   "50%"
            Index           =   1
         End
         Begin VB.Menu mnuZoomLevel 
            Caption         =   "100%"
            Index           =   2
         End
         Begin VB.Menu mnuZoomLevel 
            Caption         =   "200%"
            Index           =   3
         End
      End
   End
End
Attribute VB_Name = "frmPreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private ZoomLevel As Integer
Private frmH As Single
Private frmW As Single
Private frmT As Single
Private PaperW As Single
Private PaperH As Single
Private MarginOffset As Single
Private ImageBorder As Single
Private bNoPageChange As Boolean
Private LastExportDir As String

Private Sub cmdTurn_Click()
  If hscrPage.value < hscrPage.Max Then
    hscrPage.value = hscrPage.value + 1
  Else
    Beep
  End If
End Sub

Private Sub Form_Activate()
  Me.tbrPreview.Buttons(2).Enabled = IsPrinterAvail(True)
  Me.mnuPrint.Enabled = IsPrinterAvail(False)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Call LockWindowUpdate(Me.picPaper.hWnd)
  Select Case KeyCode
    Case 40
      If (vscrPaper.value + vscrPaper.SmallChange) <= vscrPaper.Max And vscrPaper.Enabled = True Then
        vscrPaper.value = vscrPaper.value + vscrPaper.SmallChange
      ElseIf vscrPaper.Enabled = True Then
        vscrPaper.value = vscrPaper.Max
      End If
    Case 39
     If (hscrPaper.value + hscrPaper.SmallChange) <= hscrPaper.Max And hscrPaper.Enabled = True Then
       hscrPaper.value = hscrPaper.value + hscrPaper.SmallChange
     ElseIf hscrPaper.Enabled = True Then
         hscrPaper.value = hscrPaper.Max
     End If
    Case 38
      If (vscrPaper.value - vscrPaper.SmallChange) >= vscrPaper.Min And vscrPaper.Enabled = True Then
        vscrPaper.value = vscrPaper.value - vscrPaper.SmallChange
      ElseIf vscrPaper.Enabled = True Then
        vscrPaper.value = vscrPaper.Min
      End If
    Case 37
      If (hscrPaper.value - hscrPaper.SmallChange) >= hscrPaper.Min And hscrPaper.Enabled = True Then
        hscrPaper.value = hscrPaper.value - hscrPaper.SmallChange
      ElseIf hscrPaper.Enabled = True Then
        hscrPaper.value = hscrPaper.Min
      End If
    Case 33 'PAGE UP
      If hscrPage.value > hscrPage.Min Then
        hscrPage.value = hscrPage.value - 1
      End If
    Case 34 'PAGE DOWN
      If hscrPage.value < hscrPage.Max Then
        hscrPage.value = hscrPage.value + 1
      End If
  End Select
  picPaper.Refresh
  Call LockWindowUpdate(0)
End Sub

Private Sub Form_Load()
  LastExportDir = AppPath
End Sub

Private Sub Form_Resize()
  Dim w As Long, h As Long

  If (Me.ScaleHeight > 0) And (Me.ScaleWidth > 0) Then
    bNoPageChange = True
     Me.hscrPage.Min = 1
     Me.hscrPage.value = 1
     Me.hscrPage.Max = ReportControl.Pages_N
    bNoPageChange = False
    
    ' resize code
    Call LockWindowUpdate(Me.hWnd)
    frmT = tbrPreview.Height
    frmH = Max(Me.ScaleHeight - frmT, 0)
    frmW = Me.ScaleWidth
  
    vscrPaper.Left = frmW - vscrPaper.Width
    vscrPaper.Top = frmT
    vscrPaper.Height = Max(frmH - cmdTurn.Height, 0)
  
    hscrPaper.Left = 0
    hscrPaper.Top = frmH + frmT - hscrPaper.Height
    hscrPaper.Width = (frmW * 2) / 3!
  
    hscrPage.Top = hscrPaper.Top
    hscrPage.Left = hscrPaper.Width
    hscrPage.Width = frmW - Me.hscrPaper.Width - cmdTurn.Width
  
    cmdTurn.Left = hscrPaper.Width + hscrPage.Width
    cmdTurn.Top = tbrPreview.Height + vscrPaper.Height
  
    picPage.Left = frmW - picPage.Width
    
    Call ScalePaper
    Call LockWindowUpdate(0)
  End If
End Sub

Private Sub mnuClose_Click()
  Unload Me
End Sub

Private Sub mnuExport_Click()
  Dim sFilter As String, SaveFile As String
  Dim eType As REPORT_EXPORTS
  Dim i As Long, j As Long
  Dim FilterIndex As Long
  
  For i = [_REPORT_EXPORTS_FIRST] To L_LAST_EXPORT
    If ExportAvailableEx(i) Then sFilter = sFilter & ExportTypeStrEx(i) & " (*" & ExportTypeExtEx(i) & ")|*" & ExportTypeExtEx(i) & "|"
  Next i
  sFilter = Left$(sFilter, Len(sFilter) - 1)
  SaveFile = FileSaveAsDlgFilter(FilterIndex, "Export Report", sFilter, LastExportDir)
  If Len(SaveFile) > 0 Then
    j = 1
    For i = [_REPORT_EXPORTS_FIRST] To L_LAST_EXPORT
      If ExportAvailableEx(i) Then
        If j = FilterIndex Then
          eType = i
          Exit For
        End If
        j = j + 1
      End If
    Next i
    If eType <> 0 Then
      Call SplitPath(SaveFile, LastExportDir)
      Call ExportReportEx(SaveFile, eType, True, False)
    Else
      Call ECASE("No export to that type of file")
    End If
  End If
End Sub

Private Sub mnuPrint_Click()
  Call PrintDialog
End Sub

Private Sub PrintDialog()
  Dim i As Long, j As Long
  Dim Cancel As Boolean
    
  On Error Resume Next
  Load frmPrintDialog
  frmPrintDialog.txtCopies = CStr(ReportControl.PageCopies)
  frmPrintDialog.optPrint(ReportControl.PrintDlgOpt).value = True
  frmPrintDialog.updCopies.Min = 1
  frmPrintDialog.updCopies.Max = MAX_PAGE_COPIES
  If (ReportControl.PageFrom < 1) Or _
     (ReportControl.PageTo < 1) Or _
     (ReportControl.PageFrom > ReportControl.PageTo) Or _
     (ReportControl.PageTo > ReportControl.Pages_N) Then
    ReportControl.PageFrom = 1
    ReportControl.PageTo = ReportControl.Pages_N
  End If
  frmPrintDialog.txtFrom = CStr(ReportControl.PageFrom)
  frmPrintDialog.txtTo = CStr(ReportControl.PageTo)
  frmPrintDialog.updFrom.Min = 1
  frmPrintDialog.updFrom.Max = ReportControl.Pages_N
  frmPrintDialog.updTo.Min = 1
  frmPrintDialog.updTo.Max = ReportControl.Pages_N
  frmPrintDialog.Show vbModal
  Cancel = frmPrintDialog.Cancel
  If Not frmPrintDialog.Cancel Then
    ReportControl.PageCopies = CLng(frmPrintDialog.txtCopies)
    For i = PAGES_ALL To PAGES_CURRENT
      If frmPrintDialog.optPrint(i).value Then ReportControl.PrintDlgOpt = i
    Next i
    ReportControl.PageFrom = CLng(frmPrintDialog.txtFrom)
    ReportControl.PageTo = CLng(frmPrintDialog.txtTo)
    Call Me.Refresh
    Call SetCursor
    For i = 1 To ReportControl.PageCopies
      If ReportControl.PrintDlgOpt = PAGES_ALL Then
        Call PreviewPrintPageEx(1, ReportControl.Pages_N)
      ElseIf ReportControl.PrintDlgOpt = PAGES_CURRENT Then
        Call PreviewPrintPageEx(Me.hscrPage.value, Me.hscrPage.value)
      ElseIf ReportControl.PrintDlgOpt = PAGES_RANGE Then
        Call PreviewPrintPageEx(ReportControl.PageFrom, ReportControl.PageTo)
      End If
    Next i
    Call ClearCursor
  End If
  Unload frmPrintDialog
End Sub

Private Sub mnuZoomLevel_Click(Index As Integer)
  Select Case Index
    Case 0
      PaperZoom 25
    Case 1
      PaperZoom 50
    Case 2
      PaperZoom 100
    Case 3
      PaperZoom 200
    Case Else
      ECASE " ZOOM"
  End Select
End Sub

Private Sub tbrPreview_ButtonClick(ByVal Button As MSComctlLib.Button)
  Select Case Button.tag
    Case "PrintReport"
            Call PrintDialog
    Case "Toggle"
            Call PaperZoom
    Case Else
            Call ECASE("Unknown action for button")
  End Select
End Sub

Private Sub vscrPaper_Change()
  Call LockWindowUpdate(Me.hWnd)
  shpMargin.Top = frmT + ImageBorder - vscrPaper.value
  picPaper.Top = shpMargin.Top + MarginOffset
  'shpShadow.Top = shpMargin.Top + MarginOffset
  Call SetScaleParameters
  Call LockWindowUpdate(0)
End Sub


Private Sub hscrPage_Change()
  If Not bNoPageChange Then
    Me.cmdTurn.Enabled = Not (hscrPage.value = ReportControl.Pages_N)
    lblPage = "Page " & hscrPage.value & " of " & ReportControl.Pages_N
    Call PreviewPageEx(hscrPage.value, Me.picPaper)
  End If
End Sub

Private Sub hscrPaper_Change()
  Call LockWindowUpdate(Me.hWnd)
  Me.shpMargin.Left = ImageBorder - hscrPaper.value
  'Me.shpShadow.Left = Me.shpMargin.Left + MarginOffset
  Me.picPaper.Left = Me.shpMargin.Left + MarginOffset
  Call SetScaleParameters
  Call LockWindowUpdate(0)
End Sub

Private Function ScalePaper() As Boolean
  Dim tmp As Integer
  On Error GoTo ScalePaper_Err
  
  Call xSet("ScalePaper")
  Call SetCursor
  Call LockWindowUpdate(Me.hWnd)
  PaperW = (ReportControl.PageWidth * ReportControl.Zoom) / 100!
  PaperH = (ReportControl.PageHeight * ReportControl.Zoom) / 100!
  MarginOffset = Min(PaperW / 20, PaperH / 20)
  ImageBorder = MarginOffset / 4
  
  picPaper.Width = PaperW
  picPaper.Height = PaperH
  
  If (PaperH + 3 * MarginOffset + 2 * ImageBorder) > frmH Then
    tmp = vscrPaper.Max
    vscrPaper.Min = 0
    vscrPaper.Max = (PaperH + 3 * MarginOffset + 2 * ImageBorder + hscrPaper.Height) - frmH
    vscrPaper.SmallChange = MarginOffset * 2
    vscrPaper.LargeChange = MarginOffset * 4
    vscrPaper.value = 0
    vscrPaper.Enabled = True
    shpMargin.Top = frmT + ImageBorder
  Else
    vscrPaper.Enabled = False
    shpMargin.Top = frmT + (frmH - (PaperH + 3 * MarginOffset + 2 * ImageBorder)) / 2
  End If
  shpMargin.Height = PaperH + MarginOffset * 2
'  shpShadow.Height = shpMargin.Height
  picPaper.Top = shpMargin.Top + MarginOffset
  'shpShadow.Top = shpMargin.Top + MarginOffset


  If (PaperW + 3 * MarginOffset + 2 * ImageBorder) > frmW Then
    hscrPaper.Min = 0
    hscrPaper.Max = (PaperW + 3 * MarginOffset + 2 * ImageBorder + vscrPaper.Width) - frmW
    hscrPaper.SmallChange = MarginOffset * 2
    hscrPaper.LargeChange = MarginOffset * 4
    hscrPaper.value = 0
    shpMargin.Left = ImageBorder
    hscrPaper.Enabled = True
  Else
    shpMargin.Left = (frmW - (PaperW + 2 * MarginOffset + 2 * ImageBorder)) / 2
    hscrPaper.Enabled = False
  End If
  shpMargin.Width = PaperW + MarginOffset * 2
  'shpShadow.Width = shpMargin.Width
  picPaper.Left = shpMargin.Left + MarginOffset
  'shpShadow.Left = shpMargin.Left + MarginOffset
  
  Call SetScaleParameters
  lblPage = "Page " & hscrPage.value & " of " & ReportControl.Pages_N
  Call PreviewPageEx(hscrPage.value, Me.picPaper)

ScalePaper_End:
  Call LockWindowUpdate(0)
  Call ClearCursor
  Call xReturn("ScalePaper")
  Exit Function
  
ScalePaper_Err:
  Resume ScalePaper_End
End Function

Public Function PaperZoom(Optional ByVal ZoomPercent As Long = 0) As Boolean
  On Error GoTo PaperZoom_Err
  Call xSet("PaperZoom")
  If ZoomPercent <= 0 Then
    Select Case ZoomLevel
    Case 0
      ReportControl.Zoom = 40
      ZoomLevel = 1
    Case 1
      ReportControl.Zoom = 100
      ZoomLevel = 0
    End Select
  Else
    ReportControl.Zoom = ZoomPercent
  End If
  Call SetZoomLimit
  Call ScalePaper  'scale and Preview Page
  
PaperZoom_End:
  Call xReturn("PaperZoom")
  Exit Function

PaperZoom_Err:
  Resume PaperZoom_End
End Function

Private Sub SetScaleParameters()
  Me.picPaper.ScaleLeft = 0
  Me.picPaper.ScaleTop = 0
  Me.picPaper.ScaleHeight = ReportControl.PageHeight
  Me.picPaper.ScaleWidth = ReportControl.PageWidth
End Sub

