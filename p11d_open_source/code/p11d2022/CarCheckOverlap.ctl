VERSION 5.00
Object = "{8D988532-0F0C-460C-B00E-7B5637E97680}#1.0#0"; "ATC2VTEXT.OCX"
Begin VB.UserControl CarCheckOverlap 
   BackColor       =   &H8000000E&
   ClientHeight    =   4320
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7335
   KeyPreview      =   -1  'True
   ScaleHeight     =   4320
   ScaleWidth      =   7335
   Begin VB.CommandButton cmdAutoLink 
      Caption         =   "Auto link"
      Height          =   390
      Left            =   5850
      TabIndex        =   6
      Top             =   75
      Width           =   990
   End
   Begin VB.VScrollBar scroll 
      Height          =   4335
      Left            =   6960
      TabIndex        =   5
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton btnDelete 
      Appearance      =   0  'Flat
      Height          =   320
      Index           =   0
      Left            =   6600
      Picture         =   "CarCheckOverlap.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   600
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   330
   End
   Begin atc2valtext.ValText txtEnd 
      Height          =   255
      Index           =   0
      Left            =   5640
      TabIndex        =   0
      Top             =   600
      Visible         =   0   'False
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   450
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   ""
      TypeOfData      =   2
      AutoSelect      =   0
   End
   Begin atc2valtext.ValText txtStart 
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   1
      Top             =   600
      Visible         =   0   'False
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   450
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   ""
      TypeOfData      =   2
      AutoSelect      =   0
   End
   Begin VB.Line Line3 
      X1              =   5400
      X2              =   5400
      Y1              =   480
      Y2              =   360
   End
   Begin VB.Line Line2 
      X1              =   1200
      X2              =   1200
      Y1              =   495
      Y2              =   375
   End
   Begin VB.Line lnYear 
      X1              =   1200
      X2              =   5400
      Y1              =   450
      Y2              =   450
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000009&
      Caption         =   "Apr XX"
      Height          =   255
      Left            =   5160
      TabIndex        =   3
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000009&
      Caption         =   "Apr XX"
      Height          =   255
      Left            =   840
      TabIndex        =   2
      Top             =   120
      Width           =   615
   End
End
Attribute VB_Name = "CarCheckOverlap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Const MAX_TEXT_BOXES_SHOW As Long = 6
Private m_sPNum As String
Private m_sSQL As String

Private m_BarInfoMove As BARINFO
Private m_BarInfoMovedLastCarIndex As Long

Dim m_bInMove As Boolean
Dim m_bHitLeft As Boolean
Dim m_bHitRight As Boolean
Dim m_bMoveDirectionFound As Boolean
Private m_bDirty As Boolean
Private m_hRgn As Long
Dim m_ptHit As POINT

Private iSnapToFit As Long

Private Const SNG_YEAR_WIDTH  As Single = 5100


Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type size
    cx As Long
    cy As Long
End Type

Private Type POINT
    X As Long
    Y As Long
End Type

Private Type BARINFO
  BarBackRect As RECT
  BarRect As RECT
  BarLeftHit As RECT
  BarRightHit As RECT
  Registration As String
  InError As Boolean
  carIndex As Long
  DateStart As Date
  DateEnd As Date
  Replaced As Boolean
  Replacement As Boolean
  ReplacementRegistration As String
  Highlighted As Boolean
  LinkRects() As RECT
End Type

' ==================================================================
' Clipping functions:
' ==================================================================
Private Declare Function SelectClipRgn Lib "gdi32" ( _
    ByVal hdc As Long, ByVal hRgn As Long) As Long
Private Declare Function CreateRectRgn Lib "gdi32" ( _
    ByVal x1 As Long, ByVal y1 As Long, _
    ByVal x2 As Long, ByVal y2 As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" ( _
    ByVal hObject As Long) As Long
    
Private m_DragDropMode As Boolean
Public Property Get DragDropMode() As Boolean
  DragDropMode = m_DragDropMode
End Property
Public Property Let DragDropMode(ByVal NewValue As Boolean)
  m_DragDropMode = NewValue
  cmdAutoLink.Visible = m_DragDropMode
End Property
Private Property Get ResizeBufferTwips()
  ResizeBufferTwips = 3 * Screen.TwipsPerPixelX
End Property

Public Property Get Dirty() As Boolean
  Dirty = m_bDirty
End Property

Public Property Let Dirty(value As Boolean)
  m_bDirty = value
End Property



Private Function DateWithinTaxYear(dtNew As String) As Boolean
  Dim dt As Variant
  Dim bError As Boolean
  
  bError = False
  dt = ScreenToDateVal(dtNew, STDV_UNDATED)
  bError = (dt = UNDATED Or dt = UNDATED)
  If Not (bError) Then
    bError = Not DateInRange(dt, p11d32.Rates.value(TaxYearStart), p11d32.Rates.value(TaxYearEnd))
  End If

  DateWithinTaxYear = bError
End Function

Public Sub SaveOverlappingCCC(db As Database)
'Save company cars avail from and avail to dates

Dim i As Integer
Dim j As Integer
Dim ee As Employee
Dim ben As IBenefitClass, benReplaced As IBenefitClass
On Error GoTo SaveOverlappingCCC_ERR

  For i = 0 To txtStart.UBound
    If txtStart(i).Visible = True And Len(txtStart(i).Tag) > 0 Then
      If Not DateWithinTaxYear(txtStart(i).Text) And Not DateWithinTaxYear(txtEnd(i).Text) Then
        'Find employee
        Set ee = p11d32.CurrentEmployer.FindEmployee(GetPropertyFromString(txtStart(i).Tag, S_FIELD_PERSONEL_NUMBER))
        If Not ee Is Nothing Then
          'Load company car benefits
          ee.LoadBenefits (TBL_COMPANY_CARS)
          For j = 1 To ee.benefits.Count
            Set ben = ee.benefits(j)
            If Not ben Is Nothing Then
              'get correct car benefit
               If StrComp(ben.value(car_Registration_db), GetPropertyFromString(txtStart(i).Tag, S_FIELD_CAR_REGISTRATION)) = 0 Then
                 'write values in memeory
                 ben.value(Car_AvailableFrom_db) = ScreenToDateVal((txtStart(i).Text), STDV_UNDATED)
                 ben.value(Car_AvailableTo_db) = ScreenToDateVal((txtEnd(i).Text), STDV_UNDATED)
                 ben.value(car_Replaced_db) = CBoolean(GetPropertyFromString(txtStart(i).Tag, "REPLACED"))
                 ben.value(car_Replacement_db) = CBoolean(GetPropertyFromString(txtStart(i).Tag, "REPLACEMENT"))
                 ben.value(car_RegReplaced_db) = GetPropertyFromString(txtStart(i).Tag, "REPLACED_REGISTRATION")
                   
                 If m_DragDropMode Then
                   Set benReplaced = FindReplacedCarBenefit(ee, ben.value(car_RegReplaced_db))
                   If Not benReplaced Is Nothing Then
                     ben.value(car_CarReplacedMake_db) = benReplaced.value(car_Make_db)
                     ben.value(car_CarReplacedModel_db) = benReplaced.value(car_Model_db)
                     ben.value(car_CarReplacedEngineSize_db) = benReplaced.value(car_enginesize_db)
                   Else
                     ben.value(car_CarReplacedMake_db) = ""
                     ben.value(car_CarReplacedModel_db) = ""
                     ben.value(car_CarReplacedEngineSize_db) = 0
                   End If
                 End If
                 'write changes to db
                 Call ben.WriteDB
               End If
            End If
          Next j
          Call ee.KillBenefits
        Else
          Call Err.Raise(Err.Number, ErrorSource(Err, "SaveOverlappingCCC"), "Could not find employee ")
        End If
      End If
    End If
  Next i
  m_bDirty = False
SaveOverlappingCCC_END:
  Exit Sub
SaveOverlappingCCC_ERR:
  Call Err.Raise(Err.Number, ErrorSource(Err, "SaveOverlappingCCC"), Err.Description)
  Resume SaveOverlappingCCC_END
  Resume
End Sub
Private Function FindReplacedCarBenefit(ByVal ee As Employee, sRegReplaced As String) As IBenefitClass
  Dim i As Long
  Dim ben As IBenefitClass
  
 On Error GoTo err_Err
 
  For i = 1 To ee.benefits.Count
    Set ben = ee.benefits(i)
    If Not ben Is Nothing Then
        If ben.BenefitClass = BC_COMPANY_CARS_F Then
          If ben.Name = sRegReplaced Then
            Set FindReplacedCarBenefit = ben
            Exit For
          End If
        End If
    End If
  Next
  
err_End:
  Exit Function
err_Err:
  Call Err.Raise(Err.Number, ErrorSource(Err, "FindReplacedCarBenefit"), "Failed to find replacement car")
End Function
Private Sub PositionControls()
  Dim j As Long, i As Long
  
  If (txtStart.UBound < scroll.Max) Then Exit Sub
  For i = 1 To scroll.Max
    If (i < scroll.value) Then
      txtStart(i).Visible = False
      txtEnd(i).Visible = False
      btnDelete(i).Visible = False
    Else
      txtStart(i).Top = txtStart(j).Top + 400
      txtStart(i).Visible = True
      btnDelete(i).Top = btnDelete(j).Top + 400
      btnDelete(i).Visible = True
      txtEnd(i).Top = txtEnd(j).Top + 400
      txtEnd(i).Visible = True
      j = i
    End If
  Next
  Call UserControl.Refresh
End Sub

Public Sub DrawOverlaps(sSQL As String)
Dim i As Integer
Dim dtStart As Date
Dim dtEnd As Date
Dim lOffset As Long
Dim rs As Recordset
Dim irsCount As Long

On Error GoTo DrawOverlaps_ERR
  Call LockWindowUpdate(UserControl.hwnd)
  Label1.Caption = Replace(Label1.Caption, "XX", Right(DateStringEx(p11d32.Rates.value(TaxYearStart), p11d32.Rates.value(TaxYearStart)), 2))
  Label2.Caption = Replace(Label2.Caption, "XX", Right(DateStringEx(p11d32.Rates.value(TaxYearEnd), p11d32.Rates.value(TaxYearEnd)), 2))
  
  Set rs = p11d32.CurrentEmployer.db.OpenRecordset(sSQL)
  m_sSQL = sSQL
  For i = 1 To txtStart.UBound
    txtStart(i).Visible = False
    txtEnd(i).Visible = False
    btnDelete(i).Visible = False
  Next i
  i = 1
  irsCount = Records(rs)
  Do While Not rs.EOF
    If txtStart.UBound < i Then
      Load txtStart(i)
      Load txtEnd(i)
      Load btnDelete(i)
    End If
      
    txtStart(i).Tag = SetPropertiesFromString("", S_FIELD_CAR_REGISTRATION, rs(S_FIELD_CAR_REGISTRATION), S_FIELD_PERSONEL_NUMBER, rs("P_Num"), "DISPLAY", rs("Displayname"), "REPLACED", rs("Replaced"), "REPLACEMENT", rs.Fields("REPLACEMENT"), "REPLACED_REGISTRATION", rs("RegReplaced"))
    
    txtEnd(i).Text = DateValReadToScreenOnlyValidDates(rs("AvailTo"))
    txtStart(i).Text = DateValReadToScreenOnlyValidDates(rs("AvailFrom"))

    Call SetDefaultVTDate(txtStart(i))
    Call SetDefaultVTDate(txtEnd(i))
    
    i = i + 1
    rs.MoveNext
  Loop
  If (irsCount > 0) Then
    scroll.Min = 1
    scroll.Max = irsCount
    scroll.value = 1
    scroll.SmallChange = 1
    scroll.LargeChange = MAX_TEXT_BOXES_SHOW
  End If
  scroll.Visible = i > MAX_TEXT_BOXES_SHOW
  
  Call PositionControls
  m_bDirty = False
DrawOverlaps_END:
  Call LockWindowUpdate(0)
  Call UserControl.Refresh
  
  Exit Sub
DrawOverlaps_ERR:
   Call ErrorMessage(ERR_ERROR, Err, "DrawOverlaps", "Dislpay Overlapping Cars", "Error in user control - CarCheckOverlap")
   Resume DrawOverlaps_END
   Resume
End Sub
Private Sub btnDelete_Click(Index As Integer)
  Dim biFrom As BARINFO, biTo As BARINFO
  On Error GoTo err_Err
  
  If m_DragDropMode Then
    biFrom = BarInformation(Index)
    biTo = FindReplacement(Index)
    If (biTo.carIndex <> -1) Then
      Call SetRepelacement(biFrom.carIndex, biTo.carIndex, "")
    End If
    Call UserControl.Refresh

  Else
    Call DeleteCar(Index)
  End If
err_End:
  Exit Sub
err_Err:
  Call ErrorMessage(ERR_ERROR, Err, "Delete", "Failed to delete", Err.Description)
  Resume err_End
End Sub
Private Sub DeleteCar(i As Integer)
  
  Dim CC As CompanyCar
  Dim ee As Employee
  Dim ben As IBenefitClass
  Dim k As Integer
  Dim j As Integer
  
  On Error GoTo DeleteCar_ERR
  
  Select Case MultiDialog("Warning", "Are sure that you want to remove this?", "Delete", "Cancel")
    Case 2
      Exit Sub
  End Select

  Set ee = p11d32.CurrentEmployer.FindEmployee(GetPropertyFromString(txtStart(i).Tag, S_FIELD_PERSONEL_NUMBER))
  If Not ee Is Nothing Then
    'Load company car benefits
    ee.LoadBenefits (TBL_COMPANY_CARS)
    For j = 1 To ee.benefits.Count
     Set ben = ee.benefits(j)
     'get correct car benefit
     If Not ben Is Nothing Then
        If ben.value(car_Registration_db) = GetPropertyFromString(txtStart(i).Tag, S_FIELD_CAR_REGISTRATION) Then
         Set CC = ben
          Call ee.RemoveBenefitWithLinks(F_CompanyCar, ben, j, CC.Fuel, False)
          Exit For
        End If
      End If
    Next j
    Call ee.KillBenefits
  Else
    Call Err.Raise(Err.Number, ErrorSource(Err, "DeleteCar"), "Could not find employee ")
  End If

  'm_rs.Requery
  Call DrawOverlaps(m_sSQL)
DeleteCar_END:
  Exit Sub
DeleteCar_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "DeleteCar", "DeleteCar", "Error in user control - CarCheckOverlap")
  Resume DeleteCar_END
  Resume
End Sub

Private Sub cmdAutoLink_Click()
  Dim i As Long
  Dim j As Long
  
  On Error GoTo err_Err
  
  If m_DragDropMode Then
    For i = 1 To txtStart.UBound
      If Not txtStart(i).Visible Then Exit For
      For j = 1 To txtStart.UBound
        If Not txtStart(i).Visible Then Exit For
        If i <> j Then
          If ReplaceCar(i, j) Then Exit For
        End If
      Next
    Next
  End If
  Call UserControl.Refresh
  
err_End:
  Exit Sub
err_Err:
  Call ErrorMessage(ERR_ERROR, Err, ErrorSource(Err, "AutoLink"), "Auto link", "Failed to auto link the cars")
  Resume err_End
End Sub

Private Sub scroll_Change()
  Call PositionControls
End Sub

Private Sub txtEnd_Change(Index As Integer)
  m_bDirty = True
  Call UpdateReplacedReplacement(Index, False)
  UserControl.Refresh
End Sub

Private Sub txtStart_Change(Index As Integer)
  m_bDirty = True
  Call UpdateReplacedReplacement(Index, True)
  UserControl.Refresh
End Sub
Private Sub UpdateReplacedReplacement(ByVal Index As Long, bStart As Boolean)
  Dim biTo As BARINFO, biFrom As BARINFO
  Dim i As Long
    
    
  
    biTo = BarInformation(Index)
    If (biTo.Replacement And bStart) Or (biTo.Replaced And Not bStart) Then
        For i = 1 To txtStart.UBound
          If txtStart(i).Visible Then
            If (i <> Index) Then
              biFrom = BarInformation(i)
              If (biTo.Registration = biFrom.ReplacementRegistration And bStart) Or (biTo.ReplacementRegistration = biFrom.Registration And Not bStart) Then
                txtStart(Index).Tag = SetPropertyFromString(txtStart(i).Tag, "REPLACED_REGISTRATION", "")
                txtStart(Index).Tag = SetPropertyFromString(txtStart(i).Tag, "REPLACEMENT", False)
                txtStart(i).Tag = SetPropertyFromString(txtStart(i).Tag, "REPLACED", False)
                Exit For
              End If
            End If
          End If
        Next
    End If
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim pt As POINT
  Dim bi As BARINFO
  
  Call UserControl.SetFocus
  
  If (m_bInMove) Then Exit Sub
  
  pt.X = X
  pt.Y = Y
  bi = drawboxes(False, pt)
  
  m_BarInfoMovedLastCarIndex = -1
  
  If m_DragDropMode Then
    If (bi.carIndex <> -1) Then
      m_bInMove = True
      Call SetCursor(vbUpArrow)
      m_BarInfoMovedLastCarIndex = bi.carIndex
      m_BarInfoMove = bi
    End If
  Else
    m_bHitLeft = PointInRect(pt, bi.BarLeftHit)
    m_bHitRight = PointInRect(pt, bi.BarRightHit)
    If (m_bHitLeft Or m_bHitRight) Then
      If (Button = vbLeftButton) Then
        'start moving
        m_ptHit = pt
        m_bInMove = True
        m_BarInfoMove = bi
      End If
    End If
      
  End If
  
  
  
  
End Sub

'CAD TODO
Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim bi As BARINFO
  Dim pt As POINT
  Dim sz As size
  Dim dx As Long
  Dim iDays As Long
  Dim dNewDate As Date
  On Error GoTo err_Err
  
  pt.X = X
  pt.Y = Y
  
  If m_DragDropMode Then
    bi = drawboxes(False, pt)
    If Not m_bInMove Then
      If (bi.carIndex <> m_BarInfoMovedLastCarIndex) Then
        m_BarInfoMovedLastCarIndex = bi.carIndex
        Call drawboxes(True, pt, True)
      End If
    Else
      'in a move
      
      If (bi.carIndex <> m_BarInfoMovedLastCarIndex) Then
        m_BarInfoMovedLastCarIndex = bi.carIndex
        
        Call drawboxes(True, pt, True)
        
      End If
    End If
  Else
    If (m_bInMove) Then
      If Not m_bMoveDirectionFound Then
        m_bMoveDirectionFound = True
        'decide which way to move
        sz = RectSize(m_BarInfoMove.BarRect)
        If (sz.cx < (4 * 20)) Then
          m_bHitLeft = (pt.X < m_ptHit.X)
          m_bHitRight = (pt.X > m_ptHit.X)
        End If
      End If
      sz = RectSize(m_BarInfoMove.BarBackRect)
      dx = pt.X - m_ptHit.X
      iDays = (dx / sz.cx) * p11d32.Rates.value(DaysInYear)
      
      If (m_bHitLeft) Then
        dNewDate = DateAdd("d", iDays, m_BarInfoMove.DateStart)
        If DateInRange(dNewDate, p11d32.Rates.value(TaxYearStart), m_BarInfoMove.DateEnd) Then
          txtStart(m_BarInfoMove.carIndex).Text = DateValReadToScreen(dNewDate)
          GoTo err_End
        End If
      Else
        dNewDate = DateAdd("d", iDays, m_BarInfoMove.DateEnd)
        If DateInRange(dNewDate, m_BarInfoMove.DateStart, p11d32.Rates.value(TaxYearEnd)) Then
          txtEnd(m_BarInfoMove.carIndex).Text = DateValReadToScreen(dNewDate)
          GoTo err_End
        End If
        
      End If
    Else
      Call SetEWCursor(X, Y)
    End If
  End If

err_End:
  Exit Sub
err_Err:
  Resume err_End
  Resume
End Sub

Private Function CenterTextInTwipsRect(sText As String, rTwips As RECT) As POINT
  Dim pt As POINT
  Dim szRect As size, szText As size
  
  szText.cy = UserControl.TextHeight(sText)
  szText.cx = UserControl.TextWidth(sText)
  
  szRect = RectSize(rTwips)
  pt.X = ((szRect.cx - szText.cx) / 2) + rTwips.Left
  pt.Y = ((szRect.cy - szText.cy) / 2) + rTwips.Top
  CenterTextInTwipsRect = pt
End Function

Private Sub SelectClipRegionFromTwipsRect(rTwips As RECT)
  Dim rPixels As RECT
  
  rPixels = TwipsRectToPixelsRect(rTwips)
  m_hRgn = CreateRectRgn(rPixels.Left, rPixels.Top, rPixels.Right, rPixels.Bottom)
  Call SelectClipRgn(UserControl.hdc, m_hRgn)
  
End Sub
Private Sub DeleteClipRegion()
  
  Call SelectClipRgn(UserControl.hdc, 0)
  If (m_hRgn <> 0) Then DeleteObject (m_hRgn)
End Sub
Private Sub PointToCurrenyPosition(pt As POINT)
  UserControl.CurrentX = pt.X
  UserControl.CurrentY = pt.Y
  
End Sub
Private Sub DrawBar(ByRef bi As BARINFO, ByVal barColor As ColorConstants, Optional ByVal bDrawHighLightBorder As Boolean = False)
  Dim hRgn As Long
  Dim hOldRgn As Long
  Dim lcolor As Long
  Dim sngTextX As Single, sngTextY As Single
  Dim sText As String
  Dim ptText As POINT
  Dim rContainerInTwips As RECT, rBoxInTwips As RECT, rTwips As RECT
    
  sText = bi.Registration
  rContainerInTwips = bi.BarBackRect
  rBoxInTwips = bi.BarRect
  rTwips = rContainerInTwips
  'we will have 3 rects to do
  
  'first rect
  rTwips.Right = rBoxInTwips.Left
  
  UserControl.Line (rContainerInTwips.Left, rContainerInTwips.Top)-(rContainerInTwips.Right, rContainerInTwips.Bottom), UserControl.BackColor, BF
  
  UserControl.Line (rBoxInTwips.Left, rBoxInTwips.Top)-(rBoxInTwips.Right, rBoxInTwips.Bottom), barColor, BF
  
  If (bDrawHighLightBorder) Then
    UserControl.Line (rContainerInTwips.Left, rContainerInTwips.Top)-(rContainerInTwips.Right, rContainerInTwips.Bottom), vbRed, B
  End If
  
  ptText = CenterTextInTwipsRect(sText, rContainerInTwips)
  Call SelectClipRegionFromTwipsRect(rTwips)
  Call PointToCurrenyPosition(ptText)
  UserControl.Print sText
  Call DeleteClipRegion
  
  'second rect
  Call SelectClipRegionFromTwipsRect(rBoxInTwips)
  UserControl.CurrentX = sngTextX
  UserControl.CurrentY = sngTextY
  lcolor = UserControl.ForeColor
  UserControl.ForeColor = vbWhite
  Call PointToCurrenyPosition(ptText)
  UserControl.Print sText
  Call DeleteClipRegion
  UserControl.ForeColor = lcolor
  'third rect
  
  
  rTwips.Left = rBoxInTwips.Right
  rTwips.Right = rContainerInTwips.Right
  
  Call SelectClipRegionFromTwipsRect(rTwips)
  Call PointToCurrenyPosition(ptText)
  UserControl.Print sText
  Call DeleteClipRegion
  
  
End Sub

Private Function TwipsRectToPixelsRect(rTwips As RECT) As RECT
  Dim rPixels As RECT
  rPixels.Left = rTwips.Left / Screen.TwipsPerPixelX
  rPixels.Right = rTwips.Right / Screen.TwipsPerPixelX
  
  rPixels.Top = rTwips.Top / Screen.TwipsPerPixelY
  rPixels.Bottom = rTwips.Bottom / Screen.TwipsPerPixelY
  
  TwipsRectToPixelsRect = rPixels
End Function
Private Function RectSize(r As RECT) As size
  Dim sz As size
  sz.cx = (r.Right - r.Left) + 1
  sz.cy = (r.Bottom - r.Top) + 1
    
  RectSize = sz
End Function



Private Sub SetEWCursor(X As Single, Y As Single)
  Dim bi As BARINFO
  Dim pt As POINT
  
  pt.X = X
  pt.Y = Y
  bi = drawboxes(False, pt)
  If Not ((bi.carIndex = -1) Or (bi.InError)) Then
    If (PointInRect(pt, bi.BarLeftHit) Or PointInRect(pt, bi.BarRightHit)) Then
      Call SetCursor(vbSizeWE)
      Exit Sub
    End If
  End If

  Call ClearCursor

End Sub
Private Sub SetRepelacement(carIndexFrom As Long, carIndexTo As Long, ByVal sReplacementRegistration As String)
  Dim bReplaced As Boolean
  
  sReplacementRegistration = Trim$(sReplacementRegistration)
  bReplaced = Len(sReplacementRegistration) > 0
  
  txtStart(carIndexFrom).Tag = SetPropertyFromString(txtStart(carIndexFrom).Tag, "REPLACED", bReplaced)
  txtStart(carIndexTo).Tag = SetPropertyFromString(txtStart(carIndexTo).Tag, "REPLACED_REGISTRATION", sReplacementRegistration)
  txtStart(carIndexTo).Tag = SetPropertyFromString(txtStart(carIndexTo).Tag, "REPLACEMENT", bReplaced)
  Me.Dirty = True
End Sub
Private Function IsActuallyReplaced(iCarIndex As Long) As Boolean
  Dim i As Long
  Dim j As Long
  Dim bi As BARINFO
  Dim biSrc As BARINFO
  
  
  biSrc = BarInformation(iCarIndex)
  If Not biSrc.Replaced Then Exit Function
  For i = 1 To txtStart.UBound
    If Not txtStart(i).Visible Then Exit For
    If i <> iCarIndex Then
      bi = BarInformation(i)
      If StrComp(bi.ReplacementRegistration, biSrc.Registration, vbTextCompare) = 0 Then
        IsActuallyReplaced = True
        Exit Function
      End If
    End If
  Next
End Function
Private Function IsActuallyReplacement(iCarIndex As Long) As Boolean
  Dim i As Long
  Dim j As Long
  Dim bi As BARINFO
  Dim biSrc As BARINFO
  
  IsActuallyReplacement = False
  biSrc = BarInformation(iCarIndex)
  If Not biSrc.Replacement Then Exit Function
  For i = 1 To txtStart.UBound
    If Not txtStart(i).Visible Then Exit For
    If i <> iCarIndex Then
      bi = BarInformation(i)
      If StrComp(biSrc.ReplacementRegistration, bi.Registration, vbTextCompare) = 0 Then
        If bi.Replaced Then
          IsActuallyReplacement = True
          Exit Function
        End If
      End If
    End If
  Next
End Function

Private Function ReplaceCar(ByVal iSrcCarIndex As Long, ByVal iDstCarIndex As Long) As Boolean
  Dim biSource As BARINFO
  Dim biDest As BARINFO
    
  
  biSource = BarInformation(iSrcCarIndex)
  biDest = BarInformation(iDstCarIndex)
  
  
  If IsActuallyReplaced(iSrcCarIndex) Then Exit Function
  If IsActuallyReplacement(iDstCarIndex) Then Exit Function
  
  If (biSource.DateEnd = DateAdd("d", -1, biDest.DateStart)) Or (biSource.DateEnd = biDest.DateStart) Then
    If biSource.ReplacementRegistration <> biDest.Registration Or (biSource.Replaced = False) Then
      Call SetRepelacement(biSource.carIndex, biDest.carIndex, biSource.Registration)
      ReplaceCar = True
    End If
  End If

End Function
Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim pt As POINT
  Dim bi As BARINFO
  Dim biSource As BARINFO, biDest As BARINFO
  
  On Error GoTo err_Err
  
  Dim sData As String
  pt.X = X
  pt.Y = Y
  
  If (m_DragDropMode) Then
    If m_bInMove Then
      Call SetCursor(vbDefault)
      bi = drawboxes(False, pt)
      If (bi.carIndex <> -1) Then
        If bi.carIndex <> m_BarInfoMove.carIndex Then
          Call ReplaceCar(m_BarInfoMove.carIndex, bi.carIndex)
          
        End If
      End If
    End If
      m_BarInfoMove.carIndex = -1
    UserControl.Refresh
  Else
    Call SetEWCursor(X, Y)
  End If
  
  m_bInMove = False
  m_bMoveDirectionFound = False
  
err_End:
  Exit Sub
err_Err:
  Call ErrorMessage(ERR_ERROR, Err, ErrorSource(Err, "MouseUp"), "Mouse Up", "Error linking cars")
  Resume err_End
  Resume
End Sub

Private Sub UserControl_Paint()
  Dim pt As POINT
    Call drawboxes(True, pt)
End Sub
Private Function BarBackRect(iBoxNumber As Long) As RECT
  Dim r As RECT
  
  r.Left = lnYear.x1
  r.Right = lnYear.x2
  r.Top = txtStart(iBoxNumber).Top
  r.Bottom = txtStart(iBoxNumber).Top + txtStart(iBoxNumber).height
  
  BarBackRect = r
End Function

Private Function BarInformation(ByVal iBoxNumber As Long, Optional bFindLinks As Boolean = True) As BARINFO
  Dim bi As BARINFO
  Dim sz As size
  Dim i As Long
  
  bi.BarBackRect = BarBackRect(iBoxNumber)
  
  bi.carIndex = iBoxNumber
  bi.DateStart = ScreenToDateVal(txtStart(iBoxNumber).Text, STDV_UNDATED)
  
  bi.DateEnd = ScreenToDateVal(txtEnd(iBoxNumber).Text, STDV_UNDATED)
  bi.InError = (bi.DateEnd = UNDATED Or bi.DateStart = UNDATED)
  If Not (bi.InError) Then
    bi.InError = Not DateInRange(bi.DateStart, p11d32.Rates.value(TaxYearStart), p11d32.Rates.value(TaxYearEnd))
    bi.InError = bi.InError Or (Not DateInRange(bi.DateEnd, p11d32.Rates.value(TaxYearStart), p11d32.Rates.value(TaxYearEnd)))
  End If
  If Not bi.InError Then
    bi.BarRect = BarRect(bi, iBoxNumber)
    
    bi.BarLeftHit = bi.BarRect
    bi.BarLeftHit.Left = bi.BarRect.Left - ResizeBufferTwips
    bi.BarLeftHit.Right = bi.BarRect.Left + ResizeBufferTwips
    
    bi.BarRightHit = bi.BarRect
    bi.BarRightHit.Left = bi.BarRect.Right - ResizeBufferTwips
    bi.BarRightHit.Right = bi.BarRect.Right + ResizeBufferTwips
  End If
  
  bi.Registration = GetPropertyFromString(txtStart(iBoxNumber).Tag, "DISPLAY")
  bi.Replaced = GetPropertyFromString(txtStart(iBoxNumber).Tag, "REPLACED")
  bi.Replacement = GetPropertyFromString(txtStart(iBoxNumber).Tag, "REPLACEMENT")
  bi.ReplacementRegistration = GetPropertyFromString(txtStart(iBoxNumber).Tag, "REPLACED_REGISTRATION")
  
  If Len(bi.ReplacementRegistration) > 0 Then
    'find the car I am linking to
    For i = 1 To txtStart.UBound
      If GetPropertyFromString(txtStart(iBoxNumber).Tag, "DISPLAY") = bi.ReplacementRegistration Then
           
      End If
    Next
    'BI.LinkPoints = BI.BarRect.Right
  Else
    'ReDim BI.LinkPoints(0 To 0)
  End If
  
  
  BarInformation = bi
  
End Function
Private Function BarRect(bi As BARINFO, iBoxNumber As Long) As RECT
  Dim r As RECT
  Dim sz As size
  
  sz = RectSize(bi.BarBackRect)
  r = bi.BarBackRect
  If Not (bi.InError) Then
    r.Left = GetOffset(bi.DateStart, bi.BarBackRect)
    r.Right = GetOffset(bi.DateEnd, bi.BarBackRect)
    If (r.Right - r.Left <= 0) Then r.Left = r.Right
    sz = RectSize(bi.BarBackRect)
    If (sz.cx * Screen.TwipsPerPixelX < 1) Then
      r.Right = r.Left + Screen.TwipsPerPixelX
      If (r.Right > bi.BarBackRect.Right) Then
        r.Right = bi.BarBackRect.Right
        r.Left = r.Right - Screen.TwipsPerPixelX
      ElseIf (r.Left < bi.BarBackRect.Left) Then
        r.Left = bi.BarBackRect.Left
        r.Right = r.Left + Screen.TwipsPerPixelX
      End If
    End If
  End If
  BarRect = r
End Function
Private Function PointInRect(pt As POINT, r As RECT) As Boolean
  PointInRect = ((pt.X >= r.Left And pt.X <= r.Right)) And ((pt.Y >= r.Top And pt.Y <= r.Bottom))
End Function

Private Function drawboxes(bDraw As Boolean, pt As POINT, Optional ByVal highLightHitBox As Boolean = False) As BARINFO
  Dim i As Long, j As Long
  Dim bi As BARINFO
  Dim biFrom As BARINFO, biTo As BARINFO
  Dim iHit As Long
  Dim biHit As BARINFO
  Dim bHit As Boolean
  Dim bError As Boolean
  Dim iColor As Long
  
  Dim y1 As Long, y2 As Long
  Dim x1 As Long, x2 As Long
  
  
  iHit = -1
  biHit.carIndex = -1
  
  For i = 1 To txtStart.UBound
    If txtStart(i).Visible Then
      bHit = False
      bi = BarInformation(i)
      If (PointInRect(pt, bi.BarBackRect)) Then
        iHit = i
        biHit = bi
      End If
      If bDraw Then
        If (bi.InError) Then
          Call DrawErrorBox(bi)
        Else
          If (m_DragDropMode And (m_BarInfoMove.carIndex = i)) Or ((iHit > 0) And highLightHitBox) Then
            iHit = 0
            Call DrawBar(bi, vbBlue, highLightHitBox)
          Else
            Call DrawBar(bi, vbBlue)
          End If
        End If
      End If
    End If
  Next i
  If bDraw Then
    'draw the replacement replaced lines
    For i = 1 To txtStart.UBound
      If Not txtStart(i).Visible Then Exit For
      biFrom = BarInformation(i)
      biTo = FindReplacement(i)
      If biTo.carIndex <> -1 Then
        'draw the connecting lines
        y1 = biFrom.BarRect.Top + ((biFrom.BarRect.Bottom - biFrom.BarRect.Top) / 2)
        x1 = biFrom.BarRect.Left - 100
        UserControl.Line (biFrom.BarRect.Left, y1)-(x1, y1), vbGreen
        y2 = biTo.BarRect.Top + ((biTo.BarRect.Bottom - biTo.BarRect.Top) / 2)
        UserControl.Line (x1, y1)-(x1, y2), vbGreen
        UserControl.Line (x1, y2)-(biTo.BarRect.Left, y2), vbGreen
      End If
    Next
  End If
  drawboxes = biHit

End Function
Private Function FindReplacement(ByVal carIndex As Long) As BARINFO
  Dim j As Long
  Dim biFrom As BARINFO
  Dim biTo As BARINFO
  
  
  biFrom = BarInformation(carIndex)
  For j = 1 To txtStart.UBound
    If carIndex <> j Then
      biTo = BarInformation(j)
      If biFrom.Replaced And biTo.Replacement And (biFrom.Registration = biTo.ReplacementRegistration) Then
        FindReplacement = biTo
        Exit Function
      End If
    End If
  Next
    
  biFrom.carIndex = -1
  FindReplacement = biFrom
End Function
Private Sub DrawErrorBox(bi As BARINFO)
  Dim rBoxInTwips As RECT
  Dim sText As String
  Dim ptText As POINT
  
  rBoxInTwips = bi.BarBackRect
  sText = bi.Registration
  
  UserControl.Line (rBoxInTwips.Left, rBoxInTwips.Top)-(rBoxInTwips.Right, rBoxInTwips.Bottom), vbRed, B
  UserControl.Line (rBoxInTwips.Left, rBoxInTwips.Top)-(rBoxInTwips.Right, rBoxInTwips.Bottom), vbRed
  UserControl.Line (rBoxInTwips.Left, rBoxInTwips.Bottom)-(rBoxInTwips.Right, rBoxInTwips.Top), vbRed
  
  ptText = CenterTextInTwipsRect(sText, rBoxInTwips)
  Call PointToCurrenyPosition(ptText)
  UserControl.Print sText

End Sub
Private Function GetOffset(dtRef As Date, rBar As RECT) As Long
  Dim dFactor As Double
  Dim days As Long
  days = DateDiff("d", p11d32.Rates.value(TaxYearStart), dtRef)
  dFactor = days / p11d32.Rates.value(DaysInYear)
  GetOffset = (dFactor * (rBar.Right - rBar.Left)) + rBar.Left
End Function

Private Sub UserControl_Resize()
  scroll.Left = UserControl.width - scroll.width
End Sub

Private Sub UserControl_Show()
  If UserControl.Ambient.UserMode = True Then
    Set btnDelete(0).Picture = MDIMain.ImgToolbar.ListImages(13).Picture
  End If
End Sub
