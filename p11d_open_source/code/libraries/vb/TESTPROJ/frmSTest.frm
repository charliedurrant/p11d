VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmSortTest 
   Caption         =   "Sort Testing"
   ClientHeight    =   4500
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7335
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4500
   ScaleWidth      =   7335
   Begin VB.Frame fraSortType 
      Caption         =   "Sort Type"
      Height          =   1125
      Left            =   4215
      TabIndex        =   5
      Top             =   105
      Width           =   2985
      Begin VB.OptionButton SortOptions 
         Caption         =   "Comb Sort"
         Height          =   285
         Index           =   1
         Left            =   120
         TabIndex        =   7
         Top             =   645
         Width           =   2310
      End
      Begin VB.OptionButton SortOptions 
         Caption         =   "Quick Sort"
         Height          =   285
         Index           =   0
         Left            =   120
         TabIndex        =   6
         Top             =   270
         Value           =   -1  'True
         Width           =   2310
      End
   End
   Begin VB.CommandButton cmdSort 
      Caption         =   "Start Sort"
      Height          =   390
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   1620
   End
   Begin MSComCtl2.UpDown UDSort 
      Height          =   375
      Left            =   3436
      TabIndex        =   2
      Top             =   180
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   661
      _Version        =   393216
      Value           =   1
      AutoBuddy       =   -1  'True
      BuddyControl    =   "txtSortElements"
      BuddyDispid     =   196612
      OrigLeft        =   3810
      OrigTop         =   180
      OrigRight       =   4050
      OrigBottom      =   600
      Increment       =   100
      Max             =   100000
      Min             =   1
      SyncBuddy       =   -1  'True
      BuddyProperty   =   65547
      Enabled         =   -1  'True
   End
   Begin VB.TextBox txtSortElements 
      Height          =   375
      Left            =   2640
      TabIndex        =   1
      Text            =   "1000"
      Top             =   180
      Width           =   795
   End
   Begin VB.Label lblResults 
      Caption         =   "Results:"
      Height          =   2970
      Left            =   120
      TabIndex        =   4
      Top             =   1425
      Width           =   7155
   End
   Begin VB.Label lblText 
      Caption         =   "Number of Sort elements"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   180
      Width           =   2295
   End
End
Attribute VB_Name = "frmSortTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Enum TEST_SORTTYPE
  QUICK_SORT
  COMB_SORT
End Enum

Private Sub cmdSort_Click()
  Dim v As Variant
  Dim st As TEST_SORTTYPE, nElements As Long

  Call SetCursor
  Randomize
  nElements = CLng(Me.txtSortElements)
  If SortOptions(0).Value Then
    st = QUICK_SORT
  Else
    st = COMB_SORT
  End If
  v = GetRandomArrayofLongs(1, nElements)
  Call DoSort(v, 1, nElements, st)
  Call ClearCursor
End Sub

Private Function GetRandomArrayofLongs(ByVal lb As Long, ByVal ub As Long) As Variant
  Dim ar() As Long, i As Long
  
  ReDim ar(lb To ub) As Long
  For i = lb To ub
    ar(i) = Int(((ub * 2) * Rnd) + 1)
  Next i
  GetRandomArrayofLongs = ar
End Function

Private Sub DoSort(v As Variant, ByVal lb As Long, ByVal ub As Long, ByVal st As TEST_SORTTYPE)
  Dim t0 As Long, t1 As Long
  Dim d As Double, nElements As Long
  Dim Sortfn As SortLongs
  
  Me.lblResults = "Results:" & vbCrLf
  Set Sortfn = New SortLongs
  If st = QUICK_SORT Then
    t0 = GetTicks
    Call QSortEx(v, lb, ub, Sortfn)
    t1 = GetTicks
    Me.lblResults = Me.lblResults & "Quicksort: " & Format$((t1 - t0) / 1000, "#,##0.00") & " seconds."
  End If
  If st = COMB_SORT Then
    t0 = GetTicks
    Call CombSortEx(v, lb, ub, Sortfn)
    t1 = GetTicks
    Me.lblResults = Me.lblResults & "Comb Sort: " & Format$((t1 - t0) / 1000, "#,##0.00") & " seconds."
  End If
  nElements = (ub - lb + 1)
  d = Fix(nElements * Log(nElements) / Log(2))
  Me.lblResults = Me.lblResults & vbCrLf & vbCrLf & "Element Count: " & nElements & vbCrLf & "N Log(2) N = " & d & vbCrLf & "Comparison Count: " & SortElementComparisons
End Sub

