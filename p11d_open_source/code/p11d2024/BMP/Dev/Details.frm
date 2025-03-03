VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.2#0"; "COMCTL32.OCX"
Object = "{4582CA9E-1A45-11D2-8D2F-00C04FA9DD6F}#1.0#0"; "ATC2VTEXT.OCX"
Begin VB.Form F_Details 
   Caption         =   "Personal Details"
   ClientHeight    =   5520
   ClientLeft      =   1080
   ClientTop       =   2505
   ClientWidth     =   8430
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   Icon            =   "Details.frx":0000
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5520
   ScaleWidth      =   8430
   WindowState     =   2  'Maximized
   Begin ComctlLib.ListView LB 
      Height          =   2565
      Left            =   75
      TabIndex        =   0
      Tag             =   "free,font"
      Top             =   105
      Width           =   8265
      _ExtentX        =   14579
      _ExtentY        =   4524
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   7
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Reference"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   2
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "NI Number"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   3
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Status"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   4
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Group1"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   5
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Group2"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   6
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Group3"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Frame fmeDetails 
      Height          =   2775
      Left            =   90
      TabIndex        =   14
      Top             =   2700
      Width           =   8265
      Begin VB.CommandButton B_ChangePNum 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   4995
         TabIndex        =   3
         Tag             =   "FREE"
         Top             =   1665
         Width           =   315
      End
      Begin atc2valtext.ValText TB_Data 
         Height          =   285
         Index           =   6
         Left            =   1935
         TabIndex        =   5
         Tag             =   "FREE,FONT"
         Top             =   2385
         Width           =   975
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "Details.frx":030A
         Text            =   ""
         TypeOfData      =   3
      End
      Begin atc2valtext.ValText TB_Data 
         Height          =   585
         Index           =   7
         Left            =   5535
         TabIndex        =   11
         Tag             =   "FREE,FONT"
         Top             =   2115
         Width           =   2640
         _ExtentX        =   4657
         _ExtentY        =   1032
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "Details.frx":0326
         Text            =   ""
         TypeOfData      =   3
      End
      Begin VB.ComboBox CB_Status 
         DataField       =   "Status"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6630
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Tag             =   "FREE,FONT"
         Top             =   195
         Width           =   1545
      End
      Begin atc2valtext.ValText TB_Data 
         Height          =   285
         Index           =   2
         Left            =   6630
         TabIndex        =   8
         Tag             =   "FREE,FONT"
         Top             =   555
         Width           =   1545
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "Details.frx":0342
         Text            =   ""
         TypeOfData      =   3
      End
      Begin atc2valtext.ValText TB_Date 
         Height          =   285
         Index           =   1
         Left            =   6645
         TabIndex        =   9
         Tag             =   "FREE,FONT"
         Top             =   915
         Width           =   1545
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "Details.frx":035E
         Text            =   ""
         TypeOfData      =   2
         Maximum         =   "5/4/1999"
         Minimum         =   "6/4/1998"
      End
      Begin atc2valtext.ValText TB_Date 
         Height          =   285
         Index           =   2
         Left            =   6630
         TabIndex        =   10
         Tag             =   "FREE,FONT"
         Top             =   1275
         Width           =   1545
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "Details.frx":037A
         Text            =   ""
         TypeOfData      =   2
         Maximum         =   "5/4/1999"
         Minimum         =   "6/4/1998"
      End
      Begin atc2valtext.ValText TB_Data 
         Height          =   285
         Index           =   8
         Left            =   4365
         TabIndex        =   6
         Tag             =   "FREE,FONT"
         Top             =   2385
         Width           =   975
         _ExtentX        =   0
         _ExtentY        =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "Details.frx":0396
         Text            =   ""
         TypeOfData      =   3
      End
      Begin atc2valtext.ValText TB_Data 
         Height          =   285
         Index           =   3
         Left            =   1935
         TabIndex        =   4
         Tag             =   "FREE,FONT"
         Top             =   2025
         Width           =   3390
         _ExtentX        =   5980
         _ExtentY        =   503
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "Details.frx":03B2
         Text            =   ""
         TypeOfData      =   3
      End
      Begin atc2valtext.ValText TB_Data 
         Height          =   285
         Index           =   1
         Left            =   1935
         TabIndex        =   2
         Tag             =   "FREE,FONT"
         Top             =   1665
         Width           =   3075
         _ExtentX        =   5424
         _ExtentY        =   503
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Locked          =   -1  'True
         MouseIcon       =   "Details.frx":03CE
         Text            =   ""
         TypeOfData      =   3
      End
      Begin atc2valtext.ValText TB_Data 
         Height          =   285
         Index           =   0
         Left            =   540
         TabIndex        =   1
         Tag             =   "FREE,FONT"
         Top             =   180
         Width           =   690
         _ExtentX        =   1217
         _ExtentY        =   503
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "Details.frx":03EA
         Text            =   ""
         TypeOfData      =   3
      End
      Begin atc2valtext.ValText TB_Data 
         Height          =   285
         Index           =   4
         Left            =   630
         TabIndex        =   23
         Tag             =   "FREE,FONT"
         Top             =   585
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   503
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "Details.frx":0406
         Text            =   ""
         TypeOfData      =   3
      End
      Begin atc2valtext.ValText TB_Data 
         Height          =   285
         Index           =   5
         Left            =   2340
         TabIndex        =   24
         Tag             =   "FREE,FONT"
         Top             =   180
         Width           =   2985
         _ExtentX        =   5265
         _ExtentY        =   503
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "Details.frx":0422
         Text            =   ""
         TypeOfData      =   3
      End
      Begin atc2valtext.ValText TB_Data 
         Height          =   285
         Index           =   9
         Left            =   2340
         TabIndex        =   25
         Tag             =   "FREE,FONT"
         Top             =   585
         Width           =   2985
         _ExtentX        =   5265
         _ExtentY        =   503
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "Details.frx":043E
         Text            =   ""
         TypeOfData      =   3
      End
      Begin atc2valtext.ValText TB_Data 
         Height          =   285
         Index           =   10
         Left            =   3870
         TabIndex        =   29
         Tag             =   "FREE,FONT"
         Top             =   945
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "Details.frx":045A
         Text            =   ""
         TypeOfData      =   3
      End
      Begin atc2valtext.ValText TB_Data 
         Height          =   285
         Index           =   11
         Left            =   765
         TabIndex        =   31
         Tag             =   "FREE,FONT"
         Top             =   1305
         Width           =   4560
         _ExtentX        =   8043
         _ExtentY        =   503
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "Details.frx":0476
         Text            =   ""
         TypeOfData      =   3
      End
      Begin VB.Label Label13 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Email"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   135
         TabIndex        =   32
         Tag             =   "FREE,FONT"
         Top             =   1350
         Width           =   735
      End
      Begin VB.Label Label12 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Salutation"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   1440
         TabIndex        =   30
         Tag             =   "FREE,FONT"
         Top             =   990
         Width           =   705
      End
      Begin VB.Label Label9 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Surname"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   1440
         TabIndex        =   28
         Tag             =   "FREE,FONT"
         Top             =   630
         Width           =   630
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Title"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   135
         TabIndex        =   27
         Tag             =   "FREE,FONT"
         Top             =   225
         Width           =   300
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Initials"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   135
         TabIndex        =   26
         Tag             =   "FREE,FONT"
         Top             =   630
         Width           =   435
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Comments"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   2
         Left            =   5535
         TabIndex        =   21
         Tag             =   "FREE,FONT"
         Top             =   1800
         Width           =   1515
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Group Code 3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   1
         Left            =   3195
         TabIndex        =   16
         Tag             =   "FREE,FONT"
         Top             =   2430
         Width           =   1155
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "First name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   1440
         TabIndex        =   12
         Tag             =   "FREE,FONT"
         Top             =   225
         Width           =   720
      End
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Personnel Reference"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   90
         TabIndex        =   13
         Tag             =   "FREE,FONT"
         Top             =   1665
         Width           =   1755
      End
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "NI Number"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   5550
         TabIndex        =   18
         Tag             =   "FREE,FONT"
         Top             =   630
         Width           =   1095
      End
      Begin VB.Label Label10 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Date started"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   5550
         TabIndex        =   19
         Tag             =   "FREE,FONT"
         Top             =   990
         Width           =   1110
      End
      Begin VB.Label Label11 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Date leaving"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   5550
         TabIndex        =   20
         Tag             =   "FREE,FONT"
         Top             =   1350
         Width           =   1095
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Group Code 1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   135
         TabIndex        =   22
         Tag             =   "FREE,FONT"
         Top             =   2070
         Width           =   1635
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Group Code 2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   15
         Tag             =   "FREE,FONT"
         Top             =   2475
         Width           =   1755
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Status"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   5550
         TabIndex        =   17
         Tag             =   "FREE,FONT"
         Top             =   315
         Width           =   1095
      End
   End
End
Attribute VB_Name = "F_Details"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Implements IBenefitForm2
Implements IFrmGeneral

Private mclsResize As New clsFormResize
Private Const L_DES_HEIGHT = 5925
Private Const L_DES_WIDTH = 8505
Private m_InvalidVt As atc2valtext.ValText

Private Sub AddEmployee()
  Dim vancol As clsVansCollection
  Dim lst As ListItem, b As IBenefitClass
  Dim ee As clsEmployee
  Dim ben As IBenefitClass
  Dim ibf As IBenefitForm2
  Dim I As Long
  
  On Error GoTo AddEmployee_Err
  Call xSet("AddEmployee")
  
  'if there is a current employee then we must save the data for that employee
  If Not LoadEmployee(Nothing, False) Then GoTo AddEmployee_End
  
RE_TEST:
  
  F_EeNew.Show 1
  'Create new employee and fill in details...
  If F_EeNew.m_OK = False Then
    Call SelectBenefit(Me)
    GoTo AddEmployee_End
  End If
  
  Set ee = New clsEmployee
  With ee
    
    Set .Parent = CurrentEmployer
    
    .PersonelNo = F_EeNew.TxtBx(0).Text
    .Surname = F_EeNew.TxtBx(1).Text
    .Title = F_EeNew.TxtBx(2).Text
    .FirstName = F_EeNew.TxtBx(3).Text
    .Initials = F_EeNew.TxtBx(4).Text
    .NINumber = F_EeNew.TxtBx(5).Text
    .Payeref = F_EeNew.TxtBx(6).Text
    
    Call ee.LoadBenefit(TBL_VANS) ' Load Van Collection
    Set b = New clsBenShVan
    Call b.SetItem(shvan_employeereference, ee.PersonelNo)
    Set b.Parent = ee.benefits.Item(1)
    Set vancol = ee.benefits.Item(1)
    vancol.HasSharedVan = False
    Set vancol.SharedVan = b
    Set ibf = Me
    .WriteDB
            
    I = Employees.Add(ee)
    Set lst = lb.ListItems.Add(, , .Name)
      
    Call ibf.UpdateBenefitListViewItem(lst, ee, I, True)
    Call ibf.BenefitToScreen(I)
  End With
  Set ee = Nothing
    
AddEmployee_End:
  Unload F_EeNew
  Set F_EeNew = Nothing
  MDIMain.Enabled = True
  Set ee = Nothing
  Call DBEngine.Idle(dbFreeLocks)
  Call xReturn("AddEmployee")
  Exit Sub
  
AddEmployee_Err:
  Select Case Err.Number
    Case 3022
      Call ErrorMessage(ERR_ERROR, Err, "Duplicate Personnel reference", "ERR_ADDEMPOYEE", "The personnel reference you are trying to add already exists in the database." & vbCrLf & "Please use an alternative reference")
      Resume RE_TEST
    Case 3315
      Call ErrorMessage(ERR_ERROR, Err, "Empty field", "ERR_ADDEMPOYEE", "You must complete all the fields.")
      Resume RE_TEST
    Case Else
      Call ErrorMessage(ERR_ERROR, Err, "AddEmployee", "ERR_AddEmployee", "Error in AddEmployee function, called from the form " & Me.Name & ".")
      Resume AddEmployee_End
  End Select
  Resume
End Sub

Private Sub CB_Status_Click()
  Call IFrmGeneral_CheckChanged(CB_Status, True)
End Sub

Private Sub IBenefitForm2_AddBenefit()
  Call AddEmployee
End Sub

Private Property Get IBenefitForm2_benefit() As IBenefitClass
  Set IBenefitForm2_benefit = CurrentEmployee
End Property

Private Property Let IBenefitForm2_benefit(NewValue As IBenefitClass)
  
End Property

Private Function IBenefitForm2_BenefitFormState(ByVal fState As BENEFIT_FORM_STATE) As Boolean
On Error GoTo BenefitFormState_err
  Call xSet("IBenefitForm2_EnableBenefitForm")
  
  If (fState = FORM_ENABLED) Or (fState = FORM_CDB) Then
    If fState = FORM_ENABLED Then
      fmeDetails.Enabled = True
    Else
      'cbd zzzz
    End If
    lb.Enabled = True
    MDIMain.mnuBenefits.Enabled = True
    MDIMain.tbrBenefits.Enabled = True
    Call MDIMain.SetDelete
  ElseIf fState = FORM_DISABLED Then
    fmeDetails.Enabled = False
    lb.Enabled = False
    Call MDIMain.ClearDelete
    Call MDIMain.ClearConfirmUndo
    MDIMain.mnuBenefits.Enabled = False
    MDIMain.tbrBenefits.Enabled = False
  End If
  
  Call MDIMain.NavigateBarUpdate(CurrentEmployee)
  
  IBenefitForm2_BenefitFormState = True
    
BenefitFormState_end:
  Call xReturn("BenefitFormState")
  Exit Function
  
BenefitFormState_err:
  IBenefitForm2_BenefitFormState = False
  Call ErrorMessage(ERR_ERROR, Err, "BenefitFormState", "ERR_UNDEFINED", "Undefined error.")
  Resume BenefitFormState_end
  Resume
End Function

Private Function IBenefitForm2_BenefitsToListView() As Long
  Dim v As Variant
  Dim I As Long
  Dim lst As ListItem
  Dim emp As clsEmployee
  
  On Error GoTo BenefitsToListView_err
  Call xSet("BenefitsToListView")
  
  
  Call ClearForm(Me)
  Call MDIMain.SetAdd
  
  Call ColumnWidths(lb, 30, 10, 10, 10, 10, 10, 10)
  For I = 1 To Employees.count
    Set emp = Employees.Item(I)
    If Not emp Is Nothing Then
      Set lst = lb.ListItems.Add(, , emp.Name)
      lst.Tag = I
      'if this changes change in GotoScreen()
      lst.SubItems(1) = emp.PersonelNo
      lst.SubItems(2) = emp.NINumber
      lst.SubItems(3) = IIf(emp.Status, "Director", "Staff")
      lst.SubItems(4) = emp.Group1
      lst.SubItems(5) = emp.Group2
      lst.SubItems(6) = emp.Group3
    End If
    IBenefitForm2_BenefitsToListView = IBenefitForm2_BenefitsToListView + 1
  Next I
      
BenefitsToListView_end:
  Set emp = Nothing
  Set lst = Nothing
  Call xReturn("BenefitsToListView")
  Exit Function
  
BenefitsToListView_err:
  Call ErrorMessage(ERR_ERROR, Err, "BenefitsToListView", "Adding Benefits to ListView", "Error in adding benefits to listview.")
  Resume BenefitsToListView_end
End Function

Private Function IBenefitForm2_BenefitToScreen(Optional ByVal BenefitIndex As Long = -1&, Optional ByVal UpdateBenefit As Boolean = True) As IBenefitClass
  Dim lst As ListItem
  Dim emp As clsEmployee
  Dim ibf As IBenefitForm2
  
  On Error GoTo BenefitToScreen_Err
  
  Call xSet("BenefitToScreen")
  If UpdateBenefit Then Call UpdateBenefitFromTags
  
  If BenefitIndex <> -1 Then
    Set emp = Employees.Item(BenefitIndex)
    'we want to save employee details when ever moving employee
    If Not LoadEmployee(emp, False) Then GoTo BenefitToScreen_End
    If emp Is Nothing Then Call Ecase("Employee -emp is nothing check BenefitToScreen")
    
    TB_Data(0).Text = emp.Title
    TB_Data(5).Text = emp.FirstName
    TB_Data(4).Text = emp.Initials
    TB_Data(9).Text = emp.Surname
    TB_Data(10).Text = emp.Salutation
    TB_Data(11).Text = emp.Email
    
    TB_Data(1).Text = emp.PersonelNo
    TB_Data(2).Text = emp.NINumber
    TB_Data(3).Text = emp.Group1
    TB_Data(6).Text = emp.Group2
    TB_Data(7).Text = emp.Comments
    
    CB_Status = IIf(emp.Status, S_DIRECTOR, S_STAFF)
    TB_Data(8).Text = emp.Group3
    TB_Date(1).Text = DateStringEx(emp.Joined, emp.Joined)
    TB_Date(2).Text = DateStringEx(emp.Left, emp.Left)
  Else
    TB_Data(0).Text = ""
    TB_Data(1).Text = ""
    TB_Data(2).Text = ""
    TB_Data(3).Text = ""
    TB_Data(4).Text = ""
    TB_Data(5).Text = ""
    TB_Data(6).Text = ""
    TB_Data(7).Text = ""
    'CB_Status = ""
    TB_Data(8).Text = ""
    TB_Data(9).Text = ""
    TB_Data(10).Text = ""
    TB_Data(11).Text = ""
    TB_Date(1).Text = ""
    TB_Date(2).Text = ""
  End If
  
  Set ibf = Me
  If emp Is Nothing Then
    Set CurrentEmployee = Nothing
    Call ibf.BenefitFormState(FORM_DISABLED)
  Else
    emp.InvalidFields = InvalidFields(ibf)
    Call ibf.BenefitFormState(FORM_ENABLED)
  End If
   
BenefitToScreen_End:
  Set ibf = Nothing
  Set lst = Nothing
  Set emp = Nothing
  Call xReturn("BenefitToScreen")
  Exit Function

BenefitToScreen_Err:
  Call ErrorMessage(ERR_ERROR, Err, "BenefitToScreen", "ERR_UNDEFINED", "Unable to place then chosen benefit to the screen. benefit index = " & BenefitIndex & ".")
  Resume BenefitToScreen_End
  Resume
End Function

Private Property Let IBenefitForm2_bentype(ByVal RHS As benClass)

End Property

Private Property Get IBenefitForm2_bentype() As benClass

End Property

Private Property Get IBenefitForm2_lv() As ComctlLib.IListView
  Set IBenefitForm2_lv = lb
End Property

Private Function IBenefitForm2_RemoveBenefit(ByVal BenefitIndex As Long) As Boolean
  Dim NextBenefitIndex As Long
  Dim ben As IBenefitClass
  Dim ee As clsEmployee
  Dim ibf As IBenefitForm2
  On Error GoTo RemoveBenefit_ERR
  
  Call xSet("RemoveBenefit")
  
  Set ee = Employees(BenefitIndex)
  If Not ee Is Nothing Then
    NextBenefitIndex = GetNextBestListItemBenefitIndex(Me, BenefitIndex)
    Call LoadEmployee(ee, , LEMBeforeDelete)
    Set ben = ee
    Call ben.DeleteDB
    Call ben.Kill
    Call Employees.Remove(BenefitIndex)
    Set ibf = Me
    ibf.lv.ListItems.Remove (ibf.lv.SelectedItem.Index)
    Call SelectBenefit(Me, NextBenefitIndex)
    IBenefitForm2_RemoveBenefit = True
  End If
    
RemoveBenefit_END:
  Call xReturn("RemoveBenefit")
  Exit Function
  
RemoveBenefit_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "RemoveBenefit", "Remove Benefit", "Error removing benefit.")
  Resume RemoveBenefit_END
  Resume
End Function

Private Function IBenefitForm2_UpdateBenefitListViewItem(li As ComctlLib.IListItem, benefit As IBenefitClass, Optional ByVal BenefitIndex As Long = 0&, Optional ByVal SelectItem As Boolean = False) As Boolean
  Dim ee As clsEmployee
  
  On Error GoTo UpdateBenefitListViewItem_ERR
  
  Call xSet("UpdateBenefitListViewItem")
  
  If Not li Is Nothing And Not benefit Is Nothing Then
    Set ee = benefit
    If BenefitIndex > 0 Then li.Tag = BenefitIndex
    With li
      li.Text = ee.Name
      li.SubItems(1) = ee.PersonelNo
      li.SubItems(2) = ee.NINumber
      li.SubItems(3) = IIf(ee.Status, "Director", "Staff")
      li.SubItems(4) = ee.Group1
      li.SubItems(5) = ee.Group2
      li.SubItems(6) = ee.Group3
      If SelectItem Then Set IBenefitForm2_lv.SelectedItem = li
    End With
    IBenefitForm2_UpdateBenefitListViewItem = True
  End If
  
UpdateBenefitListViewItem_END:
  Set ee = Nothing
  Call xReturn("UpdateBenefitListViewItem")
  Exit Function
  
UpdateBenefitListViewItem_ERR:
  IBenefitForm2_UpdateBenefitListViewItem = False
  'If Err <> 35605 Then  's control has been deleted
    Call ErrorMessage(ERR_ERROR, Err, "UpdateBenefitListViewItem", "Updating a benefits list view text", "Unable to update the benefits list view text.")
  'End If
  Resume UpdateBenefitListViewItem_END
  Resume

End Function

Private Function IFrmGeneral_CheckChanged(C As Control, ByVal UpdateCurrentListItem As Boolean) As Boolean
  Dim lst As ListItem
  Dim vDate As Date
  Dim bDirty As Boolean
  
  On Error GoTo CheckChanged_Err
  Call xSet("CheckChanged")
  
  With C
    If CurrentEmployee Is Nothing Then
      GoTo CheckChanged_End
    End If
    'we are asking if the value has changed and if it is valid thus save
    
    Select Case .Name
        Case "TB_Data"
          Select Case .Index
            Case 0
              bDirty = StrComp(.Text, CurrentEmployee.Title, vbBinaryCompare)
              If bDirty Then CurrentEmployee.Title = .Text
            Case 5
              bDirty = StrComp(.Text, CurrentEmployee.FirstName, vbBinaryCompare)
              If bDirty Then CurrentEmployee.Title = .Text
            Case 4
              bDirty = StrComp(.Text, CurrentEmployee.Initials, vbBinaryCompare)
              If bDirty Then CurrentEmployee.Initials = .Text
            Case 9
              bDirty = StrComp(.Text, CurrentEmployee.Surname, vbBinaryCompare)
              If bDirty Then CurrentEmployee.Surname = .Text
            Case 10
              bDirty = StrComp(.Text, CurrentEmployee.Salutation, vbBinaryCompare)
              If bDirty Then CurrentEmployee.Salutation = .Text
            Case 11
              bDirty = StrComp(.Text, CurrentEmployee.Email, vbBinaryCompare)
              If bDirty Then CurrentEmployee.Email = .Text
            Case 1
              'Ignore the P_NUM field
            Case 2
              bDirty = StrComp(.Text, CurrentEmployee.NINumber, vbBinaryCompare)
              If bDirty Then CurrentEmployee.NINumber = .Text
            Case 3
              bDirty = StrComp(.Text, CurrentEmployee.Group1, vbBinaryCompare)
              If bDirty Then CurrentEmployee.Group1 = .Text
            Case 6
              bDirty = StrComp(.Text, CurrentEmployee.Group2, vbBinaryCompare)
              If bDirty Then CurrentEmployee.Group2 = .Text
            Case 7
              bDirty = StrComp(.Text, CurrentEmployee.Comments, vbBinaryCompare)
              If bDirty Then CurrentEmployee.Comments = .Text
            Case 8
              bDirty = StrComp(.Text, CurrentEmployee.Group3, vbBinaryCompare)
              If bDirty Then CurrentEmployee.Group3 = .Text
            Case Else
              Ecase "Unknown control"
          End Select
        Case "TB_Date"
          Select Case .Index
            Case 1
              bDirty = StrComp(.Text, CurrentEmployee.Joined, vbBinaryCompare)
              If bDirty Then CurrentEmployee.Joined = .Text
            Case 2
              bDirty = StrComp(.Text, CurrentEmployee.Left, vbBinaryCompare)
              If bDirty Then CurrentEmployee.Left = .Text
          End Select
        Case "CB_Status"
          bDirty = StrComp(.Text, IIf(CurrentEmployee.Status, S_DIRECTOR, S_STAFF), vbBinaryCompare)
          If bDirty Then CurrentEmployee.Status = IIf(.Text = S_DIRECTOR, True, False)
        Case Else
          Ecase ("UNKNOWN CONTROL")
     End Select
    
    'must be required in all check changed
    IFrmGeneral_CheckChanged = AfterCheckChanged(C, Me, bDirty, UpdateCurrentListItem)

  End With
  
CheckChanged_End:
  Set lst = Nothing
  Call xReturn("CheckChanged")
  Exit Function
  
CheckChanged_Err:
  IFrmGeneral_CheckChanged = False
  Call ErrorMessage(ERR_ERROR, Err, "CheckChanged", "ERR_CHECKCHANGED", "This function has failed for the form " & Me.Name & ".")
  Resume CheckChanged_End
  Resume
End Function

Public Property Get IFrmGeneral_InvalidVT() As atc2valtext.ValText
  Set IFrmGeneral_InvalidVT = m_InvalidVt
End Property

Public Property Set IFrmGeneral_InvalidVT(NewValue As atc2valtext.ValText)
  Set m_InvalidVt = NewValue
End Property


Private Sub CB_Status_Lostfocus()
  Call IFrmGeneral_CheckChanged(CB_Status, True)
End Sub




Private Sub Form_Load()
  If Not (mclsResize.InitResize(Me, L_DES_HEIGHT, L_DES_WIDTH, DESIGN, , , MDIMain)) Then ', DESIGN)) Then
    Err.Raise ERR_Application
  End If
  CB_Status.Clear
  CB_Status.AddItem S_STAFF
  CB_Status.AddItem S_DIRECTOR
End Sub
Private Sub LB_DblClick()
  Call BenefitToolBar(1, GetEmployeeIndexFromSelectedEmployee)
End Sub

Private Sub LB_ItemClick(ByVal Item As ComctlLib.ListItem)
  Call SetLastListItemSelected(Item)
  If Not (lb.SelectedItem Is Nothing) Then
    IBenefitForm2_BenefitToScreen (Item.Tag)
  End If
End Sub

Private Sub LB_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then 'Return key
    Call LB_DblClick
  End If
End Sub

Private Sub Form_Resize()
  mclsResize.Resize
  Call ColumnWidths(Me.lb, L_NAME_COL, L_REFERENCE_COL, L_NINUMBER_COL&, L_STATUS_COL&, L_GROUP1_COL&, L_GROUP2_COL&, L_GROUP3_COL&)
End Sub

Private Sub LB_ColumnClick(ByVal ColumnHeader As ComctlLib.ColumnHeader)
  Me.lb.SortKey = ColumnHeader.Index - 1
  lb.SelectedItem.EnsureVisible
End Sub


Private Sub TB_Data_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  TB_Data(Index).Tag = SetChanged
End Sub

Private Sub TB_data_Lostfocus(Index As Integer)
  Call IFrmGeneral_CheckChanged(TB_Data(Index), True)
End Sub

Private Sub TB_Date_FieldInvalid(Index As Integer, Valid As Boolean, Message As String)
  Call MDIMain.sts.SetStatus(0, Message)
End Sub

Private Sub TB_Date_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  TB_Date(Index).Tag = SetChanged
End Sub

Private Sub TB_Date_LostFocus(Index As Integer)
  Call IFrmGeneral_CheckChanged(TB_Date(Index), True)
End Sub
