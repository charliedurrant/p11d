VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{8D988532-0F0C-460C-B00E-7B5637E97680}#1.0#0"; "atc2vtext.ocx"
Begin VB.Form F_EmployerDetails 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Employer Details"
   ClientHeight    =   9510
   ClientLeft      =   3075
   ClientTop       =   3315
   ClientWidth     =   7065
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9510
   ScaleWidth      =   7065
   StartUpPosition =   1  'CenterOwner
   Begin TabDlg.SSTab tab 
      Height          =   8565
      Left            =   0
      TabIndex        =   26
      Top             =   120
      Width           =   6900
      _ExtentX        =   12171
      _ExtentY        =   15108
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      Tab             =   2
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Required"
      TabPicture(0)   =   "Erdetail.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Lab(0)"
      Tab(0).Control(1)=   "Lab(1)"
      Tab(0).Control(2)=   "Lab(2)"
      Tab(0).Control(3)=   "TxtBx(1)"
      Tab(0).Control(4)=   "TxtBx(0)"
      Tab(0).Control(5)=   "TxtBx(2)"
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "Electronic Submission"
      TabPicture(1)   =   "Erdetail.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Lab(11)"
      Tab(1).Control(1)=   "Lab(12)"
      Tab(1).Control(2)=   "Lab(14)"
      Tab(1).Control(3)=   "Lab(15)"
      Tab(1).Control(4)=   "Lab(16)"
      Tab(1).Control(5)=   "Lab(17)"
      Tab(1).Control(6)=   "Lab(18)"
      Tab(1).Control(7)=   "Lab(19)"
      Tab(1).Control(8)=   "lblDemoElectronicFields"
      Tab(1).Control(9)=   "TxtBx(30)"
      Tab(1).Control(10)=   "TxtBx(29)"
      Tab(1).Control(11)=   "TxtBx(28)"
      Tab(1).Control(12)=   "TxtBx(27)"
      Tab(1).Control(13)=   "TxtBx(12)"
      Tab(1).Control(14)=   "TxtBx(11)"
      Tab(1).ControlCount=   15
      TabCaption(2)   =   "P11D(b)"
      TabPicture(2)   =   "Erdetail.frx":0038
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Label1"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "General"
      TabPicture(3)   =   "Erdetail.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Lab(3)"
      Tab(3).Control(1)=   "Lab(5)"
      Tab(3).Control(2)=   "Lab(13)"
      Tab(3).Control(3)=   "Lab(4)"
      Tab(3).Control(4)=   "Lab(8)"
      Tab(3).Control(5)=   "Lab(7)"
      Tab(3).Control(6)=   "Lab(6)"
      Tab(3).Control(7)=   "Lab(10)"
      Tab(3).Control(8)=   "Lab(9)"
      Tab(3).Control(9)=   "Label18"
      Tab(3).Control(10)=   "Label17"
      Tab(3).Control(11)=   "Label16"
      Tab(3).Control(12)=   "TxtBx(33)"
      Tab(3).Control(13)=   "TxtBx(32)"
      Tab(3).Control(14)=   "TxtBx(31)"
      Tab(3).Control(15)=   "TxtBx(10)"
      Tab(3).Control(16)=   "TxtBx(9)"
      Tab(3).Control(17)=   "TxtBx(8)"
      Tab(3).Control(18)=   "TxtBx(7)"
      Tab(3).Control(19)=   "TxtBx(6)"
      Tab(3).Control(20)=   "TxtBx(13)"
      Tab(3).Control(21)=   "TxtBx(5)"
      Tab(3).Control(22)=   "TxtBx(3)"
      Tab(3).Control(23)=   "TxtBx(4)"
      Tab(3).Control(24)=   "ChkBx(0)"
      Tab(3).Control(25)=   "ChkBx(1)"
      Tab(3).Control(26)=   "ChkBx(2)"
      Tab(3).ControlCount=   27
      Begin VB.CheckBox ChkBx 
         Alignment       =   1  'Right Justify
         Caption         =   "Treat all loans as non 'Taxable Cheap Loans'"
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   2
         Left            =   -74820
         TabIndex        =   17
         Top             =   2520
         Width           =   6555
      End
      Begin VB.CheckBox ChkBx 
         Alignment       =   1  'Right Justify
         Caption         =   "&Treat employee cars as under approved MARORS"
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   1
         Left            =   -74820
         TabIndex        =   16
         Top             =   2205
         Width           =   6555
      End
      Begin VB.CheckBox ChkBx 
         Alignment       =   1  'Right Justify
         Caption         =   "&CT treatment of employee entertainment"
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   0
         Left            =   -74820
         TabIndex        =   15
         Top             =   1890
         Width           =   6555
      End
      Begin atc2valtext.ValText TxtBx 
         Height          =   285
         Index           =   2
         Left            =   -73245
         TabIndex        =   2
         Top             =   1350
         Width           =   2325
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   255
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxLength       =   6
         MouseIcon       =   "Erdetail.frx":0070
         Text            =   ""
         TypeOfData      =   3
         AllowEmpty      =   0   'False
      End
      Begin atc2valtext.ValText TxtBx 
         Height          =   285
         Index           =   0
         Left            =   -73245
         TabIndex        =   0
         Top             =   645
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   503
         BackColor       =   255
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxLength       =   255
         MouseIcon       =   "Erdetail.frx":008C
         Text            =   ""
         TypeOfData      =   3
         AllowEmpty      =   0   'False
      End
      Begin atc2valtext.ValText TxtBx 
         Height          =   285
         Index           =   1
         Left            =   -73245
         TabIndex        =   1
         Top             =   1035
         Width           =   2325
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
         MaxLength       =   50
         MouseIcon       =   "Erdetail.frx":00A8
         Text            =   ""
         TypeOfData      =   4
         AllowEmpty      =   0   'False
      End
      Begin atc2valtext.ValText TxtBx 
         Height          =   285
         Index           =   11
         Left            =   -72840
         TabIndex        =   5
         Top             =   840
         Width           =   1425
         _ExtentX        =   2514
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
         MaxLength       =   8
         MouseIcon       =   "Erdetail.frx":00C4
         Text            =   ""
         Minimum         =   "0"
         TXTAlign        =   2
         AutoSelect      =   0
      End
      Begin atc2valtext.ValText TxtBx 
         Height          =   285
         Index           =   12
         Left            =   -72840
         TabIndex        =   6
         Top             =   1200
         Width           =   4410
         _ExtentX        =   7779
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
         MaxLength       =   35
         MouseIcon       =   "Erdetail.frx":00E0
         Text            =   ""
         TypeOfData      =   3
      End
      Begin atc2valtext.ValText TxtBx 
         Height          =   285
         Index           =   4
         Left            =   -70560
         TabIndex        =   12
         Top             =   810
         Width           =   2325
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
         MaxLength       =   255
         MouseIcon       =   "Erdetail.frx":00FC
         Text            =   ""
         TypeOfData      =   3
      End
      Begin atc2valtext.ValText TxtBx 
         Height          =   285
         Index           =   3
         Left            =   -70560
         TabIndex        =   11
         Top             =   495
         Width           =   2325
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
         MaxLength       =   50
         MouseIcon       =   "Erdetail.frx":0118
         Text            =   ""
         TypeOfData      =   3
      End
      Begin atc2valtext.ValText TxtBx 
         Height          =   285
         Index           =   5
         Left            =   -70560
         TabIndex        =   13
         Top             =   1125
         Width           =   2325
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
         MaxLength       =   50
         MouseIcon       =   "Erdetail.frx":0134
         Text            =   ""
         TypeOfData      =   3
      End
      Begin atc2valtext.ValText TxtBx 
         Height          =   285
         Index           =   13
         Left            =   -69720
         TabIndex        =   14
         Top             =   1440
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   503
         BackColor       =   255
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxLength       =   50
         MouseIcon       =   "Erdetail.frx":0150
         Text            =   ""
         TypeOfData      =   2
         AllowEmpty      =   0   'False
      End
      Begin atc2valtext.ValText TxtBx 
         Height          =   285
         Index           =   6
         Left            =   -70560
         TabIndex        =   18
         Top             =   2880
         Width           =   2325
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
         MaxLength       =   50
         MouseIcon       =   "Erdetail.frx":016C
         Text            =   ""
         TypeOfData      =   3
      End
      Begin atc2valtext.ValText TxtBx 
         Height          =   285
         Index           =   7
         Left            =   -70560
         TabIndex        =   19
         Top             =   3195
         Width           =   2325
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
         MaxLength       =   50
         MouseIcon       =   "Erdetail.frx":0188
         Text            =   ""
         TypeOfData      =   3
      End
      Begin atc2valtext.ValText TxtBx 
         Height          =   285
         Index           =   8
         Left            =   -70560
         TabIndex        =   20
         Top             =   3510
         Width           =   2325
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
         MaxLength       =   50
         MouseIcon       =   "Erdetail.frx":01A4
         Text            =   ""
         TypeOfData      =   3
      End
      Begin atc2valtext.ValText TxtBx 
         Height          =   285
         Index           =   9
         Left            =   -70560
         TabIndex        =   21
         Top             =   3825
         Width           =   2325
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
         MaxLength       =   50
         MouseIcon       =   "Erdetail.frx":01C0
         Text            =   ""
         TypeOfData      =   3
      End
      Begin atc2valtext.ValText TxtBx 
         Height          =   285
         Index           =   10
         Left            =   -70560
         TabIndex        =   22
         Top             =   4095
         Width           =   2325
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
         MaxLength       =   50
         MouseIcon       =   "Erdetail.frx":01DC
         Text            =   ""
         TypeOfData      =   3
      End
      Begin atc2valtext.ValText TxtBx 
         Height          =   285
         Index           =   27
         Left            =   -72840
         TabIndex        =   7
         Top             =   2025
         Width           =   1665
         _ExtentX        =   2937
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
         MaxLength       =   35
         MouseIcon       =   "Erdetail.frx":01F8
         Text            =   ""
         TypeOfData      =   3
         AutoSelect      =   0
      End
      Begin atc2valtext.ValText TxtBx 
         Height          =   285
         Index           =   28
         Left            =   -72840
         TabIndex        =   8
         Top             =   2400
         Width           =   1665
         _ExtentX        =   2937
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
         MaxLength       =   35
         MouseIcon       =   "Erdetail.frx":0214
         Text            =   ""
         TypeOfData      =   3
         AutoSelect      =   0
      End
      Begin atc2valtext.ValText TxtBx 
         Height          =   285
         Index           =   29
         Left            =   -72840
         TabIndex        =   9
         Top             =   2760
         Width           =   1665
         _ExtentX        =   2937
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
         MaxLength       =   35
         MouseIcon       =   "Erdetail.frx":0230
         Text            =   ""
         TypeOfData      =   3
         AutoSelect      =   0
      End
      Begin atc2valtext.ValText TxtBx 
         Height          =   285
         Index           =   30
         Left            =   -72840
         TabIndex        =   10
         Top             =   3120
         Width           =   4290
         _ExtentX        =   7567
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
         MaxLength       =   255
         MouseIcon       =   "Erdetail.frx":024C
         Text            =   ""
         TypeOfData      =   3
         AutoSelect      =   0
      End
      Begin atc2valtext.ValText TxtBx 
         Height          =   285
         Index           =   31
         Left            =   -70560
         TabIndex        =   23
         Top             =   4590
         Width           =   2325
         _ExtentX        =   4101
         _ExtentY        =   503
         BackColor       =   255
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxLength       =   50
         MouseIcon       =   "Erdetail.frx":0268
         Text            =   ""
         TypeOfData      =   3
         Minimum         =   "1"
         AllowEmpty      =   0   'False
      End
      Begin atc2valtext.ValText TxtBx 
         Height          =   285
         Index           =   32
         Left            =   -70560
         TabIndex        =   24
         Top             =   4905
         Width           =   2325
         _ExtentX        =   4101
         _ExtentY        =   503
         BackColor       =   255
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxLength       =   50
         MouseIcon       =   "Erdetail.frx":0284
         Text            =   ""
         TypeOfData      =   3
         Minimum         =   "1"
         AllowEmpty      =   0   'False
      End
      Begin atc2valtext.ValText TxtBx 
         Height          =   285
         Index           =   33
         Left            =   -70560
         TabIndex        =   25
         Top             =   5265
         Width           =   2325
         _ExtentX        =   4101
         _ExtentY        =   503
         BackColor       =   255
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxLength       =   50
         MouseIcon       =   "Erdetail.frx":02A0
         Text            =   ""
         TypeOfData      =   3
         Minimum         =   "1"
         AllowEmpty      =   0   'False
      End
      Begin VB.Label Label1 
         Caption         =   "P11D(b) information has moved. Please load the employer and use the menu Employer->P11D(b)"
         Height          =   480
         Left            =   345
         TabIndex        =   51
         Top             =   585
         Width           =   5895
      End
      Begin VB.Label lblDemoElectronicFields 
         Caption         =   "lblDemoElectronicFields"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   1455
         Left            =   -74820
         TabIndex        =   50
         Top             =   3780
         Width           =   5235
      End
      Begin VB.Label Label16 
         Caption         =   "Group code 1 alias"
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   -74775
         TabIndex        =   49
         Tag             =   "free,font"
         Top             =   4635
         Width           =   1545
      End
      Begin VB.Label Label17 
         Caption         =   "Group code 2 alias"
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   -74775
         TabIndex        =   48
         Tag             =   "free,font"
         Top             =   4950
         Width           =   1545
      End
      Begin VB.Label Label18 
         Caption         =   "Group code 3 alias"
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   -74775
         TabIndex        =   47
         Tag             =   "free,font"
         Top             =   5265
         Width           =   1500
      End
      Begin VB.Label Lab 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Online Services"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   19
         Left            =   -74760
         TabIndex        =   46
         Top             =   1800
         Width           =   1350
      End
      Begin VB.Label Lab 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Magnetic Media"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   18
         Left            =   -74760
         TabIndex        =   45
         Top             =   600
         Width           =   1365
      End
      Begin VB.Label Lab 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Email"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   17
         Left            =   -74520
         TabIndex        =   44
         Top             =   3165
         Width           =   375
      End
      Begin VB.Label Lab 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Submitter's name"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   16
         Left            =   -74520
         TabIndex        =   43
         Top             =   2790
         Width           =   1200
      End
      Begin VB.Label Lab 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Password"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   15
         Left            =   -74520
         TabIndex        =   42
         Top             =   2445
         Width           =   690
      End
      Begin VB.Label Lab 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "User ID"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   14
         Left            =   -74520
         TabIndex        =   41
         Top             =   2070
         Width           =   540
      End
      Begin VB.Label Lab 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Address line 4"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   9
         Left            =   -74775
         TabIndex        =   40
         Top             =   3870
         Width           =   990
      End
      Begin VB.Label Lab 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Postcode"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   10
         Left            =   -74775
         TabIndex        =   39
         Top             =   4185
         Width           =   675
      End
      Begin VB.Label Lab 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Address line 1"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   6
         Left            =   -74760
         TabIndex        =   38
         Top             =   2925
         Width           =   990
      End
      Begin VB.Label Lab 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Address line 2"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   7
         Left            =   -74775
         TabIndex        =   37
         Top             =   3240
         Width           =   990
      End
      Begin VB.Label Lab 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Address line 3"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   8
         Left            =   -74775
         TabIndex        =   36
         Top             =   3555
         Width           =   990
      End
      Begin VB.Label Lab 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Contact &number"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   4
         Left            =   -74760
         TabIndex        =   35
         Top             =   1170
         Width           =   1125
      End
      Begin VB.Label Lab 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Employee letter response date"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   13
         Left            =   -74760
         TabIndex        =   34
         Top             =   1530
         Width           =   2130
      End
      Begin VB.Label Lab 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Contact name"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   5
         Left            =   -74775
         TabIndex        =   33
         Top             =   810
         Width           =   990
      End
      Begin VB.Label Lab 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Signatory"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   3
         Left            =   -74760
         TabIndex        =   32
         Top             =   500
         Width           =   660
      End
      Begin VB.Label Lab 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Submitter name"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   12
         Left            =   -74520
         TabIndex        =   31
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label Lab 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Submitter ref"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   11
         Left            =   -74520
         TabIndex        =   30
         Top             =   880
         Width           =   885
      End
      Begin VB.Label Lab 
         BackStyle       =   0  'Transparent
         Caption         =   "&Filename"
         ForeColor       =   &H00800000&
         Height          =   435
         Index           =   2
         Left            =   -74760
         TabIndex        =   29
         Top             =   1350
         Width           =   1740
      End
      Begin VB.Label Lab 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&PAYE"
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   1
         Left            =   -74760
         TabIndex        =   28
         Top             =   1000
         Width           =   1755
      End
      Begin VB.Label Lab 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Name"
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   0
         Left            =   -74760
         TabIndex        =   27
         Top             =   650
         Width           =   1755
      End
   End
   Begin VB.CommandButton B_Ok 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   350
      Left            =   4680
      TabIndex        =   3
      Top             =   9000
      Width           =   1095
   End
   Begin VB.CommandButton B_Cancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   350
      Left            =   5880
      TabIndex        =   4
      Top             =   9000
      Width           =   1095
   End
End
Attribute VB_Name = "F_EmployerDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Implements IFrmGeneral
Private m_InvalidVT As Control
Public m_OK As Boolean
Private Sub B_Cancel_Click()
  m_OK = False
  Me.Hide
End Sub
Private Sub B_OK_Click()
  m_OK = True
  Call CheckValidity(Me)
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If AllowElectronicFieldsToBeset Then
    If lows.IsControlShiftE(KeyCode, Shift) Then
      'm media
      TxtBx(0).Text = "SCRIBE AND CO"
      TxtBx(6).Text = "The Office"
      TxtBx(7).Text = "The Road"
      TxtBx(8).Text = "The City"
      TxtBx(9).Text = ""
      TxtBx(10).Text = "TF3 4ER"
      
      TxtBx(11).Text = "702784" 'submitter ref
      TxtBx(12).Text = "Deloitte & Touche" 'submitter name
      'paye online
      TxtBx(27).Text = "isv234" 'id
      TxtBx(28).Text = "testing1" 'password
      TxtBx(29).Text = "deloitte" 'username
      TxtBx(30).Text = "chdurrant@deloitte.co.uk" 'email
      
      TxtBx(1).Text = "999/A234" 'paye ref
      p11d32.PAYEonline.ExtraSubmissionPropertiesMenu = True
      p11d32.PAYEonline.Efiler_Proceed_Submission = True
      p11d32.PAYEonline.Efiler_Test_Submission = True
    End If
  End If
End Sub
Private Property Get AllowElectronicFieldsToBeset() As Boolean
  AllowElectronicFieldsToBeset = p11d32.LicenceType = LT_DEMO Or IsRunningInIDE
End Property
Private Sub Form_Load()
  Dim s As String
  Me.KeyPreview = True
  Me.tab.tab = 0
  ChkBx(1).Visible = True
  TxtBx(13).Minimum = DateValReadToScreen(p11d32.Rates.value(TaxYearStart))
  If AllowElectronicFieldsToBeset Then
    s = "Press SHIFT CTRL E to fill with valid values"
  End If
  lblDemoElectronicFields.Caption = s
  lblDemoElectronicFields.Visible = IsRunningInIDE()
  
  
  
End Sub
Private Function IFrmGeneral_CheckChanged(c As Control) As Boolean
  
End Function
Public Property Get IFrmGeneral_InvalidVT() As Control
  Set IFrmGeneral_InvalidVT = m_InvalidVT
End Property
Public Property Set IFrmGeneral_InvalidVT(NewValue As Control)
  Set m_InvalidVT = NewValue
End Property

Private Sub Label20_Click()
End Sub

Private Sub Label19_Click()
End Sub

Private Sub lblEmployerDeclaration_Click()
End Sub

Private Sub Label3_Click()

End Sub

Private Sub Label11_Click()
End Sub

Private Sub Label7_Click()

End Sub

Private Sub TxtBx_FieldInvalid(Index As Integer, Valid As Boolean, Message As String)
  Call SetPanel2(Message)
End Sub
Private Sub TxtBx_GotFocus(Index As Integer)
  TxtBx(Index).lValidate
End Sub
Private Sub TxtBx_UserValidate(Index As Integer, Valid As Boolean, Message As String, sTextEntered As String)
  Valid = False
  If (Len(Trim$(sTextEntered)) = 0) Then
    Exit Sub
  End If
  Valid = ValidatePAYE(sTextEntered)
  If (Not Valid) Then Message = "PAYE reference is invalid"
End Sub


