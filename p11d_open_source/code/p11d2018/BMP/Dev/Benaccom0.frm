VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{8D988532-0F0C-460C-B00E-7B5637E97680}#1.0#0"; "ATC2VTEXT.OCX"
Begin VB.Form F_Accommodation 
   Caption         =   " "
   ClientHeight    =   5685
   ClientLeft      =   585
   ClientTop       =   1785
   ClientWidth     =   8280
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
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5685
   ScaleWidth      =   8280
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin TabDlg.SSTab tabAccom 
      Height          =   3480
      Left            =   45
      TabIndex        =   26
      Tag             =   "FREE,FONT"
      Top             =   2160
      Width           =   8205
      _ExtentX        =   14473
      _ExtentY        =   6138
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Details"
      TabPicture(0)   =   "Benaccom0.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Lab(15)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Lab(21)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Lab(5)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Lab(25)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Lab(31)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Lab(32)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "TxtBx(2)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "TxtBx(8)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "TxtBx(15)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "TxtBx(19)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "TxtBx(20)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "ChkBx(4)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "ChkBx(0)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "ChkBx(1)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "fmeApportion"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "ChkBx(3)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).ControlCount=   16
      TabCaption(1)   =   "Expensive Accommodation"
      TabPicture(1)   =   "Benaccom0.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Lab(19)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Lab(8)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Lab(7)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Lab(18)"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Lab(14)"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Lab(24)"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Lab(23)"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Lab(29)"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "Lab(28)"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "TxtBx(3)"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "TxtBx(13)"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "TxtBx(14)"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "TxtBx(17)"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "TxtBx(9)"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "TxtBx(11)"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "TxtBx(18)"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "ChkBx(2)"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).ControlCount=   17
      TabCaption(2)   =   "Extra Expenses"
      TabPicture(2)   =   "Benaccom0.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Lab(17)"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Lab(9)"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Lab(1)"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Lab(27)"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "Lab(4)"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "Lab(12)"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "Lab(22)"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "Lab(30)"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "Lab(0)"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).Control(9)=   "TxtBx(5)"
      Tab(2).Control(9).Enabled=   0   'False
      Tab(2).Control(10)=   "TxtBx(6)"
      Tab(2).Control(10).Enabled=   0   'False
      Tab(2).Control(11)=   "TxtBx(16)"
      Tab(2).Control(11).Enabled=   0   'False
      Tab(2).Control(12)=   "TxtBx(0)"
      Tab(2).Control(12).Enabled=   0   'False
      Tab(2).Control(13)=   "TxtBx(10)"
      Tab(2).Control(13).Enabled=   0   'False
      Tab(2).Control(14)=   "TxtBx(12)"
      Tab(2).Control(14).Enabled=   0   'False
      Tab(2).Control(15)=   "cmdN1A(0)"
      Tab(2).Control(15).Enabled=   0   'False
      Tab(2).Control(16)=   "cmdN1A(1)"
      Tab(2).Control(16).Enabled=   0   'False
      Tab(2).Control(17)=   "cmdN1A(2)"
      Tab(2).Control(17).Enabled=   0   'False
      Tab(2).ControlCount=   18
      Begin VB.CommandButton cmdN1A 
         Caption         =   "L - Asset placed at employees disposal"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   -68280
         TabIndex        =   56
         Top             =   1800
         Width           =   1335
      End
      Begin VB.CommandButton cmdN1A 
         Caption         =   "M - Other items (non Class 1A)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   -68280
         TabIndex        =   55
         Top             =   1200
         Width           =   1335
      End
      Begin VB.CommandButton cmdN1A 
         Caption         =   "M - Other items (Class 1A)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   -68280
         TabIndex        =   54
         Top             =   720
         Width           =   1335
      End
      Begin VB.CheckBox ChkBx 
         Alignment       =   1  'Right Justify
         Caption         =   "Is the amount above an amount subjected to PAYE?"
         DataField       =   "IsRent"
         DataSource      =   "DB"
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
         Height          =   285
         Index           =   3
         Left            =   90
         TabIndex        =   7
         Tag             =   "free,font"
         Top             =   3050
         Width           =   4065
      End
      Begin VB.CheckBox ChkBx 
         Alignment       =   1  'Right Justify
         Caption         =   "Does the accommodation meet the six year rule?"
         DataField       =   "SixYear"
         DataSource      =   "DB"
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
         Height          =   285
         Index           =   2
         Left            =   -71775
         TabIndex        =   17
         Tag             =   "free,font"
         Top             =   1080
         Width           =   4600
      End
      Begin VB.Frame fmeApportion 
         Caption         =   "Note: Only annualised values require apportionment."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   1155
         Left            =   4200
         TabIndex        =   27
         Tag             =   "free,font"
         Top             =   2040
         Width           =   3900
         Begin atc2valtext.ValText TxtBx 
            Height          =   315
            Index           =   4
            Left            =   2235
            TabIndex        =   11
            Tag             =   "free,font"
            Top             =   675
            Width           =   1065
            _ExtentX        =   1879
            _ExtentY        =   556
            BackColor       =   255
            ForeColor       =   64
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MouseIcon       =   "Benaccom0.frx":0054
            Text            =   ""
            TypeOfData      =   2
            Maximum         =   "5/4/1999"
            Minimum         =   "6/4/1998"
            AllowEmpty      =   0   'False
         End
         Begin atc2valtext.ValText TxtBx 
            Height          =   315
            Index           =   1
            Left            =   2235
            TabIndex        =   10
            Tag             =   "free,font"
            Top             =   270
            Width           =   1065
            _ExtentX        =   1879
            _ExtentY        =   556
            BackColor       =   255
            ForeColor       =   64
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MouseIcon       =   "Benaccom0.frx":0070
            Text            =   ""
            TypeOfData      =   2
            Maximum         =   "5/4/1999"
            Minimum         =   "6/4/1998"
            AllowEmpty      =   0   'False
         End
         Begin VB.Label Lab 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Available from"
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
            Index           =   3
            Left            =   270
            TabIndex        =   29
            Tag             =   "free,font"
            Top             =   315
            Width           =   990
         End
         Begin VB.Label Lab 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Available to"
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
            Index           =   10
            Left            =   270
            TabIndex        =   28
            Tag             =   "free,font"
            Top             =   720
            Width           =   825
         End
      End
      Begin VB.CheckBox ChkBx 
         Alignment       =   1  'Right Justify
         Caption         =   "Is the accommodation job-related?"
         DataField       =   "JobRelated"
         DataSource      =   "DB"
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
         Height          =   285
         Index           =   1
         Left            =   4320
         TabIndex        =   8
         Tag             =   "free,font"
         Top             =   1035
         Width           =   3200
      End
      Begin VB.CheckBox ChkBx 
         Alignment       =   1  'Right Justify
         Caption         =   "Is the accommodation owned by the employer?"
         DataField       =   "ErOwn"
         DataSource      =   "DB"
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
         Height          =   375
         Index           =   0
         Left            =   60
         TabIndex        =   2
         Tag             =   "free,font"
         Top             =   920
         Width           =   4095
      End
      Begin VB.CheckBox ChkBx 
         Alignment       =   1  'Right Justify
         Caption         =   "Is the above figure rent?"
         DataField       =   "IsRent"
         DataSource      =   "DB"
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
         Height          =   285
         Index           =   4
         Left            =   90
         TabIndex        =   5
         Tag             =   "free,font"
         Top             =   2350
         Width           =   4065
      End
      Begin atc2valtext.ValText TxtBx 
         Height          =   315
         Index           =   20
         Left            =   3105
         TabIndex        =   6
         Tag             =   "free,font"
         Top             =   2700
         Width           =   1065
         _ExtentX        =   1323
         _ExtentY        =   476
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "Benaccom0.frx":008C
         Text            =   "0"
         AllowEmpty      =   0   'False
         TXTAlign        =   2
      End
      Begin atc2valtext.ValText TxtBx 
         Height          =   315
         Index           =   19
         Left            =   6465
         TabIndex        =   9
         Tag             =   "free,font"
         Top             =   1365
         Width           =   1065
         _ExtentX        =   1323
         _ExtentY        =   476
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "Benaccom0.frx":00A8
         Text            =   "0"
         Maximum         =   "100"
         Minimum         =   "0"
         AllowEmpty      =   0   'False
         TXTAlign        =   2
      End
      Begin atc2valtext.ValText TxtBx 
         Height          =   315
         Index           =   15
         Left            =   3105
         TabIndex        =   4
         Tag             =   "free,font"
         Top             =   2000
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "Benaccom0.frx":00C4
         Text            =   "0"
         AllowEmpty      =   0   'False
         TXTAlign        =   2
      End
      Begin atc2valtext.ValText TxtBx 
         Height          =   315
         Index           =   8
         Left            =   3870
         TabIndex        =   3
         Tag             =   "free,font"
         Top             =   1300
         Width           =   285
         _ExtentX        =   503
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "Benaccom0.frx":00E0
         Text            =   "0"
         AllowEmpty      =   0   'False
         TXTAlign        =   2
      End
      Begin atc2valtext.ValText TxtBx 
         Height          =   315
         Index           =   2
         Left            =   765
         TabIndex        =   1
         Tag             =   "free,font"
         Top             =   600
         Width           =   7005
         _ExtentX        =   12356
         _ExtentY        =   556
         BackColor       =   255
         ForeColor       =   128
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
         MouseIcon       =   "Benaccom0.frx":00FC
         Text            =   ""
         TypeOfData      =   3
         AllowEmpty      =   0   'False
      End
      Begin atc2valtext.ValText TxtBx 
         Height          =   315
         Index           =   18
         Left            =   -73290
         TabIndex        =   16
         Tag             =   "free,font"
         Top             =   2300
         Width           =   1065
         _ExtentX        =   1323
         _ExtentY        =   476
         BackColor       =   255
         ForeColor       =   128
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "Benaccom0.frx":0118
         Text            =   ""
         TypeOfData      =   2
         AllowEmpty      =   0   'False
      End
      Begin atc2valtext.ValText TxtBx 
         Height          =   315
         Index           =   11
         Left            =   -73290
         TabIndex        =   14
         Tag             =   "free,font"
         Top             =   1500
         Width           =   1065
         _ExtentX        =   1323
         _ExtentY        =   476
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "Benaccom0.frx":0134
         Text            =   "0"
         AllowEmpty      =   0   'False
         TXTAlign        =   2
      End
      Begin atc2valtext.ValText TxtBx 
         Height          =   315
         Index           =   9
         Left            =   -73290
         TabIndex        =   13
         Tag             =   "free,font"
         Top             =   1100
         Width           =   1065
         _ExtentX        =   1323
         _ExtentY        =   476
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "Benaccom0.frx":0150
         Text            =   "0"
         AllowEmpty      =   0   'False
         TXTAlign        =   2
      End
      Begin atc2valtext.ValText TxtBx 
         Height          =   315
         Index           =   17
         Left            =   -68220
         TabIndex        =   19
         Tag             =   "free,font"
         Top             =   2300
         Width           =   1065
         _ExtentX        =   1323
         _ExtentY        =   476
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "Benaccom0.frx":016C
         Text            =   "0"
         AllowEmpty      =   0   'False
         TXTAlign        =   2
      End
      Begin atc2valtext.ValText TxtBx 
         Height          =   315
         Index           =   14
         Left            =   -73290
         TabIndex        =   15
         Tag             =   "free,font"
         Top             =   1900
         Width           =   1065
         _ExtentX        =   1323
         _ExtentY        =   476
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "Benaccom0.frx":0188
         Text            =   "0"
         AllowEmpty      =   0   'False
         TXTAlign        =   2
      End
      Begin atc2valtext.ValText TxtBx 
         Height          =   315
         Index           =   13
         Left            =   -68220
         TabIndex        =   18
         Tag             =   "free,font"
         Top             =   1900
         Width           =   1065
         _ExtentX        =   1323
         _ExtentY        =   476
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "Benaccom0.frx":01A4
         Text            =   "0"
         AllowEmpty      =   0   'False
         TXTAlign        =   2
      End
      Begin atc2valtext.ValText TxtBx 
         Height          =   315
         Index           =   3
         Left            =   -73305
         TabIndex        =   12
         Tag             =   "free,font"
         Top             =   700
         Width           =   1065
         _ExtentX        =   1323
         _ExtentY        =   476
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "Benaccom0.frx":01C0
         Text            =   "0"
         AllowEmpty      =   0   'False
         TXTAlign        =   2
      End
      Begin atc2valtext.ValText TxtBx 
         Height          =   315
         Index           =   12
         Left            =   -68175
         TabIndex        =   24
         Tag             =   "free,font"
         Top             =   2100
         Visible         =   0   'False
         Width           =   1150
         _ExtentX        =   2037
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "Benaccom0.frx":01DC
         Text            =   "0"
         AllowEmpty      =   0   'False
         TXTAlign        =   2
      End
      Begin atc2valtext.ValText TxtBx 
         Height          =   315
         Index           =   10
         Left            =   -68175
         TabIndex        =   23
         Tag             =   "free,font"
         Top             =   1750
         Visible         =   0   'False
         Width           =   1150
         _ExtentX        =   2037
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "Benaccom0.frx":01F8
         Text            =   "0"
         AllowEmpty      =   0   'False
         TXTAlign        =   2
      End
      Begin atc2valtext.ValText TxtBx 
         Height          =   315
         Index           =   0
         Left            =   -68175
         TabIndex        =   20
         Tag             =   "free,font"
         Top             =   700
         Visible         =   0   'False
         Width           =   1150
         _ExtentX        =   2037
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "Benaccom0.frx":0214
         Text            =   "0"
         AllowEmpty      =   0   'False
         TXTAlign        =   2
      End
      Begin atc2valtext.ValText TxtBx 
         Height          =   315
         Index           =   16
         Left            =   -68175
         TabIndex        =   25
         Tag             =   "free,font"
         Top             =   2450
         Visible         =   0   'False
         Width           =   1150
         _ExtentX        =   2037
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "Benaccom0.frx":0230
         Text            =   "0"
         AllowEmpty      =   0   'False
         TXTAlign        =   2
      End
      Begin atc2valtext.ValText TxtBx 
         Height          =   315
         Index           =   6
         Left            =   -68175
         TabIndex        =   22
         Tag             =   "free,font"
         Top             =   1400
         Visible         =   0   'False
         Width           =   1150
         _ExtentX        =   2037
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "Benaccom0.frx":024C
         Text            =   "0"
         AllowEmpty      =   0   'False
         TXTAlign        =   2
      End
      Begin atc2valtext.ValText TxtBx 
         Height          =   315
         Index           =   5
         Left            =   -68175
         TabIndex        =   21
         Tag             =   "free,font"
         Top             =   1050
         Visible         =   0   'False
         Width           =   1150
         _ExtentX        =   2037
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "Benaccom0.frx":0268
         Text            =   "0"
         AllowEmpty      =   0   'False
         TXTAlign        =   2
      End
      Begin VB.Label Lab 
         Caption         =   $"Benaccom0.frx":0284
         ForeColor       =   &H00800000&
         Height          =   600
         Index           =   0
         Left            =   -74880
         TabIndex        =   53
         Tag             =   "free,font"
         Top             =   840
         Width           =   6495
      End
      Begin VB.Label Lab 
         Caption         =   "Note: Any assets provided as part of the accommodation should be entered under section L"
         ForeColor       =   &H00800000&
         Height          =   480
         Index           =   30
         Left            =   -74880
         TabIndex        =   52
         Tag             =   "free,font"
         Top             =   1800
         Width           =   6495
      End
      Begin VB.Label Lab 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Other, ancillary expenses not covered in a) or b) above"
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
         Index           =   22
         Left            =   -74895
         TabIndex        =   51
         Tag             =   "free,font"
         Top             =   2100
         Width           =   3870
      End
      Begin VB.Label Lab 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Employee expenses reimbursed under a) and b) above, but not already included "
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
         Index           =   12
         Left            =   -74910
         TabIndex        =   50
         Tag             =   "free,font"
         Top             =   1400
         Width           =   5670
      End
      Begin VB.Label Lab 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Heating, lighting, cleaning"
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
         Index           =   4
         Left            =   -74910
         TabIndex        =   49
         Tag             =   "free,font"
         Top             =   705
         Width           =   1830
      End
      Begin VB.Label Lab 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "If the accommodation qualifies as 'job-related', s315, the employee's net earnings"
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
         Index           =   27
         Left            =   -74880
         TabIndex        =   48
         Tag             =   "free,font"
         Top             =   2445
         Width           =   5685
      End
      Begin VB.Label Lab 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Actual values"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   1
         Left            =   -68175
         TabIndex        =   47
         Tag             =   "free,font"
         Top             =   405
         Width           =   1170
      End
      Begin VB.Label Lab 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Repairs, maintenance, decoration"
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
         Index           =   9
         Left            =   -74910
         TabIndex        =   46
         Tag             =   "free,font"
         Top             =   1050
         Width           =   2385
      End
      Begin VB.Label Lab 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Costs made good by employee of employer's expenses under a) or b) above"
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
         Index           =   17
         Left            =   -74910
         TabIndex        =   45
         Tag             =   "free,font"
         Top             =   1750
         Width           =   5325
      End
      Begin VB.Label Lab 
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Improvements between FIRSTOCC and TAX_YR"
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
         Height          =   435
         Index           =   28
         Left            =   -71640
         TabIndex        =   44
         Tag             =   "free,font"
         Top             =   2300
         Width           =   3315
      End
      Begin VB.Label Lab 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Date first occupied"
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
         Index           =   29
         Left            =   -74820
         TabIndex        =   43
         Tag             =   "free,font"
         Top             =   2295
         Width           =   1335
      End
      Begin VB.Label Lab 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Market Value when first occupied"
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
         Index           =   23
         Left            =   -71610
         TabIndex        =   42
         Tag             =   "free,font"
         Top             =   1900
         Width           =   2370
      End
      Begin VB.Label Lab 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Payment for tenancy"
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
         Index           =   24
         Left            =   -74820
         TabIndex        =   41
         Tag             =   "free,font"
         Top             =   1900
         Width           =   1455
      End
      Begin VB.Label Lab 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Improvements"
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
         Index           =   14
         Left            =   -74820
         TabIndex        =   40
         Tag             =   "free,font"
         Top             =   1100
         Width           =   990
      End
      Begin VB.Label Lab 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "If so:"
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
         Index           =   18
         Left            =   -71760
         TabIndex        =   39
         Tag             =   "free,font"
         Top             =   1600
         Width           =   345
      End
      Begin VB.Label Lab 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Accommodation first occupied after 31 March 1983"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   7
         Left            =   -71775
         TabIndex        =   38
         Tag             =   "free,font"
         Top             =   600
         Width           =   4350
      End
      Begin VB.Label Lab 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Purchase cost"
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
         Index           =   8
         Left            =   -74820
         TabIndex        =   37
         Tag             =   "free,font"
         Top             =   700
         Width           =   1020
      End
      Begin VB.Label Lab 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Capital contribution"
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
         Index           =   19
         Left            =   -74820
         TabIndex        =   36
         Tag             =   "free,font"
         Top             =   1500
         Width           =   1350
      End
      Begin VB.Label Lab 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Business use (percentage)"
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
         Index           =   32
         Left            =   4365
         TabIndex        =   35
         Tag             =   "free,font"
         Top             =   1395
         Width           =   1875
      End
      Begin VB.Label Lab 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Actual amount made good, or amount subjected to PAYE"
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
         Height          =   390
         Index           =   31
         Left            =   90
         TabIndex        =   34
         Tag             =   "free,font"
         Top             =   2620
         Width           =   2805
         WordWrap        =   -1  'True
      End
      Begin VB.Label Lab 
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Higher of rent paid by employer or annual value of the property"
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
         Height          =   405
         Index           =   25
         Left            =   90
         TabIndex        =   33
         Tag             =   "free,font"
         Top             =   1920
         Width           =   2655
      End
      Begin VB.Label Lab 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
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
         Index           =   5
         Left            =   90
         TabIndex        =   32
         Tag             =   "free,font"
         Top             =   600
         Width           =   810
      End
      Begin VB.Label Lab 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Basic charge"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   21
         Left            =   45
         TabIndex        =   31
         Tag             =   "free,font"
         Top             =   1695
         Width           =   1125
      End
      Begin VB.Label Lab 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Number of other employees sharing this residence"
         DataSource      =   "DB"
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
         Index           =   15
         Left            =   90
         TabIndex        =   30
         Tag             =   "free,font"
         Top             =   1300
         Width           =   3510
      End
   End
   Begin MSComctlLib.ListView LB 
      Height          =   2085
      Left            =   45
      TabIndex        =   0
      Tag             =   "free,font"
      Top             =   45
      Width           =   8145
      _ExtentX        =   14367
      _ExtentY        =   3678
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
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
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Accommodation Reference"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   1
         Text            =   "Benefit"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "F_Accommodation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Implements IBenefitForm2
Implements IFrmGeneral

Public benefit As IBenefitClass
Private m_InvalidVT As Control

Private mclsResize As New clsFormResize

Private Const L_DES_HEIGHT  As Long = 6090
Private Const L_DES_WIDTH  As Long = 8445
Private ml_CurrentTab As Long 'used to store the current tab

Private Sub ChkBx_Click(Index As Integer)
  Call IFrmGeneral_CheckChanged(ChkBx(Index))
End Sub

Private Sub cmdN1A_Click(Index As Integer)
  
  'EK removal of accomodation expenses TTP#28
Select Case Index
  Case 0
    Call BenScreenSwitch(BC_CLASS_1A_M)
  Case 1
    Call BenScreenSwitch(BC_NON_CLASS_1A_M)
  Case 2
    Call BenScreenSwitch(BC_ASSETSATDISPOSAL_L)
End Select

End Sub

Private Sub Form_Load()
  If Not (mclsResize.InitResize(Me, L_DES_HEIGHT, L_DES_WIDTH, DESIGN, , , MDIMain)) Then
    Err.Raise ERR_Application
  End If

  Call SetDefaultVTDate(TxtBx(1))
  Call SetDefaultVTDate(TxtBx(4))
  
  'EK removing section O accomodation beenfits, TTP#28
  Lab(1).Visible = False
  Lab(4).Visible = False
  Lab(9).Visible = False
  Lab(12).Visible = False
  Lab(17).Visible = False
  Lab(22).Visible = False
  Lab(27).Visible = False
  TxtBx(0).Enabled = False
  TxtBx(5).Enabled = False
  TxtBx(6).Enabled = False
  TxtBx(10).Enabled = False
  TxtBx(12).Enabled = False
  TxtBx(16).Enabled = False

  Me.tabAccom.tab = 0
End Sub

Private Sub Form_Resize()
  Call mclsResize.Resize
  Call ColumnWidths(LB, 50, 25, 25)
End Sub
Private Function SixYearaccommodationToScreen(b As Boolean) As Boolean

  On Error GoTo SixYearaccommodationToScreen_ERR
  Call xSet("SixYearaccommodationToScreen")
  
  TxtBx(13).Enabled = b
  TxtBx(17).Enabled = b

  SixYearaccommodationToScreen = True
  
SixYearaccommodationToScreen_END:
  Call xReturn("SixYearaccommodationToScreen")
  Exit Function
SixYearaccommodationToScreen_ERR:
  SixYearaccommodationToScreen = False
  Call ErrorMessage(ERR_ERROR, Err, "SixYearaccommodationToScreen", "Six Year accommodation To Screen", "Unable to set the controls for the six year accommodation check.")
  Resume SixYearaccommodationToScreen_END

End Function

Private Sub IBenefitForm2_AddBenefit()
  Dim ben As Accommodation
  
  On Error GoTo AddBenefit_Err

  Call xSet("AddBenefit")
  
  Set ben = New Accommodation
  Call AddBenefitHelper(Me, ben)
  

AddBenefit_End:
  Set ben = Nothing
  Call xReturn("AddBenefit")
  Exit Sub
AddBenefit_Err:
  Call ErrorMessage(ERR_ERROR, Err, "AddBenefit", "ERR_ADDBENEFIT", "Error in AddBenefit function, called from the form " & Me.Name & ".")
  Resume AddBenefit_End
  Resume
End Sub

Private Function IBenefitForm2_AddBenefitSetDefaults(ben As IBenefitClass) As Long
  With ben
    Call StandardReadData(ben)
    .value(accom_item_db) = "Please enter address"
    .value(accom_rent_db) = 0
    .value(accom_Business_db) = 0
'MP DB (not used)             .value(accom_Unavailable) = 0
'MP DB - removed accom_RelevantDays_db as in ReadDB but not in use elsewhere
'   .value(accom_RelevantDays_db) = 0
    .value(accom_ConsiderationForUse_db) = 0
    .value(accom_Price_db) = 0
    .value(accom_Improv_db) = 0
    .value(accom_CapContrib_db) = 0
    .value(accom_Tenancy_db) = 0
    .value(accom_MVFirstOcc_db) = 0
    .value(accom_RecentImprov_db) = 0
'MP DB    .value(accom_Assets) = 0
'MP DB    .value(accom_AvailDays_db) = 0
    .value(accom_nemployees_db) = 0
    
  'EK removal of accomodation expenses TTP#28
'    .value(accom_Utilities) = 0
'    .value(accom_Repairs) = 0
'    .value(accom_Reimbursements) = 0
'    .value(accom_NetEmoluments) = 0
'    .value(accom_Expenses_MadeGood) = 0
'    .value(accom_Ancillary) = 0
    
    Call SetAvaialbleRange(ben, ben.Parent, accom_availablefrom_db, accom_availableto_db)
    .value(accom_FirstOcc_db) = p11d32.Rates.value(accomFirstOccupiedDef)
  End With
End Function

Private Property Set IBenefitForm2_benefit(NewValue As IBenefitClass)
  Set benefit = NewValue
End Property

Private Property Get IBenefitForm2_benefit() As IBenefitClass
  Set IBenefitForm2_benefit = benefit
End Property

Private Function IBenefitForm2_BenefitFormState(ByVal fState As BENEFIT_FORM_STATE) As Boolean

  On Error GoTo BenefitFormState_err
  Call xSet("IBenefitForm2_EnableBenefitForm")
  
  If (fState = FORM_ENABLED) Or (fState = FORM_CDB) Then
    If fState = FORM_ENABLED Then
      tabAccom.Enabled = True
    Else
      ECASE ("Car CBD?") 'CAD
    End If
    Call MDIMain.SetDelete
  ElseIf fState = FORM_DISABLED Then
    Set benefit = Nothing
    tabAccom.Enabled = False
    Call MDIMain.ClearDelete
    Call MDIMain.ClearConfirmUndo
  End If
  
  IBenefitForm2_BenefitFormState = True
    
BenefitFormState_end:
  Call xReturn("BenefitFormState")
  Exit Function
  
BenefitFormState_err:
  IBenefitForm2_BenefitFormState = False
  Call ErrorMessage(ERR_ERROR, Err, "BenefitFormState", "Benefit Form State", "Error setting the benefit form state.")
  Resume BenefitFormState_end
End Function

Private Function IBenefitForm2_BenefitOff() As Boolean
    TxtBx(1).Text = ""
    TxtBx(2).Text = ""
    TxtBx(3).Text = ""
    TxtBx(4).Text = ""
    TxtBx(8).Text = ""
    TxtBx(9).Text = ""
    TxtBx(11).Text = ""
    TxtBx(14).Text = ""
    TxtBx(18).Text = ""
    Lab(28).Caption = ""
    TxtBx(13).Text = ""
    TxtBx(17).Text = ""
    TxtBx(15).Text = ""
    TxtBx(19).Text = ""
    TxtBx(20).Text = ""
    
  
    ChkBx(0) = vbUnchecked
    ChkBx(1) = vbUnchecked
    ChkBx(2) = vbUnchecked
    ChkBx(4) = vbUnchecked

End Function

Private Function IBenefitForm2_BenefitOn() As Boolean
  TxtBx(2).Text = benefit.value(accom_item_db)
  TxtBx(3).Text = benefit.value(accom_Price_db)
  TxtBx(1).Text = DateValReadToScreen(benefit.value(accom_availablefrom_db))
  TxtBx(4).Text = DateValReadToScreen(benefit.value(accom_availableto_db))
  TxtBx(8).Text = benefit.value(accom_nemployees_db)
  TxtBx(9).Text = benefit.value(accom_Improv_db)
  TxtBx(11).Text = benefit.value(accom_CapContrib_db)
  TxtBx(14).Text = benefit.value(accom_Tenancy_db)
  TxtBx(18).Text = DateValReadToScreen(benefit.value(accom_FirstOcc_db))
  Lab(28).Caption = "Improvements between " & benefit.value(accom_FirstOcc_db) & " and " & p11d32.Rates.value(TaxYearStart)
  TxtBx(13).Text = benefit.value(accom_MVFirstOcc_db)
  TxtBx(17).Text = benefit.value(accom_RecentImprov_db)
  TxtBx(15).Text = benefit.value(accom_rent_db)
  TxtBx(19).Text = benefit.value(accom_Business_db)
  TxtBx(20).Text = benefit.value(accom_ConsiderationForUse_db)
  
  ' EK remvoing accomodation benefits for section O, TTP#28
  ' TxtBx(0).Text = benefit.value(accom_Utilities)
  ' TxtBx(5).Text = benefit.value(accom_Repairs)
  ' TxtBx(6).Text = benefit.value(accom_Reimbursements)
  ' TxtBx(10).Text = benefit.value(accom_Expenses_MadeGood)
  ' TxtBx(12).Text = benefit.value(accom_Ancillary)
  ' TxtBx(16).Text = benefit.value(accom_NetEmoluments)
  
  ChkBx(0) = IIf(benefit.value(accom_erown_db), vbChecked, vbUnchecked)
  ChkBx(1) = IIf(benefit.value(accom_JobRelated_db), vbChecked, vbUnchecked)
  ChkBx(2) = IIf(benefit.value(accom_SixYear_db), vbChecked, vbUnchecked)
  ChkBx(3) = BoolToChkBox(benefit.value(ITEM_MADEGOOD_IS_TAXDEDUCTED))
  tabAccom.TabEnabled(1) = benefit.value(accom_erown_db)
    
  Call SixYearaccommodationToScreen(benefit.value(accom_SixYear_db))
  
  ChkBx(4) = IIf(benefit.value(accom_isrent_db), vbChecked, vbUnchecked)
  fmeApportion.Enabled = Not benefit.value(accom_isrent_db)

End Function

Private Function IBenefitForm2_BenefitsToListView() As Long
  IBenefitForm2_BenefitsToListView = BenefitsToListView(Me)
End Function

Private Function IBenefitForm2_BenefitToListView(ben As IBenefitClass, ByVal lBenefitIndex As Long) As Long
  IBenefitForm2_BenefitToListView = BenefitToListView(ben, Me, lBenefitIndex)
End Function

Private Function IBenefitForm2_BenefitToScreen(Optional ByVal BenefitIndex As Long = -1&, Optional ByVal UpdateBenefit As Boolean = True) As Boolean
  IBenefitForm2_BenefitToScreen = BenefitToScreenHelper(Me, BenefitIndex, UpdateBenefit)
End Function

Private Property Let IBenefitForm2_benclass(ByVal NewValue As BEN_CLASS)
End Property

Private Property Get IBenefitForm2_benclass() As BEN_CLASS
  IBenefitForm2_benclass = BC_LIVING_ACCOMMODATION_D
End Property

'Private Property Get IBenefitForm2_ControlDefault() As Control
'  Set IBenefitForm2_ControlDefault = TxtBx(2)
'End Property

Private Property Get IBenefitForm2_lv() As MSComctlLib.IListView
  Set IBenefitForm2_lv = LB
End Property

Private Function IBenefitForm2_RemoveBenefit(ByVal BenefitIndex As Long) As Boolean
  
  On Error GoTo RemoveBenefit_ERR
  Call xSet("RemoveBenefit")
  
  IBenefitForm2_RemoveBenefit = p11d32.CurrentEmployer.CurrentEmployee.RemoveBenefit(Me, benefit, BenefitIndex)
RemoveBenefit_END:
  Call xReturn("RemoveBenefit")
  Exit Function
RemoveBenefit_ERR:
  IBenefitForm2_RemoveBenefit = False
  Call ErrorMessage(ERR_ERROR, Err, "RemoveBenefit", "Remove Benefit", "Error removing the selected benefit.")
  Resume RemoveBenefit_END
End Function

Private Function IBenefitForm2_UpdateBenefitListViewItem(li As MSComctlLib.IListItem, benefit As IBenefitClass, Optional ByVal BenefitIndex As Long = 0&, Optional ByVal SelectItem As Boolean = False) As Long
  IBenefitForm2_UpdateBenefitListViewItem = UpdateBenefitListViewItem(li, benefit, BenefitIndex, SelectItem)
 
End Function

Private Function IBenefitForm2_ValididateBenefit(ben As IBenefitClass) As Boolean
  If ben.BenefitClass = BC_LIVING_ACCOMMODATION_D Then IBenefitForm2_ValididateBenefit = True
End Function

Private Function IFrmGeneral_CheckChanged(c As Control) As Boolean
  Dim lst As ListItem
  Dim i As Long
  Dim bDirty As Boolean
  
On Error GoTo CheckChanged_Err

  Call xSet("CheckChanged")
  
  If p11d32.CurrentEmployeeIsNothing Then GoTo CheckChanged_End
  If benefit Is Nothing Then GoTo CheckChanged_End
  
  With c
    Select Case .Name
      Case "TxtBx"
        Select Case .Index
          Case 1
            bDirty = CheckTextInput(.Text, benefit, accom_availablefrom_db)
          Case 2
            bDirty = CheckTextInput(.Text, benefit, accom_item_db)
          Case 3
            bDirty = CheckTextInput(.Text, benefit, accom_Price_db)
          Case 4
            bDirty = CheckTextInput(.Text, benefit, accom_availableto_db)
          Case 8
            bDirty = CheckTextInput(.Text, benefit, accom_nemployees_db)
          Case 9
            bDirty = CheckTextInput(.Text, benefit, accom_Improv_db)
          Case 11
            bDirty = CheckTextInput(.Text, benefit, accom_CapContrib_db)
          Case 13
            bDirty = CheckTextInput(.Text, benefit, accom_MVFirstOcc_db)
          Case 14
            bDirty = CheckTextInput(.Text, benefit, accom_Tenancy_db)
          Case 15
            bDirty = CheckTextInput(.Text, benefit, accom_rent_db)
          Case 17
            bDirty = CheckTextInput(.Text, benefit, accom_RecentImprov_db)
          Case 18
            bDirty = CheckTextInput(.Text, benefit, accom_FirstOcc_db)
            If bDirty Then Lab(28).Caption = "Improvements between " & benefit.value(accom_FirstOcc_db) & " and " & p11d32.Rates.value(TaxYearStart)
          Case 19
            bDirty = CheckTextInput(.Text, benefit, accom_Business_db)
          Case 20
            bDirty = CheckTextInput(.Text, benefit, accom_ConsiderationForUse_db)
          'EK removal of accomodation expenses TTP#28
'          Case 0
'            bDirty = CheckTextInput(.Text, benefit, accom_Utilities)
'          Case 5
'            bDirty = CheckTextInput(.Text, benefit, accom_Repairs)
'          Case 6
'            bDirty = CheckTextInput(.Text, benefit, accom_Reimbursements)
'          Case 10
'            bDirty = CheckTextInput(.Text, benefit, accom_Expenses_MadeGood)
'          Case 12
'            bDirty = CheckTextInput(.Text, benefit, accom_Ancillary)
'          Case 16
'            bDirty = CheckTextInput(.Text, benefit, accom_NetEmoluments)
          Case Else
            ECASE "Unknown control"
            GoTo CheckChanged_End
        End Select
      Case "ChkBx"
        Select Case .Index
          Case 0
            bDirty = CheckCheckBoxInput(.value, benefit, accom_erown_db)
            tabAccom.TabEnabled(1) = benefit.value(accom_erown_db)
          Case 1
            bDirty = CheckCheckBoxInput(.value, benefit, accom_JobRelated_db)
          Case 2
            bDirty = CheckCheckBoxInput(.value, benefit, accom_SixYear_db)
            Call SixYearaccommodationToScreen(benefit.value(accom_SixYear_db))
          Case 3
            bDirty = CheckCheckBoxInput(.value, benefit, ITEM_MADEGOOD_IS_TAXDEDUCTED)
          Case 4
            bDirty = CheckCheckBoxInput(.value, benefit, accom_isrent_db)
            If bDirty Then
              fmeApportion.Enabled = Not benefit.value(accom_isrent_db)
            End If
          Case 5
          
          Case Else
            ECASE "Unknown control"
            GoTo CheckChanged_End
        End Select
      Case Else
        ECASE "Unknown control"
    End Select
  End With
  IFrmGeneral_CheckChanged = AfterCheckChanged(c, Me, bDirty)
      
CheckChanged_End:
  Call xReturn("CheckChanged")
  Exit Function

CheckChanged_Err:
  Call ErrorMessage(ERR_ERROR, Err, "CheckChanged", "ERR_CHECKCHANGED", "This function has failed for the form " & Me.Name & ".")
  Resume CheckChanged_End
End Function
Private Function EnableDates(b As Boolean) As Boolean
    fmeApportion.Enabled = b
End Function
Private Property Get IFrmGeneral_InvalidVT() As Control
  Set IFrmGeneral_InvalidVT = m_InvalidVT
End Property

Private Property Set IFrmGeneral_InvalidVT(NewValue As Control)
  Set m_InvalidVT = NewValue
End Property

Private Sub LB_ItemClick(ByVal Item As MSComctlLib.ListItem)
  Call SetLastListItemSelected(Item)
  If Not (LB.SelectedItem Is Nothing) Then
    Call IBenefitForm2_BenefitToScreen(Item.Tag)
  End If
End Sub

Private Sub lb_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
  Call SetSortOrder(LB, ColumnHeader)
End Sub

Private Sub LB_KeyDown(KeyCode As Integer, Shift As Integer)
  Call LVKeyDown(KeyCode, Shift)
End Sub

Private Sub SSTab1_DblClick()

End Sub

Private Sub TxtBx_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  TxtBx(Index).Tag = SetChanged
End Sub

Private Sub TxtBx_Validate(Index As Integer, Cancel As Boolean)
  Call IFrmGeneral_CheckChanged(TxtBx(Index))
End Sub
