VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{8D988532-0F0C-460C-B00E-7B5637E97680}#1.0#0"; "atc2vtext.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form F_CompanyCar 
   Caption         =   " f"
   ClientHeight    =   6540
   ClientLeft      =   1680
   ClientTop       =   3015
   ClientWidth     =   8685
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
   ScaleHeight     =   6540
   ScaleWidth      =   8685
   Tag             =   "FREE,FONT"
   Begin TabDlg.SSTab tab 
      Height          =   3885
      Left            =   3960
      TabIndex        =   39
      Tag             =   "FREE,FONT"
      Top             =   2280
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   6853
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      ForeColor       =   8388608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "General"
      TabPicture(0)   =   "Bencar.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "L_Data(9)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "L_Data(8)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "L_Data(7)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "L_Data(19)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label4(16)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lblAccessories"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label2(2)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "TB_DATA(13)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "TB_DATA(7)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "TB_DATA(9)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "TB_DATA(8)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "B_Make"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "TB_DATA(16)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "fraCO2"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "B_Acc"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).ControlCount=   15
      TabCaption(1)   =   "P46 Details"
      TabPicture(1)   =   "Bencar.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Op_Data(6)"
      Tab(1).Control(1)=   "CB_FuelType(0)"
      Tab(1).Control(2)=   "Op_Data(1)"
      Tab(1).Control(3)=   "Op_Data(5)"
      Tab(1).Control(4)=   "Op_Data(2)"
      Tab(1).Control(5)=   "CB_CARLIST"
      Tab(1).Control(6)=   "Op_Data(3)"
      Tab(1).Control(7)=   "Label2(1)"
      Tab(1).Control(8)=   "Label1"
      Tab(1).Control(9)=   "L_Data(18)"
      Tab(1).Control(10)=   "L_Data(17)"
      Tab(1).Control(11)=   "L_Data(16)"
      Tab(1).Control(12)=   "L_Data(11)"
      Tab(1).Control(13)=   "L_Data(5)"
      Tab(1).Control(14)=   "L_Data(14)"
      Tab(1).Control(15)=   "L_Data(15)"
      Tab(1).ControlCount=   16
      TabCaption(2)   =   "Fuel Benefit"
      TabPicture(2)   =   "Bencar.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Op_Data(12)"
      Tab(2).Control(1)=   "CB_FuelType(1)"
      Tab(2).Control(2)=   "fraFuelBenefit"
      Tab(2).Control(3)=   "Op_Data(10)"
      Tab(2).Control(4)=   "Op_Data(9)"
      Tab(2).Control(5)=   "Op_Data(8)"
      Tab(2).Control(6)=   "TB_DATA(14)"
      Tab(2).Control(7)=   "TB_DATA(17)"
      Tab(2).Control(8)=   "L_Data(6)"
      Tab(2).Control(9)=   "Label2(3)"
      Tab(2).Control(10)=   "lblFuelType"
      Tab(2).ControlCount=   11
      Begin VB.CheckBox Op_Data 
         Alignment       =   1  'Right Justify
         Caption         =   "Is the above an amount subjected to PAYE?"
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
         Left            =   -74940
         TabIndex        =   56
         Tag             =   "free,font"
         Top             =   950
         Width           =   4200
      End
      Begin VB.ComboBox CB_FuelType 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   315
         Index           =   1
         ItemData        =   "Bencar.frx":0054
         Left            =   -73455
         List            =   "Bencar.frx":0056
         Style           =   2  'Dropdown List
         TabIndex        =   57
         Tag             =   "free,font"
         Top             =   1200
         Width           =   2730
      End
      Begin VB.CommandButton B_Acc 
         Appearance      =   0  'Flat
         Caption         =   "&Accessories..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3015
         TabIndex        =   13
         Top             =   1395
         Width           =   1215
      End
      Begin VB.Frame fraCO2 
         BorderStyle     =   0  'None
         ClipControls    =   0   'False
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
         Height          =   1005
         Left            =   60
         TabIndex        =   61
         Top             =   1710
         Width           =   4305
         Begin VB.CheckBox Op_Data 
            Alignment       =   1  'Right Justify
            Caption         =   "Tick, if no approved carbon dioxide emissions figure"
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
            Index           =   11
            Left            =   1
            TabIndex        =   15
            Tag             =   "free,font"
            Top             =   360
            Width           =   4215
         End
         Begin atc2valtext.ValText TB_DATA 
            Height          =   285
            Index           =   12
            Left            =   3000
            TabIndex        =   14
            Tag             =   "free,font"
            Top             =   45
            Width           =   1215
            _ExtentX        =   2143
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
            MouseIcon       =   "Bencar.frx":0058
            Text            =   ""
            Minimum         =   "0"
            AllowEmpty      =   0   'False
            TXTAlign        =   2
         End
         Begin atc2valtext.ValText TB_DATA 
            Height          =   285
            Index           =   11
            Left            =   3000
            TabIndex        =   16
            Tag             =   "free,font"
            Top             =   675
            Width           =   1200
            _ExtentX        =   2117
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
            Text            =   "0"
            AllowEmpty      =   0   'False
            TXTAlign        =   2
         End
         Begin VB.Label Label2 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Engine cc"
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
            Index           =   0
            Left            =   0
            TabIndex        =   63
            Tag             =   "free,font"
            Top             =   675
            Width           =   720
         End
         Begin VB.Label lblEmissions 
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Carbon dioxide emissions (g/km)"
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
            Left            =   0
            TabIndex        =   62
            Tag             =   "free,font"
            Top             =   90
            Width           =   2400
         End
      End
      Begin VB.Frame fraFuelBenefit 
         Caption         =   "Fuel benefit availability (not electric only)"
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
         Height          =   1200
         Left            =   -74950
         TabIndex        =   59
         Top             =   1750
         Width           =   4215
         Begin VB.CheckBox Op_Data 
            Alignment       =   1  'Right Justify
            Caption         =   "Was fuel reinstated?"
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
            Left            =   200
            TabIndex        =   67
            Tag             =   "free,font"
            Top             =   540
            Width           =   3900
         End
         Begin VB.CheckBox Op_Data 
            Alignment       =   1  'Right Justify
            Caption         =   "Do the car's days unavailable also relate to fuel?"
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
            Index           =   7
            Left            =   200
            TabIndex        =   68
            Tag             =   "free,font"
            Top             =   805
            Width           =   3900
         End
         Begin atc2valtext.ValText TB_DATA 
            Height          =   285
            Index           =   15
            Left            =   3150
            TabIndex        =   66
            Tag             =   "FREE,FONT"
            Top             =   230
            Width           =   945
            _ExtentX        =   1667
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
            MouseIcon       =   "Bencar.frx":0074
            Text            =   ""
            TypeOfData      =   2
            Maximum         =   "5/4/2001"
            Minimum         =   "6/4/2000"
            AllowEmpty      =   0   'False
         End
         Begin VB.Label L_Data 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Date fuel benefit withdrawn"
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
            Left            =   200
            TabIndex        =   60
            Tag             =   "FREE,FONT"
            Top             =   260
            Width           =   1935
         End
      End
      Begin atc2valtext.ValText TB_DATA 
         Height          =   255
         Index           =   16
         Left            =   3870
         TabIndex        =   17
         Tag             =   "FREE,FONT"
         Top             =   2745
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   450
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
         Text            =   "0"
         Minimum         =   "1"
         AllowEmpty      =   0   'False
         TXTAlign        =   2
      End
      Begin VB.CommandButton B_Make 
         Appearance      =   0  'Flat
         Caption         =   "&Select Make"
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3000
         TabIndex        =   10
         Top             =   420
         Width           =   1215
      End
      Begin VB.CheckBox Op_Data 
         Alignment       =   1  'Right Justify
         Caption         =   "If so, was the total cost actually made good?"
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
         Height          =   240
         Index           =   10
         Left            =   -74940
         TabIndex        =   55
         Tag             =   "free,font"
         Top             =   700
         Width           =   4200
      End
      Begin VB.CheckBox Op_Data 
         Alignment       =   1  'Right Justify
         Caption         =   "If so, was the total cost required to be made good?"
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
         Left            =   -74940
         TabIndex        =   54
         Tag             =   "free,font"
         Top             =   500
         Width           =   4200
      End
      Begin VB.CheckBox Op_Data 
         Alignment       =   1  'Right Justify
         Caption         =   "Was fuel provided for private use?"
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
         Left            =   -74940
         TabIndex        =   53
         Tag             =   "free,font"
         Top             =   250
         Width           =   4200
      End
      Begin VB.CheckBox Op_Data 
         Alignment       =   1  'Right Justify
         Caption         =   "Tick, if no approved carbon dioxide emissions figure"
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
         Index           =   6
         Left            =   -74940
         TabIndex        =   25
         Tag             =   "free,font"
         Top             =   2720
         Width           =   4085
      End
      Begin VB.ComboBox CB_FuelType 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   315
         Index           =   0
         Left            =   -72840
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Tag             =   "free,font"
         Top             =   2030
         Width           =   2055
      End
      Begin VB.CheckBox Op_Data 
         Alignment       =   1  'Right Justify
         Caption         =   "Replaced"
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
         Left            =   -74940
         TabIndex        =   19
         Tag             =   "free,font"
         Top             =   360
         Width           =   1695
      End
      Begin VB.CheckBox Op_Data 
         Alignment       =   1  'Right Justify
         Caption         =   "Force print on P46"
         DataField       =   "ForceP46"
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
         Height          =   330
         Index           =   5
         Left            =   -72570
         TabIndex        =   22
         Tag             =   "free,font"
         Top             =   550
         Width           =   1815
      End
      Begin VB.CheckBox Op_Data 
         Alignment       =   1  'Right Justify
         Caption         =   "Second car"
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
         Left            =   -72570
         TabIndex        =   21
         Tag             =   "free,font"
         Top             =   350
         Width           =   1815
      End
      Begin VB.ComboBox CB_CARLIST 
         Appearance      =   0  'Flat
         DataField       =   "regreplaced"
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
         Height          =   315
         Left            =   -74910
         TabIndex        =   23
         Tag             =   "free,font"
         Text            =   "CB_CARLIST"
         Top             =   1125
         Width           =   2265
      End
      Begin VB.CheckBox Op_Data 
         Alignment       =   1  'Right Justify
         Caption         =   "Replacement"
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
         Index           =   3
         Left            =   -74940
         TabIndex        =   20
         Tag             =   "free,font"
         Top             =   550
         Width           =   1695
      End
      Begin atc2valtext.ValText TB_DATA 
         Height          =   285
         Index           =   8
         Left            =   1035
         TabIndex        =   12
         Tag             =   "free,font"
         Top             =   1085
         Width           =   1575
         _ExtentX        =   2778
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
         MouseIcon       =   "Bencar.frx":0090
         Text            =   "1/8/95"
         TypeOfData      =   4
         AllowEmpty      =   0   'False
      End
      Begin atc2valtext.ValText TB_DATA 
         Height          =   285
         Index           =   9
         Left            =   645
         TabIndex        =   11
         Tag             =   "free,font"
         Top             =   745
         Width           =   2300
         _ExtentX        =   4048
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
         MaxLength       =   50
         MouseIcon       =   "Bencar.frx":00AC
         Text            =   ""
         TypeOfData      =   3
         InvalidColor    =   -2147483643
      End
      Begin atc2valtext.ValText TB_DATA 
         Height          =   285
         Index           =   7
         Left            =   645
         TabIndex        =   9
         Tag             =   "free,font"
         Top             =   405
         Width           =   2300
         _ExtentX        =   4048
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
         MouseIcon       =   "Bencar.frx":00C8
         Text            =   ""
         TypeOfData      =   3
         AllowEmpty      =   0   'False
      End
      Begin atc2valtext.ValText TB_DATA 
         Height          =   315
         Index           =   10
         Left            =   -72360
         TabIndex        =   27
         Tag             =   "free,font"
         Top             =   2390
         Width           =   1565
         _ExtentX        =   2752
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
         MaxLength       =   50
         MouseIcon       =   "Bencar.frx":00E4
         Text            =   ""
         TXTAlign        =   2
      End
      Begin atc2valtext.ValText TB_DATA 
         Height          =   285
         Index           =   13
         Left            =   3015
         TabIndex        =   18
         Tag             =   "free,font"
         Top             =   3060
         Width           =   1215
         _ExtentX        =   2143
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
         MouseIcon       =   "Bencar.frx":0100
         Text            =   ""
         Minimum         =   "0"
         AllowEmpty      =   0   'False
         TXTAlign        =   2
      End
      Begin atc2valtext.ValText TB_DATA 
         Height          =   285
         Index           =   14
         Left            =   -71990
         TabIndex        =   70
         Tag             =   "free,font"
         Top             =   3000
         Width           =   1215
         _ExtentX        =   2143
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
         MouseIcon       =   "Bencar.frx":011C
         Text            =   ""
         Minimum         =   "0"
         AllowEmpty      =   0   'False
         TXTAlign        =   2
      End
      Begin atc2valtext.ValText TB_DATA 
         Height          =   285
         Index           =   17
         Left            =   -71800
         TabIndex        =   69
         Tag             =   "FREE,FONT"
         Top             =   1500
         Width           =   1065
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
         MouseIcon       =   "Bencar.frx":0138
         Text            =   "0"
         Minimum         =   "1"
         AllowEmpty      =   0   'False
         TXTAlign        =   2
      End
      Begin VB.Label L_Data 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Electric range (miles)"
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
         Index           =   6
         Left            =   -74940
         TabIndex        =   71
         Tag             =   "FREE,FONT"
         Top             =   1500
         Width           =   1455
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "OpRA amount forgone"
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
         Left            =   -74940
         TabIndex        =   65
         Tag             =   "free,font"
         Top             =   3000
         Width           =   1590
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "OpRA amount forgone"
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
         Left            =   60
         TabIndex        =   64
         Tag             =   "free,font"
         Top             =   3060
         Width           =   1590
      End
      Begin VB.Label lblAccessories 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1020
         TabIndex        =   40
         Tag             =   "free,font"
         Top             =   1395
         Width           =   1575
      End
      Begin VB.Label Label4 
         Caption         =   "No of employees using this car (ESC A71)"
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
         Height          =   255
         Index           =   16
         Left            =   45
         TabIndex        =   58
         Tag             =   "free,font"
         Top             =   2745
         Width           =   3855
      End
      Begin VB.Label lblFuelType 
         Caption         =   "Fuel or power used"
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
         Height          =   255
         Left            =   -74940
         TabIndex        =   52
         Tag             =   "free,font"
         Top             =   1200
         Width           =   1725
      End
      Begin VB.Label Label2 
         Caption         =   "Carbon dioxide emissions (gm/km)"
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
         Height          =   255
         Index           =   1
         Left            =   -74940
         TabIndex        =   51
         Tag             =   "free,font"
         Top             =   2420
         Width           =   2415
      End
      Begin VB.Label Label1 
         Caption         =   "Type of fuel or power used"
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
         Height          =   255
         Left            =   -74940
         TabIndex        =   50
         Tag             =   "free,font"
         Top             =   2055
         Width           =   2415
      End
      Begin VB.Label L_Data 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Replacement registration"
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
         Left            =   -72570
         TabIndex        =   48
         Tag             =   "free,font"
         Top             =   900
         Width           =   1755
      End
      Begin VB.Label L_Data 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Car Replaced"
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
         Index           =   17
         Left            =   -72525
         TabIndex        =   26
         Tag             =   "FREE,FONT"
         Top             =   1100
         Width           =   1770
      End
      Begin VB.Label L_Data 
         Appearance      =   0  'Flat
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
         Index           =   16
         Left            =   -72030
         TabIndex        =   47
         Tag             =   "free,font"
         Top             =   1450
         Width           =   825
      End
      Begin VB.Label L_Data 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Replacement make and model"
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
         Index           =   11
         Left            =   -74940
         TabIndex        =   46
         Tag             =   "free,font"
         Top             =   1455
         Width           =   2160
      End
      Begin VB.Label L_Data 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Replacements available"
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
         Left            =   -74940
         TabIndex        =   45
         Tag             =   "free,font"
         Top             =   900
         Width           =   1695
      End
      Begin VB.Label L_Data 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Replacement Make and Model"
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
         Index           =   14
         Left            =   -74940
         TabIndex        =   28
         Tag             =   "FREE,FONT"
         Top             =   1650
         Width           =   2745
      End
      Begin VB.Label L_Data 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Available To"
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
         Index           =   15
         Left            =   -72030
         TabIndex        =   29
         Tag             =   "FREE,FONT"
         Top             =   1650
         Width           =   1275
      End
      Begin VB.Label L_Data 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Model"
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
         Index           =   19
         Left            =   90
         TabIndex        =   44
         Tag             =   "free,font"
         Top             =   750
         Width           =   435
      End
      Begin VB.Label L_Data 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Make"
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
         Index           =   7
         Left            =   90
         TabIndex        =   43
         Tag             =   "free,font"
         Top             =   405
         Width           =   615
      End
      Begin VB.Label L_Data 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Registered"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   8
         Left            =   90
         TabIndex        =   42
         Tag             =   "free,font"
         Top             =   1080
         Width           =   780
      End
      Begin VB.Label L_Data 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Accessories"
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
         Index           =   9
         Left            =   90
         TabIndex        =   41
         Tag             =   "free,font"
         Top             =   1395
         Width           =   975
      End
   End
   Begin MSComctlLib.ListView LB 
      Height          =   2085
      Left            =   90
      TabIndex        =   72
      Tag             =   "free,font"
      Top             =   90
      Width           =   8520
      _ExtentX        =   15028
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
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Car Reference"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Available From"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Available To"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Car Benefit"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Fuel Benefit"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Frame fmeInput 
      Height          =   4215
      Left            =   120
      TabIndex        =   37
      Top             =   2160
      Width           =   8460
      Begin VB.CheckBox Op_Data 
         Alignment       =   1  'Right Justify
         Caption         =   "Is the above an amount subjected to PAYE?"
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
         Height          =   345
         Index           =   4
         Left            =   90
         TabIndex        =   7
         Tag             =   "free,font"
         Top             =   2565
         Width           =   3470
      End
      Begin VB.ComboBox CB_Frequency 
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
         ForeColor       =   &H00000080&
         Height          =   315
         ItemData        =   "Bencar.frx":0154
         Left            =   2520
         List            =   "Bencar.frx":0156
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Tag             =   "free,font"
         Top             =   2925
         Width           =   1065
      End
      Begin atc2valtext.ValText TB_DATA 
         Height          =   285
         Index           =   1
         Left            =   2520
         TabIndex        =   1
         Tag             =   "FREE,FONT"
         Top             =   630
         Width           =   1065
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
         MaxLength       =   7
         MouseIcon       =   "Bencar.frx":0158
         Text            =   "0"
         Minimum         =   "0"
         AllowEmpty      =   0   'False
         TXTAlign        =   2
      End
      Begin atc2valtext.ValText TB_DATA 
         Height          =   285
         Index           =   2
         Left            =   2520
         TabIndex        =   2
         Tag             =   "FREE,FONT"
         Top             =   945
         Width           =   1065
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
         MouseIcon       =   "Bencar.frx":0174
         Text            =   "6/4/2000"
         TypeOfData      =   4
         AllowEmpty      =   0   'False
      End
      Begin atc2valtext.ValText TB_DATA 
         Height          =   285
         Index           =   3
         Left            =   2520
         TabIndex        =   3
         Tag             =   "FREE,FONT"
         Top             =   1275
         Width           =   1065
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
         MouseIcon       =   "Bencar.frx":0190
         Text            =   "5/4/2001"
         TypeOfData      =   2
         Maximum         =   "5/4/2001"
         Minimum         =   "6/4/2000"
         AllowEmpty      =   0   'False
      End
      Begin atc2valtext.ValText TB_DATA 
         Height          =   285
         Index           =   4
         Left            =   2520
         TabIndex        =   4
         Tag             =   "FREE,FONT"
         Top             =   1590
         Width           =   1065
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
         MouseIcon       =   "Bencar.frx":01AC
         Text            =   "0"
         Maximum         =   "365"
         Minimum         =   "0"
         AllowEmpty      =   0   'False
         TXTAlign        =   2
      End
      Begin atc2valtext.ValText TB_DATA 
         Height          =   285
         Index           =   5
         Left            =   2520
         TabIndex        =   5
         Tag             =   "FREE,FONT"
         Top             =   1920
         Width           =   1065
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
         MaxLength       =   6
         MouseIcon       =   "Bencar.frx":01C8
         Text            =   "0"
         Minimum         =   "0"
         AllowEmpty      =   0   'False
         TXTAlign        =   2
      End
      Begin atc2valtext.ValText TB_DATA 
         Height          =   285
         Index           =   6
         Left            =   2520
         TabIndex        =   6
         Tag             =   "FREE,FONT"
         Top             =   2235
         Width           =   1065
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
         MaxLength       =   6
         MouseIcon       =   "Bencar.frx":01E4
         Text            =   "0"
         Minimum         =   "0"
         AllowEmpty      =   0   'False
         TXTAlign        =   2
      End
      Begin VB.Frame fmeTAB 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2415
         Index           =   1
         Left            =   3600
         TabIndex        =   38
         Top             =   585
         Visible         =   0   'False
         Width           =   4455
      End
      Begin atc2valtext.ValText TB_DATA 
         Height          =   315
         Index           =   0
         Left            =   1425
         TabIndex        =   0
         Tag             =   "free,font"
         Top             =   240
         Width           =   2160
         _ExtentX        =   3810
         _ExtentY        =   556
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
         MaxLength       =   30
         MouseIcon       =   "Bencar.frx":0200
         Text            =   ""
         TypeOfData      =   3
         AllowEmpty      =   0   'False
      End
      Begin VB.Label L_Data 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Frequency of contributions towards private running costs"
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
         Index           =   20
         Left            =   90
         TabIndex        =   49
         Tag             =   "free,font"
         Top             =   2880
         Width           =   2445
         WordWrap        =   -1  'True
      End
      Begin VB.Label L_Data 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Registration"
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
         Left            =   120
         TabIndex        =   30
         Tag             =   "FREE,FONT"
         Top             =   315
         Width           =   975
      End
      Begin VB.Label L_Data 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "List price "
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
         Left            =   120
         TabIndex        =   31
         Tag             =   "FREE,FONT"
         Top             =   645
         Width           =   675
      End
      Begin VB.Label L_Data 
         Appearance      =   0  'Flat
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
         Index           =   2
         Left            =   120
         TabIndex        =   32
         Tag             =   "FREE,FONT"
         Top             =   960
         Width           =   1350
      End
      Begin VB.Label L_Data 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Days unavailable"
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
         Left            =   120
         TabIndex        =   34
         Tag             =   "FREE,FONT"
         Top             =   1605
         Width           =   1470
      End
      Begin VB.Label L_Data 
         Appearance      =   0  'Flat
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
         Index           =   3
         Left            =   120
         TabIndex        =   33
         Tag             =   "FREE,FONT"
         Top             =   1275
         Width           =   1350
      End
      Begin VB.Label L_Data 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Capital contribution towards cost"
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
         Left            =   120
         TabIndex        =   35
         Tag             =   "FREE,FONT"
         Top             =   1935
         Width           =   2460
         WordWrap        =   -1  'True
      End
      Begin VB.Label L_Data 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Contribution towards private running costs/amount subjected to PAYE"
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
         Height          =   585
         Index           =   13
         Left            =   120
         TabIndex        =   36
         Tag             =   "FREE,FONT"
         Top             =   2175
         Width           =   2355
         WordWrap        =   -1  'True
      End
   End
End
Attribute VB_Name = "F_CompanyCar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Implements IBenefitForm2
Implements IFrmGeneral

Public benefit As IBenefitClass

Private mclsResize As New clsFormResize
Private m_InvalidVT As Control

Private Const L_DES_HEIGHT  As Long = 6090
Private Const L_DES_WIDTH  As Long = 8445

'EK 1/04 TTP #8 save list price when click on Make

Public Enum LV_CAR_SUBITEMS
  LV_CAR_AVAILABLE_FROM = 1
  LV_CAR_AVAILABLE_TO = 2
  LV_CAR_BENEFIT = 3
  LV_CAR_FUEL_BENEFIT = 4
  'LV_CAR_AVAIALABELFROM_SORT = 5
  'LV_CAR_AVAIALABELTO_SORT = 6
End Enum

Private Const L_REGISTRATION_DATE_INDEX = 8
Private Const L_AVAILABLE_FROM_INDEX = 2

Private Sub B_Make_Click()
  Call DialogToScreen(F_CompanyCarCO2Emissions, Nothing, 0, Me, p11d32.CurrentEmployer.CurrentEmployee.benefits.ItemIndex(benefit), False)
  Call CO2StuffChanged
  
End Sub


Private Sub CB_CARLIST_Click()
  Call GetReplacementCar(CB_CARLIST.ListIndex)
End Sub

Private Sub CB_CARLIST_KeyPress(KeyAscii As Integer)
  'Check for return key - tab to next field
  If KeyAscii = 13 Then Call SendKeysEx(vbTab)
End Sub

Private Sub CB_CARLIST_Validate(Cancel As Boolean)
  Call IFrmGeneral_CheckChanged(CB_CARLIST)
End Sub

Private Sub CB_Frequency_KeyDown(KeyCode As Integer, Shift As Integer)
  CB_Frequency.Tag = SetChanged
End Sub
Private Sub CB_Frequency_Click()
  Call IFrmGeneral_CheckChanged(CB_Frequency)
End Sub

Private Sub CB_Frequency_KeyPress(KeyAscii As Integer)
  'Check for return key - tab to next field
  If KeyAscii = 13 Then Call SendKeysEx(vbTab)
End Sub

Private Sub CB_Frequency_Validate(Cancel As Boolean)
  Call IFrmGeneral_CheckChanged(CB_Frequency)
End Sub

'KM 26/02/02 - Decision was made that tick box should no longer be used / appear after v2000
'KM - added function to ensure the "Is this car diesel?" check box
'is checked when Diesel is selected in the FuelType combo box
Private Sub CB_FuelType_Click(Index As Integer)
    
    Call IFrmGeneral_CheckChanged(CB_FuelType(1))

End Sub

Private Sub CB_FuelType_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  CB_FuelType(0).Tag = SetChanged 'AM
  CB_FuelType(1).Tag = SetChanged
End Sub

Private Sub CB_FuelType_KeyPress(Index As Integer, KeyAscii As Integer)
  'Check for return key - tab to next field
  If KeyAscii = 13 Then Call SendKeysEx(vbTab)
End Sub

Private Sub CB_FuelType_LostFocus(Index As Integer)
  Call IFrmGeneral_CheckChanged(CB_FuelType(1))
End Sub

Private Sub CB_FuelType_Validate(Index As Integer, Cancel As Boolean)
  Call IFrmGeneral_CheckChanged(CB_FuelType(1))
End Sub

Private Sub Form_Resize()
  mclsResize.Resize
  Call ColumnWidths(LB, 20, 20, 20, 20, 20, 0, 0)
End Sub

Private Sub B_Acc_Click()
  Call DialogToScreen(F_CompanyCarAcc, lblAccessories, car_Accessories, Me, p11d32.CurrentEmployer.CurrentEmployee.benefits.ItemIndex(benefit))
End Sub

Private Function FuelTypeDescription(ccft As COMPANY_CAR_FUEL_TYPE) As String
  Dim p46s As P46_FUEL_TYPE_STRINGS
  p46s = P46FuelTypeStrings(ccft)
  If (Len(p46s.Letter)) > 0 Then
    FuelTypeDescription = "(" & p46s.Letter & ") " & p46s.Description
  Else
    FuelTypeDescription = p46s.Description
  End If
End Function
Private Sub Form_Load()
  If Not (mclsResize.InitResize(Me, L_DES_HEIGHT, L_DES_WIDTH, DESIGN, , , MDIMain)) Then
    Err.Raise ERR_Application
  End If

  CB_Frequency.AddItem (S_ANNUALLY) 'lisindex = 1
  CB_Frequency.AddItem (S_QUARTERLY)
  CB_Frequency.AddItem (S_MONTHLY)
  CB_Frequency.AddItem (S_WEEKLY)
  CB_Frequency.AddItem (S_ACTUAL) 'so lisindex = 4
  
  Call AddComboItemAndItemData(CB_FuelType(1), FuelTypeDescription(CCFT_NONE), CCFT_NONE)
  Call AddComboItemAndItemData(CB_FuelType(1), FuelTypeDescription(CCFT_PETROL), CCFT_PETROL)
  Call AddComboItemAndItemData(CB_FuelType(1), FuelTypeDescription(CCFT_DIESEL), CCFT_DIESEL)
  Call AddComboItemAndItemData(CB_FuelType(1), FuelTypeDescription(CCFT_EUROIVDIESEL), CCFT_EUROIVDIESEL)
  Call AddComboItemAndItemData(CB_FuelType(1), FuelTypeDescription(CCFT_RDE2_DIESEL), CCFT_RDE2_DIESEL)
  Call AddComboItemAndItemData(CB_FuelType(1), FuelTypeDescription(CCFT_E85_BIO_ENTHANOL_AND_PETROL), CCFT_E85_BIO_ENTHANOL_AND_PETROL)
  Call AddComboItemAndItemData(CB_FuelType(1), FuelTypeDescription(CCFT_ELECTRIC), CCFT_ELECTRIC)
  Call AddComboItemAndItemData(CB_FuelType(1), FuelTypeDescription(CCFT_HYBRID), CCFT_HYBRID)
  Call AddComboItemAndItemData(CB_FuelType(1), FuelTypeDescription(CCFT_GAS_ONLY), CCFT_GAS_ONLY)
  Call AddComboItemAndItemData(CB_FuelType(1), FuelTypeDescription(CCFT_BIFUEL_WITH_CO2_FOR_GAS), CCFT_BIFUEL_WITH_CO2_FOR_GAS)
  Call AddComboItemAndItemData(CB_FuelType(1), FuelTypeDescription(CCFT_BIFUEL_CONVERSION_OTHER_NOT_WITHIN_TYPE_B), CCFT_BIFUEL_CONVERSION_OTHER_NOT_WITHIN_TYPE_B)
  
    
  'Call SetSortOrderToColumn(lb, LV_CAR_AVAIALABELFROM_SORT + 1, lvwAscending)
  Call SetDefaultVTDate(TB_Data(L_AVAILABLE_FROM_INDEX))
  Call SetDefaultVTDate(TB_Data(3))
  Call SetDefaultVTDate(TB_Data(15))
  Call SetDefaultVTDate(TB_Data(L_REGISTRATION_DATE_INDEX), UNDATED, UNDATED, True) 'registration date
  
  Me.tab.tab = 0
  Op_Data(7).Caption = S_COMPANY_CAR_DAYS_UNAVAILABLE_FUEL_DESCRIPTION
  
  Me.Label1.Visible = False
  Me.CB_FuelType(0).Visible = False
  Me.Label2(1).Visible = False
  Me.TB_Data(10).Visible = False
  Me.Op_Data(6).Visible = False
  
  TB_Data(4).Maximum = p11d32.Rates.value(DaysInYearLeap)
  TB_Data(12).MaxLength = p11d32.BenDataLinkMMFieldSize(BC_COMPANY_CARS_F, car_p46CarbonDioxide_db)
    
  Call SetupOpraInput(Label2(2), TB_Data(13))
  Call SetupOpraInput(Label2(3), TB_Data(14))

End Sub

Private Sub IBenefitForm2_AddBenefit()
  
  Dim benCar As CompanyCar
  Dim ben As IBenefitClass
On Error GoTo AddBenefit_Err

  Call xSet("AddBenefit")
  
  Set benCar = New CompanyCar
  Set ben = benCar
  Set ben.Parent = p11d32.CurrentEmployer.CurrentEmployee
  Call benCar.AddFuelBenefit
  
  Call AddBenefitHelper(Me, benCar)
'MP DB ToDo confirm - Fuel Enum does not have matching BEN_ITEMS for StandardRead, so commented below line
'MP DB  Call StandardReadData(ben.Fuel)
  'km - set make to default on CO2Emissions dialog
  
  
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
  Dim car As CompanyCar
  
  Call StandardReadData(ben)
  
  ben.value(car_Registration_db) = "Please enter registration..."
  ben.value(car_ListPrice_db) = 0
  ben.value(car_UnavailableDays_db) = 0
  ben.value(car_capitalcontribution_db) = 0
  ben.value(car_MadeGood_db) = 0
  ben.value(car_CheapAccessories_db) = 0
  ben.value(car_AccessoriesOriginal_db) = 0
  ben.value(car_AccessoriesNew_db) = 0
  ben.value(car_enginesize_db) = 0
  Call SetAvaialbleRange(ben, ben.Parent, Car_AvailableFrom_db, Car_AvailableTo_db)
  ben.value(car_Make_db) = "Please enter make..."
  ben.value(car_Model_db) = ""
  ben.value(car_p46PaymentFrequency_db) = P46_PAYMENT_FREQUENCY.P46PF_ACTUAL
  ben.value(car_p46FuelType_db) = CCFT_NONE
  ben.value(car_p46CarbonDioxide_db) = 0
  ben.value(car_Registrationdate_db) = p11d32.Rates.value(CarRegDateDef)
  'RK default fuel to car available dates
  ' ben.value(car_fuelavailablefrom) = ""
  ben.value(Car_FuelAvailableTo_db) = p11d32.Rates.value(TaxYearEnd)
  ben.value(car_NumberOfUsers_db) = 1
  ben.value(Car_HasFuelUnavailableDays_db) = False
  ben.value(car_Second_db) = False
  ben.value(car_Replaced_db) = False
  ben.value(car_Replacement_db) = False
  ben.value(car_ForceP46_db) = False
  ben.value(ITEM_OPRA_AMOUNT_FOREGONE) = 0
  ben.value(car_FuelOPRA_Ammount_Foregone_db) = 0
  ben.value(car_ElectricRangeMiles_db) = 0
  
  Set car = ben
  Call StandardReadData(car.Fuel)
    
End Function

Private Property Set IBenefitForm2_benefit(NewValue As IBenefitClass)
  Set benefit = NewValue
End Property

Private Property Get IBenefitForm2_benefit() As IBenefitClass
  Set IBenefitForm2_benefit = benefit
End Property

Private Function IBenefitForm2_BenefitFormState(ByVal fState As BENEFIT_FORM_STATE) As Boolean
  IBenefitForm2_BenefitFormState = BenefitFormStateEx(fState, benefit, fmeInput, Me.tab)
End Function

Private Function IBenefitForm2_BenefitOff() As Boolean
    TB_Data(0).Text = ""
    TB_Data(1).Text = ""
    TB_Data(2).Text = ""
    TB_Data(3).Text = ""
    TB_Data(4).Text = ""
    TB_Data(5).Text = ""
    TB_Data(6).Text = ""
    TB_Data(7).Text = ""
    TB_Data(8).Text = ""
    TB_Data(9).Text = ""
    TB_Data(11).Text = ""
    TB_Data(12).Text = "" 'AM
    TB_Data(15).Text = ""
    TB_Data(16).Text = ""   'IK 10/04/2004
    TB_Data(17).Text = ""   'IK 10/04/2004
    
    Op_Data(1) = vbUnchecked
    Op_Data(2) = vbUnchecked
    Op_Data(3) = vbUnchecked
    Call ReplacementCarVisible("", Op_Data(3))
    Op_Data(5) = vbUnchecked
    Op_Data(8) = vbUnchecked
    Op_Data(9) = vbUnchecked
    Op_Data(10) = vbUnchecked
    Op_Data(11) = vbUnchecked 'AM op_data(6) is op_data(11) in 2001
    
    ' EK 2/04 new controls TTP#212
    Op_Data(0) = vbUnchecked
    Op_Data(7) = vbUnchecked
    
    CB_CARLIST.Clear
    lblAccessories = ""
    TB_Data(13).Text = ""
    TB_Data(14).Text = ""
    
  
    
End Function
Private Function IBenefitForm2_BenefitOn() As Boolean
  
On Error GoTo BenefitOn_Err
  
  Call xSet("BenefitOn")
  
  
  With benefit
    TB_Data(0).Text = .value(car_Registration_db)
    TB_Data(1).Text = .value(car_ListPrice_db)
    TB_Data(2).Text = DateValReadToScreen(.value(Car_AvailableFrom_db))
    TB_Data(3).Text = DateValReadToScreen(.value(Car_AvailableTo_db))
    TB_Data(4).Text = .value(car_UnavailableDays_db)
    TB_Data(5).Text = .value(car_capitalcontribution_db)
    TB_Data(6).Text = .value(car_MadeGood_db)
    TB_Data(7).Text = .value(car_Make_db)
    TB_Data(9).Text = .value(car_Model_db)
    TB_Data(8).Text = DateValReadToScreen(.value(car_Registrationdate_db))
    TB_Data(11).Text = .value(car_enginesize_db)
    TB_Data(12).Text = .value(car_p46CarbonDioxide_db) 'AM
    TB_Data(15).Text = DateValReadToScreen(.value(Car_FuelAvailableTo_db))  'AM
    TB_Data(16).Text = .value(car_NumberOfUsers_db)
    TB_Data(17).Text = .value(car_ElectricRangeMiles_db)
    
    Op_Data(1) = IIf(.value(car_Replaced_db), vbChecked, vbUnchecked)
    Op_Data(2) = IIf(.value(car_Second_db), vbChecked, vbUnchecked)
    'Add these cars to an object list
    Call ReplacementCarsToList
    Op_Data(3) = IIf(.value(car_Replacement_db), vbChecked, vbUnchecked)
    Call ReplacementCarVisible(.value(car_RegReplaced_db), Op_Data(3))
    Op_Data(4) = BoolToChkBox(.value(ITEM_MADEGOOD_IS_TAXDEDUCTED))
    Op_Data(5) = IIf(.value(car_ForceP46_db), vbChecked, vbUnchecked)
    Op_Data(8) = IIf(.value(car_privatefuel_db), vbChecked, vbUnchecked)
    Op_Data(9) = IIf(.value(car_requiredmakegood_db), vbChecked, vbUnchecked)
    Op_Data(10) = IIf(.value(car_actualmadegood_db), vbChecked, vbUnchecked)
    Op_Data(11) = IIf(.value(car_p46NoApprovedCO2Figure_db), vbChecked, vbUnchecked)
    Op_Data(0) = IIf(.value(car_fuelreinstated_db), vbChecked, vbUnchecked)
    Op_Data(7) = IIf(.value(Car_HasFuelUnavailableDays_db), vbChecked, vbUnchecked)
    Op_Data(12) = BoolToChkBox(.value(car_FuelMadeGoodIsTaxDeducted_db))
    
    CB_Frequency.ListIndex = .value(car_p46PaymentFrequency_db) 'so
    Call ComboBoxItemDataToScreen(CB_FuelType(1), .value(car_p46FuelType_db))
    lblAccessories = .value(car_Accessories)
  
    Call CO2StuffChanged
    
    Op_Data(11).Visible = True
     
    TB_Data(13).Text = .value(ITEM_OPRA_AMOUNT_FOREGONE)
    TB_Data(14).Text = .value(car_FuelOPRA_Ammount_Foregone_db)
    Call electricRangeMiles
  
  End With
BenefitOn_End:
  Call xReturn("BenefitOn")
  Exit Function
BenefitOn_Err:
  Call ErrorMessage(ERR_ERROR, Err, "BenefitOn", "ERR_BenefitOn", "Error loading the car details onto the form.")
  Resume BenefitOn_End
  Resume
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
  IBenefitForm2_benclass = BC_COMPANY_CARS_F
End Property
Private Property Get IBenefitForm2_lv() As MSComctlLib.IListView
  Set IBenefitForm2_lv = LB
End Property
Private Function IBenefitForm2_RemoveBenefit(ByVal BenefitIndex As Long) As Boolean
  Dim CC As CompanyCar
  
  On Error GoTo RemoveBenefit_ERR
  
  Call xSet("RemoveBenefit")
  
  Set CC = benefit
  IBenefitForm2_RemoveBenefit = p11d32.CurrentEmployer.CurrentEmployee.RemoveBenefitWithLinks(Me, benefit, BenefitIndex, CC.Fuel)
  
RemoveBenefit_END:
  Set CC = Nothing
  Call xReturn("RemoveBenefit")
  Exit Function
RemoveBenefit_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "RemoveBenefit", "Remove Benefit", "Error removing a company car benefit.")
  Resume RemoveBenefit_END
End Function
Private Function IBenefitForm2_UpdateBenefitListViewItem(li As MSComctlLib.IListItem, benefit As IBenefitClass, Optional ByVal BenefitIndex As Long = 0&, Optional ByVal SelectItem As Boolean = False) As Long
  On Error GoTo UpdateBenefitListViewItemCompanyCar_ERR
  
  Call xSet("UpdateBenefitListViewItemCompanyCar")
  
  If Not li Is Nothing And Not benefit Is Nothing Then
    If BenefitIndex > 0 Then li.Tag = BenefitIndex
    li.SmallIcon = benefit.ImageListKey
    li.SubItems(LV_CAR_SUBITEMS.LV_CAR_BENEFIT) = FormatWN(benefit.Calculate)
    li.Text = benefit.Name
    li.SubItems(LV_CAR_AVAILABLE_FROM) = DateValReadToScreen(benefit.value(Car_AvailableFrom_db))
    li.SubItems(LV_CAR_AVAILABLE_TO) = DateValReadToScreen(benefit.value(Car_AvailableTo_db))
    li.SubItems(LV_CAR_FUEL_BENEFIT) = FormatWN(benefit.value(car_FuelBenefit))
    If SelectItem Then li.Selected = SelectItem
    IBenefitForm2_UpdateBenefitListViewItem = li.Index
    
    Call BenefitInErrorRow(benefit, li)
  End If

UpdateBenefitListViewItemCompanyCar_END:
  Call xReturn("UpdateBenefitListViewItemCompanyCar")
  Exit Function
UpdateBenefitListViewItemCompanyCar_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "UpdateBenefitListViewItemCompanyCar", "Update Benefit List View Item Company Car", "Error updating the benefit list view item for a company car.")
  Resume UpdateBenefitListViewItemCompanyCar_END
  Resume
End Function

'called on checkchanged of official CO2 and the fuel type
Private Function IBenefitForm2_ValididateBenefit(ben As IBenefitClass) As Boolean
  If ben.BenefitClass = BC_COMPANY_CARS_F Then IBenefitForm2_ValididateBenefit = True
End Function
Private Function IFrmGeneral_CheckChanged(c As Control) As Boolean
  Dim lst As ListItem
  Dim i As Long
  Dim bDirty As Boolean
  Dim car As CompanyCar
  
  On Error GoTo CheckChanged_Err
  Call xSet("CheckChanged")
  
  If p11d32.CurrentEmployeeIsNothing Then GoTo CheckChanged_End
  If benefit Is Nothing Then GoTo CheckChanged_End
  
  With c
    Select Case .Name
      Case "CB_CARLIST"
      Case "TB_DATA"
        Select Case .Index
          Case 0
            bDirty = CheckTextInput(.Text, benefit, car_Registration_db)
          Case 1
            bDirty = CheckTextInput(.Text, benefit, car_ListPrice_db)
          Case 2
            bDirty = CheckTextInput(.Text, benefit, Car_AvailableFrom_db)
          Case 3
            bDirty = CheckTextInput(.Text, benefit, Car_AvailableTo_db)
          Case 4
            bDirty = CheckTextInput(.Text, benefit, car_UnavailableDays_db)
          Case 5
            bDirty = CheckTextInput(.Text, benefit, car_capitalcontribution_db)
          Case 6
            bDirty = CheckTextInput(.Text, benefit, car_MadeGood_db)
          Case 7
            bDirty = CheckTextInput(.Text, benefit, car_Make_db)
          Case 8
            bDirty = CheckTextInput(.Text, benefit, car_Registrationdate_db)
            If (bDirty) Then
              If IsDate(benefit.value(car_Registrationdate_db)) Then
                If benefit.value(car_Registrationdate_db) < p11d32.Rates.value(TaxYearStart) Then
                  TB_Data(L_AVAILABLE_FROM_INDEX).Minimum = DateStringEx(p11d32.Rates.value(TaxYearStart), UNDATED)
                Else
                  TB_Data(L_AVAILABLE_FROM_INDEX).Minimum = DateStringEx(benefit.value(car_Registrationdate_db), UNDATED)
                End If
              Else
                TB_Data(L_AVAILABLE_FROM_INDEX).Minimum = DateStringEx(p11d32.Rates.value(TaxYearStart), UNDATED)
              End If
            End If
            
          Case 9
            bDirty = CheckTextInput(.Text, benefit, car_Model_db)
          Case 11
            bDirty = CheckTextInput(.Text, benefit, car_enginesize_db)
          Case 12
            bDirty = CheckTextInput(.Text, benefit, car_p46CarbonDioxide_db)
          Case 13
            bDirty = CheckTextInput(.Text, benefit, ITEM_OPRA_AMOUNT_FOREGONE)
          Case 14
            bDirty = CheckTextInput(.Text, benefit, car_FuelOPRA_Ammount_Foregone_db)
                      
          Case 15
            bDirty = CheckTextInput(.Text, benefit, Car_FuelAvailableTo_db)
          Case 16
            bDirty = CheckTextInput(.Text, benefit, car_NumberOfUsers_db)
          Case 17
            bDirty = CheckTextInput(.Text, benefit, car_ElectricRangeMiles_db)
          Case Else
            ECASE "Unknown control"
            GoTo CheckChanged_End
        End Select
      Case "Op_Data"
        Select Case .Index
          Case 1
            bDirty = CheckCheckBoxInput(.value, benefit, car_Replaced_db)
          Case 2
            bDirty = CheckCheckBoxInput(.value, benefit, car_Second_db)
          Case 3
            bDirty = CheckCheckBoxInput(.value, benefit, car_Replacement_db)
            Call ReplacementCarVisible(benefit.value(car_RegReplaced_db), c)
          Case 4
            bDirty = CheckCheckBoxInput(.value, benefit, ITEM_MADEGOOD_IS_TAXDEDUCTED)
          Case 5
            bDirty = CheckCheckBoxInput(.value, benefit, car_ForceP46_db)
          Case 8  'km
            bDirty = CheckCheckBoxInput(.value, benefit, car_privatefuel_db)
            If (bDirty) Then Call CO2StuffChanged
          Case 9  'km
            bDirty = CheckCheckBoxInput(.value, benefit, car_requiredmakegood_db)
          Case 10 'km
            bDirty = CheckCheckBoxInput(.value, benefit, car_actualmadegood_db)
          Case 11 'AM
              bDirty = CheckCheckBoxInput(.value, benefit, car_p46NoApprovedCO2Figure_db)
              If bDirty Then Call CO2StuffChanged
          Case 12
            bDirty = CheckCheckBoxInput(.value, benefit, car_FuelMadeGoodIsTaxDeducted_db)
          Case 0
            bDirty = CheckCheckBoxInput(.value, benefit, car_fuelreinstated_db)
          Case 7
            bDirty = CheckCheckBoxInput(.value, benefit, Car_HasFuelUnavailableDays_db)
          Case Else
            ECASE "Unknown control"
            GoTo CheckChanged_End
        End Select
      Case "CB_Frequency"
        bDirty = CB_Frequency.ListIndex <> benefit.value(car_p46PaymentFrequency_db)
        benefit.value(car_p46PaymentFrequency_db) = CB_Frequency.ListIndex
      Case "CB_FuelType"
           i = CB_FuelType(1).ItemData(CB_FuelType(1).ListIndex)
           If (i = CCFT_NONE) Then
            CB_FuelType(1).BackColor = TB_Data(1).InvalidColor
            
           Else
            CB_FuelType(1).BackColor = TB_Data(1).ValidColor
           End If
           
           bDirty = i <> benefit.value(car_p46FuelType_db)
           If (bDirty) Then
             benefit.value(car_p46FuelType_db) = i
             Call CO2StuffChanged
           End If
      Case "L_Data"
        Select Case .Index
          Case 17
            'reg of replacemendt car changed
            bDirty = CheckTextInput(.Caption, benefit, car_RegReplaced_db)
            Call CheckTextInput(L_Data(14), benefit, car_CarReplaced_db)
            Call CheckTextInput(L_Data(15), benefit, car_dateReplaced_db)
        Case Else
        End Select
      Case Else
        ECASE "Unknown control"
        GoTo CheckChanged_End
    End Select
  End With
  IFrmGeneral_CheckChanged = AfterCheckChanged(c, Me, bDirty)
  If (bDirty) Then
    lblAccessories = benefit.value(car_Accessories)
    
    Call electricRangeMiles
  End If
  
    
CheckChanged_End:
  Set car = Nothing
  Call xReturn("CheckChanged")
  Exit Function

CheckChanged_Err:
  Call ErrorMessage(ERR_ERROR, Err, "CheckChanged", "ERR_CHECKCHANGED", "This function has failed for the form " & Me.Name & ".")
  Resume CheckChanged_End
  Resume
End Function

Private Function GetReplacementCar(ByVal ListIndex As Long) As Boolean
  Dim ben As IBenefitClass

On Error GoTo GetReplacementCar_ERR
  Call xSet("GetReplacementCar")
    
  
  If ListIndex <> -1 Then
    
    Set ben = p11d32.CurrentEmployer.CurrentEmployee.benefits(CB_CARLIST.ItemData(ListIndex))
    
    L_Data(17) = ben.value(car_Registration_db)
    L_Data(14) = IIf(ben.value(car_Model_db) = "", ben.value(car_Make_db), (ben.value(car_Make_db) & " " & ben.value(car_Model_db)))
    
    benefit.value(car_CarReplacedMake_db) = ben.value(car_Make_db)
    benefit.value(car_CarReplacedModel_db) = ben.value(car_Model_db)
    benefit.value(car_CarReplacedEngineSize_db) = ben.value(car_enginesize_db)
    L_Data(15) = DateValReadToScreen(ben.value(Car_AvailableTo_db))
    Call IFrmGeneral_CheckChanged(L_Data(17))
  Else
    L_Data(17) = ""
    L_Data(14) = ""
    L_Data(15) = ""
  End If
  
  GetReplacementCar = True
  
GetReplacementCar_END:
  Call xReturn("GetReplacementCar")
  Exit Function
  
GetReplacementCar_ERR:
  GetReplacementCar = False
  Call ErrorMessage(ERR_ERROR, Err, "GetReplacementCar", "Get Replacement Car", "Error getting the replacement car from the cars object list. Index = " & ListIndex)
  Resume GetReplacementCar_END
  Resume
End Function
Private Function SetReplacementCarColor(ByVal lCol As Long) As Boolean

On Error GoTo SetReplacementCarColor_ERR
  
  Call xSet(" SetReplacementCarColor")
  
  If L_Data(17).ForeColor <> lCol Then
    L_Data(17).ForeColor = lCol
    L_Data(14).ForeColor = lCol
    L_Data(15).ForeColor = lCol
  End If
  
  SetReplacementCarColor = True
    
SetReplacementCarColor_END:
  Call xReturn(" SetReplacementCarColor")
  Exit Function
  
SetReplacementCarColor_ERR:
  SetReplacementCarColor = False
  Call ErrorMessage(ERR_ERROR, Err, " SetReplacementCarColor", "Get Replacement Car", "Error setting the replacement car details color.")
  Resume SetReplacementCarColor_END
  Resume
  
End Function

Private Function BlankReplacementCar() As Boolean
'RK old function?
On Error GoTo BlankReplacementCar_ERR

  Call xSet("BlankReplacementCar")
  
  benefit.value(car_RegReplaced_db) = ""
  benefit.value(car_CarReplaced_db) = ""
  benefit.value(car_dateReplaced_db) = ""
   
  BlankReplacementCar = True
  
BlankReplacementCar_END:
  Call xReturn("BlankReplacementCar")
  Exit Function
BlankReplacementCar_ERR:
  BlankReplacementCar = False
  Call ErrorMessage(ERR_ERROR, Err, "BlankReplacementCar", "Blank Replacement Car", "Error setting the replacement cars details to zero length strings.")
  Resume BlankReplacementCar_END
End Function

Private Function ReplacementCarVisible(sCurrentReg As String, cB As CheckBox) As Boolean
  Dim l As Long
  Dim bFound As Boolean
  
  
On Error GoTo ReplacementCarVisible_ERR
  
  Call xSet("ReplacementCarVisible")
  
  Call ReplacementCarsToList
  
  If cB.value = vbChecked Then
    If CB_CARLIST.ListCount Then
      CB_CARLIST.Visible = True
      'data
      L_Data(14).Visible = True
      L_Data(15).Visible = True
      L_Data(17).Visible = True
      'titles
      L_Data(18).Visible = True
      L_Data(5).Visible = True
      L_Data(11).Visible = True
      L_Data(16).Visible = True
      'has one been in the system before
      
      If Len(sCurrentReg) Then
        For l = 0 To CB_CARLIST.ListCount - 1
          If StrComp(sCurrentReg, CB_CARLIST.List(l)) = 0 Then
              CB_CARLIST.Text = sCurrentReg
              bFound = True
            Exit For
          End If
        Next
        If bFound Then
          'yes so select it in the combo box
           Call GetReplacementCar(l)
        Else
          'give warning to the user that their previous car is not available
          Call ErrorMessage(ERR_INFO, Err, "ReplacementCarVisible", "Replacement Car Visible", "The car with registration number of " & benefit.value(car_RegReplaced_db) & " is no longer available please check available cars. Selecting the first available car")
          Call GetReplacementCar(0) 'select th first one
        End If
      Else
        'there has not been a previous car so select the first in the combo
        Call GetReplacementCar(0)
      End If
    Else
      Call ErrorMessage(ERR_INFO, Err, "ReplacementCarVisible", "Replacement Car Visible", "There are no replacement cars available. Please check the available to and from date for the current car and the car you wish to replace.")
      Op_Data(3) = vbUnchecked
    End If
  Else
    CB_CARLIST.Visible = False
    CB_CARLIST.Text = ""
    
    L_Data(14) = ""
    L_Data(15) = ""
    L_Data(17) = ""
    
    L_Data(14).Visible = False
    L_Data(15).Visible = False
    L_Data(17).Visible = False
    L_Data(18).Visible = False
    L_Data(5).Visible = False
    L_Data(11).Visible = False
    L_Data(16).Visible = False
    'print make on p46

  End If
  
ReplacementCarVisible_END:
  Call xReturn("ReplacementCarVisible")
  Exit Function
ReplacementCarVisible_ERR:
  ReplacementCarVisible = False
  Call ErrorMessage(ERR_ERROR, Err, "ReplacementCarVisible", "Replacement Car Visible", "Error setting the controls for the replacement car flag.")
  Resume ReplacementCarVisible_END
  Resume
  
End Function
Private Property Get IFrmGeneral_InvalidVT() As Control
  Set IFrmGeneral_InvalidVT = m_InvalidVT
End Property

Private Property Set IFrmGeneral_InvalidVT(NewValue As Control)
  Set m_InvalidVT = NewValue
End Property

Private Sub LB_ItemClick(ByVal Item As MSComctlLib.ListItem)
  Call SetLastListItemSelected(Item)
  F_CompanyCarCO2Emissions.CB_Make.ListIndex = 0
  Call IBenefitForm2_BenefitToScreen(Item.Tag)
End Sub

Private Sub lb_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
  Call SetSortOrder(LB, ColumnHeader)
End Sub

Private Sub LB_KeyDown(KeyCode As Integer, Shift As Integer)
  Call LVKeyDown(KeyCode, Shift)
End Sub

Private Sub lb_KeyPress(KeyAscii As Integer)
  'Check for return key - tab to next field
  If KeyAscii = 13 Then Call SendKeysEx(vbTab)
End Sub

Private Sub Op_Data_Click(Index As Integer)
  Call IFrmGeneral_CheckChanged(Op_Data(Index))
End Sub

Private Sub Op_Data_KeyPress(Index As Integer, KeyAscii As Integer)
  'Check for return key - tab to next field
  If KeyAscii = 13 Then Call SendKeysEx(vbTab)
End Sub

Private Sub TB_Data_FieldInvalid(Index As Integer, Valid As Boolean, Message As String)
  'RK Special handling for CO2 validations
  Call SetPanel2(Message)
End Sub
Private Sub TB_Data_Validate(Index As Integer, Cancel As Boolean)
  Call IFrmGeneral_CheckChanged(TB_Data(Index))
End Sub
Private Sub TB_Data_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  TB_Data(Index).Tag = SetChanged
End Sub
Private Function ReplacementCarsToList() As Boolean
  Dim i As Long
  Dim ben As IBenefitClass


  On Error GoTo ReplacementCarsToList_Err
  Call xSet("ReplacementCarsToList")

  If benefit Is Nothing Then GoTo ReplacementCarsToList_End
  
  CB_CARLIST.Clear
  For i = 1 To p11d32.CurrentEmployer.CurrentEmployee.benefits.Count
    Set ben = p11d32.CurrentEmployer.CurrentEmployee.benefits(i)
    If Not (ben Is Nothing) Then
      If ben.BenefitClass = BC_COMPANY_CARS_F And Not ben Is benefit Then
        If ben.value(Car_AvailableTo_db) = benefit.value(Car_AvailableFrom_db) Or DateAdd("d", 1, ben.value(Car_AvailableTo_db)) = benefit.value(Car_AvailableFrom_db) Then
          CB_CARLIST.AddItem (ben.Name)
          CB_CARLIST.ItemData(CB_CARLIST.ListCount - 1) = i
        End If
      End If
    End If
  Next i


ReplacementCarsToList_End:
  Call xReturn("ReplacementCarsToList")
  Exit Function

ReplacementCarsToList_Err:
  Call ErrorMessage(ERR_ERROR, Err, "ReplacementCarsToList", "Replacement Cars To List", "Error placing any available replacement cars to the replacement cars list combo box.")
  Resume ReplacementCarsToList_End
  Resume
End Function

Private Sub electricRangeMiles()
  Dim b As Boolean
  
  b = benefit.value(car_p46FuelType_db) = COMPANY_CAR_FUEL_TYPE.CCFT_HYBRID
  
  TB_Data(17).Visible = b
  L_Data(6).Visible = b
  TB_Data(17).Validate = b
  

End Sub
Private Sub CO2StuffChanged()
   Dim bCO2Required As Boolean
   Dim bNoOfficalCO2 As Boolean
   Dim ccft As COMPANY_CAR_FUEL_TYPE
   Dim ifg As IFrmGeneral
   
   ccft = benefit.value(car_p46FuelType_db)
      
   Select Case ccft
    Case CCFT_PETROL, CCFT_DIESEL, CCFT_HYBRID, CCFT_BIFUEL_WITH_CO2_FOR_GAS, CCFT_EUROIVDIESEL, CCFT_BIFUEL_CONVERSION_OTHER_NOT_WITHIN_TYPE_B, CCFT_GAS_ONLY, CCFT_E85_BIO_ENTHANOL_AND_PETROL, CCFT_RDE2_DIESEL, CCFT_NONE
        bCO2Required = True
     Case CCFT_ELECTRIC
        bCO2Required = False
     Case Else
      Call Err.Raise(ERR_CAR_CO2, "CO2Changed", "invalid fuel type")
   End Select
      
   If ccft = CCFT_ELECTRIC Then
     Call EnableFrame(Me, fraFuelBenefit, False)
   Else
     Call EnableFrame(Me, fraFuelBenefit, benefit.value(car_privatefuel_db))
   End If
   
   If (Not fraFuelBenefit.Enabled) Then 'WE ARE ELECTRIC
    
    'we must validate the date withdrawn if invalid
     If (TB_Data(15).FieldInvalid) Then 'fuel available to
       If (Not TB_Data(3).FieldInvalid) Then 'car available to
         TB_Data(15).Text = TB_Data(3).Text
       Else
         TB_Data(15).Text = DateValReadToScreen(p11d32.Rates.value(TaxYearEnd)) 'SET BACK TO ITS ORIGINAL DEFAULT
       End If
     End If
     If (TB_Data(11).FieldInvalid) Then 'CC
      TB_Data(11).Text = "0"
      Call IFrmGeneral_CheckChanged(TB_Data(11))
     End If
   End If
      
   Call EnableFrame(Me, fraCO2, bCO2Required)
      
      
   TB_Data(11).Enabled = NoOfficalCO2.Enabled = lblEmissions.Enabled = CO2Level.Enabled = bCO2Required
   
   TB_Data(11).Validate = CO2Level.Validate = bCO2Required
End Sub
Private Property Get CO2Level() As ValText
  Set CO2Level = TB_Data(12)
End Property
Private Property Get NoOfficalCO2() As CheckBox
  Set NoOfficalCO2 = Op_Data(11)
End Property

