VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{89056D22-ECDA-4A64-B90B-25EBB3AE8DB8}#1.0#0"; "atc2hook.ocx"
Object = "{8D988532-0F0C-460C-B00E-7B5637E97680}#1.0#0"; "atc2vtext.ocx"
Begin VB.Form F_Employees 
   Caption         =   "Personal Details"
   ClientHeight    =   5745
   ClientLeft      =   60
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
   Icon            =   "F_Emloyees.frx":0000
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5745
   ScaleWidth      =   8430
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ListView LB 
      Height          =   2400
      Left            =   45
      TabIndex        =   0
      Tag             =   "free,font"
      Top             =   45
      Width           =   8310
      _ExtentX        =   14658
      _ExtentY        =   4233
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      Checkboxes      =   -1  'True
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
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Reference"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "NI Number"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Status"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Group1"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Group2"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Group3"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Frame fmeDetails 
      Height          =   3270
      Left            =   90
      TabIndex        =   24
      Top             =   2430
      Width           =   8265
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   3030
         Left            =   45
         ScaleHeight     =   3030
         ScaleWidth      =   8160
         TabIndex        =   25
         Top             =   180
         Width           =   8160
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
            Left            =   6540
            Style           =   2  'Dropdown List
            TabIndex        =   16
            Tag             =   "FREE,FONT"
            Top             =   45
            Width           =   1545
         End
         Begin VB.CommandButton cmdChangePNum 
            Appearance      =   0  'Flat
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "System"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   4905
            TabIndex        =   8
            Tag             =   "FREE"
            Top             =   1330
            Width           =   350
         End
         Begin VB.CommandButton cmdSelectByGroup 
            Height          =   285
            Index           =   0
            Left            =   4905
            Picture         =   "F_Emloyees.frx":030A
            Style           =   1  'Graphical
            TabIndex        =   10
            TabStop         =   0   'False
            ToolTipText     =   "Select by group code 1"
            Top             =   1665
            Width           =   350
         End
         Begin VB.CommandButton cmdSelectByGroup 
            Height          =   285
            Index           =   1
            Left            =   2655
            Picture         =   "F_Emloyees.frx":040C
            Style           =   1  'Graphical
            TabIndex        =   12
            TabStop         =   0   'False
            ToolTipText     =   "Select by group code 2"
            Top             =   1965
            Width           =   350
         End
         Begin VB.CommandButton cmdSelectByGroup 
            Height          =   285
            Index           =   2
            Left            =   4905
            Picture         =   "F_Emloyees.frx":050E
            Style           =   1  'Graphical
            TabIndex        =   14
            TabStop         =   0   'False
            ToolTipText     =   "Select by group code 3"
            Top             =   1965
            Width           =   350
         End
         Begin VB.CommandButton cmdAddress 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   450
            Left            =   6510
            Picture         =   "F_Emloyees.frx":0610
            Style           =   1  'Graphical
            TabIndex        =   23
            Top             =   2505
            Width           =   1545
         End
         Begin VB.CheckBox chk_Class1AEmployeeIsNotSubjectTo 
            Alignment       =   1  'Right Justify
            Caption         =   "Employee not subject to Class 1A NIC"
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
            Height          =   360
            Left            =   5445
            TabIndex        =   22
            Tag             =   "FREE,FONT"
            Top             =   2115
            Width           =   2655
         End
         Begin atc2hook.HOOK HOOK 
            Left            =   1215
            Top             =   2250
            _ExtentX        =   847
            _ExtentY        =   847
         End
         Begin atc2valtext.ValText TB_Data 
            Height          =   285
            Index           =   6
            Left            =   1755
            TabIndex        =   11
            Tag             =   "FREE,FONT"
            Top             =   1965
            Width           =   900
            _ExtentX        =   1588
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
            MouseIcon       =   "F_Emloyees.frx":07A2
            Text            =   ""
            TypeOfData      =   3
         End
         Begin atc2valtext.ValText TB_Data 
            Height          =   285
            Index           =   2
            Left            =   6540
            TabIndex        =   17
            Tag             =   "FREE,FONT"
            Top             =   405
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
            MaxLength       =   50
            MouseIcon       =   "F_Emloyees.frx":07BE
            Text            =   ""
            TypeOfData      =   4
         End
         Begin atc2valtext.ValText TB_Date 
            Height          =   285
            Index           =   1
            Left            =   6540
            TabIndex        =   18
            Tag             =   "FREE,FONT"
            Top             =   720
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
            MouseIcon       =   "F_Emloyees.frx":07DA
            Text            =   ""
            TypeOfData      =   2
            Maximum         =   "5/4/1999"
         End
         Begin atc2valtext.ValText TB_Date 
            Height          =   285
            Index           =   2
            Left            =   6540
            TabIndex        =   19
            Tag             =   "FREE,FONT"
            Top             =   1035
            Width           =   1545
            _ExtentX        =   2725
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
            MouseIcon       =   "F_Emloyees.frx":07F6
            Text            =   ""
            TypeOfData      =   2
            Maximum         =   "5/4/1999"
            Minimum         =   "6/4/1998"
         End
         Begin atc2valtext.ValText TB_Data 
            Height          =   285
            Index           =   8
            Left            =   4095
            TabIndex        =   13
            Tag             =   "FREE,FONT"
            Top             =   1965
            Width           =   810
            _ExtentX        =   1429
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
            MouseIcon       =   "F_Emloyees.frx":0812
            Text            =   ""
            TypeOfData      =   3
         End
         Begin atc2valtext.ValText TB_Data 
            Height          =   285
            Index           =   3
            Left            =   1755
            TabIndex        =   9
            Tag             =   "FREE,FONT"
            Top             =   1650
            Width           =   3150
            _ExtentX        =   5556
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
            MouseIcon       =   "F_Emloyees.frx":082E
            Text            =   ""
            TypeOfData      =   3
         End
         Begin atc2valtext.ValText TB_Data 
            Height          =   285
            Index           =   1
            Left            =   1755
            TabIndex        =   7
            Tag             =   "FREE,FONT"
            Top             =   1320
            Width           =   3150
            _ExtentX        =   5556
            _ExtentY        =   503
            BackColor       =   255
            Enabled         =   0   'False
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
            MaxLength       =   255
            MouseIcon       =   "F_Emloyees.frx":084A
            Text            =   ""
            TypeOfData      =   3
            AllowEmpty      =   0   'False
         End
         Begin atc2valtext.ValText TB_Data 
            Height          =   285
            Index           =   0
            Left            =   540
            TabIndex        =   1
            Tag             =   "FREE,FONT"
            Top             =   45
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
            MaxLength       =   50
            MouseIcon       =   "F_Emloyees.frx":0866
            Text            =   ""
            TypeOfData      =   3
         End
         Begin atc2valtext.ValText TB_Data 
            Height          =   285
            Index           =   4
            Left            =   540
            TabIndex        =   2
            Tag             =   "FREE,FONT"
            Top             =   360
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
            MaxLength       =   50
            MouseIcon       =   "F_Emloyees.frx":0882
            Text            =   ""
            TypeOfData      =   3
         End
         Begin atc2valtext.ValText TB_Data 
            Height          =   285
            Index           =   5
            Left            =   2250
            TabIndex        =   3
            Tag             =   "FREE,FONT"
            Top             =   45
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
            MaxLength       =   50
            MouseIcon       =   "F_Emloyees.frx":089E
            Text            =   ""
            TypeOfData      =   3
         End
         Begin atc2valtext.ValText TB_Data 
            Height          =   285
            Index           =   9
            Left            =   2250
            TabIndex        =   4
            Tag             =   "FREE,FONT"
            Top             =   360
            Width           =   2985
            _ExtentX        =   5265
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
            MouseIcon       =   "F_Emloyees.frx":08BA
            Text            =   ""
            TypeOfData      =   3
            AllowEmpty      =   0   'False
         End
         Begin atc2valtext.ValText TB_Data 
            Height          =   285
            Index           =   10
            Left            =   2250
            TabIndex        =   5
            Tag             =   "FREE,FONT"
            Top             =   690
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
            MouseIcon       =   "F_Emloyees.frx":08D6
            Text            =   ""
            TypeOfData      =   3
         End
         Begin atc2valtext.ValText TB_Data 
            Height          =   285
            Index           =   11
            Left            =   675
            TabIndex        =   6
            Tag             =   "FREE,FONT"
            Top             =   1005
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
            MaxLength       =   255
            MouseIcon       =   "F_Emloyees.frx":08F2
            Text            =   ""
            TypeOfData      =   3
         End
         Begin atc2valtext.ValText TB_Data 
            Height          =   285
            Index           =   12
            Left            =   6540
            TabIndex        =   21
            Tag             =   "FREE,FONT"
            Top             =   1740
            Width           =   1545
            _ExtentX        =   2725
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
            MouseIcon       =   "F_Emloyees.frx":090E
            Text            =   ""
            TypeOfData      =   3
         End
         Begin atc2valtext.ValText TB_Data 
            Height          =   285
            Index           =   13
            Left            =   6540
            TabIndex        =   20
            Tag             =   "FREE,FONT"
            Top             =   1425
            Width           =   1545
            _ExtentX        =   2725
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
            MouseIcon       =   "F_Emloyees.frx":092A
            Text            =   ""
            TypeOfData      =   3
         End
         Begin atc2valtext.ValText TB_Data 
            Height          =   675
            Index           =   7
            Left            =   1755
            TabIndex        =   15
            Tag             =   "FREE,FONT"
            Top             =   2280
            Width           =   3495
            _ExtentX        =   6165
            _ExtentY        =   1191
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
            MouseIcon       =   "F_Emloyees.frx":0946
            Text            =   ""
            TypeOfData      =   3
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
            Left            =   5460
            TabIndex        =   42
            Tag             =   "FREE,FONT"
            Top             =   90
            Width           =   1095
         End
         Begin VB.Label lblGroupCode2 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Group code 2"
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
            Left            =   45
            TabIndex        =   41
            Tag             =   "FREE,FONT"
            Top             =   2010
            Width           =   975
         End
         Begin VB.Label lblGroupCode1 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Group code 1"
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
            Left            =   45
            TabIndex        =   40
            Tag             =   "FREE,FONT"
            Top             =   1695
            Width           =   975
         End
         Begin VB.Label Label11 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Date left"
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
            Left            =   5460
            TabIndex        =   39
            Tag             =   "FREE,FONT"
            Top             =   1080
            Width           =   600
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
            Left            =   5460
            TabIndex        =   38
            Tag             =   "FREE,FONT"
            Top             =   765
            Width           =   1110
         End
         Begin VB.Label Label8 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "NI number"
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
            Left            =   5460
            TabIndex        =   37
            Tag             =   "FREE,FONT"
            Top             =   450
            Width           =   735
         End
         Begin VB.Label Label7 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Personnel number"
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
            Left            =   45
            TabIndex        =   36
            Tag             =   "FREE,FONT"
            Top             =   1380
            Width           =   1275
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
            Left            =   1350
            TabIndex        =   35
            Tag             =   "FREE,FONT"
            Top             =   90
            Width           =   720
         End
         Begin VB.Label lblGroupCode3 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Group code 3"
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
            Left            =   3060
            TabIndex        =   34
            Tag             =   "FREE,FONT"
            Top             =   2010
            Width           =   975
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
            Left            =   45
            TabIndex        =   33
            Tag             =   "FREE,FONT"
            Top             =   2340
            Width           =   1005
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
            Left            =   45
            TabIndex        =   32
            Tag             =   "FREE,FONT"
            Top             =   420
            Width           =   435
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
            Left            =   45
            TabIndex        =   31
            Tag             =   "FREE,FONT"
            Top             =   90
            Width           =   300
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
            Left            =   1350
            TabIndex        =   30
            Tag             =   "FREE,FONT"
            Top             =   420
            Width           =   630
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
            Left            =   1350
            TabIndex        =   29
            Tag             =   "FREE,FONT"
            Top             =   735
            Width           =   705
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
            Left            =   45
            TabIndex        =   28
            Tag             =   "FREE,FONT"
            Top             =   1050
            Width           =   735
         End
         Begin VB.Label Label14 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Intranet password"
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
            Index           =   0
            Left            =   5460
            TabIndex        =   27
            Tag             =   "FREE,FONT"
            Top             =   1725
            Width           =   1005
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label14 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Intranet username"
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
            Index           =   1
            Left            =   5460
            TabIndex        =   26
            Tag             =   "FREE,FONT"
            Top             =   1305
            Width           =   1005
            WordWrap        =   -1  'True
         End
      End
   End
End
Attribute VB_Name = "F_Employees"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Implements IBenefitForm2
Implements IFrmGeneral
Public SortOrder As ListSortOrderConstants

Private Const L_DES_HEIGHT  As Long = 5925
Private Const L_DES_WIDTH  As Long = 8505

Private mclsResize As New clsFormResize
Private m_InvalidVT As Control
Public benefit As IBenefitClass

Private m_LastSearch As String

Private Sub chk_Class1AEmployeeIsNotSubjectTo_Click()
  Call IFrmGeneral_CheckChanged(chk_Class1AEmployeeIsNotSubjectTo)
End Sub

Private Sub cmdChangePNum_Click()
  Dim ibf As IBenefitForm2
  Dim ee As Employee
  
On Error GoTo cmdChangePNum_ERR
  Call xSet("cmdChangePNum_Click")
  
  Set ibf = Me
  If Not ibf.lv.SelectedItem Is Nothing Then
    Set ee = p11d32.CurrentEmployer.employees(ibf.lv.SelectedItem.Tag)
    Call ee.PersonnelNumberChange
    Call ibf.UpdateBenefitListViewItem(ibf.lv.SelectedItem, ee)
    Call ibf.BenefitOn
  End If
  
cmdChangePNum_END:
  Call xSet("cmdChangePNum_Click")
  Exit Sub
cmdChangePNum_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "cmdChangePNum", "cmd Change PNum", "Error in cmd change.")
  Resume cmdChangePNum_END
End Sub

Public Sub cmdSelectByGroup_Click(Index As Integer)

  Select Case Index
    Case 0
      Call SelectItems(F_Employees.lb, SELECT_MODE.SELECT_GROUP_1, TB_Data(3).Text)
    Case 1
      Call SelectItems(F_Employees.lb, SELECT_MODE.SELECT_GROUP_2, TB_Data(6).Text)
    Case 2
      Call SelectItems(F_Employees.lb, SELECT_MODE.SELECT_GROUP_3, TB_Data(8).Text)
    
  End Select
  
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Set m_InvalidVT = Nothing
  Set mclsResize = Nothing
  Set benefit = Nothing
End Sub

Private Sub CB_Status_Validate(Cancel As Boolean)
   Call IFrmGeneral_CheckChanged(CB_Status)
End Sub

Private Sub cmdAddress_Click()
   Call DialogToScreen(F_Addresses, Nothing, 0, Me, F_Employees.lb.SelectedItem.Tag)
End Sub


Private Sub FramePicture1_GotFocus()

End Sub

Private Sub HOOK_WndProc(Discard As Boolean, MsgReturn As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long)
  Select Case Msg
    Case WM_LBUTTONDOWN
      If Not lb.SelectedItem Is Nothing Then
        Discard = Not p11d32.CurrentEmployer.SaveCurrentEmployee
      End If
  End Select
  
End Sub

Private Sub IBenefitForm2_AddBenefit()
  Dim ben As IBenefitClass
  Dim lst As ListItem, i As Long
  Dim ibf As IBenefitForm2
  Dim ee As Employee
  Dim CurrentEmployee As IBenefitClass
  
On Error GoTo Employee_AddBenefit_Err
  
  Call xSet("Employee_AddBenefit")
  
  Set CurrentEmployee = p11d32.CurrentEmployer.CurrentEmployee
  
  If Not p11d32.CurrentEmployer.LoadEmployeeEx(ben) Then GoTo Employee_AddBenefit_End
  
RE_TEST:
  
'  F_EmployeeNew.Show vbModal
  Call p11d32.Help.ShowForm(F_EmployeeNew, vbModal)
  If F_EmployeeNew.m_OK = False Then
    Set p11d32.CurrentEmployer.CurrentEmployee = CurrentEmployee
    GoTo Employee_AddBenefit_End
  End If
  
  If Not p11d32.CurrentEmployer.ValidatePersonnelNumber(F_EmployeeNew.TxtBx(0).Text) Then
    GoTo RE_TEST
  End If
  
  Set ben = New Employee
  With ben
    Set .Parent = p11d32.CurrentEmployer
    Set ibf = Me
    Call ibf.AddBenefitSetDefaults(ben)
    ben.Dirty = True
    .WriteDB
    i = p11d32.CurrentEmployer.employees.Add(ben)
    
    Set lst = lb.listitems.Add(, , .Name)
    Call ibf.UpdateBenefitListViewItem(lst, ben, i, True)
    Call SelectBenefitByListItem(ibf, lst)
    ben.ReadFromDB = True
  End With
  
  
Employee_AddBenefit_End:
  
  Call UpdateEmployerEmployeesCount
  Set ben = Nothing
  Set lst = Nothing
  Set ibf = Nothing
  Unload F_EmployeeNew
  Set F_EmployeeNew = Nothing
  MDIMain.Enabled = True
  Set ee = Nothing
  Call DBEngine.Idle(dbFreeLocks)
  Call xSet("Employee_AddBenefit")
  Exit Sub
Employee_AddBenefit_Err:
  Select Case Err.Number
    Case 3022
      Call ErrorMessage(ERR_ERROR, Err, "Duplicate Personnel reference", "ERR_ADDEMPOYEE", "The personnel reference you are trying to add already exists in the database." & vbCrLf & "Please use a unique personnel reference for each employee.")
      Resume RE_TEST
    Case 3315
      Call ErrorMessage(ERR_ERROR, Err, "Empty field", "ERR_ADDEMPOYEE", "You must complete all the fields.")
      Resume RE_TEST
    Case Else
      Call ErrorMessage(ERR_ERROR, Err, "AddEmployee", "ERR_AddEmployee", "Error in AddEmployee function, called from the form " & Me.Name & ".")
      Resume Employee_AddBenefit_End
  End Select
  Resume
End Sub
Private Sub UpdateEmployerEmployeesCount()
  Dim ben As IBenefitClass
  Dim ibf As IBenefitForm2
  
  On Error GoTo UpdateEmployerEmployeesCount_ERR
  
  Call xSet("UpdateEmployerEmployeesCount")
  Set ibf = Me
  Set ben = p11d32.CurrentEmployer
  ben.value(employer_EmployeesCount) = ibf.lv.listitems.Count
  
UpdateEmployerEmployeesCount_END:
  Call xReturn("UpdateEmployerEmployeesCount")
  Exit Sub
UpdateEmployerEmployeesCount_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "UpdateEmployerEmployeesCount", "Update Employer Employees Count", "Unable to set the employer count.")
  Resume UpdateEmployerEmployeesCount_END
End Sub
Private Function IBenefitForm2_AddBenefitSetDefaults(ben As IBenefitClass) As Long
  With ben
    .value(ee_PersonnelNumber_db) = F_EmployeeNew.TxtBx(0).Text
    .value(ee_Surname_db) = F_EmployeeNew.TxtBx(1).Text
    .value(ee_Title_db) = F_EmployeeNew.TxtBx(2).Text
    .value(ee_Firstname_db) = F_EmployeeNew.TxtBx(3).Text
    .value(ee_Initials_db) = F_EmployeeNew.TxtBx(4).Text
    .value(ee_NINumber_db) = F_EmployeeNew.TxtBx(5).Text
    
    'FC - Class1A
    .value(ee_Class1AEmployeeIsNotSubjectTo_db) = False
    
    .value(ee_OneOrMoreSharedVanAvailable_db) = False
    .value(ee_RelevantDaysForDailySharedVanCalc_db) = 0
    .value(ee_PaymentsForPrivateUseOfSharedVans_db) = 0
    .value(ee_NonSharedVanAvailableAtSameTimeAsSharedVan_db) = False
    .value(ee_ReportyDailyCalculationOfSharedVans_db) = False
    .value(ee_Selected) = False
  End With
End Function

Private Property Get IBenefitForm2_benefit() As IBenefitClass
  Set IBenefitForm2_benefit = benefit
End Property

Private Property Set IBenefitForm2_benefit(NewValue As IBenefitClass)
  Set benefit = NewValue
End Property

Private Function IBenefitForm2_BenefitFormState(ByVal fState As BENEFIT_FORM_STATE) As Boolean
On Error GoTo BenefitFormState_err
  Call xSet("IBenefitForm2_EnableBenefitForm")
  
  If (fState = FORM_ENABLED) Or (fState = FORM_CDB) Then
    If fState = FORM_ENABLED Then
      fmeDetails.Enabled = True
    End If
    Call SetLVEnabled(lb, True)
    
    Call MDIMain.SetDelete
    Call MDIMain.SetAdd
    cmdChangePNum.Enabled = True
    MDIMain.mnuBenefits.Enabled = True
    MDIMain.tbrBenefits.Enabled = True
    Call MDIMain.SetDelete
  ElseIf fState = FORM_DISABLED Then
    fmeDetails.Enabled = False
    Call SetLVEnabled(lb, False)
    cmdChangePNum.Enabled = False
    Call MDIMain.ClearDelete
    Call MDIMain.ClearConfirmUndo
    MDIMain.mnuBenefits.Enabled = False
    MDIMain.tbrBenefits.Enabled = False
  End If
  
  Call MDIMain.NavigateBarUpdate(benefit)
  
  IBenefitForm2_BenefitFormState = True
    
BenefitFormState_end:
  Call xReturn("BenefitFormState")
  Exit Function
  
BenefitFormState_err:
  IBenefitForm2_BenefitFormState = False
  Call ErrorMessage(ERR_ERROR, Err, "BenefitFormState", "Benefit Form State", "Error setting the benefit form state for the employees form.")
  Resume BenefitFormState_end
  Resume
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
  'CB_Status = ""
  TB_Data(8).Text = ""
  TB_Data(9).Text = ""
  TB_Data(10).Text = ""
  TB_Data(11).Text = ""
  TB_Data(12).Text = ""
  TB_Date(1).Text = ""
  TB_Date(2).Text = ""
  
  'FC - Class1A
  chk_Class1AEmployeeIsNotSubjectTo.value = vbUnchecked
  
  MDIMain.mnuEmployeeItems(MNU_EMPLOYEE_DELETE).Enabled = False
  MDIMain.mnuEmployeeItems(MNU_EMPLOYEE_GOTO).Enabled = False
  Call MDIMain.NavigateBarUpdate(benefit)
End Function

Private Function IBenefitForm2_BenefitOn() As Boolean
  With benefit
    TB_Data(0).Text = .value(ee_Title_db)
    TB_Data(5).Text = .value(ee_Firstname_db)
    TB_Data(4).Text = .value(ee_Initials_db)
    TB_Data(9).Text = .value(ee_Surname_db)
    TB_Data(10).Text = .value(ee_Salutation_db)
    TB_Data(11).Text = .value(ee_Email_db)
    TB_Data(12).Text = .value(ee_Password_db)
    TB_Data(13).Text = .value(ee_Username_db)
    
    TB_Data(1).Text = .value(ee_PersonnelNumber_db)
    TB_Data(2).Text = .value(ee_NINumber_db)
    TB_Data(3).Text = .value(ee_Group1_db)
    TB_Data(6).Text = .value(ee_Group2_db)
    TB_Data(8).Text = .value(ee_Group3_db)
    TB_Data(7).Text = .value(ee_Comments_db)
    
    CB_Status = IIf(.value(ee_Director_db), S_DIRECTOR, S_STAFF)
    
    TB_Date(1).Text = IIf(.value(ee_joined_db) = UNDATED, "", DateValReadToScreen(.value(ee_joined_db)))
    TB_Date(2).Text = IIf(.value(ee_left_db) = UNDATED, "", DateValReadToScreen(.value(ee_left_db)))
    
    'FC - Class1A
    chk_Class1AEmployeeIsNotSubjectTo.value = BoolToChkBox(.value(ee_Class1AEmployeeIsNotSubjectTo_db))
    
    MDIMain.mnuEmployeeItems(MNU_EMPLOYEE_DELETE).Enabled = True
    MDIMain.mnuEmployeeItems(MNU_EMPLOYEE_GOTO).Enabled = True
  End With
  Call MDIMain.NavigateBarUpdate(benefit)
End Function

Private Function IBenefitForm2_BenefitsToListView() As Long
  Dim i As Long, li As ListItem
  Dim ben As IBenefitClass
  Dim ibf As IBenefitForm2
  
  On Error GoTo BenefitsToListView_err
  Call xSet("BenefitsToListView")
  Call SetCursor
  Set ibf = Me
  Call ibf.BenefitToScreen(, False)
  
  Call MDIMain.SetAdd
  ibf.lv.Sorted = False

  Set ibf.lv.SmallIcons = MDIMain.imlListViewBenefits
  Call p11d32.CurrentEmployer.employees.Compact
  Call AllocateListview(ibf.lv, p11d32.CurrentEmployer.employees.Count)
  For i = 1 To p11d32.CurrentEmployer.employees.Count
    Set ben = p11d32.CurrentEmployer.employees(i)
    Set li = lb.listitems.Item(i)
    li.Key = "EE" & ben.value(ee_PersonnelNumber_db)
    Call ibf.UpdateBenefitListViewItem(li, ben, i)
  Next i
  IBenefitForm2_BenefitsToListView = ibf.lv.listitems.Count
      
      
  Call SetSortOrder(ibf.lv, ibf.lv.ColumnHeaders(p11d32.EmployeeSortOrderColumn + 1), p11d32.EmployeeSortOrder)
  'ibf.lv.Sorted = True
  
BenefitsToListView_end:
  Set ben = Nothing
  Call ClearCursor
  Call xReturn("BenefitsToListView")
  Exit Function
  
BenefitsToListView_err:
  Call ErrorMessage(ERR_ERROR, Err, "BenefitsToListView", "Adding Benefits to ListView", "Error in adding benefits to listview.")
  Resume BenefitsToListView_end
  Resume
End Function

Private Function IBenefitForm2_BenefitToListView(ben As IBenefitClass, ByVal lBenefitIndex As Long) As Long
  Call ECASE("BenefitToListView - not implemented")
'  Dim ibf As IBenefitForm2
'  Dim lst As ListItem
'
'  Set ibf = Me
'  If Not ben Is Nothing Then 'if this changes change in GotoScreen()
'    Set lst = lb.ListItems.Add(, "EE" & ben.value(ee_PersonnelNumber))
'    Call ibf.UpdateBenefitListViewItem(lst, ben, lBenefitIndex)
'  End If
'  IBenefitForm2_BenefitToListView = 1 ' cd don't change
End Function

Private Function IBenefitForm2_BenefitToScreen(Optional ByVal BenefitIndex As Long = -1&, Optional ByVal UpdateBenefit As Boolean = True) As Boolean
  Dim ben As IBenefitClass, oldcur As IBenefitClass
  Dim ibf As IBenefitForm2
   
  On Error GoTo Employee_BenefitToScreen_Err
  Call xSet("Employee_BenefitToScreen")
  Set ibf = Me
  IBenefitForm2_BenefitToScreen = False
  If UpdateBenefit Then Call UpdateBenefitFromTags
  Set ben = Nothing
  If BenefitIndex <> -1 Then Set ben = p11d32.CurrentEmployer.employees(BenefitIndex)
  If Not p11d32.CurrentEmployeeIsNothing Then
    If Not p11d32.CurrentEmployer.CurrentEmployee Is ben Then
      Set oldcur = p11d32.CurrentEmployer.CurrentEmployee
      Call p11d32.CurrentEmployer.LoadEmployeeEx(ben, True)
    End If
  End If
  
  If Not ben Is Nothing Then
    If ben.BenefitClass <> ibf.benclass Then Call Err.Raise(ERR_INVALIDBENCLASS, "BenefitToScreen", "Benefit type invalid")
    Set ibf.benefit = ben
    Call ibf.BenefitOn
  Else
    Set ibf.benefit = Nothing
    Call ibf.BenefitOff
  End If
  Set p11d32.CurrentEmployer.CurrentEmployee = ben
  Call SetBenefitFormState(ibf)
  IBenefitForm2_BenefitToScreen = True
   
Employee_BenefitToScreen_End:
  Set ibf = Nothing
  Set ben = Nothing
  Call xReturn("Employee_BenefitToScreen")
  Exit Function
  
Employee_BenefitToScreen_Err:
  IBenefitForm2_BenefitToScreen = False
  Call ErrorMessage(ERR_ERROR, Err, "Employee_BenefitToScreen", "Employee BenefitToScreen", "Unable to place the chosen benefit onto the screen. Benefit index = " & BenefitIndex & ".")
  Resume Employee_BenefitToScreen_End
  Resume
End Function

Private Property Let IBenefitForm2_benclass(ByVal NewValue As BEN_CLASS)
End Property

Private Property Get IBenefitForm2_benclass() As BEN_CLASS
  IBenefitForm2_benclass = BC_EMPLOYEE
End Property
Private Property Get IBenefitForm2_lv() As MSComctlLib.IListView
  Set IBenefitForm2_lv = lb
End Property

Private Function IBenefitForm2_RemoveBenefit(ByVal BenefitIndex As Long) As Boolean
  Dim NextBenefitIndex As Long
  Dim ben As IBenefitClass
  Dim ee As Employee
  Dim ibf As IBenefitForm2
  Dim li As ListItem
  
  On Error GoTo RemoveBenefit_ERR
  
  Call xSet("RemoveBenefit")
  
  Set ibf = Me
  
  Set ee = p11d32.CurrentEmployer.employees(BenefitIndex)
  If Not ee Is Nothing Then
    Call GetNextBestListItem(li, ibf.lv, ibf.lv.SelectedItem)
    Set ben = ee
    
    Call ben.DeleteDB
    Call ben.Kill
    Set p11d32.CurrentEmployer.CurrentEmployee = Nothing
    Call p11d32.CurrentEmployer.employees.Remove(BenefitIndex)
    Set ibf = Me
    ibf.lv.listitems.Remove (ibf.lv.SelectedItem.Index)
    Set ibf.lv.SelectedItem = Nothing
    Set benefit = Nothing
    If li Is Nothing Then
      Call p11d32.CurrentEmployer.employees.Compact
    End If
    Call SelectBenefitByListItem(ibf, li)
    IBenefitForm2_RemoveBenefit = True
  End If
    
RemoveBenefit_END:
  Call UpdateEmployerEmployeesCount
  Set ee = Nothing
  Set ben = Nothing
  Call xReturn("RemoveBenefit")
  Exit Function
  
RemoveBenefit_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "RemoveBenefit", "Remove Benefit", "Error removing benefit.")
  Resume RemoveBenefit_END
  Resume
End Function

Private Function IBenefitForm2_UpdateBenefitListViewItem(li As MSComctlLib.IListItem, benefit As IBenefitClass, Optional ByVal BenefitIndex As Long = 0&, Optional ByVal SelectItem As Boolean = False) As Long
  Dim ee As Employee
  
  On Error GoTo UpdateBenefitListViewItem_ERR
  Call xSet("UpdateBenefitListViewItem")
  
  If Not li Is Nothing And Not benefit Is Nothing Then
    Set ee = benefit
    If BenefitIndex > 0 Then li.Tag = BenefitIndex
    With li
      .SmallIcon = benefit.ImageListKey
      .Checked = benefit.value(ee_Selected)
      .Text = ee.FullName
      .SubItems(LV_EE_PERSONNEL_NUMBER) = ee.PersonnelNumber
      .SubItems(LV_EE_NI_NUMBER) = benefit.value(ee_NINumber_db)
      .SubItems(LV_EE_STATUS) = IIf(benefit.value(ee_Director_db), "Director", "Staff")
      .SubItems(LV_EE_GROUP1) = benefit.value(ee_Group1_db)
      .SubItems(LV_EE_GROUP2) = benefit.value(ee_Group2_db)
      .SubItems(LV_EE_GROUP3) = benefit.value(ee_Group3_db)
      If SelectItem Then Set IBenefitForm2_lv.SelectedItem = li
    End With
    IBenefitForm2_UpdateBenefitListViewItem = li.Index
  End If
  
UpdateBenefitListViewItem_END:
  Set ee = Nothing
  Call xReturn("UpdateBenefitListViewItem")
  Exit Function
  
UpdateBenefitListViewItem_ERR:
  IBenefitForm2_UpdateBenefitListViewItem = False
  Call ErrorMessage(ERR_ERROR, Err, "UpdateBenefitListViewItem", "Updating a benefits list view text", "Unable to update the benefits list view text.")
  Resume UpdateBenefitListViewItem_END
  Resume
End Function

Private Function IBenefitForm2_ValididateBenefit(ben As IBenefitClass) As Boolean

End Function

Private Function IFrmGeneral_CheckChanged(c As Control) As Boolean
  Dim lst As ListItem
  Dim ben As IBenefitClass
  Dim bDirty As Boolean
  Dim ee As Employee
  On Error GoTo CheckChanged_Err
  Call xSet("CheckChanged")
  
    If p11d32.CurrentEmployeeIsNothing Then
      GoTo CheckChanged_End
    End If
    'we are asking if the value has changed and if it is valid thus save
    With c
      Select Case .Name
        Case "TB_Data"
          Select Case .Index
            Case 0
              bDirty = CheckTextInput(.Text, benefit, ee_Title_db)
            Case 5
              bDirty = CheckTextInput(.Text, benefit, ee_Firstname_db)
            Case 4
              bDirty = CheckTextInput(.Text, benefit, ee_Initials_db)
            Case 9
              bDirty = CheckTextInput(.Text, benefit, ee_Surname_db)
            Case 10
              bDirty = CheckTextInput(.Text, benefit, ee_Salutation_db)
            Case 11
              bDirty = CheckTextInput(.Text, benefit, ee_Email_db)
            Case 12
              bDirty = CheckTextInput(.Text, benefit, ee_Password_db)
            Case 13
              bDirty = CheckTextInput(.Text, benefit, ee_Username_db)
            Case 1
              'Ignore the P_NUM field
            Case 2
              bDirty = CheckTextInput(.Text, benefit, ee_NINumber_db)
            Case 3
              bDirty = CheckTextInput(.Text, benefit, ee_Group1_db)
            Case 6
              bDirty = CheckTextInput(.Text, benefit, ee_Group2_db)
            Case 7
              bDirty = CheckTextInput(.Text, benefit, ee_Comments_db)
            Case 8
              bDirty = CheckTextInput(.Text, benefit, ee_Group3_db)
            Case Else
              ECASE "Unknown control"
          End Select
        Case "TB_Date"
          Select Case .Index
            Case 1
              bDirty = CheckTextInput(.Text, benefit, ee_joined_db)
            Case 2
              bDirty = CheckTextInput(.Text, benefit, ee_left_db)
          End Select
          
        'FC - Class1A
        Case "chk_Class1AEmployeeIsNotSubjectTo"
          bDirty = CheckCheckBoxInput(.value, benefit, ee_Class1AEmployeeIsNotSubjectTo_db)
          If bDirty Then
            'need to tell all my benefits to require recalc
            Set ee = benefit
            Call ee.EnumBenefitsDoAction(True, BET_NEED_TO_CALCULATE)
          End If
        Case "CB_Status"
          bDirty = StrComp(IIf(benefit.value(ee_Director_db), S_DIRECTOR, S_STAFF), c.Text)
          If StrComp(c.Text, S_DIRECTOR) = 0 Then
            benefit.value(ee_Director_db) = True
          Else
            benefit.value(ee_Director_db) = False
          End If
        Case Else
          ECASE ("UNKNOWN CONTROL")
     End Select
     IFrmGeneral_CheckChanged = AfterCheckChanged(c, Me, bDirty)
     If bDirty Then Call MDIMain.NavigateBarUpdate(benefit)
    End With
  
CheckChanged_End:
  Set ben = Nothing
  Set lst = Nothing
  Call xReturn("CheckChanged")
  Exit Function
  
CheckChanged_Err:
  IFrmGeneral_CheckChanged = False
  Call ErrorMessage(ERR_ERROR, Err, "CheckChanged", "ERR_CHECKCHANGED", "This function has failed for the form " & Me.Name & ".")
  Resume CheckChanged_End
  Resume
End Function

Public Property Get IFrmGeneral_InvalidVT() As Control
  Set IFrmGeneral_InvalidVT = m_InvalidVT
End Property

Public Property Set IFrmGeneral_InvalidVT(NewValue As Control)
  Set m_InvalidVT = NewValue
End Property

Private Sub Form_Load()
  Set mclsResize = New clsFormResize
  HOOK.hWnd = lb.hWnd
  HOOK.Messages(WM_LBUTTONDOWN) = True
  
  If Not (mclsResize.InitResize(Me, L_DES_HEIGHT, L_DES_WIDTH, DESIGN, , , MDIMain)) Then ', DESIGN)) Then
    Err.Raise ERR_Application
  End If
  CB_Status.Clear
  CB_Status.AddItem S_STAFF
  CB_Status.AddItem S_DIRECTOR
  Call MDIMain.ClearConfirmUndo
  Call SetDefaultVTDate(TB_Date(1), , , True)
  Call SetDefaultVTDate(TB_Date(2))
  Call SetSortOrder(lb, lb.ColumnHeaders(p11d32.EmployeeSortOrderColumn + 1), p11d32.EmployeeSortOrder)
  lb.ColumnHeaders(L_LV_COL_INDEX_EMPLOYEE_REFERENCE).Tag = TYPE_LONG
  
  
End Sub

Private Sub lb_DblClick()
  Call BenefitToolBar(1, GetEmployeeIndexFromSelectedEmployee)
End Sub

Public Sub LB_ItemCheck(ByVal Item As MSComctlLib.ListItem)
  Call p11d32.CurrentEmployer.EmployeeCheck(Item.Checked, Item.Tag)
End Sub

Private Sub LB_ItemClick(ByVal Item As MSComctlLib.ListItem)
  Call IBenefitForm2_BenefitToScreen(Item.Tag)
End Sub
Public Sub FoundEmployee(ByVal KeyCode As Integer, ByVal KeyAscii As Integer)
  Dim lRet As Long
  Dim ibf As IBenefitForm2
  
  On Error GoTo FoundEmployee_ERR
  
  lb.SetFocus
  
  lRet = ListViewFastKey(lb, p11d32.EmployeeSortOrderColumn, KeyCode, KeyAscii, m_LastSearch)
  If lRet > 0 Then Call LB_ItemClick(lb.listitems(lRet))
  
FoundEmployee_END:

  Exit Sub
FoundEmployee_ERR:
  Call ErrorMessage(ERR_ERROR, Err, "Found Employee", "Found Employee", "Error finding employees from the employee screen.")
  Resume FoundEmployee_END
End Sub
Private Sub LB_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF3 Then
    Call FoundEmployee(KeyCode, 0) 'EK 1/04 TTP#12
    KeyCode = 0
  End If
End Sub

Private Sub lb_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    Call lb_DblClick
  Else
    Call FoundEmployee(0, KeyAscii) 'EK 1/04 TTP#12
    KeyAscii = 0
  End If
End Sub

Private Sub Form_Resize()
  mclsResize.Resize
  Call ColumnWidths(Me.lb, L_NAME_COL, L_REFERENCE_COL, L_NINUMBER_COL&, L_STATUS_COL&, L_GROUP1_COL&, L_GROUP2_COL&, L_GROUP3_COL&)
End Sub
Private Sub lb_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
  Call SetSortOrderEmployees(lb, ColumnHeader)
  Call MDIMain.NavigateBarUpdate(p11d32.CurrentEmployer.CurrentEmployee)
End Sub

Private Sub TB_Data_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  TB_Data(Index).Tag = SetChanged
End Sub

Private Sub TB_Data_UserValidate(Index As Integer, Valid As Boolean, Message As String, sTextEntered As String)
  If Index = L_NI_NUMBER_TEXT_BOX_INDEX Then
    sTextEntered = Trim$(sTextEntered)
    If (Len(sTextEntered) = 0) Then
      Valid = True
      Exit Sub
    Else
      If p11d32.ValidateNINumberOnEmployeeScreen Then
        Valid = ValidateNI(sTextEntered, True)
      Else
        Valid = True
      End If
    End If
    If Not Valid Then Message = "NI Number is invalid"
  End If
End Sub

Private Sub TB_Data_Validate(Index As Integer, Cancel As Boolean)
  Call IFrmGeneral_CheckChanged(TB_Data(Index))
End Sub
Private Sub TB_Date_FieldInvalid(Index As Integer, Valid As Boolean, Message As String)
  Call SetPanel2(Message)
End Sub
Private Sub TB_Date_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  TB_Date(Index).Tag = SetChanged
End Sub
Private Sub TB_Date_Validate(Index As Integer, Cancel As Boolean)
  Call IFrmGeneral_CheckChanged(TB_Date(Index))
End Sub
Public Sub UpdateGroupCodeLables()
  lblGroupCode1.Caption = p11d32.CurrentEmployer.GroupCode1Alias
  lblGroupCode2.Caption = p11d32.CurrentEmployer.GroupCode2Alias
  lblGroupCode3.Caption = p11d32.CurrentEmployer.GroupCode3Alias
  lb.ColumnHeaders(L_LV_COL_INDEX_EMPLOYEE_GROUP1).Text = lblGroupCode1.Caption
  lb.ColumnHeaders(L_LV_COL_INDEX_EMPLOYEE_GROUP2).Text = lblGroupCode2.Caption
  lb.ColumnHeaders(L_LV_COL_INDEX_EMPLOYEE_GROUP3).Text = lblGroupCode3.Caption
  With p11d32
    .BenDataLinkFieldDetails(BC_EMPLOYEE, ee_Group1_db).Description = .CurrentEmployer.GroupCode1Alias
    .BenDataLinkFieldDetails(BC_EMPLOYEE, ee_Group2_db).Description = .CurrentEmployer.GroupCode2Alias
    .BenDataLinkFieldDetails(BC_EMPLOYEE, ee_Group3_db).Description = .CurrentEmployer.GroupCode3Alias
  End With
  With MDIMain
    .mnuViewSelectGroup1.Caption = p11d32.CurrentEmployer.GroupCode1Alias
    .mnuViewGroupSortByGroup1.Caption = p11d32.CurrentEmployer.GroupCode1Alias
    
    .mnuViewSelectGroup2.Caption = p11d32.CurrentEmployer.GroupCode2Alias
    .mnuViewGroupSortByGroup2.Caption = p11d32.CurrentEmployer.GroupCode2Alias
    
    .mnuViewSelectGroup3.Caption = p11d32.CurrentEmployer.GroupCode3Alias
    .mnuViewGroupSortByGroup3.Caption = p11d32.CurrentEmployer.GroupCode3Alias
    
    
  End With
End Sub
