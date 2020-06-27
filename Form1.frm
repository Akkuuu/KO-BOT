VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.4#0"; "TASARIM.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8250
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4740
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8250
   ScaleWidth      =   4740
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   2400
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   80
      TabStop         =   0   'False
      Top             =   7680
      Width           =   495
      Begin VB.Timer Recv 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   0
         Top             =   0
      End
   End
   Begin XtremeSuiteControls.PushButton PushButton8 
      Height          =   495
      Left            =   2880
      TabIndex        =   79
      Top             =   7680
      Width           =   1815
      _Version        =   851972
      _ExtentX        =   3201
      _ExtentY        =   873
      _StockProps     =   79
      Caption         =   "Connect"
      UseVisualStyle  =   -1  'True
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Segoe Condensed"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   73
      Text            =   "Knight OnLine Client"
      Top             =   7680
      Width           =   2175
   End
   Begin VB.Timer Timer4 
      Interval        =   1000
      Left            =   5640
      Top             =   8160
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   5160
      Top             =   8160
   End
   Begin XtremeSuiteControls.TabControl TabControl1 
      Height          =   7575
      Left            =   0
      TabIndex        =   32
      Top             =   0
      Width           =   4695
      _Version        =   851972
      _ExtentX        =   8281
      _ExtentY        =   13361
      _StockProps     =   68
      ItemCount       =   4
      SelectedItem    =   2
      Item(0).Caption =   "Main"
      Item(0).ControlCount=   22
      Item(0).Control(0)=   "Check1"
      Item(0).Control(1)=   "ComboSpeed"
      Item(0).Control(2)=   "CheckSpeed"
      Item(0).Control(3)=   "Frame2"
      Item(0).Control(4)=   "Frame1"
      Item(0).Control(5)=   "chkPeriKutu"
      Item(0).Control(6)=   "chkWallHack"
      Item(0).Control(7)=   "Label5"
      Item(0).Control(8)=   "Label4"
      Item(0).Control(9)=   "Label3"
      Item(0).Control(10)=   "Label2"
      Item(0).Control(11)=   "Label1(0)"
      Item(0).Control(12)=   "PushButton3"
      Item(0).Control(13)=   "PushButton4"
      Item(0).Control(14)=   "PushButton5"
      Item(0).Control(15)=   "PushButton6"
      Item(0).Control(16)=   "PushButton7"
      Item(0).Control(17)=   "Text5"
      Item(0).Control(18)=   "Command4"
      Item(0).Control(19)=   "Command10"
      Item(0).Control(20)=   "PushButton9"
      Item(0).Control(21)=   "PushButton11"
      Item(1).Caption =   "Boss Search"
      Item(1).ControlCount=   9
      Item(1).Control(0)=   "lstMap"
      Item(1).Control(1)=   "lstMobName"
      Item(1).Control(2)=   "txtMobName"
      Item(1).Control(3)=   "MobSearchBut"
      Item(1).Control(4)=   "MobSearch(0)"
      Item(1).Control(5)=   "MobSearch(1)"
      Item(1).Control(6)=   "MobSearch(2)"
      Item(1).Control(7)=   "MobSearch(3)"
      Item(1).Control(8)=   "lwMob"
      Item(2).Caption =   "Pazar"
      Item(2).ControlCount=   12
      Item(2).Control(0)=   "PazarList"
      Item(2).Control(1)=   "ProgressBar1"
      Item(2).Control(2)=   "Command6"
      Item(2).Control(3)=   "Command5"
      Item(2).Control(4)=   "List1"
      Item(2).Control(5)=   "PushButton12"
      Item(2).Control(6)=   "Text6"
      Item(2).Control(7)=   "Text7"
      Item(2).Control(8)=   "Text8"
      Item(2).Control(9)=   "Text9"
      Item(2).Control(10)=   "Text10"
      Item(2).Control(11)=   "Text11"
      Item(3).Caption =   "Upgrade"
      Item(3).ControlCount=   9
      Item(3).Control(0)=   "canta"
      Item(3).Control(1)=   "Label1(1)"
      Item(3).Control(2)=   "Combo1"
      Item(3).Control(3)=   "PushButton10"
      Item(3).Control(4)=   "PushButton2"
      Item(3).Control(5)=   "PushButton1"
      Item(3).Control(6)=   "Text21"
      Item(3).Control(7)=   "Text20"
      Item(3).Control(8)=   "Label13"
      Begin VB.TextBox Text11 
         Height          =   375
         Left            =   2760
         TabIndex        =   92
         Text            =   "Text11"
         Top             =   6960
         Width           =   1695
      End
      Begin VB.TextBox Text10 
         Height          =   375
         Left            =   1680
         TabIndex        =   91
         Text            =   "Text10"
         Top             =   6600
         Width           =   1695
      End
      Begin VB.TextBox Text9 
         Height          =   375
         Left            =   1680
         TabIndex        =   90
         Text            =   "Text9"
         Top             =   6240
         Width           =   1695
      End
      Begin VB.TextBox Text8 
         Height          =   375
         Left            =   240
         TabIndex        =   89
         Text            =   "Text8"
         Top             =   6960
         Width           =   2175
      End
      Begin VB.TextBox Text7 
         Height          =   375
         Left            =   240
         TabIndex        =   88
         Text            =   "Text7"
         Top             =   6600
         Width           =   1335
      End
      Begin VB.TextBox Text6 
         Height          =   375
         Left            =   240
         TabIndex        =   87
         Text            =   "Text6"
         Top             =   6240
         Width           =   1335
      End
      Begin XtremeSuiteControls.PushButton PushButton12 
         Height          =   495
         Left            =   3360
         TabIndex        =   86
         Top             =   4440
         Width           =   1095
         _Version        =   851972
         _ExtentX        =   1931
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "Yerlestir"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton PushButton11 
         Height          =   375
         Left            =   -67000
         TabIndex        =   85
         Top             =   5400
         Visible         =   0   'False
         Width           =   1575
         _Version        =   851972
         _ExtentX        =   2778
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Oto Kutu"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton PushButton9 
         Height          =   375
         Left            =   -67000
         TabIndex        =   84
         Top             =   2640
         Visible         =   0   'False
         Width           =   1575
         _Version        =   851972
         _ExtentX        =   2778
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Relog"
         UseVisualStyle  =   -1  'True
      End
      Begin VB.CommandButton Command10 
         Caption         =   "GetItemID slot1-2-3"
         Height          =   480
         Left            =   -68800
         TabIndex        =   83
         Top             =   1080
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Command4"
         Height          =   480
         Left            =   -69880
         TabIndex        =   82
         Top             =   1080
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox Text5 
         Height          =   375
         Left            =   -69880
         TabIndex        =   81
         Text            =   "Text5"
         Top             =   1680
         Visible         =   0   'False
         Width           =   2415
      End
      Begin XtremeSuiteControls.PushButton PushButton3 
         Height          =   495
         Left            =   -67000
         TabIndex        =   74
         Top             =   3960
         Visible         =   0   'False
         Width           =   1575
         _Version        =   851972
         _ExtentX        =   2778
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "AP SC"
         UseVisualStyle  =   -1  'True
      End
      Begin VB.Timer Timer6 
         Interval        =   2000
         Left            =   4200
         Top             =   5640
      End
      Begin VB.TextBox Text20 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -68680
         TabIndex        =   71
         Text            =   "27"
         Top             =   6720
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox Text21 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -68680
         TabIndex        =   70
         Top             =   6960
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.ComboBox Combo1 
         BackColor       =   &H80000004&
         Height          =   315
         ItemData        =   "Form1.frx":0000
         Left            =   -69400
         List            =   "Form1.frx":0010
         Style           =   2  'Dropdown List
         TabIndex        =   66
         Top             =   6000
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.ListBox List1 
         Height          =   1035
         ItemData        =   "Form1.frx":0051
         Left            =   120
         List            =   "Form1.frx":0053
         TabIndex        =   63
         Top             =   5040
         Width           =   4335
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Tarat"
         Height          =   480
         Left            =   120
         TabIndex        =   62
         Top             =   4440
         Width           =   1575
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Pazar Oku"
         Height          =   480
         Left            =   1800
         TabIndex        =   61
         Top             =   4440
         Width           =   1455
      End
      Begin VB.TextBox txtMobName 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   495
         Left            =   -69880
         TabIndex        =   52
         Top             =   3720
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.ListBox lstMobName 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   3150
         ItemData        =   "Form1.frx":0055
         Left            =   -67360
         List            =   "Form1.frx":0057
         TabIndex        =   51
         Top             =   480
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.ListBox lstMap 
         Appearance      =   0  'Flat
         Height          =   3150
         ItemData        =   "Form1.frx":0059
         Left            =   -69880
         List            =   "Form1.frx":005B
         TabIndex        =   50
         Top             =   480
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.CheckBox chkWallHack 
         BackColor       =   &H80000004&
         Caption         =   "WallHack"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -69880
         TabIndex        =   44
         Top             =   4680
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CheckBox chkPeriKutu 
         BackColor       =   &H80000004&
         Caption         =   "PeriKutu"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -69880
         TabIndex        =   43
         Top             =   4440
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Frame Frame1 
         Caption         =   "Zone Change"
         Height          =   1575
         Left            =   -67240
         TabIndex        =   40
         Top             =   5880
         Visible         =   0   'False
         Width           =   1815
         Begin VB.TextBox Text2 
            Appearance      =   0  'Flat
            Height          =   375
            Left            =   120
            TabIndex        =   42
            Text            =   "NO"
            Top             =   360
            Width           =   1575
         End
         Begin VB.CommandButton cmdPaket 
            Caption         =   "Paket"
            Height          =   600
            Left            =   120
            TabIndex        =   41
            Top             =   840
            Width           =   1575
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Kendine Trade"
         Height          =   1575
         Left            =   -69880
         TabIndex        =   36
         Top             =   5880
         Visible         =   0   'False
         Width           =   1695
         Begin VB.CommandButton Command2 
            Caption         =   "Kendine Trade At"
            Height          =   360
            Left            =   120
            TabIndex        =   39
            Top             =   360
            Width           =   1455
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Kabul et"
            Height          =   360
            Left            =   120
            TabIndex        =   38
            Top             =   720
            Width           =   1455
         End
         Begin VB.CommandButton Command3 
            Caption         =   "Trade OK"
            Height          =   360
            Left            =   120
            TabIndex        =   37
            Top             =   1080
            Width           =   1455
         End
      End
      Begin VB.CheckBox CheckSpeed 
         Caption         =   "Check speed"
         Height          =   255
         Left            =   -69880
         TabIndex        =   35
         Top             =   5040
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.ComboBox ComboSpeed 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "Form1.frx":005D
         Left            =   -69880
         List            =   "Form1.frx":0076
         Style           =   2  'Dropdown List
         TabIndex        =   34
         Top             =   5400
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.CheckBox Check1 
         Caption         =   "NODC"
         Height          =   255
         Left            =   -69880
         TabIndex        =   33
         Top             =   4200
         Visible         =   0   'False
         Width           =   975
      End
      Begin XtremeSuiteControls.PushButton MobSearchBut 
         Height          =   495
         Left            =   -67360
         TabIndex        =   53
         Top             =   4680
         Visible         =   0   'False
         Width           =   1815
         _Version        =   851972
         _ExtentX        =   3201
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "Search the list"
         BackColor       =   -2147483644
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton MobSearch 
         Height          =   855
         Index           =   0
         Left            =   -69880
         TabIndex        =   54
         Top             =   4320
         Visible         =   0   'False
         Width           =   2175
         _Version        =   851972
         _ExtentX        =   3836
         _ExtentY        =   1508
         _StockProps     =   79
         Caption         =   "Search the map!"
         BackColor       =   -2147483644
         Appearance      =   6
         EnableMarkup    =   -1  'True
         BorderGap       =   0
      End
      Begin XtremeSuiteControls.PushButton MobSearch 
         Height          =   495
         Index           =   1
         Left            =   -68560
         TabIndex        =   55
         Top             =   3720
         Visible         =   0   'False
         Width           =   735
         _Version        =   851972
         _ExtentX        =   1296
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "Add"
         BackColor       =   -2147483644
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton MobSearch 
         Height          =   495
         Index           =   2
         Left            =   -67360
         TabIndex        =   56
         Top             =   3720
         Visible         =   0   'False
         Width           =   1815
         _Version        =   851972
         _ExtentX        =   3201
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "Clear"
         BackColor       =   -2147483644
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton MobSearch 
         Height          =   495
         Index           =   3
         Left            =   -67360
         TabIndex        =   57
         Top             =   4200
         Visible         =   0   'False
         Width           =   1815
         _Version        =   851972
         _ExtentX        =   3201
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "Get X, Y"
         BackColor       =   -2147483644
         UseVisualStyle  =   -1  'True
      End
      Begin MSComctlLib.ListView lwMob 
         Height          =   2175
         Left            =   -69880
         TabIndex        =   58
         Top             =   5280
         Visible         =   0   'False
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   3836
         View            =   3
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   0   'False
         HideSelection   =   0   'False
         HideColumnHeaders=   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         TextBackground  =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "ID"
            Object.Width           =   970
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Name"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "HP"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "X"
            Object.Width           =   954
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Y"
            Object.Width           =   954
         EndProperty
      End
      Begin MSComctlLib.ListView PazarList 
         Height          =   3495
         Left            =   120
         TabIndex        =   59
         Top             =   480
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   6165
         View            =   3
         Arrange         =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         PictureAlignment=   3
         _Version        =   393217
         Icons           =   "ImageList1"
         SmallIcons      =   "ImageList1"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Karakter Adi"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Item Adi"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Item Fiyati"
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "X"
            Object.Width           =   882
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Y"
            Object.Width           =   882
         EndProperty
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   375
         Left            =   120
         TabIndex        =   60
         Top             =   3960
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   661
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
      End
      Begin XtremeSuiteControls.ListBox canta 
         Height          =   5175
         Left            =   -69880
         TabIndex        =   64
         Top             =   480
         Visible         =   0   'False
         Width           =   4455
         _Version        =   851972
         _ExtentX        =   7858
         _ExtentY        =   9128
         _StockProps     =   77
         BackColor       =   16777215
         BackColor       =   16777215
         MultiSelect     =   1
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.PushButton PushButton10 
         Height          =   375
         Left            =   -67120
         TabIndex        =   67
         Top             =   5760
         Visible         =   0   'False
         Width           =   1695
         _Version        =   851972
         _ExtentX        =   2990
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Yenile"
         BackColor       =   -2147483644
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton PushButton2 
         Height          =   615
         Left            =   -67120
         TabIndex        =   68
         Top             =   6840
         Visible         =   0   'False
         Width           =   1695
         _Version        =   851972
         _ExtentX        =   2990
         _ExtentY        =   1085
         _StockProps     =   79
         Caption         =   "Hizli Upgrade"
         BackColor       =   -2147483644
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton PushButton1 
         Height          =   735
         Left            =   -67120
         TabIndex        =   69
         Top             =   6120
         Visible         =   0   'False
         Width           =   1695
         _Version        =   851972
         _ExtentX        =   2990
         _ExtentY        =   1296
         _StockProps     =   79
         Caption         =   "Upgrade"
         BackColor       =   -2147483644
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton PushButton4 
         Height          =   495
         Left            =   -67000
         TabIndex        =   75
         Top             =   4440
         Visible         =   0   'False
         Width           =   1575
         _Version        =   851972
         _ExtentX        =   2778
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "DEF SC"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton PushButton5 
         Height          =   495
         Left            =   -67000
         TabIndex        =   76
         Top             =   4920
         Visible         =   0   'False
         Width           =   1575
         _Version        =   851972
         _ExtentX        =   2778
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "DROP SC"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton PushButton6 
         Height          =   495
         Left            =   -67000
         TabIndex        =   77
         Top             =   3480
         Visible         =   0   'False
         Width           =   1575
         _Version        =   851972
         _ExtentX        =   2778
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "Attack SC"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton PushButton7 
         Height          =   495
         Left            =   -67000
         TabIndex        =   78
         Top             =   3000
         Visible         =   0   'False
         Width           =   1575
         _Version        =   851972
         _ExtentX        =   2778
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "Quest Bug"
         UseVisualStyle  =   -1  'True
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Scroll Slot NO: "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -69880
         TabIndex        =   72
         Top             =   6720
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000004&
         BackStyle       =   0  'Transparent
         Caption         =   "SC :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   -69880
         TabIndex        =   65
         Top             =   6000
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name :"
         BeginProperty Font 
            Name            =   "Segoe UI Semibold"
            Size            =   8.25
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   -69880
         TabIndex        =   49
         Top             =   480
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NameShow"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   -69280
         TabIndex        =   48
         Top             =   480
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label3"
         Height          =   195
         Left            =   -67960
         TabIndex        =   47
         Top             =   6600
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label4"
         Height          =   195
         Left            =   -67960
         TabIndex        =   46
         Top             =   6840
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label5"
         Height          =   255
         Left            =   -67960
         TabIndex        =   45
         Top             =   7080
         Visible         =   0   'False
         Width           =   495
      End
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Command9"
      Height          =   600
      Left            =   5280
      TabIndex        =   31
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Timer Timer5 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   6240
      Top             =   360
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Warp"
      Height          =   360
      Left            =   6960
      TabIndex        =   30
      Top             =   3240
      Width           =   990
   End
   Begin VB.ComboBox copennpc 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "Form1.frx":009A
      Left            =   5280
      List            =   "Form1.frx":00BF
      Style           =   2  'Dropdown List
      TabIndex        =   29
      Top             =   2880
      Width           =   2655
   End
   Begin VB.ListBox List5 
      Height          =   645
      Left            =   5280
      TabIndex        =   28
      Top             =   2040
      Width           =   1215
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   5280
      TabIndex        =   27
      Text            =   "Text4"
      Top             =   840
      Width           =   1095
   End
   Begin VB.ListBox List4 
      Height          =   1035
      Left            =   8040
      TabIndex        =   26
      Top             =   3240
      Width           =   2535
   End
   Begin VB.ListBox List3 
      Height          =   4935
      Left            =   19080
      TabIndex        =   25
      Top             =   4800
      Width           =   2295
   End
   Begin VB.CommandButton Command7 
      Caption         =   "pointerbul"
      Height          =   360
      Left            =   20400
      TabIndex        =   24
      Top             =   9840
      Width           =   990
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   8040
      TabIndex        =   23
      Text            =   "Text3"
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Timer pazaroku 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   5760
      Top             =   360
   End
   Begin VB.ListBox List2 
      Height          =   1230
      Left            =   8040
      TabIndex        =   22
      Top             =   1680
      Width           =   2535
   End
   Begin VB.Timer tmroku 
      Enabled         =   0   'False
      Interval        =   20
      Left            =   5280
      Top             =   360
   End
   Begin VB.Timer tmrFind2 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   6120
      Top             =   8640
   End
   Begin VB.Timer tmrFind 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   6120
      Top             =   8160
   End
   Begin XtremeSuiteControls.GroupBox GroupBox6 
      Height          =   4575
      Left            =   18960
      TabIndex        =   0
      Top             =   120
      Width           =   2415
      _Version        =   851972
      _ExtentX        =   4260
      _ExtentY        =   8070
      _StockProps     =   79
      Caption         =   "Boss Check"
      BackColor       =   -2147483644
      UseVisualStyle  =   -1  'True
      Begin VB.Timer OtoSaatKayýt 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   120
         Top             =   4080
      End
      Begin VB.Timer Timer3 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   120
         Top             =   2160
      End
      Begin VB.Timer frm5c 
         Interval        =   1000
         Left            =   120
         Top             =   3600
      End
      Begin VB.Timer Alarm2 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   120
         Top             =   1200
      End
      Begin VB.Timer Hpleriekle 
         Interval        =   1000
         Left            =   120
         Top             =   1680
      End
      Begin VB.Timer Alarm 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   120
         Top             =   720
      End
      Begin VB.Timer AlarmCal 
         Enabled         =   0   'False
         Interval        =   2000
         Left            =   120
         Top             =   240
      End
      Begin VB.Timer ara 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   120
         Top             =   3120
      End
      Begin VB.Timer Timer2 
         Enabled         =   0   'False
         Interval        =   2000
         Left            =   120
         Top             =   2640
      End
      Begin VB.ListBox Hpler 
         Appearance      =   0  'Flat
         Height          =   1785
         Left            =   600
         TabIndex        =   4
         Top             =   2640
         Width           =   1575
      End
      Begin VB.TextBox HpText 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   600
         TabIndex        =   3
         Text            =   "50000"
         Top             =   2280
         Width           =   735
      End
      Begin VB.TextBox txtFid 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   600
         TabIndex        =   2
         Text            =   "10000"
         Top             =   1920
         Width           =   735
      End
      Begin VB.TextBox txtLid 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   600
         TabIndex        =   1
         Text            =   "30000"
         Top             =   1560
         Width           =   735
      End
   End
   Begin VB.Label Labelboss 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100"
      Height          =   255
      Index           =   8
      Left            =   1080
      TabIndex        =   21
      Top             =   10080
      Width           =   615
   End
   Begin VB.Label Labelboss 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100"
      Height          =   255
      Index           =   7
      Left            =   1080
      TabIndex        =   20
      Top             =   9720
      Width           =   615
   End
   Begin VB.Label Labelboss 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100"
      Height          =   255
      Index           =   6
      Left            =   1080
      TabIndex        =   19
      Top             =   9360
      Width           =   615
   End
   Begin VB.Label Labelboss 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100"
      Height          =   255
      Index           =   5
      Left            =   1080
      TabIndex        =   18
      Top             =   9000
      Width           =   615
   End
   Begin VB.Label Labelboss 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100"
      Height          =   255
      Index           =   4
      Left            =   1080
      TabIndex        =   17
      Top             =   8760
      Width           =   615
   End
   Begin VB.Label Labelboss 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100"
      Height          =   255
      Index           =   2
      Left            =   1800
      TabIndex        =   16
      Top             =   9480
      Width           =   615
   End
   Begin VB.Label Labelboss 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100"
      Height          =   255
      Index           =   1
      Left            =   1800
      TabIndex        =   15
      Top             =   9120
      Width           =   615
   End
   Begin VB.Label Labelboss 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100"
      Height          =   255
      Index           =   0
      Left            =   1800
      TabIndex        =   14
      Top             =   8880
      Width           =   615
   End
   Begin VB.Label Labelboss 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100"
      Height          =   255
      Index           =   3
      Left            =   1080
      TabIndex        =   13
      Top             =   8520
      Width           =   615
   End
   Begin VB.Label Labelboss 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100"
      Height          =   255
      Index           =   9
      Left            =   2400
      TabIndex        =   12
      Top             =   8760
      Width           =   615
   End
   Begin VB.Label Labelboss 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100"
      Height          =   255
      Index           =   10
      Left            =   2400
      TabIndex        =   11
      Top             =   9000
      Width           =   615
   End
   Begin VB.Label Labelboss 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100"
      Height          =   255
      Index           =   11
      Left            =   2400
      TabIndex        =   10
      Top             =   9240
      Width           =   615
   End
   Begin VB.Label Labelboss 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100"
      Height          =   255
      Index           =   12
      Left            =   2400
      TabIndex        =   9
      Top             =   9480
      Width           =   615
   End
   Begin VB.Label Labelboss 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100"
      Height          =   255
      Index           =   13
      Left            =   2400
      TabIndex        =   8
      Top             =   9720
      Width           =   615
   End
   Begin VB.Label Labelboss 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100"
      Height          =   255
      Index           =   14
      Left            =   2400
      TabIndex        =   7
      Top             =   9960
      Width           =   615
   End
   Begin VB.Label Labelboss 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100"
      Height          =   255
      Index           =   15
      Left            =   1920
      TabIndex        =   6
      Top             =   9960
      Width           =   615
   End
   Begin VB.Label Labelboss 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100"
      Height          =   255
      Index           =   16
      Left            =   1800
      TabIndex        =   5
      Top             =   8640
      Width           =   615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Check1_Click()
WriteLong KO_NODC, 1
End Sub

Private Sub CheckSpeed_Click()
If CheckSpeed.value = 1 Then SpeedLe Else SpeedAl
End Sub

Private Sub chkPeriKutu_Click()
If chkPeriKutu.value = 1 Then
    periaç
    ByteYaz (LongOku(LongOku(KO_PTR_CHR) + &H58)) + &H5C4, 1
    ByteYaz (LongOku(LongOku(KO_PTR_CHR) + &H58)) + &H5C6, 1
    ByteYaz (LongOku(LongOku(KO_PTR_CHR) + &H58)) + &H5C7, 1
Else
    ByteYaz (LongOku(LongOku(KO_PTR_CHR) + &H58)) + &H5C4, 0
    ByteYaz (LongOku(LongOku(KO_PTR_CHR) + &H58)) + &H5C6, 0
    ByteYaz (LongOku(LongOku(KO_PTR_CHR) + &H58)) + &H5C7, 0
End If
End Sub

Private Sub chkWallHack_Click()
If WallHackCheck.value = 1 Then
WriteLong KO_ADR_CHR + KO_OFF_WH, "0"
Else
WriteLong KO_ADR_CHR + KO_OFF_WH, "1"
End If
End Sub







Private Sub ComboSpeed_Click()
If CheckSpeed.value = 1 Then
SeriHiz = 16256 + Val(ComboSpeed.Text)
SeriTimer = 1.5 - (FormatNumber(Val(ComboSpeed.Text) / 185, 1) + 0.1)
SpeedLe
End If
End Sub

Private Sub Command1_Click()
Paket "300201"
End Sub

Private Sub Command10_Click()
Text5.Text = GetItemID(7)
End Sub

Private Sub Command11_Click()
Form1.Width = 16515
End Sub

Private Sub Command2_Click()
Paket "3001" + CharId + "01"
End Sub

Private Sub Command3_Click()
Paket "3005"
End Sub



Private Sub Command4_Click()
ItemIDbul
End Sub

Private Sub Command5_Click()
'pazaroku.Enabled = True
tmroku.Enabled = True
End Sub

Private Sub Command6_Click()
'Dim L As Long
'Dim MobID As Long
'
'If lstMobName.ListCount > 0 Then
'        lwMob.ListItems.Clear
'
'        For i = 0 To lstMobName.ListCount - 1
'            For L = 0 To lstMap.ListCount - 1
'                If InStr(LCase(lstMap.List(L)), LCase(lstMobName.List(i))) > 0 Then
'
'                        MobID = lstMap.ItemData(L)
'                        MobBase = GetTargetBase(MobID)
'                        Paket "1D0100" & FormatHex(Hex(MobID), 4)
'                        Paket "22" & FormatHex(Hex(MobID), 4)
'                        If ReadStringAuto(MobBase + KO_OFF_NAME) <> "" Then
'                            Set lstview = lwMob.ListItems.Add(, , MobID)
'                            lstview.ListSubItems.Add , , ReadStringAuto(MobBase + KO_OFF_NAME)
'                            lstview.ListSubItems.Add , , ReadLong(MobBase + KO_OFF_HP)
'                            lstview.ListSubItems.Add , , ReadFloat(MobBase + KO_OFF_X)
'                            lstview.ListSubItems.Add , , ReadFloat(MobBase + KO_OFF_Y)
'                        End If
'
'                End If
'            Next
'        Next
'    End If
pazaroku.Enabled = True
End Sub

Private Sub Command7_Click()
List3.AddItem "KO_FLPZ = " & "&H" & Hex(FindPointer("CCCCCCCCCCCC8BC18B0D", &H401000, &H5B2000)) '
List3.AddItem Hex(KO_FLDB + &H40 + &H4 + &H4)
End Sub



Private Sub Command8_Click()
Dim pbytes(1 To 3) As Byte
Select Case copennpc.Text
    Case "MORADON"
        pbytes(1) = &H4B
        pbytes(2) = &HD6
        pbytes(3) = &H0
    Case "ABYSS"
        pbytes(1) = &H4B
        pbytes(2) = &HEF
        pbytes(3) = &HB
    Case "DELOS"
        pbytes(1) = &H4B
        pbytes(2) = &H4D
        pbytes(3) = &H8
    Case "RONARK LAND"
        pbytes(1) = &H4B
        pbytes(2) = &HD5
        pbytes(3) = &H0
    Case "ARDREAM"
    Case "EL MORAD CASTLE"
        pbytes(1) = &H4B
        pbytes(2) = &HD1
        pbytes(3) = &H1B
    Case "DODA CAMP"
        pbytes(1) = &H4B
        pbytes(2) = &HD7
        pbytes(3) = &H0
    Case "KALLUGA"
        pbytes(1) = &H4B
        pbytes(2) = &HE8
        pbytes(3) = &H0
    Case "ASGA"
        pbytes(1) = &H4B
        pbytes(2) = &HD3
        pbytes(3) = &H0
    Case "RAIBA"
        pbytes(1) = &H4B
        pbytes(2) = &HDE
        pbytes(3) = &H0
    Case "RUNAR GATE"
        pbytes(1) = &H4B
        pbytes(2) = &H46
        pbytes(3) = &H8
    Case "ESLAND (HUMAN)"
        pbytes(1) = &H4B
        pbytes(2) = &HFE
        pbytes(3) = &H0
    Case "LUFERSON"
    Case "LAON CAMP"
    Case "KALLUGA"
    Case "BELLUA"
    Case "LINATE"
    Case "RUNAR GATE"
    Case "ESLAND (ORC)"
End Select
If pbytes(1) <> 0 Then Paket "pBytes"
End Sub

Private Sub Command9_Click()
'Dim L As Long
'Dim MobID As Long
'
'
'                            Set lstview = lwMob.ListItems.Add(, , MobID)
'                            lstview.ListSubItems.Add , , ReadStringAuto(FormatHex(Hex(List1.ItemData(List1.ListIndex)), 4) + KO_ITOB)
'                            lstview.ListSubItems.Add , , ReadLong(FormatHex(Hex(List1.ItemData(List1.ListIndex)), 4) + KO_OFF_ITEMBASE)
'                            lstview.ListSubItems.Add , , ReadFloat(FormatHex(Hex(List1.ItemData(List1.ListIndex)), 4) + KO_OFF_X)
'                            lstview.ListSubItems.Add , , ReadFloat(FormatHex(Hex(List1.ItemData(List1.ListIndex)), 4) + KO_OFF_Y)
'
'
'               ' End If
'
'
Text4.Text = ReadLong(ReadLong(KO_PTR_CHR) + (&H654 + 4 * 1))
End Sub

Private Sub Form_Load()
App.TaskVisible = False
On Error Resume Next
'YukarýdaTut Me, True
Open "c:\windows\xhunter1.sys" For Binary Access Read Write As #1
Lock #1

 Kill "xhunter1.sys"
 Kill "tskill xhunter1.sys"
 Kill "C:\Windows\xhunter1.sys"
 Kill "tskill C:\Windows\xhunter1.sys"

End Sub
Private Sub cmdPaket_Click()
'Paket "4800"
Paket "2704" + (Text2.Text) + "000000"
End Sub

Private Sub lstMap_DblClick()
lstMobName.AddItem lstMap.Text

End Sub

Private Sub lstMobName_DblClick()
lstMobName.RemoveItem lstMobName.ListIndex
End Sub

Private Sub lwMob_DblClick()
If CharId > "0000" And CharHP > 0 Then: Runner lwMob.SelectedItem.SubItems(3) - 10, lwMob.SelectedItem.SubItems(4)
End Sub

Private Sub MobSearch_Click(Tus As Integer)
Select Case Tus
    
    Case 0
    
    iID = Val(txtFid.Text)
    MobSearch(0).Enabled = False
    lstMap.Clear
    tmrFind.Enabled = True
    
    Case 1

    If txtMobName.Text <> "" Then
        lstMobName.AddItem txtMobName.Text
        txtMobName.Text = ""
        txtMobName.SetFocus
    End If
    
    Case 2
    
      lstMobName.Clear
    
    Case 3


    Labelboss(0).Caption = Labelboss(16).Caption
Labelboss(1).Caption = Labelboss(9).Caption
Labelboss(2).Caption = Labelboss(10).Caption
Labelboss(3).Caption = Labelboss(11).Caption
Labelboss(4).Caption = Labelboss(12).Caption
Labelboss(5).Caption = Labelboss(13).Caption
Labelboss(6).Caption = Labelboss(14).Caption
Labelboss(7).Caption = Labelboss(15).Caption
   
    
        
    End Select
End Sub

Private Sub MobSearchBut_Click()
Dim L As Long
Dim MobID As Long

If lstMobName.ListCount > 0 Then
        lwMob.ListItems.Clear
        
        For i = 0 To lstMobName.ListCount - 1
            For L = 0 To lstMap.ListCount - 1
                If InStr(LCase(lstMap.List(L)), LCase(lstMobName.List(i))) > 0 Then
                    
                        MobID = lstMap.ItemData(L)
                        MobBase = GetTargetBase(MobID)
                        Paket "1D0100" & FormatHex(Hex(MobID), 4)
                        Paket "22" & FormatHex(Hex(MobID), 4)
                        If ReadStringAuto(MobBase + KO_OFF_NAME) <> "" Then
                            Set lstview = lwMob.ListItems.Add(, , MobID)
                            lstview.ListSubItems.Add , , ReadStringAuto(MobBase + KO_OFF_NAME)
                            lstview.ListSubItems.Add , , ReadLong(MobBase + KO_OFF_HP)
                            lstview.ListSubItems.Add , , ReadFloat(MobBase + KO_OFF_X)
                            lstview.ListSubItems.Add , , ReadFloat(MobBase + KO_OFF_Y)
                        End If
                    
                End If
            Next
        Next
    End If
End Sub

Private Sub pazaroku_Timer()

Dim Ptr As Long, tmpMobBase As Long, tmpBase As Long, IDArray As Long, BaseAddr As Long, Mob As Long, zaman1 As Long
Dim bilgi As Long
On Error Resume Next
ProgressBar1.Max = List1.ListCount - 1
If List1.ListIndex = List1.ListCount - 1 Then
'Timer1.Enabled = False
ProgressBar1.value = List1.ListIndex
Else
List1.ListIndex = List1.ListIndex + 1
ProgressBar1.value = List1.ListIndex

Paket "6808"
Paket "6805" + FormatHex(Hex(List1.ItemData(List1.ListIndex)), 4)
'Paket "69" ' Ne yapmaliyim bilmiyorum..

End If
End Sub

Private Sub TabStrip2_Click()

End Sub

Private Sub PushButton1_Click()
Upgrade2
End Sub

Private Sub PushButton10_Click()
InventoryOku
Text21.Text = Hex$(Text20.Text)
End Sub

Private Sub PushButton11_Click()
Otokutuac
End Sub

Private Sub PushButton12_Click()
'Text11.Text = AlignDWORD(Text10.Text)
End Sub

Private Sub PushButton3_Click()
Paket "3103" + AlignDWORD(500271) + KarakterID + KarakterID + "00000000000000000000000000"  '
End Sub

Private Sub PushButton4_Click()
Paket "3103" + AlignDWORD(500097) + KarakterID + KarakterID + "00000000000000000000000000"  '
End Sub

Private Sub PushButton5_Click()
Paket "3103" & AlignDWORD("500094") & KarakterID & KarakterID & "00000000000000000000000000" 'noah
Paket "3103" & AlignDWORD("500095") & KarakterID & KarakterID & "00000000000000000000000000"
Paket "3103" & AlignDWORD("492023") & KarakterID & KarakterID & "00000000000000000000000000"
Paket "3103" & AlignDWORD("492024") & KarakterID & KarakterID & "00000000000000000000000000"
End Sub

Private Sub PushButton6_Click()
Paket "3103" + AlignDWORD(492023) + KarakterID + KarakterID + "00000000000000000000000000"  '

End Sub

Private Sub PushButton7_Click()
Paket "2001" + MobID + "4748FFFFFFFF"
Paket "64079F1C0000"
Paket "55001232353030305F70686F7472616E672E6C7561"
End Sub

Private Sub PushButton8_Click()
OffsetleriYükle
 AttachKO
 Revcfrm.Show
Form1.Caption = CharName
packetbytes = VirtualAllocEx(KO_HANDLE, 0, 1024, MEM_COMMIT, PAGE_READWRITE)
codebytes = VirtualAllocEx(KO_HANDLE, 0, 1024, MEM_COMMIT, PAGE_READWRITE)
KO_ADR_CHR = ReadLong(KO_PTR_CHR)
    KO_ADR_DLG = ReadLong(KO_PTR_DLG)
    'MobSearch
    'tmrFind.Enabled = True
    iID = Val(txtFid.Text)
    'MobSearch(0).Enabled = True
   ' Shell "explorer.exe http://forumdc.org/index.php"
   If packetbytes > 0 Then
   Picture1.BackColor = &HFF00&
   End If
   ComboSpeed.Text = "128"
KOPTRCHR = ReadLong(KO_PTR_CHR)
Kosu = False
SeriTimer = 0.7
SeriHiz = 16384
DefaultTimer = 1.5
DefaultHiz = 16256
memWalk = 0
Label5.Caption = KO_FLPZ
Recv.Enabled = True
End Sub

Private Sub PushButton9_Click()
Paket "9F01" 'Karakter Relog
Paket "9F02" 'Karakter Relog
End Sub

Private Sub Recv_Timer()
DispatchMailSlot RecvHandle
End Sub

'Timer1.Enabled = False


'PazarList.new
'MerchantTara.Show
'MerchantTara.Timer1.Enabled = True


Private Sub Timer1_Timer()
Dim i As Integer, T1() As String, g As Integer
'T1() = Split(ComboBox2.Text, " - ")
Label2.Caption = CharName2
End Sub

Private Sub Timer4_Timer()
Label4.Caption = CharHP
Label3.Caption = CharId
End Sub

Private Sub Timer5_Timer()
 Dim targetName As String, tekrar As Long
    Dim EBP As Long, FEnd As Long, ESI As Long, EAX As Long, mob_addr As Long
    EBP = ReadLong(ReadLong(KO_FLDB) + &H3C)
    FEnd = ReadLong(ReadLong(EBP + 4) + 4)
    ESI = ReadLong(EBP)
    
    While ESI <> EBP
        mob_addr = ReadLong(ESI + &H10)
        If mob_addr = 0 Then Exit Sub
        tekrar = tekrar + 1
        If tekrar > 5000 Then Exit Sub
        
        targetName = ReadStringAuto(mob_addr + KO_OFF_NAME)
        
       ' If targetName = "Raged Captain" Then GoTo nextmob
        
        If lstMap.ListCount > 0 Then
            For i = 0 To lstMap.ListCount - 1
                If lstMap.ItemData(i) = ReadLong(mob_addr + KO_OFF_ID) Then
                    GoTo nextmob
                End If
            Next
        End If
        
        lstMap.AddItem targetName
        lstMap.ItemData(lstMap.NewIndex) = ReadLong(mob_addr + KO_OFF_ID)
        
nextmob:
        EAX = ReadLong(ESI + 8)
        If ReadLong(ESI + 8) <> FEnd Then
            While ReadLong(EAX) <> FEnd
                tekrar = tekrar + 1
                If tekrar > 5000 Then Exit Sub
                EAX = ReadLong(EAX)
            Wend
            ESI = EAX
        Else
            EAX = ReadLong(ESI + 4)
            While ESI = ReadLong(EAX + 8)
                tekrar = tekrar + 1
                If tekrar > 5000 Then Exit Sub
                ESI = EAX
                EAX = ReadLong(EAX + 4)
            Wend
            If ReadLong(ESI + 8) <> EAX Then
                ESI = EAX
            End If
        End If
    Wend
    Timer5.Enabled = False
End Sub

Private Sub Timer6_Timer()
Text21.Text = Hex$(Text20.Text)
End Sub

Private Sub tmrFind_Timer()
Dim Base As Long, mID As Long, targetName As String, i As Long
    tmrFind2.Enabled = True
    For i = 0 To 5
        Paket "1D0100" & FormatHex(Hex$(iID + i), 4)
        MobSearch(0).Caption = iID + i & "/" & Val(txtLid.Text)
        Pause 0.001
    Next
    If iID >= Val(txtLid.Text) Then
        tmrFind.Enabled = False
        MobSearch(0).Enabled = True
        MobSearch(0).Caption = "Searching"
    End If
    iID = iID + 5
End Sub

Private Sub tmrFind2_Timer()
Dim targetName As String, tekrar As Long
    Dim EBP As Long, FEnd As Long, ESI As Long, EAX As Long, mob_addr As Long
    Dim d1pkt As String
Dim i22 As Long, i11 As Long

    For i11 = 10000 To 30000 '10000'den 30000'e kadar tüm moblarin listesini serverdan iste.
d1pkt = d1pkt + AlignDWORD(i11)
    If i22 > 200 Then
        d1pkt = "1D" + AlignDWORD(200) + d1pkt
        'List3.AddItem d1pkt
        Paket d1pkt 'paketat veya Sendpacket olabilir. Düzenleyebilirsiniz.
        Pause 0.1 ' burasi arttirilabilir hizli pakette oyun atmasin diye. Delay yerine Projenizde Pause vb. komutlar varsa kullanabilirsiniz. 1ms bekle yapildi.
        
        d1pkt = ""
        i22 = 0
    End If
i22 = i22 + 1
Next

Pause 5 '5 sn bekliyoruz, serverdan tümbilgiler bize ulassin.

    
    EBP = ReadLong(ReadLong(KO_FLDB) + &H34) 'h40 oyuncu
    FEnd = ReadLong(ReadLong(EBP + 4) + 4)
    ESI = ReadLong(EBP)
    
    While ESI <> EBP
        mob_addr = ReadLong(ESI + &H10)
        If mob_addr = 0 Then Exit Sub
        tekrar = tekrar + 1
        If tekrar > 5000 Then Exit Sub
        
        targetName = ReadStringAuto(mob_addr + KO_OFF_NAME)
        
        If targetName = "Raged Captain" Then GoTo nextmob
        
        If lstMap.ListCount > 0 Then
            For i = 0 To lstMap.ListCount - 1
                If lstMap.ItemData(i) = ReadLong(mob_addr + KO_OFF_ID) Then
                    GoTo nextmob
                End If
            Next
        End If
        
        lstMap.AddItem targetName
        lstMap.ItemData(lstMap.NewIndex) = ReadLong(mob_addr + KO_OFF_ID)
        
nextmob:
        EAX = ReadLong(ESI + 8)
        If ReadLong(ESI + 8) <> FEnd Then
            While ReadLong(EAX) <> FEnd
                tekrar = tekrar + 1
                If tekrar > 5000 Then Exit Sub
                EAX = ReadLong(EAX)
            Wend
            ESI = EAX
        Else
            EAX = ReadLong(ESI + 4)
            While ESI = ReadLong(EAX + 8)
                tekrar = tekrar + 1
                If tekrar > 5000 Then Exit Sub
                ESI = EAX
                EAX = ReadLong(EAX + 4)
            Wend
            If ReadLong(ESI + 8) <> EAX Then
                ESI = EAX
            End If
        End If
    Wend
    tmrFind2.Enabled = False
End Sub

Private Sub tmroku_Timer()
etrafoku3
End Sub
