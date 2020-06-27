VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form MerchantTara 
   Caption         =   "Adsýz - Not Defteri"
   ClientHeight    =   6210
   ClientLeft      =   4920
   ClientTop       =   5550
   ClientWidth     =   9900
   Icon            =   "MerchantTara.frx":0000
   LinkTopic       =   "Form3"
   ScaleHeight     =   6210
   ScaleWidth      =   9900
   Begin VB.Timer Dis 
      Interval        =   1000
      Left            =   480
      Top             =   1080
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   0
      Top             =   1080
   End
   Begin VB.TextBox Text1 
      Height          =   1965
      Left            =   4200
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   840
      Width           =   855
   End
   Begin VB.ListBox List1 
      Height          =   3375
      ItemData        =   "MerchantTara.frx":F172
      Left            =   1440
      List            =   "MerchantTara.frx":F174
      TabIndex        =   1
      Top             =   840
      Width           =   2535
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Kurulan Pazarlarýn Itemleri Taranýyor..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   4575
   End
End
Attribute VB_Name = "MerchantTara"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Private Sub Dis_Timer()
'On Error Resume Next
'Client(1).PazarTara
'End Sub
'
'Private Sub Form_Load()
'FormÜstte Me, 1
'End Sub
'
'Private Sub Timer1_Timer()
'On Error Resume Next
'ProgressBar1.Max = List1.ListCount - 1
'If List1.ListIndex = List1.ListCount - 1 Then
'Timer1.Enabled = False
'ProgressBar1.value = List1.ListIndex
'MerchantOku = 0
'Merchant.Show
'Unload Me
'Else
'List1.ListIndex = List1.ListIndex + 1
'ProgressBar1.value = List1.ListIndex
'Client(1).paketYolla "6808"
'Client(1).paketYolla "6805" + Client(1).FormatHex(Hex(List1.ItemData(List1.ListIndex)), 4)
'Timer1.Enabled = False
'End If
'End Sub
