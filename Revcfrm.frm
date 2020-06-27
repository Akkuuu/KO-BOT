VERSION 5.00
Begin VB.Form Revcfrm 
   Caption         =   "Form2"
   ClientHeight    =   8910
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   14340
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   ScaleHeight     =   8910
   ScaleWidth      =   14340
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check2 
      Caption         =   "Check2"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   8520
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   8655
      Left            =   2520
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   120
      Width           =   11655
   End
   Begin VB.ListBox List1 
      Height          =   8250
      ItemData        =   "Revcfrm.frx":0000
      Left            =   120
      List            =   "Revcfrm.frx":000A
      TabIndex        =   0
      Top             =   120
      Width           =   2295
   End
End
Attribute VB_Name = "Revcfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
