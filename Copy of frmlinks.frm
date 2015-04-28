VERSION 5.00
Object = "{ED6EFBE9-2DE5-11D2-9B4A-006097731E48}#1.0#0"; "HLink.ocx"
Begin VB.Form frmlinks 
   BackColor       =   &H80000008&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cool Links"
   ClientHeight    =   975
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5535
   Enabled         =   0   'False
   Icon            =   "frmlinks.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   975
   ScaleWidth      =   5535
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "D&one"
      Height          =   375
      Left            =   1680
      TabIndex        =   4
      Top             =   600
      Width           =   2295
   End
   Begin HyperLinkControl.HLink HLink5 
      Height          =   255
      Left            =   2520
      TabIndex        =   3
      Top             =   240
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   450
      BackColor       =   -2147483640
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "frmlinks.frx":0442
      Caption         =   "Star Wars Database"
      URL             =   "http://www.swdatabase.com"
   End
   Begin HyperLinkControl.HLink HLink3 
      Height          =   255
      Left            =   1080
      TabIndex        =   2
      Top             =   240
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   450
      BackColor       =   -2147483640
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "frmlinks.frx":075C
      Caption         =   "My Star Wars Page"
      URL             =   "http://homepages.go.com/~exarkunxx/Korriban1.htm"
   End
   Begin HyperLinkControl.HLink HLink2 
      Height          =   255
      Left            =   4080
      TabIndex        =   1
      Top             =   240
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   450
      BackColor       =   -2147483640
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "frmlinks.frx":0A76
      Caption         =   "Hompages.go.com"
      URL             =   "http://hompages.go.com"
   End
   Begin HyperLinkControl.HLink HLink1 
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   240
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   450
      BackColor       =   -2147483640
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "frmlinks.frx":0D90
      Caption         =   "Star Wars.com"
      URL             =   "http://starwars.com"
   End
End
Attribute VB_Name = "frmlinks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
frmClock.Enabled = True
frmClock.Show
Unload Me
End Sub
