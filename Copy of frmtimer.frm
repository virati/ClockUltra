VERSION 5.00
Begin VB.Form frmtimer 
   BackColor       =   &H80000007&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Timer"
   ClientHeight    =   1755
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5130
   Icon            =   "frmtimer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1755
   ScaleWidth      =   5130
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "To C&lock"
      Height          =   375
      Left            =   3240
      TabIndex        =   5
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "St&op Timer"
      Height          =   375
      Left            =   1920
      TabIndex        =   4
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   0
      Top             =   1680
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   135
      Left            =   120
      TabIndex        =   2
      Top             =   2160
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "S&tart Timer"
      Height          =   375
      Left            =   600
      TabIndex        =   1
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   360
      Top             =   1680
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000009&
      X1              =   0
      X2              =   5160
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000007&
      Caption         =   "Timer"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   33.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   735
      Left            =   1680
      TabIndex        =   3
      Top             =   0
      Width           =   1695
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   1200
      TabIndex        =   0
      Top             =   840
      Width           =   2655
   End
End
Attribute VB_Name = "frmtimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim clicktime As Integer
Dim showtime As Integer
Dim changetime As Integer
Dim intime As Integer

Private Sub Command1_Click()
Timer1.Enabled = True
Timer2.Enabled = True
Command1.Enabled = False
Command2.Enabled = True
End Sub

Private Sub Command2_Click()
Timer1.Enabled = False
Timer2.Enabled = False
Command1.Enabled = True
Command2.Enabled = False
HScroll1.Value = 0
End Sub

Private Sub Command3_Click()
Unload Me
frmClock.Enabled = True
frmClock.Show
End Sub

Private Sub Form_Load()
Label1.Caption = "Click Start Timer"
Command2.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
frmClock.Enabled = True
frmClock.Show
End Sub

Private Sub Timer1_Timer()
HScroll1.Value = HScroll1.Value + 1
End Sub

Private Sub Timer2_Timer()
Label1.Caption = HScroll1.Value
End Sub
