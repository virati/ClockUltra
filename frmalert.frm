VERSION 5.00
Begin VB.Form frmalert 
   BackColor       =   &H80000007&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Alert Form"
   ClientHeight    =   1710
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4215
   Icon            =   "frmalert.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1710
   ScaleWidth      =   4215
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "E&xit"
      Height          =   375
      Left            =   2280
      TabIndex        =   3
      Top             =   1320
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Clo&se"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   1815
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   0
      Top             =   0
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000012&
      Caption         =   "What  would you like to do now."
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   960
      TabIndex        =   4
      Top             =   1080
      Width           =   2295
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   345
      Left            =   2040
      TabIndex        =   1
      Top             =   600
      Width           =   75
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000008&
      Caption         =   "Alert!!"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   1410
      TabIndex        =   0
      Top             =   0
      Width           =   1335
   End
End
Attribute VB_Name = "frmalert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
frmClock.Enabled = True
frmClock.Show
End Sub

Private Sub Command2_Click()
Unload Me
frmClock.Enabled = True
frmClock.Show
End Sub

Private Sub Form_Load()
Label2.Caption = "The time is " & frmClock.lbltime

End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload Me
frmClock.Enabled = True
frmClock.Show
End Sub

Private Sub Timer1_Timer()
If Label1.ForeColor = vbWhite Then
Label1.ForeColor = vbBlue
ElseIf Label1.ForeColor = vbBlue Then
Label1.ForeColor = vbRed
ElseIf Label1.ForeColor = vbRed Then
Label1.ForeColor = vbWhite
End If
End Sub
