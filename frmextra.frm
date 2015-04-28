VERSION 5.00
Begin VB.Form frmextra 
   BackColor       =   &H80000008&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Extra Functions"
   ClientHeight    =   3615
   ClientLeft      =   2640
   ClientTop       =   2040
   ClientWidth     =   5325
   Icon            =   "frmextra.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3615
   ScaleWidth      =   5325
   Begin VB.CommandButton Command1 
      Caption         =   "Change"
      Height          =   375
      Left            =   3360
      TabIndex        =   6
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Timer Timer2 
      Interval        =   100
      Left            =   4800
      Top             =   360
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   4800
      Top             =   0
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000012&
      Caption         =   "Seconds from Mid-night"
      ForeColor       =   &H8000000E&
      Height          =   1695
      Left            =   240
      TabIndex        =   3
      Top             =   1800
      Width           =   4455
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "frmextra.frx":0742
         Left            =   360
         List            =   "frmextra.frx":0752
         TabIndex        =   5
         Text            =   "(Speed)"
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label Label2 
         BorderStyle     =   1  'Fixed Single
         Height          =   615
         Left            =   360
         TabIndex        =   4
         Top             =   480
         Width           =   3615
      End
   End
   Begin VB.Frame franow 
      BackColor       =   &H80000012&
      Caption         =   "Current Time"
      ForeColor       =   &H00FFFFFF&
      Height          =   1335
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   4455
      Begin VB.Label Label1 
         BackColor       =   &H80000012&
         Caption         =   "Now:"
         ForeColor       =   &H00C0FFFF&
         Height          =   255
         Left            =   360
         TabIndex        =   2
         Top             =   360
         Width           =   375
      End
      Begin VB.Label lbldisplay 
         BorderStyle     =   1  'Fixed Single
         Height          =   615
         Left            =   360
         TabIndex        =   1
         Top             =   600
         Width           =   3615
      End
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   4800
      Picture         =   "frmextra.frx":0772
      Stretch         =   -1  'True
      Top             =   3000
      Width           =   495
   End
End
Attribute VB_Name = "frmextra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Combo1.Text = "Slow" Then
Timer2.Interval = 500
ElseIf Combo1.Text = "Normal" Then
Timer2.Interval = 100
ElseIf Combo1.Text = "Fast" Then
Timer2.Interval = 50
ElseIf Combo1.Text = "Faster" Then
Timer2.Interval = 5
End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
frmClock.Enabled = True
frmClock.Show
End Sub

Private Sub Image1_Click()
frmextra.Hide
frmextr2.Show
End Sub

Private Sub Label2_Click()
lbldisplay.Caption = timer
End Sub

Private Sub lblnow_Click()
lbldisplay.Caption = Now
End Sub

Private Sub Timer1_Timer()
lbldisplay.Caption = Now
End Sub

Private Sub Timer2_Timer()
Label2.Caption = timer
End Sub
