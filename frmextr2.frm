VERSION 5.00
Begin VB.Form frmextr2 
   BackColor       =   &H80000007&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Extra Functions"
   ClientHeight    =   600
   ClientLeft      =   4875
   ClientTop       =   2850
   ClientWidth     =   4185
   Icon            =   "frmextr2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   600
   ScaleWidth      =   4185
   Begin VB.CommandButton Command1 
      Caption         =   "Change"
      Height          =   255
      Left            =   3480
      TabIndex        =   3
      Top             =   360
      Width           =   735
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "frmextr2.frx":0742
      Left            =   3480
      List            =   "frmextr2.frx":0752
      TabIndex        =   2
      Text            =   "(speed)"
      Top             =   0
      Width           =   735
   End
   Begin VB.Timer Timer2 
      Interval        =   500
      Left            =   4920
      Top             =   0
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   4920
      Top             =   360
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000012&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   1920
      TabIndex        =   1
      Top             =   0
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000012&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1935
   End
   Begin VB.Menu change 
      Caption         =   "Change"
      Visible         =   0   'False
      Begin VB.Menu css 
         Caption         =   "Change seconds speed."
      End
      Begin VB.Menu d1 
         Caption         =   "-"
      End
      Begin VB.Menu btbm 
         Caption         =   "Back to big mode."
      End
   End
End
Attribute VB_Name = "frmextr2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btbm_Click()
frmextra.Show
frmextr2.Hide
End Sub

Private Sub Command1_Click()
frmextr2.PopupMenu change
End Sub

Private Sub css_Click()
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

Private Sub Timer1_Timer()
Label1.Caption = Now
End Sub

Private Sub Timer2_Timer()
Label2.Caption = timer
End Sub
