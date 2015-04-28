VERSION 5.00
Object = "{6E2C9294-8D14-11D1-8213-0060975EACAF}#3.0#0"; "VBUMidiPlayer.ocx"
Begin VB.Form frmthrone 
   BackColor       =   &H80000007&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1185
   ClientLeft      =   540
   ClientTop       =   6705
   ClientWidth     =   4815
   ControlBox      =   0   'False
   Enabled         =   0   'False
   Icon            =   "frmthrone.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1185
   ScaleWidth      =   4815
   Begin VB.CommandButton Command1 
      Caption         =   "X"
      Height          =   315
      Left            =   4440
      TabIndex        =   3
      Top             =   0
      Width           =   375
   End
   Begin VB.CommandButton cmdplay 
      Caption         =   "Pl&ay"
      Height          =   255
      Left            =   1320
      TabIndex        =   2
      Top             =   840
      Width           =   855
   End
   Begin VB.CommandButton cmdstop 
      Caption         =   "&Stop"
      Height          =   255
      Left            =   2640
      TabIndex        =   1
      Top             =   840
      Width           =   855
   End
   Begin VBUMidiFilePlayer.VBUMidiPlayer thronecplay 
      Left            =   0
      Top             =   0
      _ExtentX        =   794
      _ExtentY        =   688
      Filename        =   "I:\Star Wars\Music\Thronec.mid"
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000007&
      Caption         =   "Star Wars Music"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   720
      TabIndex        =   0
      Top             =   120
      Width           =   3255
   End
End
Attribute VB_Name = "frmthrone"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command3_Click()

End Sub

Private Sub cmdpause_Click()

End Sub

Private Sub cmdplay_Click()
thronecplay.MidiPlay
End Sub

Private Sub cmdstop_Click()
thronecplay.MidiStop
End Sub

Private Sub Command1_Click()
Unload Me
frmClock.Enabled = True
frmClock.Show
thronecplay.MidiStop
End Sub

Private Sub Form_Unload(Cancel As Integer)
frmClock.Enabled = True
frmClock.Show
thronecplay.MidiStop
End Sub

Private Sub MMControl1_Done(NotifyCode As Integer)

End Sub

Private Sub music1_PlayClick(Cancel As Integer)
End Sub

Private Sub music1_StopClick(Cancel As Integer)
End Sub

