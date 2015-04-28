VERSION 5.00
Object = "{5BD6A9E6-E8CD-11D2-A471-004005423446}#1.0#0"; "limitTextBox.ocx"
Begin VB.Form frmAlarm 
   BackColor       =   &H80000008&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Alarm Set"
   ClientHeight    =   1350
   ClientLeft      =   4635
   ClientTop       =   3090
   ClientWidth     =   4080
   Icon            =   "Alarm.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1350
   ScaleWidth      =   4080
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "Alarm.frx":0442
      Left            =   2640
      List            =   "Alarm.frx":044F
      TabIndex        =   5
      Top             =   480
      Width           =   615
   End
   Begin limitTextBox.textBox msktime 
      Height          =   315
      Left            =   720
      TabIndex        =   4
      Top             =   480
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      allowedText     =   "0123456789PM: AM"
   End
   Begin VB.CommandButton Command3 
      Caption         =   "View current alarm Time"
      Height          =   375
      Left            =   1080
      TabIndex        =   3
      Top             =   960
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Can&cel"
      Height          =   375
      Left            =   3120
      TabIndex        =   2
      Top             =   960
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "S&et"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000012&
      Caption         =   "What time do you want the alarm?"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   840
      TabIndex        =   0
      Top             =   120
      Width           =   2535
   End
End
Attribute VB_Name = "frmAlarm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim alarmtime
Private Sub MaskEdBox1_Change()

End Sub

Private Sub Command1_Click()
If msktime.Text <> "" Then
frmClock.Timer3.Enabled = True
frmClock.Label3.Caption = msktime.Text + " " + Combo1.Text
Command3.Visible = True
Command2.Caption = "To clock"
    If Combo1.Text = "" Then
    MsgBox "You have to select either AM or PM.", , "Not selected"
    End If
Else
MsgBox "That was not a valid time", , "Not valid"
End If
End Sub

Private Sub Command2_Click()
Unload Me
frmClock.Enabled = True
frmClock.Show

End Sub

Private Sub Command3_Click()
MsgBox "The current alarm time is " & frmClock.Label3.Caption
End Sub

Private Sub Form_Unload(Cancel As Integer)
frmClock.Enabled = True
frmClock.Show
End Sub

