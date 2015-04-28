VERSION 5.00
Object = "{A37D0E58-B6D0-11D2-971F-EC500970267D}#7.0#0"; "AnalogClock.ocx"
Begin VB.Form frmanalog 
   BackColor       =   &H80000012&
   Caption         =   "Analog"
   ClientHeight    =   1095
   ClientLeft      =   5835
   ClientTop       =   2220
   ClientWidth     =   1095
   Icon            =   "frmanalog.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1095
   ScaleWidth      =   1095
   Begin VB.Timer Timer3 
      Interval        =   10
      Left            =   240
      Top             =   2160
   End
   Begin VB.Timer Timer2 
      Interval        =   100
      Left            =   600
      Top             =   1440
   End
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   120
      Top             =   1440
   End
   Begin AnalogClockControl.AnalogClock AnalogClock1 
      Height          =   1095
      Left            =   0
      Top             =   0
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1931
      Value           =   0.5
      BackColor       =   -2147483630
   End
   Begin VB.Menu amanolog 
      Caption         =   "Anolog Menu"
      Visible         =   0   'False
   End
End
Attribute VB_Name = "frmanalog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
frmanalog.Hide
frmClock.Enabled = True
frmClock.Show
End Sub

Private Sub Command2_Click()

End Sub

Private Sub Form_Load()
AnalogClock1.Value = Time
frmClock.Enabled = True
frmClock.Show
AnalogClock1.Height = frmanalog.Height
AnalogClock1.Width = frmanalog.Width
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
AnalogClock1.Height = frmanalog.Height
AnalogClock1.Width = frmanalog.Width
End Sub

Private Sub Form_Resize()
AnalogClock1.Height = frmanalog.Height
AnalogClock1.Width = frmanalog.Width
End Sub

Private Sub Timer1_Timer()
AnalogClock1.Value = Time
End Sub

Private Sub Timer2_Timer()
AnalogClock1.ToolTipText = Time
End Sub

Private Sub Timer3_Timer()
AnalogClock1.Height = frmanalog.Height
AnalogClock1.Width = frmanalog.Width
End Sub
