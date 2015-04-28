VERSION 5.00
Begin VB.Form frmcolorlbl 
   BackColor       =   &H80000008&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Text Select"
   ClientHeight    =   3255
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7185
   Icon            =   "frmcolorlbl.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3255
   ScaleWidth      =   7185
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "S&et"
      Height          =   375
      Left            =   1080
      TabIndex        =   12
      Top             =   2760
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4200
      TabIndex        =   11
      Top             =   2760
      Width           =   1335
   End
   Begin VB.HScrollBar HScroll3 
      Height          =   255
      Left            =   960
      Max             =   240
      TabIndex        =   6
      Top             =   2280
      Width           =   4335
   End
   Begin VB.HScrollBar HScroll2 
      Height          =   255
      Left            =   960
      Max             =   240
      TabIndex        =   5
      Top             =   1800
      Width           =   4335
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   960
      Max             =   240
      TabIndex        =   4
      Top             =   1320
      Width           =   4335
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   6720
      Top             =   120
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   5400
      TabIndex        =   10
      Top             =   1320
      Width           =   45
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   5400
      TabIndex        =   9
      Top             =   1800
      Width           =   45
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   5400
      TabIndex        =   8
      Top             =   2280
      Width           =   45
   End
   Begin VB.Label lblcolorer 
      Height          =   1215
      Left            =   6240
      TabIndex        =   7
      Top             =   1320
      Width           =   735
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000009&
      X1              =   5640
      X2              =   6000
      Y1              =   1800
      Y2              =   1800
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000009&
      X1              =   5640
      X2              =   6000
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000012&
      Caption         =   "Red"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   480
      TabIndex        =   3
      Top             =   1320
      Width           =   375
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000012&
      Caption         =   "Green"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   1800
      Width           =   495
   End
   Begin VB.Label Label4 
      BackColor       =   &H80000012&
      Caption         =   "Blue"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   480
      TabIndex        =   1
      Top             =   2280
      Width           =   375
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000009&
      X1              =   120
      X2              =   7200
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000007&
      Caption         =   "Select the Color"
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
      Left            =   1920
      TabIndex        =   0
      Top             =   240
      Width           =   3255
   End
End
Attribute VB_Name = "frmcolorlbl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
frmClock.lbltime.ForeColor = lblcolorer.BackColor
frmClock.lbltimecol.ForeColor = lblcolorer.BackColor
Unload Me
frmClock.Enabled = True
frmClock.Show

End Sub

Private Sub Command2_Click()
frmClock.Enabled = True
frmClock.Show
Unload Me
End Sub

Private Sub Form_Load()
Label5.Caption = 0
Label6.Caption = 0
Label7.Caption = 0
lblcolorer.BackColor = RGB(Label5.Caption, Label6.Caption, Label7.Caption)
End Sub

Private Sub Form_Unload(Cancel As Integer)
frmClock.Enabled = True
frmClock.Show

End Sub

Private Sub HScroll1_Change()
Label5.Caption = HScroll1.Value
lblcolorer.BackColor = RGB(Label5.Caption, Label6.Caption, Label7.Caption)
Label1.ForeColor = lblcolorer.BackColor
End Sub

Private Sub HScroll2_Change()
Label6.Caption = HScroll2.Value
lblcolorer.BackColor = RGB(Label5.Caption, Label6.Caption, Label7.Caption)
Label1.ForeColor = lblcolorer.BackColor

End Sub

Private Sub HScroll3_Change()
Label7.Caption = HScroll3.Value
lblcolorer.BackColor = RGB(Label5.Caption, Label6.Caption, Label7.Caption)
Label1.ForeColor = lblcolorer.BackColor

End Sub

Private Sub Timer1_Timer()
If Label1.ForeColor = vbWhite Then
Label1.ForeColor = vbRed
Else
Label1.ForeColor = vbWhite
End If
End Sub

