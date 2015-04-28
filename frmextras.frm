VERSION 5.00
Begin VB.Form frmextras 
   BackColor       =   &H80000007&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Extras"
   ClientHeight    =   2880
   ClientLeft      =   3450
   ClientTop       =   2505
   ClientWidth     =   3480
   Icon            =   "frmextras.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2880
   ScaleWidth      =   3480
   Begin VB.CommandButton Command3 
      Caption         =   "Clear All"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   1920
      Width           =   3255
   End
   Begin VB.ListBox List1 
      Columns         =   3
      Height          =   735
      ItemData        =   "frmextras.frx":0442
      Left            =   120
      List            =   "frmextras.frx":0444
      Sorted          =   -1  'True
      Style           =   1  'Checkbox
      TabIndex        =   5
      ToolTipText     =   "You can check which passwords you want."
      Top             =   1080
      Width           =   3255
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Exit"
      Height          =   375
      Left            =   1800
      TabIndex        =   4
      Top             =   2520
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Try it out"
      Height          =   375
      Left            =   600
      TabIndex        =   3
      Top             =   2520
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H80000008&
      ForeColor       =   &H80000005&
      Height          =   285
      Left            =   1200
      TabIndex        =   2
      Top             =   720
      Width           =   1935
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000005&
      X1              =   0
      X2              =   3480
      Y1              =   2400
      Y2              =   2400
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000012&
      Caption         =   "Password:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000008&
      Caption         =   "Extras"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   615
      Left            =   960
      TabIndex        =   0
      Top             =   0
      Width           =   1455
   End
End
Attribute VB_Name = "frmextras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Text1.Text = "coolinks" Then
frmlinks.Enabled = True
frmClock.links.Visible = True
frmClock.l7.Visible = True
List1.AddItem "Cool Links"
MsgBox "Your password is now enabled", , "Cool"
Text1.Text = ""
Else
MsgBox "That was not a valid password", , "Sorry!"
End If
End Sub

Private Sub Command2_Click()
Unload Me
frmClock.Enabled = True
frmClock.Show

End Sub

Private Sub Command3_Click()
List1.Clear
frmClock.sc.Visible = False
frmClock.l7.Visible = False
frmClock.links.Visible = False
frmClock.music.Visible = False
frmcode.Enabled = False
frmClock.music.Visible = False
frmthrone.Enabled = False
frmlinks.Enabled = False
frmlinks.Visible = False
End Sub

Private Sub Form_Load()
'the things that go here are like passwords that help get
'the sorce code and other extras for the people who know.
' this thing also lets the person choose the passwords you
'want and you can check them to enable them like diddy kong racing
'there is a button that let's you delete all the codes.
'there is an extra that let's you hear the star Wars music
End Sub

Private Sub Form_Unload(Cancel As Integer)
frmClock.Enabled = True
frmClock.Show

End Sub

