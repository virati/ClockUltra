VERSION 5.00
Begin VB.Form frmcode 
   BackColor       =   &H80000012&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Source Code"
   ClientHeight    =   3810
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5655
   Enabled         =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3810
   ScaleWidth      =   5655
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   2775
      Left            =   0
      TabIndex        =   4
      Top             =   480
      Width           =   5655
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Change"
      Height          =   255
      Left            =   3840
      TabIndex        =   3
      Top             =   120
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Exit"
      Height          =   255
      Left            =   3840
      TabIndex        =   2
      Top             =   3480
      Width           =   1695
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "frmcode.frx":0000
      Left            =   2040
      List            =   "frmcode.frx":000D
      TabIndex        =   1
      Text            =   "(Choose the form)"
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000012&
      Caption         =   "Source Code:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "frmcode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
frmcode.Hide
frmClock.Enabled = True
frmClock.Show

End Sub

Private Sub Document1_GotFocus()

End Sub

Private Sub Form_Unload(Cancel As Integer)
frmClock.Enabled = True
frmClock.Show

End Sub

Private Sub RichTextBox1_DblClick()
'if you double click, a form comes up telling info on which
'program was used to make Clock ultra and copyright info.
End Sub

