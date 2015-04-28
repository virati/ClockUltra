VERSION 5.00
Begin VB.Form frmxtrap 
   BackColor       =   &H80000008&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Extra Programs"
   ClientHeight    =   2940
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6225
   Icon            =   "frmxtrap.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2940
   ScaleWidth      =   6225
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtextra 
      Enabled         =   0   'False
      Height          =   495
      Left            =   720
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Text            =   "frmxtrap.frx":0742
      Top             =   1080
      Width           =   4695
   End
   Begin VB.Label lblpro2 
      BackColor       =   &H80000012&
      Caption         =   "Percent Finder"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   2280
      TabIndex        =   5
      Top             =   2040
      Width           =   1575
   End
   Begin VB.Label lblpro4 
      BackColor       =   &H80000012&
      Caption         =   "Calculater"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   2520
      TabIndex        =   4
      Top             =   2400
      Width           =   1095
   End
   Begin VB.Label lblpro3 
      BackColor       =   &H80000012&
      Caption         =   "Moving objects"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   4080
      TabIndex        =   3
      Top             =   2040
      Width           =   1575
   End
   Begin VB.Label lblpro1 
      BackColor       =   &H80000012&
      Caption         =   "Color Palette"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   720
      TabIndex        =   2
      Top             =   2040
      Width           =   1455
   End
   Begin VB.Label lblextra 
      BackColor       =   &H80000008&
      Caption         =   "Extra Programs"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   855
      Left            =   840
      TabIndex        =   0
      Top             =   120
      Width           =   4575
   End
End
Attribute VB_Name = "frmxtrap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub lblpro3_Click()
End Sub
