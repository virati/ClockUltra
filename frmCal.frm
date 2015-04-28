VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "MSCAL.OCX"
Begin VB.Form frmCal 
   BackColor       =   &H8000000B&
   BorderStyle     =   0  'None
   Caption         =   "Calender"
   ClientHeight    =   4380
   ClientLeft      =   4710
   ClientTop       =   2220
   ClientWidth     =   2715
   Icon            =   "frmCal.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4380
   ScaleWidth      =   2715
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command3 
      Caption         =   "Find Date"
      Height          =   495
      Left            =   0
      TabIndex        =   5
      ToolTipText     =   "click to view the date above in the calender"
      Top             =   2880
      Width           =   2655
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   0
      TabIndex        =   4
      ToolTipText     =   "Choose a date and click find date to see the date in the calender"
      Top             =   2400
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   661
      _Version        =   393216
      Format          =   22937601
      CurrentDate     =   36267
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Ch&ange Calender"
      Height          =   495
      Left            =   0
      TabIndex        =   1
      ToolTipText     =   "click to change the calender"
      Top             =   3360
      Width           =   2655
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ba&ck"
      Height          =   495
      Left            =   0
      TabIndex        =   0
      ToolTipText     =   "click to go back to the main clock"
      Top             =   3840
      Width           =   2655
   End
   Begin MSACAL.Calendar Calendar1 
      Height          =   2415
      Left            =   0
      TabIndex        =   2
      ToolTipText     =   "This calender is to find a day in the future/past"
      Top             =   0
      Visible         =   0   'False
      Width           =   4455
      _Version        =   524288
      _ExtentX        =   7858
      _ExtentY        =   4260
      _StockProps     =   1
      BackColor       =   -2147483637
      Year            =   1999
      Month           =   4
      Day             =   16
      DayLength       =   1
      MonthLength     =   2
      DayFontColor    =   0
      FirstDay        =   1
      GridCellEffect  =   1
      GridFontColor   =   10485760
      GridLinesColor  =   -2147483632
      ShowDateSelectors=   -1  'True
      ShowDays        =   -1  'True
      ShowHorizontalGrid=   -1  'True
      ShowTitle       =   -1  'True
      ShowVerticalGrid=   -1  'True
      TitleFontColor  =   10485760
      ValueIsNull     =   0   'False
      BeginProperty DayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty GridFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComCtl2.MonthView MonthView1 
      Height          =   2370
      Left            =   0
      TabIndex        =   3
      ToolTipText     =   "This calender is to find a day in the future/past"
      Top             =   0
      Width           =   2700
      _ExtentX        =   4763
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483638
      Appearance      =   1
      StartOfWeek     =   22937601
      CurrentDate     =   36266
   End
   Begin VB.Image Image1 
      Height          =   735
      Left            =   2880
      Picture         =   "frmCal.frx":0442
      Stretch         =   -1  'True
      Top             =   2520
      Width           =   1335
   End
End
Attribute VB_Name = "frmCal"
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
If MonthView1.Visible = True Then
Calendar1.Visible = True
MonthView1.Visible = False
Else
MonthView1.Visible = True
Calendar1.Visible = False
End If
If frmCal.Width = 2715 Then
frmCal.Width = 4455
Else
frmCal.Width = 2715
End If
End Sub

Private Sub Command3_Click()
If frmCal.Width = 2715 Then
MonthView1.Year = DTPicker1.Year
MonthView1.Month = DTPicker1.Month
MonthView1.Day = DTPicker1.Day
Else
Calendar1.Year = DTPicker1.Year
Calendar1.Month = DTPicker1.Month
Calendar1.Day = DTPicker1.Day
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
frmClock.Enabled = True
frmClock.Show
End Sub

Private Sub Image1_Click()
Calendar1.Visible = False
MonthView1.Visible = True
frmCal.Width = 2715
End Sub

