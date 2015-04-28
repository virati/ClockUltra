VERSION 5.00
Object = "{3AF740FA-BF7F-11D2-8C26-444553540001}#2.0#0"; "SWBTopForm.ocx"
Object = "{A37D0E58-B6D0-11D2-971F-EC500970267D}#7.0#0"; "AnalogClock.ocx"
Object = "{B3ECB8C8-B046-11D2-B522-D0863DEC394B}#2.0#0"; "MagicFrm.ocx"
Begin VB.Form frmClock 
   BackColor       =   &H80000012&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Clock Ultra"
   ClientHeight    =   780
   ClientLeft      =   2040
   ClientTop       =   2505
   ClientWidth     =   3705
   Icon            =   "Clock.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   780
   ScaleWidth      =   3705
   Begin VB.Timer Timer8 
      Interval        =   500
      Left            =   2160
      Top             =   3600
   End
   Begin VB.Timer Timer7 
      Interval        =   500
      Left            =   4560
      Top             =   120
   End
   Begin AnalogClockControl.AnalogClock anolog1 
      Height          =   735
      Left            =   3840
      Top             =   0
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   1296
      Value           =   0.5
      HandsColor      =   16777215
      BackColor       =   -2147483630
   End
   Begin VB.Timer Timer6 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   840
      Top             =   4680
   End
   Begin VB.Timer Timer5 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   2040
      Top             =   4200
   End
   Begin VB.Timer Timer4 
      Interval        =   500
      Left            =   3240
      Top             =   3720
   End
   Begin MagicForms.SystemTray system1 
      Left            =   120
      Top             =   3720
      _ExtentX        =   2223
      _ExtentY        =   847
      TextoDica       =   "Clock Ultra 2.00"
      Icon            =   "Clock.frx":0442
   End
   Begin SWBTopFormCtl.SWBTopForm ontop1 
      Left            =   7080
      Top             =   0
      _ExtentX        =   423
      _ExtentY        =   423
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   6600
      Top             =   0
   End
   Begin VB.Timer Timer2 
      Interval        =   10000
      Left            =   0
      Top             =   360
   End
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   0
      Top             =   0
   End
   Begin VB.Label Label3 
      Height          =   375
      Left            =   6000
      TabIndex        =   8
      Top             =   240
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Image Image1 
      Height          =   255
      Left            =   3360
      Picture         =   "Clock.frx":0894
      Stretch         =   -1  'True
      Top             =   0
      Width           =   255
   End
   Begin VB.Label Label7 
      BackColor       =   &H80000012&
      Caption         =   "Date>>"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3120
      TabIndex        =   7
      Top             =   960
      Width           =   495
   End
   Begin VB.Label lbltimecol 
      BackColor       =   &H80000012&
      Caption         =   "Time color"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   1920
      TabIndex        =   6
      ToolTipText     =   "This shows the color of the time"
      Top             =   0
      Width           =   855
   End
   Begin VB.Label Label5 
      BackColor       =   &H80000012&
      Caption         =   "Extra"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   3240
      TabIndex        =   5
      ToolTipText     =   "This brings the extra menu"
      Top             =   480
      Width           =   375
   End
   Begin VB.Label Label4 
      BackColor       =   &H80000008&
      Caption         =   "Main>>"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   3120
      TabIndex        =   4
      ToolTipText     =   "Click to get the Main menu\Double click to get the Settings Menu"
      Top             =   240
      Width           =   615
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000008&
      Caption         =   "Pause"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   1200
      TabIndex        =   3
      ToolTipText     =   "The clock is paused"
      Top             =   0
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000012&
      Caption         =   "Date"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   600
      TabIndex        =   2
      ToolTipText     =   "The date is showing"
      Top             =   0
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lbldate 
      BackColor       =   &H80000012&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   960
      TabIndex        =   1
      Top             =   720
      Width           =   1695
   End
   Begin VB.Label lbltime 
      BackColor       =   &H80000012&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   240
      TabIndex        =   0
      ToolTipText     =   "Double Click to resume\pause"
      Top             =   240
      Width           =   2895
   End
   Begin VB.Menu main 
      Caption         =   "Ma&in"
      Begin VB.Menu dat 
         Caption         =   "D&ate"
         Shortcut        =   ^D
      End
      Begin VB.Menu al 
         Caption         =   "A&larm"
         Shortcut        =   ^L
      End
      Begin VB.Menu aclock 
         Caption         =   "A&nalog Clock"
         Shortcut        =   ^A
      End
      Begin VB.Menu timerer 
         Caption         =   "T&imer"
      End
      Begin VB.Menu OF 
         Caption         =   "Other functions"
      End
      Begin VB.Menu l3 
         Caption         =   "-"
      End
      Begin VB.Menu calender 
         Caption         =   "C&alender"
         Shortcut        =   ^C
      End
      Begin VB.Menu l1 
         Caption         =   "-"
      End
      Begin VB.Menu ex 
         Caption         =   "E&xit"
         Shortcut        =   {F12}
      End
   End
   Begin VB.Menu sett 
      Caption         =   "Sett&ings"
      Begin VB.Menu loca 
         Caption         =   "Location"
         Begin VB.Menu usa 
            Caption         =   "U.S."
            Checked         =   -1  'True
         End
         Begin VB.Menu india 
            Caption         =   "India"
         End
         Begin VB.Menu kuw 
            Caption         =   "Kuwait"
         End
      End
      Begin VB.Menu l4 
         Caption         =   "-"
      End
      Begin VB.Menu tf 
         Caption         =   "Time Format"
         Begin VB.Menu tfsec 
            Caption         =   "Seconds"
            Checked         =   -1  'True
         End
         Begin VB.Menu tfmin 
            Caption         =   "Minutes"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu datefom 
         Caption         =   "Date Format"
         Begin VB.Menu date1 
            Caption         =   "1-12-99"
            Checked         =   -1  'True
         End
         Begin VB.Menu date2 
            Caption         =   "12-1-99"
         End
         Begin VB.Menu date3 
            Caption         =   "99-12-1"
         End
      End
      Begin VB.Menu d2 
         Caption         =   "-"
      End
      Begin VB.Menu wsett 
         Caption         =   "Window Settings"
         Begin VB.Menu aotop 
            Caption         =   "Always on Top"
         End
         Begin VB.Menu normalwin 
            Caption         =   "Normal"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu l11 
         Caption         =   "-"
      End
      Begin VB.Menu TO 
         Caption         =   "Time Options"
         Begin VB.Menu aeh 
            Caption         =   "Alert Every Hour"
         End
         Begin VB.Menu buaf 
            Caption         =   "Bring up Alert Form"
         End
         Begin VB.Menu da 
            Caption         =   "Disable All"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu l10 
         Caption         =   "-"
      End
      Begin VB.Menu f 
         Caption         =   "F&ont"
         Begin VB.Menu col 
            Caption         =   "Colo&rs"
            Begin VB.Menu blue 
               Caption         =   "B&lue"
            End
            Begin VB.Menu red 
               Caption         =   "R&ed"
               Checked         =   -1  'True
            End
            Begin VB.Menu green 
               Caption         =   "Gree&n"
            End
            Begin VB.Menu L9 
               Caption         =   "-"
            End
            Begin VB.Menu custom 
               Caption         =   "&Custom"
            End
         End
      End
   End
   Begin VB.Menu extra 
      Caption         =   "E&xtra"
      Begin VB.Menu PCl 
         Caption         =   "P&ause Clock"
      End
      Begin VB.Menu RCl 
         Caption         =   "R&esume Clock"
         Enabled         =   0   'False
      End
      Begin VB.Menu d3 
         Caption         =   "-"
      End
      Begin VB.Menu sback 
         Caption         =   "Select Background"
      End
      Begin VB.Menu l8 
         Caption         =   "-"
      End
      Begin VB.Menu sist 
         Caption         =   "Show in System Tray"
         Checked         =   -1  'True
      End
      Begin VB.Menu l14 
         Caption         =   "-"
      End
      Begin VB.Menu extras 
         Caption         =   "Extras"
      End
      Begin VB.Menu l7 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu sc 
         Caption         =   "Source Code"
         Visible         =   0   'False
      End
      Begin VB.Menu music 
         Caption         =   "Music"
         Visible         =   0   'False
      End
      Begin VB.Menu links 
         Caption         =   "Links"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu he 
      Caption         =   "H&elp"
      Begin VB.Menu conthelp 
         Caption         =   "&Contents"
         Shortcut        =   {F1}
      End
      Begin VB.Menu l2 
         Caption         =   "-"
      End
      Begin VB.Menu about 
         Caption         =   "A&bout"
      End
   End
   Begin VB.Menu ex2 
      Caption         =   "Extra"
      Visible         =   0   'False
      Begin VB.Menu timer 
         Caption         =   "Timer"
         Begin VB.Menu pc 
            Caption         =   "Pause Clock"
         End
         Begin VB.Menu rc 
            Caption         =   "Resume Clock"
            Enabled         =   0   'False
         End
      End
      Begin VB.Menu l5 
         Caption         =   "-"
      End
      Begin VB.Menu font 
         Caption         =   "Font"
      End
      Begin VB.Menu l6 
         Caption         =   "-"
      End
      Begin VB.Menu acu 
         Caption         =   "About Clock Ultra"
      End
   End
   Begin VB.Menu dt 
      Caption         =   "Date"
      Visible         =   0   'False
      Begin VB.Menu ft 
         Caption         =   "Font"
         Begin VB.Menu blu 
            Caption         =   "Blue"
         End
         Begin VB.Menu gree 
            Caption         =   "Green"
         End
         Begin VB.Menu whit 
            Caption         =   "White"
            Checked         =   -1  'True
         End
      End
   End
   Begin VB.Menu cultra 
      Caption         =   "ClockUltra"
      Visible         =   0   'False
      Begin VB.Menu open1 
         Caption         =   "Open"
      End
      Begin VB.Menu wscu 
         Caption         =   "Window State"
         Begin VB.Menu restorecu 
            Caption         =   "Restore"
            Enabled         =   0   'False
         End
         Begin VB.Menu mcu 
            Caption         =   "Minimize"
         End
      End
      Begin VB.Menu l13 
         Caption         =   "-"
      End
      Begin VB.Menu aboutcu 
         Caption         =   "About ClockUltra"
      End
      Begin VB.Menu l12 
         Caption         =   "-"
      End
      Begin VB.Menu exit2 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu anologclock1 
      Caption         =   "Anolog"
      Visible         =   0   'False
      Begin VB.Menu showanalog 
         Caption         =   "Show"
         Checked         =   -1  'True
      End
      Begin VB.Menu HColor 
         Caption         =   "Hand Color"
         Begin VB.Menu whitanalog 
            Caption         =   "White"
            Checked         =   -1  'True
         End
         Begin VB.Menu bluanalog 
            Caption         =   "Blue"
         End
         Begin VB.Menu redanalog 
            Caption         =   "Red"
         End
      End
      Begin VB.Menu BColor 
         Caption         =   "Background color"
         Begin VB.Menu whitback 
            Caption         =   "White"
         End
         Begin VB.Menu greeback 
            Caption         =   "Green"
         End
         Begin VB.Menu blacback 
            Caption         =   "Black"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu TOanolog 
         Caption         =   "Time Options"
         Begin VB.Menu ssanalog 
            Caption         =   "Show Seconds"
            Checked         =   -1  'True
         End
         Begin VB.Menu smanalog 
            Caption         =   "Show Minutes"
         End
         Begin VB.Menu SHanalog 
            Caption         =   "Show Hours"
            Checked         =   -1  'True
         End
      End
   End
End
Attribute VB_Name = "frmClock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub about_Click()
frmAbout.Show
End Sub

Private Sub aboutcu_Click()
frmAbout.Show
End Sub

Private Sub aclock_Click()
If frmClock.Width = 3795 Then
frmClock.Width = 4650
aclock.Checked = True
Else
frmClock.Width = 3795
aclock.Checked = False
End If
End Sub

Private Sub acu_Click()
frmAbout.Show

End Sub

Private Sub aeh_Click()
If aeh.Checked = False Then
aeh.Checked = True
Timer5.Enabled = True
buaf.Checked = False
da.Checked = False
Else
aeh.Checked = False
Timer5.Enabled = False
buaf.Checked = True
End If
End Sub

Private Sub al_Click()
frmAlarm.Show
End Sub

Private Sub anolog1_Click()
frmClock.PopupMenu anologclock1
End Sub

Private Sub aotop_Click()
normalwin.Checked = False
aotop.Checked = True
ontop1.Enabled = True
End Sub

Private Sub blacback_Click()
blacback.Checked = True
whitback.Checked = False
greeback.Checked = False
anolog1.BackColor = vbBlack
End Sub

Private Sub blu_Click()
lbldate.ForeColor = vbBlue
blu.Checked = True
gree.Checked = False
whit.Checked = False
End Sub

Private Sub bluanalog_Click()
anolog1.HandsColor = vbBlue
bluanalog.Checked = True
redanalog.Checked = False
whitanalog.Checked = False
End Sub

Private Sub blue_Click()
lbltime.ForeColor = vbBlue
blue.Checked = True
red.Checked = False
green.Checked = False
custom.Checked = False
lbltimecol.ForeColor = lbltime.ForeColor
End Sub

Private Sub buaf_Click()
If buaf.Checked = False Then
buaf.Checked = True
Timer6.Enabled = True
aeh.Checked = False
Timer5.Enabled = False
da.Checked = False
Else
aeh.Checked = True
Timer5.Enabled = True
Timer6.Enabled = False
buaf.Checked = False
End If
End Sub

Private Sub calender_Click()
frmCal.Show
End Sub

Private Sub custom_Click()
frmcolorlbl.Show
custom.Checked = True
blue.Checked = False
red.Checked = False
green.Checked = False
End Sub

Private Sub da_Click()
If da.Checked = False Then
da.Checked = True
Timer5.Enabled = False
Timer6.Enabled = False
buaf.Checked = False
aeh.Checked = False
Else
da.Checked = False
Timer5.Enabled = True
aeh.Checked = True
buaf.Checked = False
Timer6.Enabled = False
End If
End Sub

Private Sub dat_Click()
If frmClock.Height = 1395 Then
        frmClock.Height = 1830
    Else
        frmClock.Height = 1395
End If
If frmClock.Height = 1830 Then
        dat.Checked = True
    Else
        dat.Checked = False
End If
If Label1.Visible = False Then
        Label1.Visible = True
    Else
        Label1.Visible = False
End If
End Sub

Private Sub EPS_Click()
frmxtrap.Show

End Sub

Private Sub date1_Click()
date1.Checked = True
date2.Checked = False
date3.Checked = False
lbldate.Caption = Format(Now, "d/m/yy")
End Sub

Private Sub date2_Click()
date2.Checked = True
date1.Checked = False
date3.Checked = False
lbldate.Caption = Format(Now, "m/d/yy")
End Sub

Private Sub date3_Click()
date3.Checked = True
date2.Checked = False
date1.Checked = False
lbldate.Caption = Format(Now, "yy/m/d")
End Sub

Private Sub ex_Click()
End
Unload Me
End Sub

Private Sub exit2_Click()
End
End Sub

Private Sub extras_Click()
frmextras.Show
End Sub

Private Sub Form_Load()
lbltime.Caption = Time
lbldate.Caption = Format(Now, "d/m/yy")
ontop1.Enabled = False
'there is goung to be a sample of lock Ultra in this Clock ultra
'release.
'there will also be links to the web.
'there could also be a form about Clock ultra that says that a donation
'of $5 would be nice.
'there is also going to be a midi\sound player and a browser
'in other releases there is going to be a wav sound thing and color
'things too!
system1.Show
End Sub

Private Sub options_Click()
frmOptions.Show
End Sub

Private Sub Form_LostFocus()
frmClock.Enabled = False

End Sub

Private Sub Form_Unload(Cancel As Integer)
system1.Hide
End Sub

Private Sub gree_Click()
lbldate.ForeColor = vbGreen
gree.Checked = True
whit.Checked = False
blu.Checked = False
End Sub

Private Sub greeback_Click()
greeback.Checked = True
whitback.Checked = False
blacback.Checked = False
anolog1.BackColor = vbGreen
End Sub

Private Sub green_Click()
lbltime.ForeColor = vbGreen
green.Checked = True
red.Checked = False
blue.Checked = False
custom.Checked = False
lbltimecol.ForeColor = lbltime.ForeColor

End Sub

Private Sub Image1_Click()
frmClock.PopupMenu he
End Sub

Private Sub india_Click()
india.Checked = True
usa.Checked = False
kuw.Checked = False
MsgBox "Sorry, This feature is not available in this version of ClockUltra", , "Sorry"
'insert the code here
End Sub

Private Sub kuw_Click()
kuw.Checked = True
india.Checked = False
usa.Checked = False
MsgBox "Sorry, This feature is not available in this version of ClockUltra", , "Sorry"
'insert code here
End Sub

Private Sub Label4_Click()
frmClock.PopupMenu main

End Sub

Private Sub Label4_DblClick()
frmClock.PopupMenu sett
End Sub

Private Sub Label5_Click()
frmClock.PopupMenu extra

End Sub

Private Sub Label7_Click()
frmClock.PopupMenu dt
End Sub

Private Sub lbldate_DblClick()
frmCal.Show
End Sub

Private Sub lbltime_DblClick()
If Timer1.Enabled = True Then
Timer1.Enabled = False
Label2.Visible = True
PCl.Enabled = False
RCl.Enabled = True
pc.Enabled = False
rc.Enabled = True
Else
Timer1.Enabled = True
Label2.Visible = False
RCl.Enabled = False
PCl.Enabled = True
pc.Enabled = True
rc.Enabled = False
End If
End Sub

Private Sub lbltimecol_Click()
frmClock.PopupMenu f
End Sub

Private Sub nf_Click()
If nf.Caption = "No frame" Then
frmClock.BorderStyle = 0
nf.Caption = "Frame"
Else
frmClock.BorderStyle = 1
nf.Caption = "No frame"
End If
End Sub

Private Sub links_Click()
frmlinks.Show
End Sub

Private Sub mcu_Click()
frmClock.WindowState = 1
mcu.Enabled = False
restorecu.Enabled = True
End Sub

Private Sub music_Click()
frmthrone.Show
End Sub

Private Sub normalwin_Click()
aotop.Checked = False
ontop1.Enabled = False
End Sub

Private Sub OF_Click()
frmextra.Show
End Sub

Private Sub open1_Click()
frmClock.WindowState = Normal
frmClock.SetFocus
End Sub

Private Sub pc_Click()
Timer1.Enabled = False
PCl.Enabled = False
RCl.Enabled = True
pc.Enabled = False
rc.Enabled = True
Label2.Visible = True

End Sub

Private Sub PCl_Click()
Timer1.Enabled = False
PCl.Enabled = False
RCl.Enabled = True
pc.Enabled = False
rc.Enabled = True
Label2.Visible = True
Timer7.Enabled = False
End Sub

Private Sub rc_Click()
Timer1.Enabled = True
RCl.Enabled = False
PCl.Enabled = True
rc.Enabled = False
pc.Enabled = True
Label2.Visible = False
'Timer7.Enabled = True
End Sub

Private Sub RCl_Click()
Timer1.Enabled = True
RCl.Enabled = False
PCl.Enabled = True
rc.Enabled = False
pc.Enabled = True
Label2.Visible = False
Timer7.Enabled = True
End Sub

Private Sub re_Click()
lbldate.ForeColor = vbRed
re.Checked = True
gree.Checked = False
blu.Checked = False
End Sub

Private Sub red_Click()
lbltime.ForeColor = vbRed
red.Checked = True
green.Checked = False
blue.Checked = False
custom.Checked = False
lbltimecol.ForeColor = lbltime.ForeColor
End Sub

Private Sub redanalog_Click()
anolog1.HandsColor = vbRed
redanalog.Checked = True
bluanalog.Checked = False
whitanalog.Checked = False
End Sub

Private Sub restorecu_Click()
frmClock.WindowState = Normal
restorecu.Enabled = False
mcu.Enabled = True
End Sub

Private Sub sback_Click()
frmcolor.Show
End Sub

Private Sub sc_Click()
frmcode.Show
End Sub

Private Sub sdwanalog_Click()
frmanalog.Show
End Sub

Private Sub SHanalog_Click()
If SHanalog.Checked = True Then
SHanalog.Checked = False
anolog1.ShowHours = False
Else
SHanalog.Checked = True
anolog1.ShowHours = True
End If
End Sub

Private Sub showanalog_Click()
showanalog.Checked = False
frmClock.Width = 3795
End Sub

Private Sub sist_Click()
If sist.Checked = True Then
sist.Checked = False
system1.Hide
Else
sist.Checked = True
system1.Show
End If
End Sub

Private Sub smanalog_Click()
If smanalog.Checked = False Then
smanalog.Checked = True
anolog1.ShowMinutes = True
Else
smanalog.Checked = False
anolog1.ShowMinutes = False
End If
End Sub

Private Sub ssanalog_Click()
If ssanalog.Checked = True Then
ssanalog.Checked = False
anolog1.ShowSecondsHand = False
Else
anolog1.ShowSecondsHand = True
ssanalog.Checked = True
End If
End Sub

Private Sub system1_LeftButtonDBlClick()
frmClock.PopupMenu sett
End Sub

Private Sub system1_RightButtonUp()
frmClock.PopupMenu cultra
End Sub

Private Sub tfmin_Click()
If tfmin.Checked = True Then
tfmin.Checked = False
Else
tfmin.Checked = True
End If
End Sub

Private Sub tfsec_Click()
If tfsec.Checked = True Then
tfsec.Checked = False
Else
tfsec.Checked = True
End If
End Sub

Private Sub Timer1_Timer()
lbltime.Caption = Time
End Sub

Private Sub Timer2_Timer()
lbldate.Caption = Date
End Sub

Private Sub Timer3_Timer()
If lbltime.Caption = Label3.Caption Then
Beep
MsgBox "Alarm at " & Label3.Caption, vbOKOnly
Label3.Caption = ""
frmAlarm.msktime.Text = ""
frmAlarm.Command3.Visible = False
frmAlarm.Combo1.Text = ""
frmAlarm.Command2.Caption = "Can&cel"
End If
End Sub

Private Sub Timer4_Timer()
system1.TextoDica = Time
End Sub

Private Sub Timer5_Timer()
If Label1.Caption = "1:00:00 PM" Then
MsgBox "The time is 1:00 PM", , "Alert"
ElseIf Label1.Caption = "2:00:00 PM" Then
MsgBox "The time is 2:00 PM", , "Alert"
ElseIf Label1.Caption = "3:00:00 PM" Then
MsgBox "The time is 3:00 PM", , "Alert"
ElseIf Label1.Caption = "4:00:00 PM" Then
MsgBox "The time is 4:00 PM", , "Alert"
ElseIf Label1.Caption = "5:00:00 PM" Then
MsgBox "The time is 5:00 PM", , "Alert"
ElseIf Label1.Caption = "6:00:00 PM" Then
MsgBox "The time is 6:00 PM", , "Alert"
ElseIf Label1.Caption = "7:00:00 PM" Then
MsgBox "The time is 7:00 PM", , "Alert"
ElseIf Label1.Caption = "8:00:00 PM" Then
MsgBox "The time is 8:00 PM", , "Alert"
ElseIf Label1.Caption = "9:00:00 PM" Then
MsgBox "The time is 9:00 PM", , "Alert"
ElseIf Label1.Caption = "10:00:00 PM" Then
MsgBox "The time is 10:00 PM", , "Alert"
ElseIf Label1.Caption = "11:00:00 PM" Then
MsgBox "The time is 11:00 PM", , "Alert"
ElseIf Label1.Caption = "12:00:00 PM" Then
MsgBox "The time is 12:00 PM", , "Alert"
ElseIf Label1.Caption = "1:00:00 AM" Then
MsgBox "The time is 1:00 AM", , "Alert"
ElseIf Label1.Caption = "2:00:00 AM" Then
MsgBox "The time is 2:00 AM", , "Alert"
ElseIf Label1.Caption = "3:00:00 AM" Then
MsgBox "The time is 3:00 AM", , "Alert"
ElseIf Label1.Caption = "4:00:00 AM" Then
MsgBox "The time is 4:00 AM", , "Alert"
ElseIf Label1.Caption = "5:00:00 AM" Then
MsgBox "The time is 5:00 AM", , "Alert"
ElseIf Label1.Caption = "6:00:00 AM" Then
MsgBox "The time is 6:00 AM", , "Alert"
ElseIf Label1.Caption = "7:00:00 AM" Then
MsgBox "The time is 7:00 AM", , "Alert"
ElseIf Label1.Caption = "8:00:00 AM" Then
MsgBox "The time is 8:00 AM", , "Alert"
ElseIf Label1.Caption = "9:00:00 AM" Then
MsgBox "The time is 9:00 AM", , "Alert"
ElseIf Label1.Caption = "10:00:00 AM" Then
MsgBox "The time is 10:00 AM", , "Alert"
ElseIf Label1.Caption = "11:00:00 AM" Then
MsgBox "The time is 11:00 AM", , "Alert"
ElseIf Label1.Caption = "12:00:00 AM" Then
MsgBox "The time is 12:00 AM", , "Alert"
End If
End Sub

Private Sub Timer6_Timer()
If Label1.Caption = "1:00:00 PM" Then
frmalert.Show
ElseIf Label1.Caption = "2:00:00 PM" Then
frmalert.Show
ElseIf Label1.Caption = "3:00:00 PM" Then
frmalert.Show
ElseIf Label1.Caption = "4:00:00 PM" Then
frmalert.Show
ElseIf Label1.Caption = "5:00:00 PM" Then
frmalert.Show
ElseIf Label1.Caption = "6:00:00 PM" Then
frmalert.Show
ElseIf Label1.Caption = "7:00:00 PM" Then
frmalert.Show
ElseIf Label1.Caption = "8:00:00 PM" Then
frmalert.Show
ElseIf Label1.Caption = "9:00:00 PM" Then
frmalert.Show
ElseIf Label1.Caption = "10:00:00 PM" Then
frmalert.Show
ElseIf Label1.Caption = "11:00:00 PM" Then
frmalert.Show
ElseIf Label1.Caption = "12:00:00 PM" Then
frmalert.Show
ElseIf Label1.Caption = "1:00:00 AM" Then
frmalert.Show
ElseIf Label1.Caption = "2:00:00 AM" Then
frmalert.Show
ElseIf Label1.Caption = "3:00:00 AM" Then
frmalert.Show
ElseIf Label1.Caption = "4:00:00 AM" Then
frmalert.Show
ElseIf Label1.Caption = "5:00:00 AM" Then
frmalert.Show
ElseIf Label1.Caption = "6:00:00 AM" Then
frmalert.Show
ElseIf Label1.Caption = "7:00:00 AM" Then
frmalert.Show
ElseIf Label1.Caption = "8:00:00 AM" Then
frmalert.Show
ElseIf Label1.Caption = "9:00:00 AM" Then
frmalert.Show
ElseIf Label1.Caption = "10:00:00 AM" Then
frmalert.Show
ElseIf Label1.Caption = "11:00:00 AM" Then
frmalert.Show
ElseIf Label1.Caption = "12:00:00 AM" Then
frmalert.Show
End If
End Sub

Private Sub Timer7_Timer()
anolog1.Value = Time
End Sub

Private Sub Timer8_Timer()
system1.TextoDica = "" & Date & "," & Time

End Sub

Private Sub timerer_Click()
frmtimer.Show
End Sub

Private Sub usa_Click()
usa.Checked = True
india.Checked = False
kuw.Checked = False
MsgBox "Sorry, This feature is not available in this version of ClockUltra", , "Sorry"
'insert code here
End Sub

Private Sub whit_Click()
lbldate.ForeColor = vbWhite
whit.Checked = True
gree.Checked = False
blu.Checked = False
End Sub

Private Sub whitanalog_Click()
anolog1.HandsColor = vbWhite
whitanalog.Checked = True
redanalog.Checked = False
bluanalog.Checked = False
End Sub

Private Sub whitback_Click()
whitback.Checked = True
blacback.Checked = False
greeback.Checked = False
anolog1.BackColor = vbWhite
End Sub
