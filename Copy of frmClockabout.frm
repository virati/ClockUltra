VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H80000012&
   BorderStyle     =   0  'None
   Caption         =   "About Clock Ultra"
   ClientHeight    =   3975
   ClientLeft      =   2940
   ClientTop       =   1875
   ClientWidth     =   5730
   ClipControls    =   0   'False
   Icon            =   "frmClockabout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2743.616
   ScaleMode       =   0  'User
   ScaleWidth      =   5380.766
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   615
      Left            =   1440
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   7
      Text            =   "frmClockabout.frx":0742
      Top             =   2280
      Width           =   2535
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   0
      Top             =   840
   End
   Begin VB.PictureBox picIcon 
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      ClipControls    =   0   'False
      Height          =   540
      Left            =   240
      Picture         =   "frmClockabout.frx":0757
      ScaleHeight     =   337.12
      ScaleMode       =   0  'User
      ScaleWidth      =   337.12
      TabIndex        =   1
      ToolTipText     =   "Clock Ultra Version 1.0"
      Top             =   240
      Width           =   540
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   4320
      TabIndex        =   0
      ToolTipText     =   "Back to main Clock"
      Top             =   3120
      Width           =   1260
   End
   Begin VB.CommandButton cmdSysInfo 
      Caption         =   "&System Info..."
      Height          =   345
      Left            =   4320
      TabIndex        =   2
      ToolTipText     =   "Click to see system Info"
      Top             =   3600
      Width           =   1245
   End
   Begin VB.Image Image4 
      Height          =   480
      Left            =   4440
      Picture         =   "frmClockabout.frx":0B99
      Top             =   2400
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image3 
      Height          =   495
      Left            =   120
      Top             =   2160
      Width           =   495
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   5160
      Picture         =   "frmClockabout.frx":0FDB
      Top             =   2400
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label lblDescription 
      BackColor       =   &H80000012&
      Caption         =   $"frmClockabout.frx":141D
      ForeColor       =   &H00FF0000&
      Height          =   1170
      Left            =   960
      TabIndex        =   6
      Top             =   960
      Width           =   3885
   End
   Begin VB.Image Image1 
      Height          =   375
      Left            =   5280
      Picture         =   "frmClockabout.frx":1533
      Stretch         =   -1  'True
      ToolTipText     =   "Click to see help"
      Top             =   0
      Width           =   375
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   112.686
      X2              =   5337.57
      Y1              =   2070.654
      Y2              =   2070.654
   End
   Begin VB.Label lblTitle 
      BackColor       =   &H80000012&
      Caption         =   "      Clock Ultra"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   480
      Left            =   1050
      TabIndex        =   4
      ToolTipText     =   "Clock Ultra"
      Top             =   240
      Width           =   3885
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   112.686
      X2              =   5323.484
      Y1              =   2070.654
      Y2              =   2070.654
   End
   Begin VB.Label lblVersion 
      BackColor       =   &H80000012&
      Caption         =   "Version: 2.0"
      ForeColor       =   &H8000000C&
      Height          =   225
      Left            =   1050
      TabIndex        =   5
      ToolTipText     =   "Version: 1.0"
      Top             =   780
      Width           =   3885
   End
   Begin VB.Label lblwarning 
      BackColor       =   &H80000012&
      Caption         =   "Warning: This program is Copywrited. Copying this program can lead to prison and a 1000 doller fine.(Just Joking!)"
      ForeColor       =   &H00000000&
      Height          =   825
      Left            =   240
      TabIndex        =   3
      Top             =   3120
      Width           =   3975
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Reg Key Security Options...
Const READ_CONTROL = &H20000
Const KEY_QUERY_VALUE = &H1
Const KEY_SET_VALUE = &H2
Const KEY_CREATE_SUB_KEY = &H4
Const KEY_ENUMERATE_SUB_KEYS = &H8
Const KEY_NOTIFY = &H10
Const KEY_CREATE_LINK = &H20
Const KEY_ALL_ACCESS = KEY_QUERY_VALUE + KEY_SET_VALUE + _
                       KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + _
                       KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL
                     
' Reg Key ROOT Types...
Const HKEY_LOCAL_MACHINE = &H80000002
Const ERROR_SUCCESS = 0
Const REG_SZ = 1                         ' Unicode nul terminated string
Const REG_DWORD = 4                      ' 32-bit number

Const gREGKEYSYSINFOLOC = "SOFTWARE\Microsoft\Shared Tools Location"
Const gREGVALSYSINFOLOC = "MSINFO"
Const gREGKEYSYSINFO = "SOFTWARE\Microsoft\Shared Tools\MSINFO"
Const gREGVALSYSINFO = "PATH"

Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long


Private Sub cmdSysInfo_Click()
  Call StartSysInfo
End Sub

Private Sub cmdOK_Click()
  Unload Me
End Sub

Private Sub Form_Load()
    Me.Caption = "About " & App.Title
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    lblTitle.Caption = App.Title
Image3.Picture = Image2.Picture
End Sub

Public Sub StartSysInfo()
    On Error GoTo SysInfoErr
  
    Dim rc As Long
    Dim SysInfoPath As String
    
    ' Try To Get System Info Program Path\Name From Registry...
    If GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFO, gREGVALSYSINFO, SysInfoPath) Then
    ' Try To Get System Info Program Path Only From Registry...
    ElseIf GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFOLOC, gREGVALSYSINFOLOC, SysInfoPath) Then
        ' Validate Existance Of Known 32 Bit File Version
        If (Dir(SysInfoPath & "\MSINFO32.EXE") <> "") Then
            SysInfoPath = SysInfoPath & "\MSINFO32.EXE"
            
        ' Error - File Can Not Be Found...
        Else
            GoTo SysInfoErr
        End If
    ' Error - Registry Entry Can Not Be Found...
    Else
        GoTo SysInfoErr
    End If
    
    Call Shell(SysInfoPath, vbNormalFocus)
    
    Exit Sub
SysInfoErr:
    MsgBox "System Information Is Unavailable At This Time", vbOKOnly
End Sub

Public Function GetKeyValue(KeyRoot As Long, KeyName As String, SubKeyRef As String, ByRef KeyVal As String) As Boolean
    Dim i As Long                                           ' Loop Counter
    Dim rc As Long                                          ' Return Code
    Dim hKey As Long                                        ' Handle To An Open Registry Key
    Dim hDepth As Long                                      '
    Dim KeyValType As Long                                  ' Data Type Of A Registry Key
    Dim tmpVal As String                                    ' Tempory Storage For A Registry Key Value
    Dim KeyValSize As Long                                  ' Size Of Registry Key Variable
    '------------------------------------------------------------
    ' Open RegKey Under KeyRoot {HKEY_LOCAL_MACHINE...}
    '------------------------------------------------------------
    rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey) ' Open Registry Key
    
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Handle Error...
    
    tmpVal = String$(1024, 0)                             ' Allocate Variable Space
    KeyValSize = 1024                                       ' Mark Variable Size
    
    '------------------------------------------------------------
    ' Retrieve Registry Key Value...
    '------------------------------------------------------------
    rc = RegQueryValueEx(hKey, SubKeyRef, 0, _
                         KeyValType, tmpVal, KeyValSize)    ' Get/Create Key Value
                        
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Handle Errors
    
    If (Asc(Mid(tmpVal, KeyValSize, 1)) = 0) Then           ' Win95 Adds Null Terminated String...
        tmpVal = Left(tmpVal, KeyValSize - 1)               ' Null Found, Extract From String
    Else                                                    ' WinNT Does NOT Null Terminate String...
        tmpVal = Left(tmpVal, KeyValSize)                   ' Null Not Found, Extract String Only
    End If
    '------------------------------------------------------------
    ' Determine Key Value Type For Conversion...
    '------------------------------------------------------------
    Select Case KeyValType                                  ' Search Data Types...
    Case REG_SZ                                             ' String Registry Key Data Type
        KeyVal = tmpVal                                     ' Copy String Value
    Case REG_DWORD                                          ' Double Word Registry Key Data Type
        For i = Len(tmpVal) To 1 Step -1                    ' Convert Each Bit
            KeyVal = KeyVal + Hex(Asc(Mid(tmpVal, i, 1)))   ' Build Value Char. By Char.
        Next
        KeyVal = Format$("&h" + KeyVal)                     ' Convert Double Word To String
    End Select
    
    GetKeyValue = True                                      ' Return Success
    rc = RegCloseKey(hKey)                                  ' Close Registry Key
    Exit Function                                           ' Exit
    
GetKeyError:      ' Cleanup After An Error Has Occured...
    KeyVal = ""                                             ' Set Return Val To Empty String
    GetKeyValue = False                                     ' Return Failure
    rc = RegCloseKey(hKey)                                  ' Close Registry Key
End Function

Private Sub Form_LostFocus()
frmAbout.SetFocus
End Sub

Private Sub Form_Unload(Cancel As Integer)
frmClock.Enabled = True
frmClock.Show
End Sub

Private Sub Image1_Click()
frmAbout.PopupMenu frmClock.he
End Sub

Private Sub Image2_Click()
If Timer1.Enabled = True Then
Timer1.Enabled = False
lblwarning.ForeColor = vbBlue
Else
Timer1.Enabled = True
lblwarning.ForeColor = vbBlack
End If
End Sub

Private Sub Image3_Click()
If Image3.Picture = Image2.Picture Then
Timer1.Enabled = False
Image3.Picture = Image4.Picture
lblwarning.ForeColor = vbRed
Else
Image3.Picture = Image2.Picture
Timer1.Enabled = True
End If
End Sub

Private Sub SoundRec1_GotFocus()

End Sub

Private Sub lblVersion_Click()
'the version will become 1.0 when i am done.
End Sub

Private Sub Timer1_Timer()
If lblwarning.ForeColor = vbBlack Then
lblwarning.ForeColor = vbRed
Else
lblwarning.ForeColor = vbBlack
End If
End Sub
