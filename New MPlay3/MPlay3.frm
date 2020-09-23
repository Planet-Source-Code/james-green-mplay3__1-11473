VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{22D6F304-B0F6-11D0-94AB-0080C74C7E95}#1.0#0"; "MSDXM.OCX"
Begin VB.Form MPlay3 
   BorderStyle     =   0  'None
   Caption         =   "MPlay3"
   ClientHeight    =   8820
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4110
   Icon            =   "MPlay3.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "MPlay3.frx":0CCA
   ScaleHeight     =   8820
   ScaleWidth      =   4110
   Begin VB.PictureBox PicDisplay 
      BackColor       =   &H00704700&
      BorderStyle     =   0  'None
      Height          =   345
      Left            =   1470
      ScaleHeight     =   345
      ScaleWidth      =   2325
      TabIndex        =   13
      Top             =   420
      Width           =   2325
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "MPlay3"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   165
         Left            =   90
         TabIndex        =   14
         Top             =   15
         Width           =   480
      End
   End
   Begin VB.Timer Timer7 
      Interval        =   100
      Left            =   1530
      Top             =   7245
   End
   Begin VB.TextBox ext 
      Height          =   270
      Left            =   2910
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   6975
      Width           =   870
   End
   Begin VB.ListBox List1 
      Height          =   255
      Left            =   660
      Sorted          =   -1  'True
      TabIndex        =   11
      Top             =   6915
      Width           =   2085
   End
   Begin VB.Timer Timer6 
      Interval        =   50
      Left            =   3120
      Top             =   7260
   End
   Begin VB.Timer Timer4 
      Interval        =   1000
      Left            =   2745
      Top             =   7260
   End
   Begin VB.Timer Timer3 
      Interval        =   60
      Left            =   2325
      Top             =   7260
   End
   Begin VB.Timer toppicture 
      Interval        =   60
      Left            =   1935
      Top             =   7260
   End
   Begin VB.Timer TVolum 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   1140
      Top             =   7260
   End
   Begin VB.Timer Timer1 
      Interval        =   400
      Left            =   765
      Top             =   7260
   End
   Begin VB.ListBox lstfavs 
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      ForeColor       =   &H00FFFFFF&
      Height          =   2955
      Left            =   195
      TabIndex        =   5
      Top             =   2025
      Width           =   3585
   End
   Begin VB.Timer Timer5 
      Interval        =   1000
      Left            =   375
      Top             =   7260
   End
   Begin VB.Timer Timer55 
      Interval        =   1000
      Left            =   0
      Top             =   7260
   End
   Begin VB.OptionButton optrnd 
      BackColor       =   &H00000000&
      Height          =   195
      Left            =   0
      TabIndex        =   4
      Top             =   6600
      Width           =   285
   End
   Begin VB.OptionButton optres 
      BackColor       =   &H80000008&
      Caption         =   "Norm"
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Left            =   345
      TabIndex        =   3
      Top             =   6615
      Width           =   285
   End
   Begin VB.OptionButton optnon 
      BackColor       =   &H00000000&
      Caption         =   "Option1"
      Height          =   195
      Left            =   720
      TabIndex        =   2
      Top             =   6600
      Width           =   195
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   3465
      Top             =   6060
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Select A Song"
   End
   Begin VB.Image viddown 
      Height          =   165
      Left            =   600
      Picture         =   "MPlay3.frx":4CA84
      Top             =   5805
      Width           =   360
   End
   Begin VB.Image vidup 
      Height          =   165
      Left            =   3075
      Picture         =   "MPlay3.frx":4CDDE
      Top             =   840
      Width           =   360
   End
   Begin VB.Image Image9 
      Height          =   195
      Left            =   3735
      ToolTipText     =   "On Top"
      Top             =   255
      Width           =   255
   End
   Begin VB.Image balance 
      Height          =   75
      Left            =   1350
      Picture         =   "MPlay3.frx":4D138
      ToolTipText     =   "Balance"
      Top             =   330
      Width           =   195
   End
   Begin VB.Image Image16 
      Height          =   165
      Left            =   2310
      Top             =   0
      Width           =   210
   End
   Begin VB.Image Image15 
      Height          =   180
      Left            =   3870
      Top             =   1755
      Width           =   240
   End
   Begin VB.Image Image14 
      Height          =   180
      Left            =   2145
      Top             =   5295
      Width           =   240
   End
   Begin VB.Label Khs 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   2340
      TabIndex        =   10
      Top             =   735
      Width           =   660
   End
   Begin VB.Label kbps 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   180
      Left            =   1380
      TabIndex        =   9
      Top             =   735
      Width           =   570
   End
   Begin VB.Image Image13 
      Height          =   180
      Left            =   2940
      Top             =   5295
      Width           =   240
   End
   Begin VB.Image Image12 
      Height          =   165
      Left            =   2745
      Top             =   5310
      Width           =   165
   End
   Begin VB.Image Image11 
      Height          =   165
      Left            =   2535
      Top             =   5310
      Width           =   195
   End
   Begin VB.Image Image10 
      Height          =   195
      Left            =   2355
      Top             =   5295
      Width           =   165
   End
   Begin VB.Image Image8 
      Height          =   180
      Left            =   2130
      Top             =   0
      Width           =   120
   End
   Begin VB.Image Image7 
      Height          =   180
      Left            =   1905
      Top             =   0
      Width           =   180
   End
   Begin VB.Image Image6 
      Height          =   165
      Left            =   1740
      Top             =   0
      Width           =   135
   End
   Begin VB.Image Image5 
      Height          =   165
      Left            =   1530
      Top             =   0
      Width           =   165
   End
   Begin VB.Label toptime 
      BackStyle       =   0  'Transparent
      Caption         =   "00.00"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   135
      Left            =   975
      TabIndex        =   8
      Top             =   15
      Width           =   435
   End
   Begin VB.Label Label20 
      BackStyle       =   0  'Transparent
      Caption         =   "00:00"
      BeginProperty Font 
         Name            =   "LcdD"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   180
      Left            =   885
      TabIndex        =   7
      Top             =   435
      Width           =   510
   End
   Begin VB.Image toppic 
      Height          =   225
      Left            =   1455
      Picture         =   "MPlay3.frx":4D242
      Top             =   0
      Width           =   1095
   End
   Begin VB.Image Image4 
      Height          =   195
      Left            =   3825
      ToolTipText     =   "Small Window"
      Top             =   15
      Width           =   150
   End
   Begin VB.Label VolumInd 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   225
      Left            =   960
      TabIndex        =   6
      Top             =   6630
      Width           =   1095
   End
   Begin VB.Image minus 
      Height          =   60
      Left            =   450
      Picture         =   "MPlay3.frx":4DF68
      ToolTipText     =   "Decrease Vloume"
      Top             =   750
      Width           =   75
   End
   Begin VB.Image pluss 
      Height          =   90
      Left            =   450
      Picture         =   "MPlay3.frx":4DFEA
      ToolTipText     =   "Increase Vloume"
      Top             =   540
      Width           =   90
   End
   Begin VB.Line LVol 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      Index           =   1
      X1              =   180
      X2              =   182
      Y1              =   915
      Y2              =   915
   End
   Begin VB.Line LVol 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      Index           =   5
      X1              =   180
      X2              =   405
      Y1              =   435
      Y2              =   435
   End
   Begin VB.Line LVol 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      Index           =   4
      X1              =   180
      X2              =   330
      Y1              =   555
      Y2              =   555
   End
   Begin VB.Line LVol 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      Index           =   3
      X1              =   180
      X2              =   255
      Y1              =   675
      Y2              =   675
   End
   Begin VB.Line LVol 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      Index           =   2
      X1              =   180
      X2              =   210
      Y1              =   795
      Y2              =   795
   End
   Begin VB.Line LVol 
      BorderColor     =   &H00FFFFFF&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   3
      Index           =   6
      X1              =   180
      X2              =   480
      Y1              =   315
      Y2              =   315
   End
   Begin VB.Image LOADLIST 
      Height          =   240
      Left            =   1095
      ToolTipText     =   "Load Playlist"
      Top             =   5400
      Width           =   330
   End
   Begin VB.Image savelist 
      Height          =   210
      Left            =   1065
      ToolTipText     =   "Save Playlist"
      Top             =   5085
      Width           =   360
   End
   Begin VB.Image list 
      Height          =   255
      Left            =   1515
      ToolTipText     =   "Playlists"
      Top             =   5265
      Width           =   300
   End
   Begin VB.Image listbar 
      Height          =   525
      Left            =   1095
      Picture         =   "MPlay3.frx":4E0A4
      Top             =   5100
      Width           =   345
   End
   Begin VB.Image directory 
      Height          =   195
      Left            =   720
      ToolTipText     =   "Add Directory"
      Top             =   5130
      Width           =   285
   End
   Begin VB.Image Image3 
      Height          =   240
      Left            =   720
      ToolTipText     =   "Add Song"
      Top             =   5415
      Width           =   315
   End
   Begin VB.Image filebar 
      Height          =   540
      Left            =   705
      Picture         =   "MPlay3.frx":4EABE
      Top             =   5100
      Width           =   315
   End
   Begin VB.Image addfile 
      Height          =   255
      Left            =   315
      ToolTipText     =   "Add Files"
      Top             =   5280
      Width           =   345
   End
   Begin VB.Image repeat 
      Height          =   330
      Left            =   3060
      ToolTipText     =   "Repeat Song"
      Top             =   1335
      Width           =   360
   End
   Begin VB.Image repeaton 
      Height          =   135
      Left            =   3345
      Picture         =   "MPlay3.frx":4F400
      Top             =   1305
      Width           =   240
   End
   Begin VB.Image shuffleon 
      Height          =   105
      Left            =   2775
      Picture         =   "MPlay3.frx":4F5F2
      Top             =   1335
      Width           =   225
   End
   Begin VB.Image shuffle 
      Height          =   345
      Left            =   2490
      ToolTipText     =   "Shuffle Songs"
      Top             =   1320
      Width           =   330
   End
   Begin VB.Image islider 
      Height          =   90
      Left            =   120
      Picture         =   "MPlay3.frx":4F784
      Top             =   1080
      Width           =   435
   End
   Begin VB.Image previous 
      Height          =   405
      Left            =   195
      ToolTipText     =   "Previous"
      Top             =   1320
      Width           =   285
   End
   Begin VB.Image next 
      Height          =   390
      Left            =   1590
      ToolTipText     =   "Next"
      Top             =   1320
      Width           =   420
   End
   Begin VB.Image stop 
      Height          =   375
      Left            =   1275
      ToolTipText     =   "Stop"
      Top             =   1320
      Width           =   315
   End
   Begin VB.Image play 
      Height          =   390
      Left            =   510
      ToolTipText     =   "Play"
      Top             =   1320
      Width           =   375
   End
   Begin VB.Label Labelpause 
      Caption         =   "Label1"
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   6930
      Width           =   540
   End
   Begin VB.Image pause 
      Height          =   405
      Left            =   915
      ToolTipText     =   "Pause"
      Top             =   1320
      Width           =   315
   End
   Begin VB.Image openfile 
      Height          =   375
      Left            =   2070
      ToolTipText     =   "Add Songs"
      Top             =   1290
      Width           =   285
   End
   Begin VB.Image plup 
      Height          =   165
      Left            =   3480
      Picture         =   "MPlay3.frx":4F9D6
      Top             =   810
      Width           =   195
   End
   Begin MediaPlayerCtl.MediaPlayer am1 
      Height          =   435
      Left            =   75
      TabIndex        =   0
      Top             =   6120
      Width           =   3375
      AudioStream     =   -1
      AutoSize        =   0   'False
      AutoStart       =   -1  'True
      AnimationAtStart=   -1  'True
      AllowScan       =   -1  'True
      AllowChangeDisplaySize=   -1  'True
      AutoRewind      =   0   'False
      Balance         =   0
      BaseURL         =   ""
      BufferingTime   =   5
      CaptioningID    =   ""
      ClickToPlay     =   -1  'True
      CursorType      =   0
      CurrentPosition =   -1
      CurrentMarker   =   0
      DefaultFrame    =   ""
      DisplayBackColor=   0
      DisplayForeColor=   16777215
      DisplayMode     =   0
      DisplaySize     =   4
      Enabled         =   -1  'True
      EnableContextMenu=   -1  'True
      EnablePositionControls=   -1  'True
      EnableFullScreenControls=   0   'False
      EnableTracker   =   -1  'True
      Filename        =   ""
      InvokeURLs      =   -1  'True
      Language        =   -1
      Mute            =   0   'False
      PlayCount       =   1
      PreviewMode     =   0   'False
      Rate            =   1
      SAMILang        =   ""
      SAMIStyle       =   ""
      SAMIFileName    =   ""
      SelectionStart  =   -1
      SelectionEnd    =   -1
      SendOpenStateChangeEvents=   -1  'True
      SendWarningEvents=   -1  'True
      SendErrorEvents =   -1  'True
      SendKeyboardEvents=   0   'False
      SendMouseClickEvents=   0   'False
      SendMouseMoveEvents=   0   'False
      SendPlayStateChangeEvents=   -1  'True
      ShowCaptioning  =   0   'False
      ShowControls    =   -1  'True
      ShowAudioControls=   -1  'True
      ShowDisplay     =   0   'False
      ShowGotoBar     =   0   'False
      ShowPositionControls=   -1  'True
      ShowStatusBar   =   0   'False
      ShowTracker     =   -1  'True
      TransparentAtStart=   0   'False
      VideoBorderWidth=   0
      VideoBorderColor=   0
      VideoBorder3D   =   0   'False
      Volume          =   -360
      WindowlessVideo =   0   'False
   End
   Begin VB.Image pldown 
      Height          =   195
      Left            =   195
      Picture         =   "MPlay3.frx":4FBD0
      Top             =   5850
      Width           =   210
   End
   Begin VB.Image Image2 
      Height          =   135
      Left            =   3645
      ToolTipText     =   "Minimize"
      Top             =   30
      Width           =   165
   End
   Begin VB.Image Image1 
      Height          =   165
      Left            =   3975
      ToolTipText     =   "Exit"
      Top             =   30
      Width           =   105
   End
End
Attribute VB_Name = "MPlay3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Intro As Boolean, STP As Boolean, Paused As Boolean
Public TwipX As Integer, TwipY As Integer
Public PrgName, Sect
Public DragFlag, SlideFlag, PlVisFlag
Public IX, IY, TX, TY, FX, FY
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Dim Volumet, tidModus, Sleep As Boolean

Private Const BIF_RETURNONLYFSDIRS = 1
Private Const BIF_DONTGOBELOWDOMAIN = 2
Private Const MAX_PATH = 260
Function GetWindowSize()
If Me.WindowState = 1 Then
Me.Caption = "MPlay3..." & Label24.Caption
End If
If Me.WindowState = 0 Then
Me.Caption = "MPlay3"
End If
If Me.WindowState = 2 Then
Me.Caption = "MPlay3"
End If
End Function
Function cdplayer()
If ext.Text = "cda" Then
am1.Filename = "CDAUDIO:"
am1.play
End If
End Function
Private Sub balance_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If SlideFlag = False Then
        IX = X: FX = balance.Left
        TX = Screen.TwipsPerPixelX
        SlideFlag = True
          End If
          
End Sub
Private Sub balance_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If SlideFlag = True Then
    On Error Resume Next
        pos = FX + (X - IX) / TX
        If pos < 900 Then pos = 900
        If pos > 1740 Then pos = 1740
        FX = pos: balance.Left = pos
        SlideFlag = True
      End If




If balance.Left > -1335 And balance.Left < 1335 Then
Label24.Caption = "Center"
End If
If balance.Left < 1200 Then
Label24.Caption = "Left"
End If
If balance.Left > 1620 Then
Label24.Caption = "Right"
am1.balance = balance.Left
End If
    
End Sub
Private Sub balance_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
   am1.balance = balance.Left - 2500
    SlideFlag = False
End Sub




Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
SaveSetting App.EXEName, "MPlay3", "MAINTOP", Me.Top
SaveSetting App.EXEName, "MPlay3", "MAINLEFT", Me.Left
End Sub

Private Sub HScroll2_Change()
On Error GoTo hell
If HScroll2.Value > -500 And HScroll2.Value < 500 Then
Label24.Caption = "Center"
End If
If HScroll2.Value < -500 Then
Label24.Caption = "Left"
End If
If HScroll2.Value > 500 Then
Label24.Caption = "Right"
End If
am1.balance = HScroll2.Value
hell:
Exit Sub
End Sub

Private Sub HScroll2_Scroll()
On Error GoTo hell
If HScroll2.Value > -500 And HScroll2.Value < 500 Then
Label24.Caption = "Center"
End If
If HScroll2.Value < -500 Then
Label24.Caption = "Left"
End If
If HScroll2.Value > 500 Then
Label24.Caption = "Right"
End If
am1.balance = HScroll2.Value
hell:
Exit Sub
End Sub

Private Sub Image10_Click()
If pause = True Then
pause = False
End If
am1.Filename = lstfavs.Text
If lstfavs.Text = "" Then
On Error Resume Next
Me.cd1.Filter = "All Supported Formats|*.mp3;*.wav;*.wma;*.asf;*.asx;*.lsf;*.lsx;*.mid;*.midi;*.rmi;*.aif;*.aifc;*.aiff;*.au;*.snd"
Me.cd1.ShowOpen
Me.am1.Filename = Me.cd1.Filename
Me.am1.play
Me.lstfavs.AddItem StripPath(Me.cd1.Filename)
Me.cd1.Filename = ""
If cd1.Filename = "" Then Exit Sub
If pause = True Then
pause = False
am1.Filename = lstfavs.Text
End If
End If
End Sub
Private Sub Image11_Click()
If lstfavs.ListCount = 0 Then Exit Sub
If Labelpause.Caption = "Pause" Then
am1.pause
Labelpause.Caption = "Resume"
Else
am1.play
Labelpause.Caption = "Pause"
End If
End Sub
Private Sub Image12_Click()
On Error GoTo hell
am1.stop
am1.CurrentPosition = 0
hell:
Exit Sub
End Sub
Private Sub Image13_Click()
next_Click
End Sub
Private Sub Image14_Click()
 If lstfavs.ListCount = 0 Then
        Exit Sub
    Else
        If lstfavs.ListIndex - 1 > -1 Then
            lstfavs.ListIndex = lstfavs.ListIndex - 1
            am1.Filename = lstfavs.Text
        Else
            lstfavs.ListIndex = lstfavs.ListCount - 1
            am1.Filename = lstfavs.Text
        End If
    End If
End Sub
Private Sub Image15_Click()
Me.Height = 1725
End Sub
Private Sub Image16_Click()
next_Click
End Sub
Private Sub Image4_Click()
Dim smallwindow As Boolean
Dim picture As Boolean
If Me.Height = 230 Then PL = True
If Me.Height = 1725 Then PL = False
' ************************************
If PL = True Then
    Me.Height = 1725
    Else
If PL = False Then
    Me.Height = 230
'*******************************************************
'the pic bit
If Me.Height = 1725 Then
    toppic.Visible = False
    Else
If Me.Height = 230 Then
      toppic.Visible = True
'********************************************************
End If
End If
End If
End If
End Sub
Private Sub Image5_Click()
 If lstfavs.ListCount = 0 Then
        Exit Sub
    Else
        If lstfavs.ListIndex - 1 > -1 Then
            lstfavs.ListIndex = lstfavs.ListIndex - 1
            am1.Filename = lstfavs.Text
        Else
            lstfavs.ListIndex = lstfavs.ListCount - 1
            am1.Filename = lstfavs.Text
        End If
    End If
End Sub
Private Sub Image6_Click()
If pause = True Then
pause = False
End If
am1.Filename = lstfavs.Text
If lstfavs.Text = "" Then
Label24.Caption = lstfavs.Text
On Error Resume Next
Me.cd1.Filter = "All Supported Formats|*.mp3;*.cda;*.wma;*.wav;*.asf;*.asx;*.lsf;*.lsx;*.mid;*.midi;*.rmi;*.aif;*.aifc;*.aiff;*.au;*.snd"
Me.cd1.ShowOpen
ext.Text = Right$(cd1.Filename, 3)
If ext.Text = "cda" Then
am1.Filename = "CDAUDIO:"
am1.play
Me.List1.AddItem (Me.cd1.Filename)
Me.lstfavs.AddItem StripPath(Me.cd1.Filename)
cd1.Filename = ""
Else
Me.am1.Filename = Me.cd1.Filename
Me.am1.play
Me.List1.AddItem (Me.cd1.Filename)
Me.lstfavs.AddItem StripPath(Me.cd1.Filename)
Label24.Caption = StripPath(Me.cd1.Filename)
Me.cd1.Filename = ""
If cd1.Filename = "" Then Exit Sub
If pause = True Then
pause = False
am1.Filename = List1.Text
End If
End If
End If
End Sub
Private Sub Image7_Click()
If lstfavs.ListCount = 0 Then Exit Sub
If Labelpause.Caption = "Pause" Then
am1.pause
Labelpause.Caption = "Resume"
Else
am1.play
Labelpause.Caption = "Pause"
End If
End Sub
Private Sub Image8_Click()
On Error GoTo hell
am1.stop
am1.CurrentPosition = 0
hell:
Exit Sub
End Sub


Private Sub List1_DblClick()
am1.Filename = List1.Text
End Sub

Private Sub addfile_Click()
Dim addfile As Boolean
If filebar.Visible = True Then addfile = True
If filebar.Visible = False Then addfile = False
' ************************************
If addfile = True Then
    filebar.Visible = False
    Else
If addfile = False Then
      filebar.Visible = True
End If
End If
End Sub
Private Sub directory_Click()
dir.Show 1
End Sub
Private Sub Form_Load()
On Error Resume Next
Image5.Enabled = False
Image6.Enabled = False
Image7.Enabled = False
Image8.Enabled = False
Image16.Enabled = False
toptime.Visible = False
toppic.Visible = False
filebar.Visible = False
listbar.Visible = False
shuffleon.Visible = False
repeaton.Visible = False
optrnd.Value = False
optres.Value = True
pldown.Visible = False
Me.Height = 1725
Me.Top = GetSetting(App.EXEName, "MPlay3", "MAINTOP")
Me.Left = GetSetting(App.EXEName, "MPlay3", "MAINLEFT")
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
    Exit Sub
        Else
    If Button = vbMiddleButton Then
    Exit Sub
        Else
    ReleaseCapture
    SendMessage Me.hWnd, &HA1, 2, 0&
    End If
    End If
End Sub
Private Sub Image1_Click()
Unload Me
End
End Sub
Private Sub Image2_Click()
Me.WindowState = 1
End Sub
Private Sub Image3_Click()
On Error Resume Next
Me.cd1.Filter = "All Supported Formats|*.mp3;*.cda;*.wma;*.wav;*.asf;*.asx;*.lsf;*.lsx;*.mid;*.midi;*.rmi;*.aif;*.aifc;*.aiff;*.au;*.snd"
Me.cd1.ShowOpen
ext.Text = Right$(cd1.Filename, 3)
If ext.Text = "cda" Then
am1.Filename = "CDAUDIO:"
am1.play
Me.List1.AddItem (Me.cd1.Filename)
Me.lstfavs.AddItem StripPath(Me.cd1.Filename)
cd1.Filename = ""
Else
Me.List1.AddItem (Me.cd1.Filename)
Me.lstfavs.AddItem StripPath(Me.cd1.Filename)
Label24.Caption = StripPath(Me.cd1.Filename)
Me.cd1.Filename = ""
If cd1.Filename = "" Then Exit Sub
If pause = True Then
pause = False
am1.Filename = List1.Text
End If
End If
End Sub

Private Sub iSlider_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If SlideFlag = False Then
        IX = X: FX = islider.Left
        TX = Screen.TwipsPerPixelX
        SlideFlag = True
    End If
End Sub
Private Sub iSlider_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If SlideFlag = True Then
    On Error Resume Next
        pos = FX + (X - IX) / TX
        If pos < 120 Then pos = 120
        If pos > 3540 Then pos = 3540
        FX = pos: islider.Left = pos
        SlideFlag = True
    End If
End Sub
Private Sub iSlider_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    P! = Int(((islider.Left - 120) / 3540) * am1.Duration)
    am1.stop
    am1.CurrentPosition = P!
    am1.play
    SlideFlag = False
End Sub

Private Sub list_Click()
' ************************************
If listbar.Visible = True Then
    listbar.Visible = False
    Else
If listbar.Visible = False Then
      listbar.Visible = True
End If
End If
End Sub
Private Sub LOADLIST_Click()
On Error Resume Next
cd1.DialogTitle = "Load Playlist"
cd1.Filter = "Playlist(*.M3u)|*.M3u"
cd1.ShowOpen
If cd1.Filename = "" Then Exit Sub
Call ReadList(lstfavs, cd1.Filename, True)
End Sub

Private Sub lstfavs_DblClick()
On Error Resume Next
Me.List1.ListIndex = Me.lstfavs.ListIndex
ext.Text = Right$(List1.Text, 3)
If ext.Text = "cda" Then
am1.Filename = "CDAUDIO:"
am1.play
Else
If pause = True Then
pause = False
End If
am1.Filename = List1.Text
Label24.Caption = lstfavs.Text
If lstfavs.Text = "" Then
On Error Resume Next
Me.cd1.Filter = "All Supported Formats|*.mp3;*.cda;*.wav;*.wma;*.asf;*.asx;*.lsf;*.lsx;*.mid;*.midi;*.rmi;*.aif;*.aifc;*.aiff;*.au;*.snd"
Me.cd1.ShowOpen
Me.am1.Filename = Me.cd1.Filename
Me.am1.play
Me.lstfavs.AddItem StripPath(Me.cd1.Filename)
Me.cd1.Filename = ""
If cd1.Filename = "" Then Exit Sub
If pause = True Then
pause = False
am1.Filename = lstfavs.Text
End If
End If
End If
End Sub
Private Sub lstfavs_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button And vbRightButton Then
On Error Resume Next
add.PopupMenu add.addpopup
If Me.lstfavs.Text = "" Then

End If
End If
End Sub

Private Sub next_Click()
On Error Resume Next
  If lstfavs.ListCount = 0 Then
        Exit Sub
    Else
        If lstfavs.ListIndex + 1 > 1 Then
            lstfavs.ListIndex = lstfavs.ListIndex + 1
            am1.Filename = lstfavs.Text
        Else
            lstfavs.ListIndex = lstfavs.ListCount + 1
            am1.Filename = lstfavs.Text
        End If
    End If
End Sub

Private Sub pause_Click()
On Error Resume Next
If lstfavs.ListCount = 0 Then Exit Sub
If Labelpause.Caption = "Pause" Then
am1.pause

Labelpause.Caption = "Resume"
Else
am1.play
Labelpause.Caption = "Pause"
End If
End Sub

Private Sub openfile_Click()
On Error Resume Next
Me.cd1.Filter = "All Supported Formats|*.mp3;*.cda;*.wma;*.wav;*.asf;*.asx;*.lsf;*.lsx;*.mid;*.midi;*.rmi;*.aif;*.aifc;*.aiff;*.au;*.snd"
Me.cd1.ShowOpen
ext.Text = Right$(cd1.Filename, 3)
If ext.Text = "cda" Then
am1.Filename = "CDAUDIO:"
am1.play
Me.List1.AddItem (Me.cd1.Filename)
Me.lstfavs.AddItem StripPath(Me.cd1.Filename)
cd1.Filename = ""
Else
Me.am1.Filename = Me.cd1.Filename
Me.am1.play
Me.List1.AddItem (Me.cd1.Filename)
Me.lstfavs.AddItem StripPath(Me.cd1.Filename)
Label24.Caption = StripPath(Me.cd1.Filename)
Me.cd1.Filename = ""
If cd1.Filename = "" Then Exit Sub
If pause = True Then
pause = False
am1.Filename = List1.Text
End If
End If
End Sub
Private Sub play_Click()
If pause = True Then
pause = False
End If
am1.Filename = lstfavs.Text
If lstfavs.Text = "" Then
Label24.Caption = lstfavs.Text
On Error Resume Next
Me.cd1.Filter = "All Supported Formats|*.mp3;*.cda;*.wma;*.wav;*.asf;*.asx;*.lsf;*.lsx;*.mid;*.midi;*.rmi;*.aif;*.aifc;*.aiff;*.au;*.snd"
Me.cd1.ShowOpen
ext.Text = Right$(cd1.Filename, 3)
If ext.Text = "cda" Then
am1.Filename = "CDAUDIO:"
am1.play
Me.List1.AddItem (Me.cd1.Filename)
Me.lstfavs.AddItem StripPath(Me.cd1.Filename)
cd1.Filename = ""
Else
Me.am1.Filename = Me.cd1.Filename
Me.am1.play
Me.List1.AddItem (Me.cd1.Filename)
Me.lstfavs.AddItem StripPath(Me.cd1.Filename)
Label24.Caption = StripPath(Me.cd1.Filename)
Me.cd1.Filename = ""
If cd1.Filename = "" Then Exit Sub
If pause = True Then
pause = False
am1.Filename = List1.Text
End If
End If
End If
End Sub

Private Sub plup_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
plup.Visible = False
pldown.Left = 3480
pldown.Top = 810
pldown.Visible = True
End Sub
Private Sub pldown_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim PL As Boolean
plup.Visible = True
pldown.Visible = False
If Me.Height = 5640 Then PL = True
If Me.Height = 1725 Then PL = False
' ************************************
If PL = True Then
    Me.Height = 1725
    Else
If PL = False Then
    Me.Height = 5640
End If
End If
End Sub
Function StripPath(T$) As String
    Dim X%, ct%
    StripPath$ = T$
    X% = InStr(T$, "\")


    Do While X%
        ct% = X%
        X% = InStr(ct% + 1, T$, "\")
    Loop
    If ct% > 0 Then StripPath$ = Mid$(T$, ct% + 1)
End Function


Private Sub previous_Click()
  If lstfavs.ListCount = 0 Then
        Exit Sub
    Else
        If lstfavs.ListIndex - 1 > -1 Then
            lstfavs.ListIndex = lstfavs.ListIndex - 1
            am1.Filename = lstfavs.Text
        Else
            lstfavs.ListIndex = lstfavs.ListCount - 1
            am1.Filename = lstfavs.Text
        End If
    End If
End Sub

Private Sub savelist_Click()
On Error Resume Next
cd1.DialogTitle = "Save Playlist"
cd1.Filter = "Playlist(*.M3u)|*.m3u"
cd1.ShowSave
If cd1.Filename = "" Then Exit Sub
Call WriteList(List1, cd1.Filename)
End Sub
Private Sub shuffle_Click()
If optrnd.Value = True Then
optrnd.Value = False
shuffleon.Visible = True
optres.Value = True
   shuffleon.Visible = False
    Else
If optrnd.Value = False Then
optrnd.Value = True
 shuffleon.Visible = False
optres.Value = False
shuffleon.Visible = True
End If
End If
End Sub


Private Sub stop_Click()
On Error GoTo hell
am1.stop
am1.CurrentPosition = 0
hell:
Exit Sub
End Sub
Private Sub am1_EndOfStream(ByVal result As Long)
On Error Resume Next
If optrnd.Value = True Then
Randomize Timer
 MyValue = Int((lstfavs.ListCount * Rnd))
   lstfavs.ListIndex = MyValue

   am1.Filename = lstfavs.Text
    If lstfavs.Text <> "" Then
        am1.play
        
        Exit Sub
        On Error Resume Next
    End If
Else
    If optres.Value = True Then
   lstfavs.ListIndex = lstfavs.ListIndex + 1
       

       lstfavs.Text = lstfavs.Text
       
        am1.Filename = lstfavs.Text
        Label24.Caption = am1.Filename
           am1.play
 End If
End If
If optnon.Value = True Then
    am1.stop
End If
End Sub

Private Sub Timer1_Timer()
GetWindowSize
End Sub


Private Sub Timer3_Timer()
        If Me.Height = 1725 Then
    toptime.Visible = False
    Else
If Me.Height = 230 Then
      toptime.Visible = True
      End If
End If
End Sub

Private Sub Timer4_Timer()
On Error Resume Next
tinseconden = am1.CurrentPosition
Dim min As Integer
Dim SEC As Integer
min = tinseconden \ 60
SEC = tinseconden - (min * 60)
If SEC = "-1" Then SEC = "0"

timseconden = am1.Duration
Dim minm As Integer
Dim secm As Integer
minm = timseconden \ 60
secm = timseconden - (minm * 60)
If secm = "-1" Then secm = "0"
toptime.Caption = Format$(min, "00") + ":" + Format$(SEC, "00")
End Sub

Private Sub Timer6_Timer()
            If Me.Height = 1725 Then
    Image5.Enabled = False
      Image6.Enabled = False
        Image7.Enabled = False
          Image8.Enabled = False
            Image8.Enabled = False
    Else
If Me.Height = 230 Then
       Image5.Enabled = True
      Image6.Enabled = True
        Image7.Enabled = True
          Image8.Enabled = True
            Image8.Enabled = True
End If
End If
End Sub

Private Sub Timer7_Timer()
Label24.Left = Label24.Left - 50
If Label24.Left < -Label24.Width Then
Label24.Left = PicDisplay.ScaleWidth
End If
End Sub
Private Sub toppicture_Timer()
If Me.Height = 1725 Then
    toppic.Visible = False
    Else
If Me.Height = 230 Then
      toppic.Visible = True
  End If
End If
End Sub

Private Sub TVolum_Timer()
On Error GoTo nix
Dim n As Integer
If Volumet = True Then
am1.Volume = am1.Volume + 130


End If
If Volumet = False Then
am1.Volume = am1.Volume - 130


End If
Select Case am1.Volume
Case Is > -1000
For n = 6 To 1 Step -1
LVol(n).Visible = True
Next
Case Else
Select Case am1.Volume
Case Is > -1500
LVol(6).Visible = False
For n = 5 To 1 Step -1
LVol(n).Visible = True
Next
Case Else
Select Case am1.Volume
Case Is > -2000
LVol(6).Visible = False
LVol(5).Visible = False
For n = 4 To 1 Step -1
LVol(n).Visible = True
Next
Case Else
Select Case am1.Volume
Case Is > -3000
LVol(6).Visible = False
LVol(5).Visible = False
LVol(4).Visible = False
For n = 3 To 1 Step -1
LVol(n).Visible = True
Next
Case Else
Select Case am1.Volume
Case Is > -4000
LVol(6).Visible = False
LVol(5).Visible = False
LVol(4).Visible = False
LVol(3).Visible = True
LVol(2).Visible = True
LVol(1).Visible = True
Case Else
Select Case am1.Volume
Case Is > -5000
For n = 6 To 3 Step -1
LVol(n).Visible = False
Next
LVol(2).Visible = True
LVol(1).Visible = True
Case Else
Select Case am1.Volume
Case Is > -6000
For n = 6 To 2 Step -1
LVol(n).Visible = False
Next
LVol(1).Visible = True
Case Else
Select Case am1.Volume
Case Is > -7000
For n = 6 To 1 Step -1
LVol(n).Visible = False
Next
End Select
End Select
End Select
End Select
End Select
End Select
End Select
End Select
Exit Sub
nix:
If Volumet = True Then
VolumInd.Caption = "Max"
End If
If Volumet = False Then
VolumInd.Caption = "Min"
End If
End Sub

Private Sub Timer5_Timer()
On Error Resume Next
tinseconden = am1.CurrentPosition
Dim min As Integer
Dim SEC As Integer
min = tinseconden \ 60
SEC = tinseconden - (min * 60)
If SEC = "-1" Then SEC = "0"

timseconden = am1.Duration
Dim minm As Integer
Dim secm As Integer
minm = timseconden \ 60
secm = timseconden - (minm * 60)
If secm = "-1" Then secm = "0"
Label20.Caption = Format$(min, "00") + ":" + Format$(SEC, "00")
End Sub

Private Sub Timer55_Timer()

On Error Resume Next
SongLen = am1.Duration
Elapsed = am1.CurrentPosition
    If SlideFlag = False Then
        islider.Left = 120 + Int((Elapsed / SongLen) * 3540)
 End If
End Sub
Private Sub Pluss_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
TVolum.Enabled = True
Volumet = True
End Sub

Private Sub Minus_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
TVolum.Enabled = False
VolumInd.Caption = ""
End Sub
Private Sub Minus_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
TVolum.Enabled = True
Volumet = False
End Sub
Private Sub Pluss_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
TVolum.Enabled = False
VolumInd.Caption = ""
End Sub
Private Sub viddown_Mousemove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.vidup.Visible = True
viddown.Visible = False
MPlay3.am1.stop
MPlay3.am1.CurrentPosition = 0
video.Show
End Sub
Private Sub vidup_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
vidup.Visible = False
viddown.Left = 3060
viddown.Top = 840
viddown.Visible = True
End Sub
