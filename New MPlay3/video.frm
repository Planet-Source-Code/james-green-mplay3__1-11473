VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{22D6F304-B0F6-11D0-94AB-0080C74C7E95}#1.0#0"; "MSDXM.OCX"
Begin VB.Form video 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3810
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5715
   LinkTopic       =   "Form1"
   Picture         =   "video.frx":0000
   ScaleHeight     =   3810
   ScaleWidth      =   5715
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   4905
      Top             =   3570
   End
   Begin VB.PictureBox PicDisplay 
      BackColor       =   &H00704700&
      BorderStyle     =   0  'None
      Height          =   135
      Left            =   4005
      ScaleHeight     =   135
      ScaleWidth      =   1410
      TabIndex        =   1
      Top             =   3315
      Width           =   1410
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         Height          =   135
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   525
      End
   End
   Begin VB.Timer Timer55 
      Interval        =   1000
      Left            =   4275
      Top             =   3555
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3765
      Top             =   3525
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Image Image5 
      Height          =   180
      Left            =   5490
      Top             =   0
      Width           =   240
   End
   Begin VB.Image Image4 
      Height          =   225
      Left            =   3045
      Top             =   3555
      Width           =   480
   End
   Begin VB.Image Image3 
      Height          =   240
      Left            =   2445
      Top             =   3540
      Width           =   510
   End
   Begin VB.Image Image2 
      Height          =   210
      Left            =   1965
      Top             =   3570
      Width           =   360
   End
   Begin VB.Image islider 
      Height          =   90
      Left            =   240
      Picture         =   "video.frx":4734A
      Top             =   3345
      Width           =   435
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   225
      Top             =   3555
      Width           =   615
   End
   Begin MediaPlayerCtl.MediaPlayer MediaPlayer1 
      Height          =   2970
      Left            =   225
      TabIndex        =   0
      ToolTipText     =   "MPlay3"
      Top             =   315
      Visible         =   0   'False
      Width           =   5205
      AudioStream     =   -1
      AutoSize        =   0   'False
      AutoStart       =   -1  'True
      AnimationAtStart=   -1  'True
      AllowScan       =   -1  'True
      AllowChangeDisplaySize=   -1  'True
      AutoRewind      =   0   'False
      Balance         =   30
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
      Volume          =   -780
      WindowlessVideo =   0   'False
   End
End
Attribute VB_Name = "video"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public TwipX As Integer, TwipY As Integer
Public PrgName, Sect
Public DragFlag, SlideFlag, PlVisFlag
Public IX, IY, TX, TY, FX, FY
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Sub Form_Load()
On Error Resume Next
Me.Top = GetSetting(App.EXEName, "MPlay3Video", "MAINTOP")
Me.Left = GetSetting(App.EXEName, "MPlay3Video", "MAINLEFT")
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

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
SaveSetting App.EXEName, "MPlay3Video", "MAINTOP", Me.Top
SaveSetting App.EXEName, "MPlay3Video", "MAINLEFT", Me.Left
End Sub
Private Sub Image1_Click()
CommonDialog1.Filter = "All Supported Files|*.avi;*.mpeg;*.mpg;*.mpe;*.mpv;*.m1v;*.mp2;*.mpv2;*.mp2v;*.mpa;*.ivf;*.mov;*.qt;*.dat"
Me.CommonDialog1.ShowOpen
Me.MediaPlayer1.Filename = Me.CommonDialog1.Filename
Label1.Caption = StripPath(Me.CommonDialog1.Filename)
Me.CommonDialog1.Filename = ""
If CommonDialog1.Filename = "" Then Exit Sub
If pause = True Then
pause = False
End If
End Sub

Private Sub Image2_Click()
Me.MediaPlayer1.play
End Sub

Private Sub Image3_Click()
Me.MediaPlayer1.pause
End Sub

Private Sub Image4_Click()
Me.MediaPlayer1.stop
MediaPlayer1.CurrentPosition = 0
End Sub

Private Sub Image5_Click()
Unload Me
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
        If pos < 240 Then pos = 240
        If pos > 3450 Then pos = 3450
        FX = pos: islider.Left = pos
        SlideFlag = True
    End If
End Sub
Private Sub iSlider_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    P! = Int(((islider.Left - 240) / 3450) * MediaPlayer1.Duration)
    Me.MediaPlayer1.stop
    MediaPlayer1.CurrentPosition = P!
    MediaPlayer1.play
    SlideFlag = False
End Sub

Private Sub Timer1_Timer()
Label1.Left = Label1.Left - 50
If Label1.Left < -Label1.Width Then
Label1.Left = PicDisplay.ScaleWidth
End If
End Sub

Private Sub Timer55_Timer()
On Error Resume Next
SongLen = Me.MediaPlayer1.Duration
Elapsed = MediaPlayer1.CurrentPosition
    If SlideFlag = False Then
        islider.Left = 240 + Int((Elapsed / SongLen) * 3450)
 End If
End Sub
