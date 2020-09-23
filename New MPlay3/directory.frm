VERSION 5.00
Begin VB.Form dir 
   BackColor       =   &H00800000&
   BorderStyle     =   0  'None
   Caption         =   "Form3"
   ClientHeight    =   3135
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4095
   LinkTopic       =   "Form3"
   Picture         =   "directory.frx":0000
   ScaleHeight     =   3135
   ScaleWidth      =   4095
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.FileListBox File1 
      BackColor       =   &H00800000&
      ForeColor       =   &H80000005&
      Height          =   870
      Left            =   135
      Pattern         =   "*.mp3;*.wav"
      TabIndex        =   2
      Top             =   3360
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.DirListBox Dir1 
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      ForeColor       =   &H80000005&
      Height          =   1890
      Left            =   180
      TabIndex        =   1
      Top             =   705
      Width           =   3615
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   1275
      TabIndex        =   0
      Top             =   330
      Width           =   1335
   End
   Begin VB.Image Image4 
      Height          =   135
      Left            =   3915
      Top             =   30
      Width           =   195
   End
   Begin VB.Image Image3 
      Height          =   195
      Left            =   2535
      Top             =   2805
      Width           =   585
   End
   Begin VB.Image Image2 
      Height          =   225
      Left            =   645
      Top             =   2775
      Width           =   630
   End
   Begin VB.Image Image1 
      Height          =   135
      Left            =   4470
      Top             =   4410
      Width           =   210
   End
End
Attribute VB_Name = "dir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Function StripPath(T$) As String
    Dim x%, ct%
    StripPath$ = T$
    x% = InStr(T$, "\")


    Do While x%
        ct% = x%
        x% = InStr(ct% + 1, T$, "\")
    Loop
    If ct% > 0 Then StripPath$ = Mid$(T$, ct% + 1)
End Function
Private Sub Drive1_Change()
On Error Resume Next
Dir1.Path = Drive1.Drive
End Sub



Private Sub Image1_Click()
Unload Me
End Sub

Private Sub Image2_Click()
File1.Path = Dir1.Path
If File1.ListCount <> 0 Then
    For tel = 1 To File1.ListCount
        File1.ListIndex = tel - 1
        
        
        
        If Len(Dir1.Path) > 3 Then
             'MPlay3.lstfavs.AddItem Dir1.Path & "\" & File1.FileName
                  
             MPlay3.List1.AddItem Dir1.Path & "\" & File1.FileName
             MPlay3.lstfavs.AddItem StripPath(Dir1.Path & "\" & File1.FileName)
        Else
           'Exit For
            'MsgBox "You can't add a drive, only folders", vbOKOnly, "Error"
           'Exit Sub
           
        '***********MPlay3.lstfavs.AddItem Dir1.Path & File1.Filename
       
        End If
    Next tel
            Unload Me
Else
    MsgBox "No files were found in specific folder", vbOKOnly, "Error"
    Unload Me
End If
End Sub

Private Sub Image3_Click()
Unload Me
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Label1.ForeColor = &HFFFFFF
End Sub
Private Sub Label1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Label1.ForeColor = &HC00000
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Label2.ForeColor = &HFFFFFF
End Sub
Private Sub Label2_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Label2.ForeColor = &HC00000
End Sub

Private Sub Image4_Click()
Unload Me
End Sub
