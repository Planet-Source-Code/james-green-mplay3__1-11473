VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMP3 
   AutoRedraw      =   -1  'True
   Caption         =   "MP3 Searcher"
   ClientHeight    =   0
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   1560
   LinkTopic       =   "Form1"
   ScaleHeight     =   0
   ScaleWidth      =   1560
   WindowState     =   2  'Maximized
   Begin MSComDlg.CommonDialog cdgMP3 
      Left            =   5400
      Top             =   60
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer tmrMP3 
      Left            =   5895
      Top             =   1500
   End
   Begin MSComctlLib.ImageList imgLstMp32 
      Left            =   5835
      Top             =   165
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMP3.frx":0000
            Key             =   "Note"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgLstMp3 
      Left            =   5865
      Top             =   885
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMP3.frx":005E
            Key             =   "Note"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar stbMp3 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   -255
      Width           =   1560
      _ExtentX        =   2752
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   14111
            MinWidth        =   14111
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picSplit 
      BorderStyle     =   0  'None
      Height          =   5295
      Left            =   3080
      MousePointer    =   9  'Size W E
      ScaleHeight     =   5295
      ScaleWidth      =   45
      TabIndex        =   1
      Top             =   20
      Width           =   50
   End
   Begin VB.ListBox List1 
      Height          =   255
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   2055
      Visible         =   0   'False
      Width           =   1935
   End
   Begin MSComctlLib.ImageList img 
      Left            =   5955
      Top             =   1965
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   128
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMP3.frx":00BC
            Key             =   "unknown"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMP3.frx":011A
            Key             =   "fixed"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMP3.frx":0178
            Key             =   "ram"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMP3.frx":01D6
            Key             =   "remove"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMP3.frx":0234
            Key             =   "cd"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMP3.frx":0292
            Key             =   "folder"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMP3.frx":02F0
            Key             =   "open"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMP3.frx":034E
            Key             =   "remote"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picMP3 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000009&
      Height          =   2550
      Left            =   3120
      ScaleHeight     =   2490
      ScaleWidth      =   2595
      TabIndex        =   3
      Top             =   35
      Width           =   2655
      Begin MSComctlLib.ListView lstMP3 
         Height          =   5325
         Left            =   0
         TabIndex        =   4
         Top             =   0
         Visible         =   0   'False
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   9393
         View            =   3
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         OLEDragMode     =   1
         OLEDropMode     =   1
         FullRowSelect   =   -1  'True
         _Version        =   393217
         Icons           =   "imgLstMp32"
         SmallIcons      =   "imgLstMp3"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   0
         OLEDragMode     =   1
         OLEDropMode     =   1
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "MP3 Name"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   1
            Text            =   "Size"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Modified"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Full Path"
            Object.Width           =   8819
         EndProperty
      End
   End
   Begin MSComctlLib.TreeView DirTree 
      Height          =   2550
      Left            =   60
      TabIndex        =   5
      Top             =   30
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   4498
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   529
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      ImageList       =   "img"
      Appearance      =   1
      OLEDropMode     =   1
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "PopupMenu"
      Visible         =   0   'False
      Begin VB.Menu mnuDelete 
         Caption         =   "Delete"
      End
      Begin VB.Menu mnuRename 
         Caption         =   "Rename"
      End
      Begin VB.Menu mnuSelectAll 
         Caption         =   "Select All"
      End
      Begin VB.Menu mnuMakePlayList 
         Caption         =   "Make Play List"
      End
      Begin VB.Menu mnuSplit 
         Caption         =   "-"
      End
      Begin VB.Menu mnuView 
         Caption         =   "View"
         Begin VB.Menu mnuViewIcon 
            Caption         =   "Large Icons"
            Index           =   0
         End
         Begin VB.Menu mnuViewIcon 
            Caption         =   "Small Icons"
            Index           =   1
         End
         Begin VB.Menu mnuViewIcon 
            Caption         =   "List"
            Index           =   2
         End
         Begin VB.Menu mnuViewIcon 
            Caption         =   "Details"
            Checked         =   -1  'True
            Index           =   3
         End
      End
   End
   Begin VB.Menu mnuPopupMenu2 
      Caption         =   "PopupMenu2"
      Visible         =   0   'False
      Begin VB.Menu mnuRenameFolder 
         Caption         =   "Rename"
      End
      Begin VB.Menu mnuAddFolder 
         Caption         =   "New"
      End
      Begin VB.Menu mnuDeleteFolder 
         Caption         =   "Delete"
      End
      Begin VB.Menu mnuRefresh 
         Caption         =   "Refresh"
      End
   End
End
Attribute VB_Name = "frmMP3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum Action
    Copy = 1
    Move = 2
    Delete = 3
    Rename = 4
End Enum

Private DropHighlightIndex As Integer
Private EditError As Boolean    'For editing listview labels
Private FileAction As Integer
Private indrag As Boolean       ' Flag that signals a Drag Drop operation.
Private nodX As Object          ' Item that is being dragged.
Private LastSelIndex As Integer
Private Const vbGrey = &H8000000C
Private Const vbLiteGrey = &H8000000F
Private MouseEvent As Boolean
Private CurrentDrive As String
Private CancelClick As Boolean
Private nNode As Node
Private Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long

Private Sub DirTree_AfterLabelEdit(Cancel As Integer, NewString As String)
    
    Dim Path As String
    Dim Response As Long
    Dim NewFolderName As String
    Dim OldPath As New Collection
    Dim Node As MSComctlLib.Node
    
    NewString = Trim(NewString)
    
    If NewString = "" Then
        Cancel = True
    ElseIf InStr(1, NewString, "\", vbTextCompare) <> 0 Then
        Cancel = True
    ElseIf InStr(1, NewString, "/", vbTextCompare) <> 0 Then
        Cancel = True
    ElseIf InStr(1, NewString, ":", vbTextCompare) <> 0 Then
        Cancel = True
    ElseIf InStr(1, NewString, "*", vbTextCompare) <> 0 Then
        Cancel = True
    ElseIf InStr(1, NewString, "<", vbTextCompare) <> 0 Then
        Cancel = True
    ElseIf InStr(1, NewString, ">", vbTextCompare) <> 0 Then
        Cancel = True
    ElseIf InStr(1, NewString, "|", vbTextCompare) <> 0 Then
        Cancel = True
    ElseIf InStr(1, NewString, "?", vbTextCompare) <> 0 Then
        Cancel = True
    ElseIf InStr(1, NewString, """", vbTextCompare) <> 0 Then
        Cancel = True
    ElseIf Len(NewString) > 255 Then
        Cancel = True
    End If
    
    If Cancel Then
        MsgBox "Folder names cannot be zero length, be larger than 255 characters or contain '\/:*<>?|'.", vbCritical
    Else
        
        Set Node = DirTree.Nodes(DirTree.SelectedItem.Index)
        Path = Mid(Node.FullPath, InStr(1, Node.FullPath, ":") - 1, 2) & Mid(Node.FullPath, InStr(1, Node.FullPath, ":") + 2)
        
        NewFolderName = NewString
        
       

        OldPath.add Path

        

        If Response <> 0 Then
            Cancel = True
        End If
        
    End If
    
End Sub


Private Sub DirTree_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo errHandler
    
    Dim m As New clsMousePointer
    m.SetCursor
    
    If KeyCode = vbKeyF5 Then
    
        DirTree.Nodes.Clear
        lstMP3.ListItems.Clear
        LoadTreeView
    
    End If
    
    Exit Sub
errHandler:
    MsgBox Err.Description, vbCritical
    Unload Me
    
End Sub

Private Sub DirTree_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    If Button = 2 Then
        If DirTree.Nodes.Count <> 0 Then
            PopupMenu mnuPopupMenu2
        End If
    End If

End Sub

Private Sub Form_Load()
     On Error GoTo errHandler
    
    Dim m As New clsMousePointer
    m.SetCursor
    
    'Print the text in the picture box
    picMP3.Print vbCrLf & "  Click the Drive\Folder you wish " & vbCrLf & "   to search for MP3's."
    
    LoadTreeView

    Exit Sub

errHandler:
    MsgBox Err.Description, vbCritical

    Unload Me

End Sub



Sub DisplayDir(Pth, Parent)
    On Error Resume Next

    Dim j As Integer
    Dim tmp As String
    On Error Resume Next
    Pth = Pth & "\"
    tmp = dir(Pth, vbDirectory + vbSystem)
    Do Until tmp = ""
        If tmp <> "." And tmp <> ".." Then
            If GetAttr(Pth & tmp) And vbDirectory Then
                'I use ListBox with property Sorted=True to
                'alphabetize directories. Easy eh? ;-)
                List1.AddItem StrConv(tmp, vbProperCase)
                'StrConv function convert for example
                '"WINDOWS" to "Windows"
            End If
        End If
       
    Loop
    
    'Add sorted directory names to TreeView
    For j = 1 To List1.ListCount
        Set nNode = DirTree.Nodes.add(Parent, tvwChild, , List1.list(j - 1), "folder")
        nNode.ExpandedImage = "open"
    Next j
    List1.Clear

End Sub

Private Sub LoadTreeView()
    On Error Resume Next

    Dim DriveNum As String
    Dim DriveType As Long
    DriveNum = 64
    On Error Resume Next
    Do
        DriveNum = DriveNum + 1
        DriveType = GetDriveType(Chr$(DriveNum) & ":\")
        If DriveNum > 90 Then Exit Do
        Select Case DriveType
            Case 0: Set nNode = DirTree.Nodes.add(, , "root" & DriveNum, StrConv(dir(Chr$(DriveNum) & ":", vbVolume), vbProperCase) & " (" & Chr$(DriveNum) & ":)", "unknown")
                    DisplayDir Mid(DirTree.Nodes("root" & DriveNum).Text, Len(DirTree.Nodes("root" & DriveNum).Text) - 2, 2), "root" & DriveNum
            Case 2: Set nNode = DirTree.Nodes.add(, , "root" & DriveNum, "(" & Chr$(DriveNum) & ":)", "remove")
            Case 3: Set nNode = DirTree.Nodes.add(, , "root" & DriveNum, StrConv(dir(Chr$(DriveNum) & ":", vbVolume), vbProperCase) & " (" & Chr$(DriveNum) & ":)", "fixed")
                    DisplayDir Mid(DirTree.Nodes("root" & DriveNum).Text, Len(DirTree.Nodes("root" & DriveNum).Text) - 2, 2), "root" & DriveNum
            Case 4: Set nNode = DirTree.Nodes.add(, , "root" & DriveNum, StrConv(dir(Chr$(DriveNum) & ":", vbVolume), vbProperCase) & " (" & Chr$(DriveNum) & ":)", "remote")
                    DisplayDir Mid(DirTree.Nodes("root" & DriveNum).Text, Len(DirTree.Nodes("root" & DriveNum).Text) - 2, 2), "root" & DriveNum
            Case 5: Set nNode = DirTree.Nodes.add(, , "root" & DriveNum, StrConv(dir(Chr$(DriveNum) & ":", vbVolume), vbProperCase) & " (" & Chr$(DriveNum) & ":)", "cd")
                    DisplayDir Mid(DirTree.Nodes("root" & DriveNum).Text, Len(DirTree.Nodes("root" & DriveNum).Text) - 2, 2), "root" & DriveNum
            Case 6: Set nNode = DirTree.Nodes.add(, , "root" & DriveNum, StrConv(dir(Chr$(DriveNum) & ":", vbVolume), vbProperCase) & " (" & Chr$(DriveNum) & ":)", "ram")
                    DisplayDir Mid(DirTree.Nodes("root" & DriveNum).Text, Len(DirTree.Nodes("root" & DriveNum).Text) - 2, 2), "root" & DriveNum
        End Select
    Loop
End Sub

Private Sub Form_Resize()
    On Error GoTo errHandler

    'Resize the conrols
    picSplit.Left = 3060
    DirTree.Width = picSplit.Left - picSplit.Width
    picMP3.Left = picSplit.Left + picSplit.Width + 20
    picMP3.Width = Me.Width - DirTree.Width + picSplit.Width - 300
    
    'Adjust back to the middle
    picSplit.Left = picSplit.Left + 20
    
    'Set the heights
    DirTree.Height = Me.Height - 900
    picMP3.Height = Me.Height - 900
    picSplit.Height = Me.Height - 880

    Exit Sub
errHandler:
    
End Sub



Private Sub lstMP3_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    On Error Resume Next

    lstMP3.SortKey = ColumnHeader.Index - 1
    
    If lstMP3.SortOrder = lvwAscending Then
        lstMP3.SortOrder = lvwDescending
    Else
        lstMP3.SortOrder = lvwAscending
    End If
    lstMP3.Sorted = True

End Sub

Private Sub lstMP3_ItemClick(ByVal Item As MSComctlLib.ListItem)
    On Error GoTo errHandler
    
    Dim SumOfLengths As Double
    Dim l As MSComctlLib.ListItem
    Dim SumOfSelection As Long
    
    For Each l In lstMP3.ListItems
        If l.Selected Then
            SumOfLengths = SumOfLengths + Left(l.ListSubItems(1).Text, Len(l.ListSubItems(1).Text) - 2)
            SumOfSelection = SumOfSelection + 1
        End If
    Next l
    
    stbMp3.Panels(1).Text = SumOfSelection & " MP3(s) Selected " & SumOfLengths / 1000000 & " MB"
    
    Exit Sub
    
errHandler:
    MsgBox Err.Description
    stbMp3.Panels(1).Text = ""

End Sub



Private Sub lstMP3_OLECompleteDrag(Effect As Long)
    On Error Resume Next
    Me.Caption = "MP3 - File Manager"
    
End Sub

Private Sub mnuAddFolder_Click()

    Call AddDir

End Sub


Private Sub mnuEditTag_Click()

    'Show the edit tag form
    frmEditTags.StartEdit (lstMP3.SelectedItem.ListSubItems(3).Text)

End Sub

Private Sub mnuMakePlayList_Click()
    
    Dim FilePath As String
    Dim l As MSComctlLib.ListItem
    Dim s As String
    
    cdgMP3.Flags = cdlOFNHideReadOnly Or cdlOFNOverwritePrompt Or cdlOFNPathMustExist
    cdgMP3.Filter = "MP3 Playlist (*.m3u)|*.m3u"

    cdgMP3.ShowSave

    FilePath = cdgMP3.Filename
    
    If FilePath = "" Then
        'User pressed cancel
        Exit Sub
    Else
    
        For Each l In lstMP3.ListItems
            If l.Selected Then
                s = s & l.ListSubItems(3).Text & vbCrLf
            End If
        Next l
        'Strip of the last vcCrLf
        s = Left(s, Len(s) - 2)
    
        'Everthing is ok creat the list
        Open FilePath For Output As #1
            Print #1, s
        Close #1

    End If

    'Clear the file name
    cdgMP3.Filename = ""
    
End Sub

Private Sub mnuRefresh_Click()
    On Error GoTo errHandler
    
    Dim m As New clsMousePointer
    m.SetCursor
    
    DirTree.Nodes.Clear
    lstMP3.ListItems.Clear
    LoadTreeView
    DropHighlightIndex = 0
    CurrentDrive = ""
    
    Exit Sub
errHandler:
    MsgBox Err.Description, vbCritical
    Unload Me
End Sub

Private Sub mnuRename_Click()
    On Error Resume Next
    
    lstMP3.StartLabelEdit

End Sub

Private Sub mnuRenameFolder_Click()

    'Edit the label
    DirTree.StartLabelEdit

End Sub

Private Sub mnuSelectAll_Click()
    On Error GoTo errHandler
    
    Dim l As Integer
    
    For l = 1 To lstMP3.ListItems.Count
        lstMP3.ListItems(l).Selected = True
    Next l

    Exit Sub
errHandler:
    MsgBox Err.Description, vbCritical

End Sub

Private Sub mnuViewIcon_Click(Index As Integer)
    On Error Resume Next
    
    Dim i As Integer
    
    For i = 0 To mnuViewIcon.UBound
        If i = Index Then
            mnuViewIcon(i).Checked = True
        Else
            mnuViewIcon(i).Checked = False
        End If
    Next i

    lstMP3.View = Index
    lstMP3.Arrange = lvwAutoLeft
    
End Sub

Private Sub picMP3_Resize()
    lstMP3.Move 0, 0, picMP3.ScaleWidth, picMP3.ScaleHeight
End Sub

Private Sub picSplit_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    
    picSplit.BackColor = vbGrey
    MouseEvent = True

End Sub

Private Sub picSplit_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    
    'Resize the controls
    If MouseEvent Then
    
        If x + picSplit.Left < 250 Then
           Exit Sub
        ElseIf x + picSplit.Left > Me.Width - 500 Then
            Exit Sub
        End If
        
        picSplit.Left = x + picSplit.Left
    End If
    
End Sub

Private Sub picSplit_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    
    picSplit.BackColor = vbLiteGrey
    If MouseEvent Then
        DirTree.Width = picSplit.Left - picSplit.Width
        picMP3.Left = picSplit.Left + picSplit.Width + 20
        picMP3.Width = Me.Width - DirTree.Width + picSplit.Width - 300
        
        'Adjust back to the middle
        picSplit.Left = picSplit.Left + 20
    End If
    MouseEvent = False
    
End Sub


Public Function GetSong(FullPath As String) As String
    On Error Resume Next

    Dim s As String
    Dim Delimiter As Integer
    Dim i As Integer
    
    'Strip off drive letter
    s = FullPath

    For i = Len(s) To 0 Step -1
        If Mid(s, i, 1) = "\" Then
            Delimiter = Len(s) - i
            Exit For
        End If
    Next i

    s = Right(s, Delimiter)
    
    GetSong = s

End Function


Private Sub lstMp3_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   
    If Button = 2 Then
        If lstMP3.ListItems.Count <> 0 Then
            PopupMenu mnuPopup
        End If
    ElseIf Button = 1 Then
        Set nodX = lstMP3.SelectedItem ' Set the item being dragged.
    End If
End Sub

Private Sub lstMp3_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    If EditError Then
        EditError = False
        lstMP3.MultiSelect = True
        lstMP3.FullRowSelect = True
        Exit Sub
    End If

    indrag = True ' Set the flag to true.

End Sub

Private Sub lstMp3_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

    indrag = False ' Set the flag to false.
    DropHighlightIndex = 0
    
End Sub



Private Sub DirTree_OLEDragOver(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)
   On Error Resume Next
   
   If indrag = True And x <> 0 Then

        If Shift = 1 Or Shift = 0 Then
            'Copy
            Effect = 1
            Me.Caption = "Copying..."
            FileAction = 1
        ElseIf Shift = 2 Then
            'move
            Effect = 2
            Me.Caption = "Moving..."
            FileAction = 2
        Else
            'Copy
            Effect = 1
            Me.Caption = "Copying..."
            FileAction = 1
        End If
        
        ' Set DropHighlight to the mouse's coordinates.
        Set DirTree.DropHighlight = DirTree.HitTest(x, y)
        DirTree.DropHighlight.EnsureVisible

        
        If DropHighlightIndex <> 0 Then
            If DropHighlightIndex <> DirTree.DropHighlight.Index Then
                tmrMP3.Interval = 400
            End If
        End If
        
        DropHighlightIndex = DirTree.DropHighlight.Index

        LastSelIndex = DirTree.DropHighlight.Index
   ElseIf x = 0 Then
        DirTree.Nodes(LastSelIndex - 1).EnsureVisible
   End If
   
End Sub



Public Sub AddDir()
    On Error GoTo errHandler
    
    Dim Node As MSComctlLib.Node
    Dim Path As String
    Dim FolderName As String
    Dim i As Integer

    Set Node = DirTree.Nodes(DirTree.SelectedItem.Index)
    
    'Get the path of the selected node
    Path = Mid(Node.FullPath, InStr(1, Node.FullPath, ":") - 1, 2) & Mid(Node.FullPath, InStr(1, Node.FullPath, ":") + 2)
    
    FolderName = "New Folder"

    Do Until dir(Path & "\" & FolderName, vbDirectory) = ""
        i = i + 1
        FolderName = "New Folder (" & i & ")"
    Loop
    
    MkDir (Path & "\" & FolderName)

    'Remove all the nodes children'
    If Node.Children <> 0 Then
        Do Until Node.Children = 0
            For i = Node.Child.LastSibling.Index To Node.Child.FirstSibling.Index Step -1
                DirTree.Nodes.Remove (i)
            Next i
        Loop
    End If
    
    DisplayDir Path, Node.Index

    Exit Sub
    
errHandler:
    MsgBox Err.Description, vbCritical

End Sub

Public Function StripFileName(FullPath As String) As String
    On Error Resume Next

    Dim s As String
    Dim Delimiter As Integer
    Dim i As Integer
    
    'Strip off File Name
    s = FullPath

    For i = Len(s) To 0 Step -1
        If Mid(s, i, 1) = "\" Then
            Delimiter = i
            Exit For
        End If
    Next i

    s = Left(s, Delimiter)
    
    StripFileName = s

End Function



Private Sub tmrMP3_Timer()

    DirTree.Nodes(DropHighlightIndex).Expanded = True
    tmrMP3.Interval = 0
    
End Sub
