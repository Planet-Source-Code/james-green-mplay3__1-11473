VERSION 5.00
Begin VB.Form add 
   Caption         =   "menus"
   ClientHeight    =   90
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   90
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu addpopup 
      Caption         =   "popup"
      Visible         =   0   'False
      Begin VB.Menu delfile 
         Caption         =   "Remove File"
      End
      Begin VB.Menu information 
         Caption         =   "Information"
      End
   End
End
Attribute VB_Name = "add"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub delfile_Click()
Dim nEntryNum As Integer
nEntryNum = MPlay3.lstfavs.ListCount
Do While nEntryNum > 0
    nEntryNum = nEntryNum - 1
    If MPlay3.lstfavs.Selected(nEntryNum) = True Then
     MPlay3.lstfavs.RemoveItem nEntryNum
     MPlay3.List1.RemoveItem nEntryNum
          End If
  Loop
  End Sub

Private Sub information_Click()
If MPlay3.lstfavs.Text = "" Then
MsgBox ("No Media Loaded"), vbExclamation, "MPlay3"

Exit Sub
   Else
frmEditTags.GetTags (MPlay3.lstfavs.Text)

frmEditTags.StartEdit (MPlay3.lstfavs.Text)
Exit Sub
End If
End Sub
