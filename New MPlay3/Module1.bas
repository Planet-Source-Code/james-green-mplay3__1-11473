Attribute VB_Name = "Module1"

Public Sub ReadList(list As ListBox, Filename As String, Optional ClearList As Boolean)
    On Error GoTo Err
    Open Filename For Input As #1
    If ClearList = True Then list.Clear


    Do While Not EOF(1)
        Input #1, lstinput
        MPlay3.List1.AddItem lstinput
       MPlay3.lstfavs.AddItem lstinput
        Loop
    Close #1
    Exit Sub
Err:
    MsgBox "Error in ReadList" & Chr(13) & Chr(13) & Err.Number _
    & " - " & Err.Description, vbCritical, "Error"
    Exit Sub
End Sub


Public Sub WriteList(list As ListBox, Filename As String)


    If list.ListCount <= 0 Then
        MsgBox "Listbox is empty - cannot write to file!", vbCritical, "Error"
        End
    End If
    On Error GoTo Err
    Open Filename For Output As #1


    For i = 0 To list.ListCount - 1
        Print #1, list.list(i)
    Next
    Close #1
    Exit Sub
Err:
    MsgBox "Error in WriteList" & Chr(13) & Chr(13) & Err.Number _
    & " - " & Err.Description, vbCritical, "Error"
    Exit Sub
    MPlay3.cd1.Filename = ""
End Sub

Public Function StripPath(T$) As String
    Dim x%, ct%
    StripPath$ = T$
    x% = InStr(T$, "\")


    Do While x%
        ct% = x%
        x% = InStr(ct% + 1, T$, "\")
    Loop
    If ct% > 0 Then StripPath$ = Mid$(T$, ct% + 1)
End Function
