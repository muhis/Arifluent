Private Function is3d(str As String) As Boolean
  x = 0
  For i = 1 To Len(str)
      xs = Mid(str, i, 1)
      If xs = " " Then
          x = x + 1
      End If
  Next i
  If x <> 2 Then is3d = False Else is3d = True
End Function

Private Function ISorigin(str As String) As Boolean
For i = 1 To Len(str)
xs = Mid(str, i, 1)
    If xs <> "0" And xs <> "," And xs <> " " Then
        ISorigin = False
        Exit For
    Else:
        ISorigin = True
    End If
Next i
End Function

Private Function detneg(str As String) As Boolean
For i = 1 To Len(str)
xs = Mid(str, i, 1)
    If xs = "-" Then
        detneg = True
        Exit For
    Else:
        detneg = False
    End If
  Next i
End Function

****File reading line by line code from stackoverflow, written by the user "Justin"
****http://stackoverflow.com/questions/1376756/superfast-way-to-read-large-files-line-by-line-in-vba-please-critique

Public Function readLine(ByRef strFilePath As String, ByRef nLine _
    As Integer) As String
    Dim NextLine As String
    Dim n As Integer
    FileNum = FreeFile
    Open strFilePath For Input As FileNum
    Do Until EOF(FileNum)
        Line Input #FileNum, NextLine
        n = n + 1
        'If n = nLine Then
        readLine = NextLine
    Loop
Close
End Function


Private Sub Command1_Click()
On Error GoTo errcatch
xfile = InputBox("Please Input file path", "ArifLuent")
Open xfile For Input As #1
Do While Not EOF(1)
    Line Input #1, ss
    ss = Replace(ss, ",", " ")
    Text1.Text = ss
    If is3d(Text1.Text) = False Then
        Text1.Text = Text1.Text & " 0"
    End If
    
    If ISorigin(Text1.Text) = False Then
        If detneg(Text1.Text) = False Then
                List2.AddItem Text1.Text
        Else
                List3.AddItem Text1.Text
        End If
    End If
Loop
Close #1

For i = 0 To List2.ListCount - 1
Text2.Text = Text2.Text & vbCrLf & "vertex create coordinates " & List2.List(i)
Next i

For i = 0 To List3.ListCount - 1
Text2.Text = Text2.Text & vbCrLf & "vertex create coordinates " & List3.List(i)
Next i
Exit Sub
errcatch:
MsgBox "error " & Err.Description

End Sub

Private Sub Command3_Click()
On Error GoTo errcatch
sfile = InputBox("Please input save path *.jou", "Saving")

 Open sfile For Output As #1
        Print #1, "/ Journal File for GAMBIT 2.4.6, Database 2.4.4, ntx86 SP2007051421"
        Print #1, "/ Made By Arif Luent v1.0.0 beta"
        Print #1, "Identifier name " & Chr(34) & "Arifluent" & Chr(34) & " new saveprevious"
        Print #1, Text2.Text
        
    Close #1
Exit Sub
errcatch:
MsgBox "error " & Err.Description

End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
Label1.Caption = KeyAscii

End Sub
