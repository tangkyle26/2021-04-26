Attribute VB_Name = "Module1"
Option Explicit
Sub search()
Dim man As String
Dim rownum As Integer
Dim status As Boolean
Dim content As String
man = Range("G1").Value

For rownum = 2 To 7
    If (Cells(rownum, 1).Value = man) Then
        Range("G2").Value = Cells(rownum, 2).Value
        If (Cells(rownum, 3).Value = 0) Then
        status = False
        Else
        status = True
        End If
    content = man & "¥I´Úª¬ºA" & status
    MsgBox (content)
    Else
    End If
Next

End Sub
