Attribute VB_Name = "Module1"
Option Explicit

'code for colorEnums-rgb.xlsx
Private Sub PutColorsInSheet()
    Dim r1 As Range, i As Integer
    For i = 2 To 143
        Cells(i, "B").Activate
        Set r1 = Range(ActiveCell.Offset(0, 3), ActiveCell.Offset(5, 5))
        r1.Interior.Color = ActiveCell.Value
        ActiveCell.Offset(1, 0).Activate
        'Application.Wait DateAdd("s", 2, Now)
        'If i Mod 10 = 0 Then Stop
    Next i
End Sub

