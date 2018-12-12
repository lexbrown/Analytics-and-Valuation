Sub PPopt()
    Dim PP As Integer
    Dim AccCF As Double
    PP = 0
    AccCF = 0
    While AccCF + Worksheets("Денежные потоки").Cells(18, 2) < 0
        AccCF = AccCF + Worksheets("Денежные потоки").Cells(18, PP + 3)
        PP = PP + 1
    Wend
    ActiveCell.Value = PP
End Sub

Sub PPprob()
    Dim PP As Integer
    Dim AccCF As Double
    PP = 0
    AccCF = 0
    While AccCF + Worksheets("Денежные потоки").Cells(19, 2) < 0
        AccCF = AccCF + Worksheets("Денежные потоки").Cells(19, PP + 3)
        PP = PP + 1
    Wend
    ActiveCell.Value = PP
End Sub

Sub PPstress()
    Dim PP As Integer
    Dim AccCF As Double
    PP = 0
    AccCF = 0
    While AccCF + Worksheets("Денежные потоки").Cells(20, 2) < 0
        AccCF = AccCF + Worksheets("Денежные потоки").Cells(20, PP + 3)
        PP = PP + 1
    Wend
    ActiveCell.Value = PP
End Sub

Sub DPPopt()
    Dim DPP As Integer
    Dim AccCF As Double
    DPP = 0
    AccCF = 0
    If Worksheets("Денежные потоки").Cells(24, 22).Value <= 0 Then
        ActiveCell.Value = "-"
    Else
        While AccCF + Worksheets("Денежные потоки").Cells(24, 2) < 0
            AccCF = AccCF + Worksheets("Денежные потоки").Cells(24, DPP + 3)
            DPP = DPP + 1
        Wend
        ActiveCell.Value = DPP
    End If
End Sub


Sub DPPprob()
    Dim DPP As Integer
    Dim AccCF As Double
    DPP = 0
    AccCF = 0
    If Worksheets("Денежные потоки").Cells(25, 22).Value <= 0 Then
        ActiveCell.Value = "-"
    Else
        While AccCF + Worksheets("Денежные потоки").Cells(25, 2) < 0
            AccCF = AccCF + Worksheets("Денежные потоки").Cells(25, DPP + 3)
            DPP = DPP + 1
        Wend
        ActiveCell.Value = DPP
    End If
End Sub


Sub DPPstress()
    Dim DPP As Integer
    Dim AccCF As Double
    DPP = 0
    AccCF = 0
    If Worksheets("Денежные потоки").Cells(26, 22).Value <= 0 Then
        ActiveCell.Value = "-"
    Else
        While AccCF + Worksheets("Денежные потоки").Cells(26, 2) < 0
            AccCF = AccCF + Worksheets("Денежные потоки").Cells(26, DPP + 3)
            DPP = DPP + 1
        Wend
        ActiveCell.Value = DPP
    End If
End Sub