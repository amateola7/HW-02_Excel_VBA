Attribute VB_Name = "Module1"
Sub Homework2()

    Dim TotalStockVolume As Double
    Dim Ticker As String
    Dim Summary_Table_Row As Integer
    Dim Oprice As Double
    Dim Cprice As Double
    Dim Year_Change As Double
    Dim Percent_Change As Double
    
    Oprice = Cells(2, 3).Value
    
    
    
    
    
    Summary_Table_Row = 2
    
    
    Cells(1, 9).Value = "Ticker"
    Cells(1, 12).Value = "Total Stock Volume"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    n = Cells(Rows.Count, 1).End(xlUp).Row


    TotalStockVolume = 0
    
    
    
    For i = 2 To n
        
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            Ticker = Cells(i, 1).Value
            Cprice = Cells(i, 6).Value
            Year_Change = Cprice - Oprice
        
        If Oprice > 0 Then
            Percent_Change = Year_Change / Oprice
        Else: Percent_Change = (Cprice - 14.5) / 14.5
        End If
        
        If Oprice = 0 And Cells(i, 3) = 0 Then
            Percent_Change = 0
        End If
        
            Range("I" & Summary_Table_Row).Value = Ticker
            Range("J" & Summary_Table_Row).Value = Year_Change
            Range("K" & Summary_Table_Row).Value = Percent_Change
            Range("L" & Summary_Table_Row).Value = TotalStockVolume
            TotalStockVolume = 0
            Summary_Table_Row = Summary_Table_Row + 1
            

        
     
        Else
            TotalStockVolume = TotalStockVolume + Cells(i, 7).Value
        End If

    Next i
    
    For j = 2 To n
        If Abs(Cells(j, 10).Value) > 0 Then
            If Cells(j, 10).Value <= 0 Then
                Cells(j, 10).Interior.ColorIndex = 3
            Else
                Cells(j, 10).Interior.ColorIndex = 4
            End If
        
        End If
    Next j

End Sub


