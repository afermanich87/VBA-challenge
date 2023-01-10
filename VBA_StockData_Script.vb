Sub alpha_test()

Dim i As Long
Dim Open_Price As Double
Dim Close_Price As Double
Dim yearly_change As Double
Dim percent_change As Double
Dim total_stock_vol As Double
Dim ticker_alpha As Integer


ticker_alpha = 1
Ticker_num = 0
Open_Price = 0
Close_Price = 0
yearly_change = 0
percent_change = 0
total_stock_vol = 0
LastRow = Cells(Rows.Count, 1).End(xlUp).Row

Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Stock Volume"

For i = 2 To LastRow

    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        ticker_alpha = ticker_alpha + 1
        Ticker_num = Cells(i, 1).Value
        Cells(ticker_alpha, 9) = Cells(i, 1).Value
            Open_Price = Cells(i, 3).Value
            Close_Price = Cells(i, 6).Value
            yearly_change = Close_Price - Open_Price
            percent_change = ((yearly_change / Open_Price) * 100)
        Cells(ticker_alpha, 10).Value = yearly_change
        Cells(ticker_alpha, 11).Value = percent_change
        Cells(ticker_alpha, 12).Value = total_stock_vol
    End If
    

total_stock_vol = total_stock_vol + Cells(i, 7).Value

        
    If Cells(ticker_alpha, 10).Value >= 0 Then
        Cells(ticker_alpha, 10).Interior.ColorIndex = 4
        
        ElseIf Cells(ticker_alpha, 10).Value <= 0 Then
            Cells(ticker_alpha, 10).Interior.ColorIndex = 3

    End If

Next i

End Sub

