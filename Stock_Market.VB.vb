Sub StockMarket_Analysis()

Dim Ticker As String

Dim year_open As Double

Dim year_close As Double

Dim Yearly_Change As Double

Dim Total_Stock_Volume As Double

Dim Percent_Change As Double

Dim raw_data As Integer

Dim ws As Worksheet

For Each ws In Worksheets

    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"

    raw_data = 2
    initial_i = 1
    Total_Stock_Volume = 0

    EndRow = ws.Cells(Rows.Count, "A").End(xlUp).Row

        For i = 2 To EndRow

            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

            Ticker = ws.Cells(i, 1).Value
            
            initial_i = initial_i + 1

            year_open = ws.Cells(initial_i, 3).Value
            
            year_close = ws.Cells(i, 6).Value

            For j = initial_i To i

                Total_Stock_Volume = Total_Stock_Volume + ws.Cells(j, 7).Value

            Next j

            If year_open = 0 Then

                Percent_Change = year_close

            Else
                Yearly_Change = year_close - year_open

                Percent_Change = Yearly_Change / year_open

            End If
            
            ws.Cells(raw_data, 9).Value = Ticker
            
            ws.Cells(raw_data, 10).Value = Yearly_Change
            
            ws.Cells(raw_data, 11).Value = Percent_Change

            ws.Cells(raw_data, 11).NumberFormat = "0.00%"
            
            ws.Cells(raw_data, 12).Value = Total_Stock_Volume

            raw_data = raw_data + 1

            Total_Stock_Volume = 0
            Yearly_Change = 0
            Percent_Change = 0

            initial_i = i

        End If

    Next i

    lastkrow = ws.Cells(Rows.Count, "K").End(xlUp).Row

    Increase = 0
    Decrease = 0
    Greatest = 0

        For k = 3 To lastkrow

            last_k = k - 1

            k_in_use = ws.Cells(k, 11).Value

            k_initial = ws.Cells(last_k, 11).Value

            volume = ws.Cells(k, 12).Value

            initial_volume = ws.Cells(last_k, 12).Value

            If Increase > k_in_use And Increase > k_initial Then

                Increase = Increase

            ElseIf k_in_use > Increase And k_in_use > k_initial Then

                Increase = k_in_use

                increase_value = ws.Cells(k, 9).Value

            ElseIf k_initial > Increase And k_initial > k_in_use Then

                Increase = k_initial

                increase_value = ws.Cells(last_k, 9).Value

            End If

            If Decrease < k_in_use And Decrease < k_initial Then

                Decrease = Decrease


            ElseIf k_in_use < Increase And k_in_use < k_initial Then

                Decrease = k_in_use


                decrease_name = ws.Cells(k, 9).Value

            ElseIf k_initial < Increase And k_initial < k_in_use Then

                Decrease = k_initial

                decrease_name = ws.Cells(last_k, 9).Value

            End If

            If Greatest > volume And Greatest > initial_volume Then

                Greatest = Greatest

            ElseIf volume > Greatest And volume > initial_volume Then

                Greatest = volume

                greatest_value = ws.Cells(k, 9).Value

            ElseIf initial_volume > Greatest And initial_volume > volume Then

                Greatest = initial_volume

                greatest_value = ws.Cells(last_k, 9).Value

            End If

        Next k

    ws.Range("N1").Value = "Column Name"
    ws.Range("N2").Value = "Greatest % Increase"
    ws.Range("N3").Value = "Greatest % Decrease"
    ws.Range("N4").Value = "Greatest Total Volume"
    ws.Range("O1").Value = "Ticker Name"
    ws.Range("P1").Value = "Value"

    ws.Range("O2").Value = increase_value
    ws.Range("O3").Value = decrease_name
    ws.Range("O4").Value = greatest_value
    ws.Range("P2").Value = Increase
    ws.Range("P3").Value = Decrease
    ws.Range("P4").Value = Greatest



    ws.Range("P2").NumberFormat = "0.00%"
    ws.Range("P3").NumberFormat = "0.00%"


    lastjrow = ws.Cells(Rows.Count, "J").End(xlUp).Row


        For j = 2 To lastjrow

            If ws.Cells(j, 10) > 0 Then

                ws.Cells(j, 10).Interior.ColorIndex = 4

            Else

                ws.Cells(j, 10).Interior.ColorIndex = 3
            End If

        Next j


Next ws

End Sub