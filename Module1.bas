
Sub basic_results()
    
    'Variable Declaration
    Dim Ticker As String
    Dim Opening_Price As Double
    Dim Closing_Price As Double
    Dim YearlyChange As Double
    Dim Percent_Change As Double
    Dim Volume As LongLong
    
    Dim ws As Worksheet
    Set ws = ActiveSheet
    Dim lastRow_data As Long
    Dim lastRow_results As Long
    
    Dim Counter As Long
    
    
' Loop through all sheets
    For Each ws In Worksheets
        
        ' Initialize headers
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "YearlyChange"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
        ' Obtain last row number of data
        lastRow_data = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

        ' Initialize last row number of results to 2
        lastRow_results = 2
        
        ' Initialize variables for comparisong and additions.
        Volume = 0
        Opening_Price = 0
        Closing_Price = 0
        
        ' Begin to iterate to all data rows.
        For i = 2 To lastRow_data
        
            ' Check if we are still within the same ticker, if it is not...
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

                ' Set the Ticker name
                Ticker = ws.Cells(i, 1).Value

                ' Add to the Volume Total
                Volume = Volume + ws.Cells(i, 7).Value

                ' Print the Ticker in the Summary Table
                ws.Range("I" & lastRow_results).Value = Ticker

                ' Print the Ticker Volume to the Summary Table
                ws.Range("L" & lastRow_results).Value = Volume

                ' If it is unique
                If Counter = 0 Then
                    Opening_Price = ws.Cells(i, 3)
                End If
                
                ' Obtain YearlyChange and Closing price, having the next ticker different.
                Closing_Price = ws.Cells(i, 6)
                YearlyChange = Closing_Price - Opening_Price
                ws.Range("J" & lastRow_results).Value = YearlyChange
                
                ' Set value to 0, when the Opening price is 0 to avoid division by 0.
                If Opening_Price <> 0 Then
                    ws.Range("K" & lastRow_results).Value = (1 - Closing_Price / Opening_Price)
                Else
                    ws.Range("K" & lastRow_results).Value = 0
                End If
                
                ' Format cells to red when below 0, and green above 0.
                If (YearlyChange >= 0) Then
                    ws.Range("J" & lastRow_results).Interior.ColorIndex = 4
                Else
                    ws.Range("J" & lastRow_results).Interior.ColorIndex = 3
                End If
                
                ' Calculates the Results last row
                lastRow_results = lastRow_results + 1
        
                ' Reset the Counter and Totals
                Brand_Total = 0
                Opening_Price = 0
                Closing_Price = 0
                Volume = 0
                Counter = 0

            ' If the cell immediately following a row is the same Ticker...
            Else
                ' Initialize Opening_Price
                ' If counter diffent than 0, the Opening Price has been already initialized
                If Counter <> 0 Then
                    Counter = Counter + 1
                Else
                    ' If Counter is 0, and the row's cell's value for Opening_Price is different than 0
                    If ws.Cells(i, 3) <> 0 Then
                        Opening_Price = ws.Cells(i, 3)
                        Counter = Counter + 1
                    Else
                        ' If row's cell's value for Opening_Price is 0, I cannot initialize the value.
                        Counter = 0
                    End If
                End If
                
                ' Add to the Volume Total
                Volume = Volume + ws.Cells(i, 7).Value

            End If
        
        
        Next i
        
        ws.Range("K" & 2, "K" & lastRow_results + 1).NumberFormat = "0.00%"
        
        ' *************************************************
        '          BONUS
        '**************************************************
        
        'Variable Bonus
        Dim Greatest_Inc As Double
        Dim Greatest_Dec As Double
        Dim Greatest_Vol As Double
        
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        ws.Range("Q2").Value = 0   ' Initialize Greatest % Increase's Cell
        ws.Range("Q3").Value = 0   ' Initialize Greatest % Decrease's Cell
        ws.Range("Q4").Value = 0   ' Initialize Greatest Total Volume's Cell
        
        
        For i = 2 To lastRow_results
            
            If (ws.Range("K" & i).Value >= ws.Range("Q2").Value) Then
                ws.Range("Q2").Value = ws.Range("K" & i).Value
                ws.Range("P2").Value = ws.Range("I" & i).Value
                
            End If
            
            If (ws.Range("K" & i).Value <= ws.Range("Q3").Value) Then
                ws.Range("Q3").Value = ws.Range("K" & i).Value
                ws.Range("P3").Value = ws.Range("I" & i).Value
            End If
            
            If (ws.Range("L" & i).Value > ws.Range("Q4").Value) Then
                ws.Range("Q4").Value = ws.Range("L" & i).Value
                ws.Range("P4").Value = ws.Range("I" & i).Value
            End If
            
            
        
        Next i
        
        ws.Range("Q2", "Q3").NumberFormat = "0.00%"

        ws.Columns("A:Q").AutoFit
    
    Next ws
    
End Sub


