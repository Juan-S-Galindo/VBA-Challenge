Sub WolfOfWallStreet():

    Dim LastRowStock As LongLong 'Define Variables
    Dim LastRowFiltered As Long
    Dim TickerRowCounter As LongLong
    Dim SumVolume As LongLong
    Dim StartOpenPrice As Double
    Dim EndClosingPrice As DOuble
    Dim DeltaPrice As Double
    Dim PercentChange As Double
   
    

    
    For Each ws In Worksheets
    
    TickerRowCounter = 2 'Counter to keep of the ticker list rows.
    SumVolume = 0 'Initial stock volume sum
    LastRowStock = ws.Cells(Rows.Count, 1).End(xlUp).Row 'checks for last row
    ws.Cells(1, 9).Value = "Ticker" 'all these add titles to the filtered data
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
'Challange
    ws.Cells(2, 14).Value = "Greatest % Increase" 'all these add titles to the filtered data for the challenge
    ws.Cells(3, 14).Value = "Greatest % Decrease"
    ws.Cells(4, 14).Value = "Greatest Total Volume"
    ws.Cells(1, 15).Value = "Ticker"
    ws.Cells(1, 16).Value = "Value"
    
    Range("N2:N4").HorizontalAlignment = xlRight 'indent right the cells in that range
    Range("A1:P1").HorizontalAlignment = xlCenter 'Centers titles
    
    StartOpenPrice = ws.Cells(2, 3).Value 'Opening price for the 1 stock in in the first run.
    
        For i = 2 To LastRowStock 'Iteration from position 2 to the end of the LastRowStok
        
        SumVolume = SumVolume + ws.Cells(i, 7).Value 'starts adding the stock volume
        
        
            If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then 'for every iteration checks if the currenct ticker is different from the next.
                'if the condtion is True
                EndClosingPrice = ws.Cells(i, 6).Value 'Records ending close price at row index i
                DeltaPrice = EndClosingPrice - StartOpenPrice 'difference in price.

                    if StartOpenPrice <> 0 Then 'error handler for div 0.
                         PercentChange = DeltaPrice / StartOpenPrice 
                    else
                        Dim TickerLabel as string
                        TickerLabel = ws.Cells(i, 1).Value
                        PercentChange = 0
                         Msgbox("Error handler for " + TickerLabel +". Open price is 0 results in Div by 0")
                    end if
                
                ws.Cells(TickerRowCounter, 9).Value = ws.Cells(i, 1).Value 'prints the ticker name to the new column for tickers following TickRowCounter index
                ws.Cells(TickerRowCounter, 10).Value = DeltaPrice 'prints delta price
                    
                    If DeltaPrice > 0 Then 'checks if change was negative or positive and assigns red or green to interior color.
                        ws.Cells(TickerRowCounter, 10).Interior.ColorIndex = 4 'Green
                    ElseIf DeltaPrice < 0 Then
                        ws.Cells(TickerRowCounter, 10).Interior.ColorIndex = 3 'Red
                    End If
                    
                ws.Cells(TickerRowCounter, 11).Value = PercentChange 'prints % change following TickerRowCounter index
                ws.Cells(TickerRowCounter, 12).Value = SumVolume 'prints sum of stock volume  following TickerRowCounter index
                    
                TickerRowCounter = TickerRowCounter + 1 ' adds 1 to the ticker counter in preparation for next iteration.
                SumVolume = 0 'resets the sum of volume.
                StartOpenPrice = ws.Cells(i + 1, 3).Value 'stores the opening price for the next ticker symbol.
                
            End If
            
            
        Next i 'next iteration.

    'After iteration is done, we can check for the challenge info.
        LastRowFiltered = ws.Cells(Rows.Count, 9).End(xlUp).Row 'Checks for the last row in the filtered ticker data

        'prettify the data
        ws.Range("J2" & ":" & "J" & LastRowFiltered).NumberFormat = "$#.##" 'format to delta price
        ws.Range("K2" & ":" & "K" & LastRowFiltered).NumberFormat = "#.##%" 'format in % and 2 decimals
        ws.Range("L2" & ":" & "L" & LastRowFiltered).NumberFormat = "#,###.##" ' format using commas and 2 decimals

        'Greatest increase
        ws.Cells(2, 16).Value = WorksheetFunction.Max(ws.Range("k2" & ":" & "K" & LastRowFiltered)) 'finds max percent increase
        ws.Cells(2, 15).Value = WorksheetFunction.Index(ws.Range("I2" & ":" & "I" & LastRowFiltered), WorksheetFunction.Match(ws.Cells(2, 16).Value, ws.Range("K2" & ":" & "K" & LastRowFiltered),0)) 'matches max % increase. note 0 for exact match.
        ws.cells(2,16).NumberFormat = "#.##%" 'format

        'Greatest Decrease 
        ws.Cells(3, 16).Value = WorksheetFunction.Min(ws.Range("k2" & ":" & "K" & LastRowFiltered)) 
        ws.Cells(3, 15).Value = WorksheetFunction.Index(ws.Range("I2" & ":" & "I" & LastRowFiltered), WorksheetFunction.Match(ws.Cells(3, 16).Value, ws.Range("K2" & ":" & "K" & LastRowFiltered),0))
        ws.cells(3,16).NumberFormat = "#.##%" 'format 

        'Greatest Total Volume 
        ws.Cells(4, 16).Value = WorksheetFunction.Max(ws.Range("L2" & ":" & "L" & LastRowFiltered)) 
        ws.Cells(4, 15).Value = WorksheetFunction.Index(ws.Range("I2" & ":" & "I" & LastRowFiltered), WorksheetFunction.Match(ws.Cells(4, 16).Value, ws.Range("L2" & ":" & "L" & LastRowFiltered),0))
        ws.cells(4,16).NumberFormat = "#,###.##" 'format
    
    Next ws 'next worksheet.
End Sub