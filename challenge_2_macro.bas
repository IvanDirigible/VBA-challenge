Attribute VB_Name = "Module1"
Sub challenge_part_1()
    ' Defining worksheet
    Dim ws As Worksheet
    
    ' Looping over worksheets
    For Each ws In Worksheets
    
        ' Counter varible
        Dim i As Long
        
        ' Gets ticker symbol
        Dim Ticker_Symbol As Integer
        Ticker_Symbol = 2
        
        ' Manually grabs the first open values
        Dim Stock_Open As Double
        Stock_Open = ws.Cells(2, 3).Value
        
        ' Close variable
        Dim Stock_Close As Double
        
        ' Quarterly change variable
        Dim Q_Change As Double
        
        ' Total stock variable
        Dim Stock_Total As Double
        Stock_Total = 0
        
        ' Finds end row
        Dim End_Row As Long
        End_Row = ws.Cells(Rows.Count, "A").End(xlUp).Row
        
        ' Put text into cells as requested
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Quarterly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        
        ' Adjust formatting for new columns
        ws.Range("J:L, O:O").Columns.AutoFit
        ws.Range("K:K, Q2:Q3").NumberFormat = "0.00%"
        
        ' Loops through the sheet
        For i = 2 To End_Row
        
            ' Checks for new ticker symbols
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                ' Adds to stock total
                Stock_Total = Stock_Total + ws.Cells(i, 7).Value
                
                ' Takes the close value of final value for ticker symbol
                Stock_Close = ws.Cells(i, 6).Value
                
                ' Calculates quarterly change
                Q_Change = Stock_Close - Stock_Open
                
                ' Prints quarterly change
                ws.Cells(Ticker_Symbol, 10).Value = Q_Change
                ' Prints percent change
                ws.Cells(Ticker_Symbol, 11).Value = Q_Change / Stock_Open
                
                ' Selects color of quarterly change based on value
                If Q_Change > 0 Then
                    ws.Cells(Ticker_Symbol, 10).Interior.ColorIndex = 4
                ElseIf Q_Change < 0 Then
                    ws.Cells(Ticker_Symbol, 10).Interior.ColorIndex = 3
                End If
                
                ' Selects first open value of next ticker symbol
                Stock_Open = ws.Cells(i + 1, 3).Value
                
                ' Prints ticker symbol
                ws.Cells(Ticker_Symbol, 9).Value = ws.Cells(i, 1).Value
                
                'Prints stock total
                ws.Cells(Ticker_Symbol, 12).Value = Stock_Total
                
                ' Moves to next cell in symbol list
                Ticker_Symbol = Ticker_Symbol + 1
                
                ' Resets stock total
                Stock_Total = 0
                
            Else
                ' Adds to stock total
                Stock_Total = Stock_Total + ws.Cells(i, 7).Value
            
            End If
        
        Next i
        
        ' Final stats counter
        Dim j As Long
        
        ' Final stats variables
        Dim Ticker_Inc As String
        Dim Ticker_Dec As String
        Dim Ticker_Tot As String
        
        ' Grabs starting value of final stats
        Dim Great_Inc As Double
        Great_Inc = ws.Cells(2, 11).Value
        Dim Great_Dec As Double
        Great_Dec = ws.Cells(2, 11).Value
        Dim Great_Tot As Double
        Great_Tot = ws.Cells(2, 12).Value
        
        
        ' Finds end of list
        Dim End_List As Long
        End_List = ws.Cells(Rows.Count, "I").End(xlUp).Row
        
        ' Loop searches for greatest increase
        For j = 2 To End_List
            ' Grabs value and ticker symbol for greatest increase
            If ws.Cells(j, 11).Value > Great_Inc Then
                Great_Inc = ws.Cells(j, 11).Value
                Ticker_Inc = ws.Cells(j, 9).Value
            End If
            
            ' Loop searches for greatest decrease
            If ws.Cells(j, 11).Value < Great_Dec Then
                ' Grabs value and ticker symbol for greatest decrease
                Great_Dec = ws.Cells(j, 11).Value
                Ticker_Dec = ws.Cells(j, 9).Value
            End If
            
            ' Loop searches for greatest total stock
            If ws.Cells(j, 12).Value > Great_Tot Then
                ' Grabs value and ticker symbol for greatest total stock
                Great_Tot = ws.Cells(j, 12).Value
                Ticker_Tot = ws.Cells(j, 9).Value
            End If
            
        Next j
        
        ' Prints gathered final stats into cells
        ws.Cells(2, 16) = Ticker_Inc
        ws.Cells(3, 16) = Ticker_Dec
        ws.Cells(4, 16) = Ticker_Tot
        ws.Cells(2, 17) = Great_Inc
        ws.Cells(3, 17) = Great_Dec
        ws.Cells(4, 17) = Great_Tot
    
    Next ws
    
End Sub
