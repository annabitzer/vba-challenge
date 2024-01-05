Attribute VB_Name = "Module1"
Sub Stock_Market_Analysis()

'for next time: to find greatest % increase and % decrease, set all values in column
'into an array & then find max or min?

For Each ws In Worksheets

    'make column headers
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    
    'make summary table labels
    ws.Range("O1").Value = "Ticker"
    ws.Range("P1").Value = "Value"
    ws.Range("N2").Value = "Greatest % Increase"
    ws.Range("N3").Value = "Greatest % Decrease"
    ws.Range("N4").Value = "Greatest Total Volume"
    
    Dim Ticker As String
    Dim Yearly_Change As Double
    Dim Percent_Change As Double
    Dim Stock_Vol As LongLong
    Dim LastRow As Long
    
    Dim OpenPrice As Double
    Dim ClosePrice As Double
    
    Dim Table_Row_Counter As Integer
    Table_Row_Counter = 2
    
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    OpenPrice = ws.Cells(2, 3).Value
    Stock_Vol = 0
    
    'loop through all the information
    For i = 2 To LastRow
        'check if ticker name has changed
        If ws.Cells(i + 1, 1) <> ws.Cells(i, 1) Then
            'Prepare everything that needs to be put in the table
            Ticker = ws.Cells(i, 1).Value
            ClosePrice = ws.Cells(i, 6).Value
            Yearly_Change = ClosePrice - OpenPrice
            Percent_Change = Yearly_Change / OpenPrice
            Stock_Vol = Stock_Vol + ws.Cells(i, 7).Value
            
            'Print Values in Table
            ws.Range("I" & Table_Row_Counter).Value = Ticker
            ws.Range("J" & Table_Row_Counter).Value = Yearly_Change
            ws.Range("K" & Table_Row_Counter).Value = Percent_Change
            ws.Range("L" & Table_Row_Counter).Value = Stock_Vol
            
            'Conditional +/- Formatting
            Pos_Neg_Check = ws.Range("J" & Table_Row_Counter).Value
                If Sgn(Pos_Neg_Check) = 1 Then
                    ws.Range("J" & Table_Row_Counter).Interior.ColorIndex = 4
                ElseIf Sgn(Pos_Neg_Check) = -1 Then
                    ws.Range("J" & Table_Row_Counter).Interior.ColorIndex = 3
                End If
            'Prepare to move on to next ticker
            'Reset open price
            OpenPrice = ws.Cells(i + 1, 3).Value
            'Reset Stock volume
            Stock_Vol = 0
            'Move one forward in the Table Row
            Table_Row_Counter = Table_Row_Counter + 1
            
        Else: Stock_Vol = Stock_Vol + ws.Cells(i, 7).Value
            
        End If
    Next i
    
    'Find max % change
    Dim MaxPercentInc As Double
    Dim MaxPercentIncTicker As String
    Dim MaxPercentDec As Double
    Dim MaxPercentDecTicker As String
    Dim MaxVolume As LongLong
    Dim MaxVolumeTicker As String
    
    'fill variables with first row that will be checked
    MaxPercentInc = ws.Range("K2").Value
    MaxPercentDec = ws.Range("K2").Value
    MaxVolume = ws.Range("L2").Value
    
    For x = 2 To LastRow
        'check each row for max increase
        If ws.Cells(x + 1, 11) > MaxPercentInc Then
            MaxPercentInc = ws.Cells(x + 1, 11).Value
            MaxPercentIncTicker = ws.Cells(x + 1, 9).Value
            
        End If
        
        'check each row for max decrease
        If ws.Cells(x + 1, 11) < MaxPercentDec Then
            MaxPercentDec = ws.Cells(x + 1, 11).Value
            MaxPercentDecTicker = ws.Cells(x + 1, 9).Value
            
        End If
        
        'check each row for max volume
        If ws.Cells(x + 1, 12) > MaxVolume Then
            MaxVolume = ws.Cells(x + 1, 12).Value
            MaxVolumeTicker = ws.Cells(x + 1, 9).Value
            
        End If
        
    Next x
        
'print results
ws.Range("O2").Value = MaxPercentIncTicker
ws.Range("P2").Value = MaxPercentInc
ws.Range("O3").Value = MaxPercentDecTicker
ws.Range("P3").Value = MaxPercentDec
ws.Range("O4").Value = MaxVolumeTicker
ws.Range("P4").Value = MaxVolume
    
'Other Formatting
ws.Range("K2:K" & LastRow).NumberFormat = "0.00%"
ws.Range("P2:P3").NumberFormat = "0.00%"
ws.Columns("I:P").AutoFit

Next ws

End Sub

'alternate way to find maximum increase, decrease, volume- but couldn't figure out how to get ticker value for each
'sources:https://stackoverflow.com/questions/42633273/finding-max-of-a-column-in-vba
'max_increase = WorksheetFunction.Max(Range("K:K"))
'max_decrease = WorksheetFunction.Min(Range("K:K"))
'max_volume = WorksheetFunction.Max(Range("L:L"))

