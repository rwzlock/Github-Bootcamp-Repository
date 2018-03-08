Sub Stocks():

    'Variable to hold ticker name
    Dim Ticker_Name As String

    'Variable to hold total ticker volume
    Dim Ticker_Total As Double
    Ticker_Total = 0
    'Row for the ticker and its volume to be printed
    Dim Summary_Table_Row As Integer

    For Each ws In Worksheets

        
        'Identifies the last cell in the given worksheet
        LastRow = Cells(Rows.Count, 1).End(xlUp).Row
        'MsgBox (LastRow)

        'Set names for summary table columns
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Total Volume"

        'Set row to begin in all Worksheets
        Summary_Table_Row = 2
        Ticker_Total = 0
        'Scan rows in current worksheet
        For i = 2 To LastRow

            'Check that we are within the same ticker
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

                'Set the Ticker name
                Ticker_Name = ws.Cells(i, 1).Value

                ' Add to total ticker volume
                Ticker_Total = Ticker_Total + ws.Cells(i, 7).Value

                ' Print the ticker name in the summary table
                ws.Range("I" & Summary_Table_Row).Value = Ticker_Name

                ' Print the ticker total volume in the summary table
                ws.Range("J" & Summary_Table_Row).Value = Ticker_Total

                'Create a different row for the next ticker
                Summary_Table_Row = Summary_Table_Row + 1
      
                'Reset total Ticker volume
                Ticker_Total = 0

            ' If next cell is still the same ticker
            Else

                ' Add to ticker total volume
                Ticker_Total = Ticker_Total + ws.Cells(i, 7).Value

            End If
        Next i

    Next ws

End Sub
