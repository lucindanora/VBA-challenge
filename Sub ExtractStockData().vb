Sub ExtractStockData()
    
    
    Dim ws As Worksheet
    Dim LastRow As Long
    Dim Ticker As String
    Dim StartPrice As Double
    Dim EndPrice As Double
    Dim Volume As LongLong
    Dim PercentChange As Double
    Dim i As Long
    Dim j As Long
    Dim grperinticker As String
    Dim grperinchge As Double
    Dim grperdeticker As String
    Dim grperdetchg As Double
    Dim grtolvolticker As String
    Dim grtotvols As Double
    
    For Each ws In Worksheets
       
        ' Find the last row with data in column A
        LastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        
        ' Initialize variables
        Ticker = ws.Cells(2, 1).Value
        StartPrice = ws.Cells(2, 3).Value
        EndPrice = ws.Cells(2, 6).Value
        Volume = ws.Cells(2, 7).Value
        j = 2
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
        ws.Cells(1, 15).Value = "Ticker"
        ws.Cells(1, 16).Value = "Value"
        ' Loop through rows to extract data
        For i = 2 To LastRow
            
            If Ticker <> ws.Cells(i, 1).Value Then
                ' Calculate percent change
                If StartPrice <> 0 Then
                    PercentChange = (EndPrice - StartPrice) / StartPrice
                Else
                    PercentChange = 0
                End If
                If PercentChange > grperinchge And PercentChange > 0 Then
                   grperinticker = Ticker
                   grperinchge = PercentChange
                Else
                 If PercentChange < 0 Then
                   If PercentChange < grperdetchg Then
                     grperdeticker = Ticker
                     grperdetchg = PercentChange
                   End If
                  End If
                End If
                If Volume > grtotvols Then
                  grtolvolticker = Ticker
                  grtotvols = Volume
                End If
                ' Print data
                ws.Cells(j, 9).Value = Ticker
                ws.Cells(j, 10).Value = EndPrice - StartPrice
                ws.Cells(j, 11).Value = Format(PercentChange, "0.00%")
                ws.Cells(j, 12).Value = Volume
                If EndPrice - StartPrice < 0 Then
                  ws.Cells(j, 10).Interior.Color = RGB(255, 0, 0)
                Else: ws.Cells(j, 10).Interior.Color = RGB(0, 255, 0)
                 End If
                j = j + 1
                ' Reset variables for the new ticker
                Ticker = ws.Cells(i, 1).Value
                StartPrice = ws.Cells(i, 3).Value
                EndPrice = ws.Cells(i, 6).Value
                Volume = ws.Cells(i, 7).Value
            
            Else
                ' Accumulate volume
                Volume = Volume + ws.Cells(i, 7).Value
                EndPrice = ws.Cells(i, 6).Value
                End If
              
        Next i
        
        ' Clean up the last row
        ws.Cells(j, 9).Value = Ticker
        ws.Cells(j, 10).Value = EndPrice - StartPrice
        If EndPrice - StartPrice < 0 Then
           ws.Cells(j, 10).Interior.Color = RGB(255, 0, 0)
        Else: ws.Cells(j, 10).Interior.Color = RGB(0, 255, 0)
        End If
           
        ' Calculate percent change
        If StartPrice <> 0 Then
            PercentChange = (EndPrice - StartPrice) / StartPrice
        Else
            PercentChange = 0
        End If
        ws.Cells(j, 11).Value = Format(PercentChange, "0.00%")
        ws.Cells(j, 12).Value = Volume
        
        ws.Cells(4, 14).Value = "Greatest Total Volume"
        ws.Cells(4, 15).Value = grtolvolticker
        ws.Cells(4, 16).Value = grtotvols
        
        ws.Cells(2, 14).Value = "Greatest % Increase"
        ws.Cells(2, 15).Value = grperinticker
        ws.Cells(2, 16).Value = Format(grperinchge, "0.00%")
        
        ws.Cells(3, 14).Value = "Greatest % Decrease"
        ws.Cells(3, 15).Value = grperdeticker
        ws.Cells(3, 16).Value = Format(grperdetchg, "0.00%")
        'clean up columns
        ws.Columns("A:P").AutoFit
        
    Next
    
End Sub