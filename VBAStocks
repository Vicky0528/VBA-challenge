Sub WorksheetLoop()
    
Dim ws As Worksheet
    
'Loop each worksheet
For Each ws In Worksheets
    
    'Declare all the variables
    Dim lastrow As Long
    Dim x, y, I, MaxLine, MinLine, MaxRow As Integer
    Dim closeNumber, openNumber, MaxNumber, MinNumber, MaxVolume As Double
    Dim total As Double
            
    'Assign an orignal value to all the variables
    x = 2
    y = 9
    MaxNumber = 0
    MinNumber = 0
    MaxVolume = 0
    openNumber = ws.Cells(2, 3).Value
    lastrow = ws.Range("A999999").End(xlUp).Row
    total = ws.Cells(2, 7).Value
            
    'Design the excel format
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Value"
    ws.Cells(1, 17).Value = "Value"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    ws.Columns("I").ColumnWidth = 10
    ws.Columns("J:L").ColumnWidth = 20
    ws.Columns("O").ColumnWidth = 20
    ws.Columns("Q").ColumnWidth = 20
    ws.Columns("p").ColumnWidth = 10

    ws.Range("I1:Q1").HorizontalAlignment = xlCenter
    
    'Calculate the Ticker, yearly change, percentage change, and total stack value
    For I = 2 To lastrow
        If ws.Cells(I, 1).Value = ws.Cells(I + 1, 1).Value Then
            total = total + ws.Cells(I + 1, 7).Value
        End If
        If ws.Cells(I, 1).Value <> ws.Cells(I + 1, 1).Value Then
            closeNumber = ws.Cells(I, 6).Value
            ws.Cells(x, y).Value = ws.Cells(I, 1).Value
            ws.Cells(x, y + 1).Value = closeNumber - openNumber
            ws.Cells(x, y + 1).NumberFormat = "0.000000000"
                If ws.Cells(x, y + 1).Value >= 0 Then
                    ws.Cells(x, y + 1).Interior.ColorIndex = 4
                Else:
                    ws.Cells(x, y + 1).Interior.ColorIndex = 3
                End If
                If openNumber <> 0 Then
                    ws.Cells(x, y + 2).Value = (closeNumber - openNumber) / openNumber
                Else:
                    ws.Cells(x, y + 2).Value = 0
                End If
            ws.Cells(x, y + 2).NumberFormat = "0.00%"
            ws.Cells(x, y + 3).Value = total
            openNumber = ws.Cells(I + 1, 3).Value
            
            'situation when the open number is 0, then caluculaate the next open_number
            If openNumber = 0 Then
                For J = I + 1 To lastrow
                    If ws.Cells(J, 1).Value <> ws.Cells(J + 1, 1).Value Then Exit For
                    If ws.Cells(J, 3).Value > 0 Then
                        openNumber = ws.Cells(J, 3).Value
                        Exit For
                    End If
                Next J
            End If
            x = x + 1
            total = ws.Cells(I + 1, 7).Value
        End If
    Next I
    
    'Calculate the greatet % Increase
    For I = 2 To lastrow
        If ws.Cells(I, 11) > MaxNumber Then
            MaxNumber = ws.Cells(I, 11).Value
            MaxLine = I
        End If
    Next I
    ws.Cells(2, 16).Value = ws.Cells(MaxLine, 9).Value
    ws.Cells(2, 17).Value = ws.Cells(MaxLine, 11).Value
    ws.Cells(2, 17).NumberFormat = "0.00%"
    
    'Calculate the greatet % Decrease
    For I = 2 To lastrow
        If ws.Cells(I, 11) < MinNumber Then
            MinNumber = ws.Cells(I, 11).Value
            MinLine = I
        End If
    Next I
    ws.Cells(3, 16).Value = ws.Cells(MinLine, 9).Value
    ws.Cells(3, 17).Value = ws.Cells(MinLine, 11).Value
    ws.Cells(3, 17).NumberFormat = "0.00%"
    
    'Calculate the Greatest Total Value
    For I = 2 To lastrow
        If ws.Cells(I, 12) > MaxVolume Then
            MaxVolume = ws.Cells(I, 12).Value
            MaxRow = I
        End If
    Next I
    ws.Cells(4, 16).Value = ws.Cells(MaxRow, 9).Value
    ws.Cells(4, 17).Value = ws.Cells(MaxRow, 12).Value
       
Next ws
                        
End Sub