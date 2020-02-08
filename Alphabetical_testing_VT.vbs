Sub Alphabetical_testing():
'Declarations of required  Varaiables

Dim lastrow As Long
Dim ticker As String
Dim open_price As Double
Dim close_price As Double
Dim volume As Long
Dim percentage_change As Double
Dim col_num_open_price As Double
Dim col_num_close_price As Integer
Dim col_num_volume As Integer
Dim col_num_date As Integer
Dim row_num As Long
Dim total_volume As LongLong
Dim count As Long
Dim i As Long
Dim ws As Variant
Dim j As Long
Dim m As Long
Dim l As Long
Dim greatest_vol As LongLong
Dim greatest As Double
Dim lowest As Double
Dim ticker_1 As String
Dim ticker_2 As String
Dim ticker_3 As String


For Each ws In Worksheets
    'Assigning of the varaiables to be used for extracting data
    col_num_open_price = 3
    col_num_close_price = 6
    col_num_date = 2
    col_num_volume = 7
    row_num = 2
    total_volume = 0
    count = 0

    'Calculation of last row and last column
    lastrow = ws.Cells(Rows.count, 1).End(xlUp).Row
    lastcol = ws.Cells(1, Columns.count).End(xlToLeft).Column

    'Headers for extracted data
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    For i = 2 To lastrow
        If ws.Cells(i, 1).Value = ws.Cells(i + 1, 1).Value Then
            volume = ws.Cells(i, col_num_volume).Value
            total_volume = total_volume + volume
            count = count + 1
        
        ElseIf ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
            open_price = ws.Cells(i - count, col_num_open_price).Value
            'Writing the ticker value
            ticker = ws.Cells(i, 1).Value
            ws.Cells(row_num, 9).Value = ticker
           
           'Extracting the closing price and calculation of price change
            close_price = ws.Cells(i, col_num_close_price).Value
            ws.Cells(row_num, 10).Value = close_price - open_price
            
           'Calculation of percentage change
            If open_price = 0 Then
                ws.Cells(row_num, 11).Value = "Can't be Calculated"
            Else
                percentage_change = (close_price - open_price) / open_price
                ws.Cells(row_num, 11).Value = Format(percentage_change, "Percent")
            End If
            'Total Volume Calculation for this case
            volume = ws.Cells(i, col_num_volume).Value
            total_volume = total_volume + volume
            ws.Cells(row_num, 12).Value = total_volume
         
            'Incrementing the row number where data is being extracted
            row_num = row_num + 1
            total_volume = 0
            count = 0
         
        End If
    
    Next i

    'Colouring the cells green and red depending on positive or negative yearly change
    For i = 2 To row_num - 1
        If ws.Cells(i, 10).Value < 0 Then
            ws.Range("J" & i).Interior.ColorIndex = 3
        Else
            ws.Range("J" & i).Interior.ColorIndex = 4
        End If
    Next i
    
    ws.Range("N2").Value = "Greatest % increase "
    ws.Range("N3").Value = "Greatest % decrease "
    ws.Range("N4").Value = "Greatest stock volume "
    ws.Range("O1").Value = "Ticker "
    ws.Range("P1").Value = "Value "
    
    
    'Challenge part
    greatest = ws.Cells(2, 11).Value
    lowest = ws.Cells(2, 11).Value
    greatest_vol = ws.Cells(2, 12).Value
    For j = 3 To row_num
        If ws.Cells(j, 11).Value > greatest Then
            If ws.Cells(j, 11).Value = "Can't be Calculated" Then
                'do nothing and ignore value
            Else
                greatest = ws.Cells(j, 11).Value
                ticker_1 = ws.Cells(j, 9).Value
            End If
        End If
    Next j
    
    For l = 3 To row_num
        If ws.Cells(l, 11).Value < lowest Then
            lowest = ws.Cells(l, 11).Value
            ticker_2 = ws.Cells(l, 9).Value
        End If
    Next l
    
    For m = 3 To row_num
        If ws.Cells(m, 12).Value > greatest_vol Then
            greatest_vol = ws.Cells(m, 12).Value
            ticker_3 = ws.Cells(m, 9).Value
        End If
    Next m
    
    ws.Range("O2").Value = ticker_1
    ws.Range("O3").Value = ticker_2
    ws.Range("O4").Value = ticker_3
    ws.Range("P2").Value = Format(greatest, "Percent")
    ws.Range("P3").Value = Format(lowest, "Percent")
    ws.Range("P4").Value = greatest_vol
    
Next ws
        
        
        
        
        
End Sub
