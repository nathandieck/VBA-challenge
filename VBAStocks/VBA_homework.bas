
Sub master()
    
    Call sort
    
    Call format

    Call list_stocks

    Call stock_value
    
    Call pct_change
    
    Call format_cells
    
    Call superlatives
    

End Sub

'This subroutine will ensure that the ticker symbols are all sorted correctly so that the rest of the program will work even if the initial data is unsorted.

Sub sort()

    
    Dim lastrow As Long
    lastrow = Cells(Rows.Count, 1).End(xlUp).Row
    With ActiveWorkbook.ActiveSheet.sort
        .SortFields.Clear
        .SortFields.Add Key:=Range("A2:A" & lastrow), Order:=xlAscending
        .SortFields.Add Key:=Range("B2:B" & lastrow), Order:=xlAscending
        .SetRange Range("A2:G" & lastrow)
        .Header = xlNo
        .Apply
    End With
    'https://stackoverflow.com/questions/37998154/vba-excel-variable-sorting-on-multiple-keys-orders
    

End Sub

'This subroutine sets up the table headers for the requested results.

Sub format()

    Cells(1, 10).Value = "Ticker Symbol"
    
    Cells(1, 11).Value = "Annual Change"
    
    Cells(1, 12).Value = "Percent Change"
    
    Cells(1, 13).Value = "Total Stock Volume"
    
    Cells(1, 15).Value = "Ticker"
    
    Cells(1, 16).Value = "Value"
    
    Cells(2, 14).Value = "Greatest % Increase"
    
    Cells(3, 14).Value = "Greatest % Decrease"
    
    Cells(4, 14).Value = "Largest Volume"
    
    Range("J1:R1").Font.Bold = True 'https://docs.microsoft.com/en-us/office/vba/api/excel.font.bold
    
    Columns("J:R").AutoFit ' from the Wells Fargo part 2 activity
    
    Range("N1:N4").Font.Bold = True
    
    

End Sub

' This is the subroutine to populate the list of different ticker symbols.

Sub list_stocks()

    Dim lastrow As Long

    lastrow = Cells(Rows.Count, "A").End(xlUp).Row
    
    Dim j As Integer
    
    j = 2
    
    For i = 2 To lastrow
    
        If Cells(i + 1, 1) <> Cells(i, 1) Then
            
            Cells(j, 10).Value = Cells(i, 1).Value
            
            j = j + 1
            
        End If
        
    Next i
     

End Sub

Sub stock_value()

'This subroutine should supply the volume of the stock (sum of total volume)

    
    Dim lastrow As Long

    lastrow = Cells(Rows.Count, "A").End(xlUp).Row
    
    Dim j As Integer
    
    j = 2
    
    Dim v As Double
    
    v = 0
    
    Dim volume As Double
    
    volume = 0
    
    For i = 2 To lastrow
        
        If Cells(i + 1, 1) = Cells(i, 1) Then
        
            volume = Cells(i, 7).Value
            
            v = v + volume
        
            'If Cells(i, 7).Value <> 0 Then
            
            '   v = v + volume
                 
            'Else
             '   v = v + 0
        
       ElseIf Cells(i + 1, 1) <> Cells(i, 1) Then
       
            volume = Cells(i, 7).Value
            
            v = v + volume
            
            Cells(j, 13) = v
            
            volume = 0
            
            v = 0
            
            j = j + 1
       
        End If
 
            
    Next i
    
    'This is me trying to put the separators (commas) into the numbers, unsuccessfully
    'It is the least of my problems right now.
    
    
    
    'Dim column_end As Long
    
    'columnend = Cells(Rows.Count, "M").End(xlUp).Row
    
    'For k = 2 To columnend
    
     '   Print format(Cells(k, 13), "#,###,###,###")
        
    'Next k
        
        
        
            
End Sub

Sub pct_change()

'this one is for the percent change component of the exercise

    Dim lastrow As Long

    lastrow = Cells(Rows.Count, "A").End(xlUp).Row
    
    Dim j As Integer
    
    j = 2
    
    Dim day1_open As Double
    
    Dim dayx_close As Double
    
    day1_open = Cells(2, 3).Value
    
    dayx_close = 0
    
    For i = 2 To lastrow
    
        If Cells(i + 1, 1) <> Cells(i, 1) Then
            
            dayx_close = Cells(i, 6).Value
            
            Dim net_change As Double
            
            Dim pct_change As Double
            
            If dayx_close = 0 Then
            
                net_change = 0
                pct_change = 0
            
            ElseIf dayx_close > 0 Then
            
                net_change = dayx_close - day1_open
            
                pct_change = dayx_close / day1_open
            
                pct_change = pct_change - 1
            
            End If
        
            Cells(j, 11).Value = net_change
            
            Cells(j, 12).Value = pct_change
            
            Cells(j, 12).Value = FormatPercent(Cells(j, 12), , vbTrue) 'https://www.excelfunctions.net/vba-formatpercent-function.html
            
            j = j + 1
            
            day1_open = Cells(i + 1, 3).Value
            
        End If
    
    
    Next i




End Sub

'this part is intended to conditionally format the cells red and green as instructed


Sub format_cells()

    Dim lastrow As Long

    lastrow = Cells(Rows.Count, "K").End(xlUp).Row
    
    For i = 2 To lastrow
    
        If Cells(i, 11) > 0 Then
            
            Range("j" & i & ":m" & i).Interior.ColorIndex = 10
            
            Range("j" & i & ":m" & i).Font.ColorIndex = 4
            
        ElseIf Cells(i, 11) < 0 Then
            
            Range("j" & i & ":m" & i).Interior.ColorIndex = 53
            
            Range("j" & i & ":m" & i).Font.ColorIndex = 3
    End If
    
    Next i
    
    
End Sub

'this part is the part where we calculate the three 'superlatives' - top pct, bottom pct, and largest volume

Sub superlatives()

    'find the last row
    
    Dim lastrow As Long
    
    lastrow = Cells(Rows.Count, "K").End(xlUp).Row
    

    'find the values
    
    
    Dim top_pct As Single
    
    Dim top_pct_ticker As String
    
    Dim bottom_pct As Single
    
    Dim bottom_pct_ticker As String
    
    Dim top_volume As Double
    
    Dim top_volume_ticker As String

    'top_pct
    
    top_pct = 0
    
    top_pct_ticker = "Null"
    
    For i = 2 To lastrow
        
            If Cells(i, 12).Value > top_pct Then
        
                top_pct = Cells(i, 12)
                
                top_pct_ticker = Cells(i, 10)
                
                
            End If
    
        Next i
        
    Range("P2").Value = top_pct
    
    Range("O2").Value = top_pct_ticker
    
    Cells(2, 16).Value = FormatPercent(Cells(2, 16), , vbTrue)
    
    'bottom_pct
    
    bottom_pct = 0
    
    bottom_pct_ticker = "Null"
    
    For i = 2 To lastrow
        
            If Cells(i, 12).Value < bottom_pct Then
        
                bottom_pct = Cells(i, 12)
                
                bottom_pct_ticker = Cells(i, 10)
                
                
            End If
    
        Next i
        
    Range("P3").Value = bottom_pct
    
    Range("O3").Value = bottom_pct_ticker
    
    Cells(3, 16).Value = FormatPercent(Cells(3, 16), , vbTrue)
    
    'top_volume
    
    top_volume = 0
    
    top_volume_ticker = "Null"
    
    For i = 2 To lastrow
        
            If Cells(i, 13).Value > top_volume Then
        
                top_volume = Cells(i, 13)
                
                top_volume_ticker = Cells(i, 10)
                
                
            End If
    
        Next i
        
    Range("P4").Value = top_volume
    
    Range("O4").Value = top_volume_ticker

    Columns("O:P").AutoFit
        
End Sub
