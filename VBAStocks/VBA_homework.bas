
Sub master()

    'The master subroutine governs the entire program and includes the coding to make the code run on each worksheet in sucession. 
    'Most of the other functionality comes from called subroutines. 
    
    Dim ws As Worksheet
    
    Application.ScreenUpdating = False 
    
    'I learned about screen updating from https://docs.microsoft.com/en-us/office/vba/api/excel.application.screenupdating
    
    For Each ws In Worksheets
    
        ws.Activate
    
        Call sort
    
        Call format

        Call list_stocks

        Call stock_value
    
        Call pct_change
    
        Call format_cells
    
        Call superlatives
    
    Next ws
    
    Application.ScreenUpdating = True

End Sub

'This subroutine will ensure that the ticker symbols are all sorted correctly so that the rest of the program will work even if the initial data is unsorted.
'It is not part of the assignment but it is intended to make sure that the code still runs even if the stocks end up out of order somehow. 

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
    
'The above code comes from https://stackoverflow.com/questions/37998154/vba-excel-variable-sorting-on-multiple-keys-orders
    
End Sub

'The subroutine below sets up the table headers for the requested results.

Sub format()

    ActiveSheet.Cells(1, 10).Value = "Ticker Symbol"
    
    ActiveSheet.Cells(1, 11).Value = "Annual Change"
    
    ActiveSheet.Cells(1, 12).Value = "Percent Change"
    
    ActiveSheet.Cells(1, 13).Value = "Total Stock Volume"
    
    ActiveSheet.Cells(1, 15).Value = "Ticker"
    
    ActiveSheet.Cells(1, 16).Value = "Value"
    
    ActiveSheet.Cells(2, 14).Value = "Greatest % Increase"
    
    ActiveSheet.Cells(3, 14).Value = "Greatest % Decrease"
    
    ActiveSheet.Cells(4, 14).Value = "Largest Volume"
    
    ActiveSheet.Range("J1:R1").Font.Bold = True 'https://docs.microsoft.com/en-us/office/vba/api/excel.font.bold
    
    ActiveSheet.Columns("J:R").AutoFit ' from the Wells Fargo part 2 activity
    
    Range("N1:N4").Font.Bold = True

End Sub

' The subroutine below populates the list of different ticker symbols.

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

'The subroutine below supplies the total volume of each stock (sum of volumes). 

Sub stock_value()

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
        
       ElseIf Cells(i + 1, 1) <> Cells(i, 1) Then
       
            volume = Cells(i, 7).Value
            
            v = v + volume
            
            Cells(j, 13) = v
            
            volume = 0
            
            v = 0
            
            j = j + 1
       
        End If
 
            
    Next i       
            
End Sub

'The subroutine below handles the percent change component for each stock. 

Sub pct_change()

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
            
                ElseIf day1_open = 0 Then
                
                    net_change = 0
                    pct_change = 0
                    
                Else
                    
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

'The below subroutine conditionally formats the rows red and green as instructed. 

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

'this subroutine calculates the three 'superlatives' - top pct, bottom pct, and largest volume

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
