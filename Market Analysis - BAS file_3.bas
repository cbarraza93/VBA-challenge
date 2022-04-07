Attribute VB_Name = "Module1"
Sub MarketAnalysis()
    
' Insert Data Labels Via Cells
    Cells(1, 9) = "Ticker"
    Cells(1, 10) = "Yearly Change"
    Cells(1, 11) = "Percent Change"
    Cells(1, 12) = "Total Stock Volume"

' Declare the Variables

    Dim output As Integer
    Dim total As Double
    Dim starting As Double
    Dim percent As Double
    
    output = 2
    total = 0
    starting = 2
    lastrow = Cells(Rows.Count, 1).End(xlUp).Row
    
' start Iteration

    For i = 2 To lastrow
        
' combine duplicate tickers, then add to column I
        If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
            Cells(output, 9).Value = Cells(i, 1).Value

' calculate the yearly change
            change = Cells(i, 6).Value - Cells(starting, 3)
            Cells(output, 10).Value = change

' error divide by zero
            If Cells(starting, 3).Value = 0 Then
            Cells(output, 10).Value = "No data available"
            Else
                
' calculate the percent change and add to column K
            percent = Cells(output, 10).Value / Cells(starting, 3).Value
            Cells(output, 11).Value = percent
            
            ' change to percent format
                Cells(output, 11).NumberFormat = "0.00%"
            End If
' combine stock volume and add to column L
            total = total + Cells(i, 7).Value
            Cells(output, 12).Value = total

            starting = i + 1
            output = output + 1
            total = 0
    
        Else
            total = total + Cells(i, 7).Value
            
        End If
    
'move to next row
    Next i
    
' Declare the Variables
    lastrow = Cells(Rows.Count, 1).End(xlUp).Row

' declare new iteration
    For i = 2 To lastrow
    
' insert colors based on percent change
        If Cells(i, 10).Value > 0 Then
            Cells(i, 10).Interior.ColorIndex = 4
        End If
        
        If Cells(i, 10).Value < 0 Then
            Cells(i, 10).Interior.ColorIndex = 3
        End If
         
        If Cells(i, 10).Value = "No data available" Then
            Cells(i, 10).Interior.ColorIndex = 8
        End If
        
'move to next row
    Next i

End Sub


