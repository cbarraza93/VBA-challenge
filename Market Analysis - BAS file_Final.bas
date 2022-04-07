Attribute VB_Name = "Module1"
Sub MarketAnalysis()

' Run code across multiple sheets
    Application.ScreenUpdating = False
For Each sh In Worksheets
    sh.Activate
    
' Insert Data Labels Via Cells
    cells(1, 9) = "Ticker"
    cells(1, 10) = "Yearly Change"
    cells(1, 11) = "Percent Change"
    cells(1, 12) = "Total Stock Volume"

' Declare the Variables

    Dim output As Integer
    Dim total As Double
    Dim starting As Double
    Dim percent As Double
    
    output = 2
    total = 0
    starting = 2
    lastrow = cells(Rows.Count, 1).End(xlUp).Row
    
' start Iteration

    For i = 2 To lastrow
        
' combine duplicate tickers, then add to column I
        If cells(i, 1).Value <> cells(i + 1, 1).Value Then
            cells(output, 9).Value = cells(i, 1).Value

' calculate the yearly change
            change = cells(i, 6).Value - cells(starting, 3)
            cells(output, 10).Value = change

' error divide by zero
            If cells(starting, 3).Value = 0 Then
            cells(output, 10).Value = "No data available"
            Else
                
' calculate the percent change and add to column K
            percent = cells(output, 10).Value / cells(starting, 3).Value
            cells(output, 11).Value = percent
            
            ' change to percent format
                cells(output, 11).NumberFormat = "0.00%"
            End If
' combine stock volume and add to column L
            total = total + cells(i, 7).Value
            cells(output, 12).Value = total

            starting = i + 1
            output = output + 1
            total = 0
    
        Else
            total = total + cells(i, 7).Value
            
        End If
    
'move to next row
    Next i
    
' Declare the Variables
    lastrow = cells(Rows.Count, 1).End(xlUp).Row

' declare new iteration
    For i = 2 To lastrow
    
' insert colors based on percent change
        If cells(i, 10).Value > 0 Then
            cells(i, 10).Interior.ColorIndex = 4
        End If
        
        If cells(i, 10).Value < 0 Then
            cells(i, 10).Interior.ColorIndex = 3
        End If
         
        If cells(i, 10).Value = "No data available" Then
            cells(i, 10).Interior.ColorIndex = 8
        End If
        
'move to next row
    Next i
  
' Insert Data Labels Via Cells
    cells(2, 14) = "Greatest % increase"
    cells(3, 14) = "Greatest % decrease"
    cells(4, 14) = "Greatest total volume"

' Review Max/Min for % increase/decrease/volume
  
    cells(2, 16) = WorksheetFunction.Max(Range("K:K"))
        cells(2, 16).NumberFormat = "0.00%"
 
    cells(3, 16) = WorksheetFunction.Min(Range("K:K"))
        cells(3, 16).NumberFormat = "0.00%"

    cells(4, 16) = WorksheetFunction.Max(Range("L:L"))
        cells(4, 16).NumberFormat = "0"
  
' move to next sheet
    Next sh
    Application.ScreenUpdating = True
    
End Sub


