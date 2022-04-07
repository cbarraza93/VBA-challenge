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
    Dim starting As Long
    Dim closed As Long
    Dim change As Double
    Dim percent As Double
    
    total = 0
    output = 2
    starting = 2
    lastrow = Cells(Rows.Count, 1).End(xlUp).Row

End Sub
