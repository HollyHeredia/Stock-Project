Attribute VB_Name = "Module2"
Sub VBAHomework()

'Define summary box headers
Range("I1").Value = "Ticker"
Range("J1").Value = "Yearly Change"
Range("K1").Value = "Percent Change"
Range("L1").Value = "Total Stock Volume"
'Bold headers
Range("I1:L1").Font.Bold = True
    
'define Summary_Table_Row
    Dim Summary_Table_Row As Double
    Summary_Table_Row = 2

'dim StockVolume
    Dim StockVolume As Double
    StockVolume = 0
    
'define 1st BeginningValue
    Dim BeginningValue As Double
    BeginningValue = Cells(2, 3).Value

'BEGINNING OF LOOP
    For i = 2 To Cells(Rows.Count, 1).End(xlUp).Row

'BEGINNING IF STATEMENT if there is a difference in the ticker then all these things happen
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    
'create ticker
    'Define ticker name
    Dim Ticker As String
    Ticker = Cells(i, 1).Value
    
'put ticker into summary box
    Cells(Summary_Table_Row, 9).Value = Ticker
    
'yearly change
    
    'define Closing value
    Dim ClosingValue As Double
    ClosingValue = Cells(i, 6).Value

    'define Yearly change
    Dim YearlyChange As Double
    YearlyChange = ClosingValue - BeginningValue

    'Yearlychange into summary table
    Cells(Summary_Table_Row, 10).Value = YearlyChange
    
    'Conditional formatting for values IF statement
        If YearlyChange > 0 Then
            Cells(Summary_Table_Row, 10).Interior.ColorIndex = 4
        ElseIf YearlyChange < 0 Then
            Cells(Summary_Table_Row, 10).Interior.ColorIndex = 3
        'Ending if
        End If
        
'% change
    Dim PercentChange As Double
    PercentChange = (ClosingValue - BeginningValue) / BeginningValue
    
    '%Change into summary table
    Cells(Summary_Table_Row, 11).Value = PercentChange
    'format cells as percentage
    Cells(Summary_Table_Row, 11).NumberFormat = "00.00%"
    
        
'redefine BeginningValue since we're done with the former value and we need to only grab
'the new value with ticker changes
   BeginningValue = Cells(i + 1, 3).Value
    
        'Nested IF for BeginningValue of 0
        If BeginningValue = 0 Then
        'change BeginningValue to the next cell in the same column
       BeginningValue = ActiveCell.Offset(1, 0).Select
    
        Else
        BeginningValue = Cells(i + 1, 3).Value
    'end of BeginningValue Nested IF
        End If
    
    
'StockVolume in Summary Table
    Cells(Summary_Table_Row, 12).Value = StockVolume
    
'next Summary Table Row
    Summary_Table_Row = Summary_Table_Row + 1

'reset StockVolume to 0
    StockVolume = 0

'next part of IF statement if ticker is not changing
    Else
    
    'Add up StockVolume in block
    StockVolume = Cells(i, 7).Value + StockVolume
    
'End of main IF statement
    End If

'Next loop to do the whole thing again!
    Next i

'End of macro
End Sub

