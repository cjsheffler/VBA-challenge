Attribute VB_Name = "Module1"
'Create sub and declare variables
Sub Stonks()

For Each ws In Excel.Worksheets
    
    Dim ticker As String
    Dim lastrow As Long
    Dim tradevolume As Double
    Dim sumtablerow As Integer
    sumtablerow = 2
    Dim openvalue As Double
    Dim closevalue As Double
    Dim yearchange As Double
    Dim yearpercentchange As Double
    
    
    'Find the last row
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    openvalue = ws.Cells(2, 3).Value
    
    'Create the headers for new data
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    

    'Create the for loop to gather information needed.
    For i = 2 To lastrow
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        ticker = ws.Cells(i, 1).Value
        tradevolume = tradevolume + ws.Cells(i, 7).Value
        closevalue = ws.Cells(i, 6).Value
        yearchange = closevalue - openvalue
    
        'add code to eliminate Div0 errors and calculate percentage change
        If openvalue = 0 Then
        yearpercentchange = 0
        Else: yearpercentchange = (closevalue - openvalue) / openvalue
        End If
    
        'fill in the calculated results
        ws.Range("I" & sumtablerow).Value = ticker
        ws.Range("J" & sumtablerow).Value = yearchange
    
        'format the value in the calculated yearly change result
        If ws.Range("J" & sumtablerow).Value > 0 Then
            ws.Range("J" & sumtablerow).Interior.ColorIndex = 4
        ElseIf ws.Range("J" & sumtablerow).Value < 0 Then
            ws.Range("J" & sumtablerow).Interior.ColorIndex = 3
        Else: ws.Range("J" & sumtablerow).Interior.ColorIndex = 2
        End If
    
        'Change percentage to show in better format
        ws.Range("K" & sumtablerow).Value = Format(yearpercentchange, "Percent")
        If ws.Range("K" & sumtablerow).Value < 0 Then
            ws.Range("K" & sumtablerow).Font.ColorIndex = 3
        End If
    
        ws.Range("L" & sumtablerow).Value = tradevolume
    
        'reset trade volume to 0
        tradevolume = 0
    
        'add one to the sum table row so data goes to new line
        sumtablerow = sumtablerow + 1
    
        'find the new opening value
        openvalue = ws.Cells(i + 1, 3).Value
    
        'On the last matching result, add the final volume value
        Else
        tradevolume = tradevolume + ws.Cells(i, 7).Value
        End If
    Next i


    'Add in the extra bonus work
    bigchange = Application.WorksheetFunction.Max(ws.Range("K:K"))
    lowchange = Application.WorksheetFunction.Min(ws.Range("K:K"))
    volmax = Application.WorksheetFunction.Max(ws.Range("L:L"))
    ws.Range("N2").Value = "Greatest Change"
    ws.Range("N3").Value = "Greatest Decrease"
    ws.Range("N4").Value = "Greatest Volume"
    
    ws.Range("O1").Value = "Ticker"
    ws.Range("P1").Value = "Value"
    
    ws.Range("O2").Formula = "=INDEX(($I:$I),MATCH(($P$2),($K:$K),0))"
    ws.Range("O3").Formula = "=INDEX(($I:$I),MATCH(($P$3),($K:$K),0))"
    ws.Range("O4").Formula = "=INDEX(($I:$I),MATCH(($P$4),($L:$L),0))"


    ws.Range("P2").Value = bigchange
    ws.Range("P3").Value = lowchange
    ws.Range("P4").Value = volmax
    
    
    

Next ws

MsgBox ("Calculations complete!")

End Sub

