Sub VBA_HW_Regrade()
    
Dim ws As Worksheet

    For Each ws In Worksheets
    ws.Activate
                    
    ' Get the last row
        
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
    ' Create the Variables
        
        Dim Opening_Price As Double
        Dim Closing_Price As Double
        Dim Yearly_Change As Double
        Dim Ticker As String
        Dim Percent_Change As Double
        Dim Volume As Double
        Volume = 0
        Dim Row As Double
        Row = 2
        Dim Column As Integer
        Column = 1
        Dim i As Long
        
    ' Create the summary table labels
        
        ws.Cells(1, "I").Value = "Ticker"
        ws.Cells(1, "J").Value = "Yearly Change"
        ws.Cells(1, "K").Value = "Percent Change"
        ws.Cells(1, "L").Value = "Total Stock Volume"
        
'Set the Open Price
    Opening_Price = Cells(2, "C").Value
        
' Loop through ticker names
        
    For i = 2 To LastRow
        
         
        If Cells(i + 1, Column).Value <> Cells(i, Column).Value Then
                
            ' Put ticker name in summary table
                
            Ticker = Cells(i, Column).Value
            Cells(Row, "I").Value = Ticker
                
' Set the Close Price
            Closing_Price = Cells(i, "F").Value
                
' Add Yearly Change by subtracting closing price from opening price
                
            Yearly_Change = Closing_Price - Opening_Price
            Cells(Row, "J").Value = Yearly_Change
                
' Calculate the percentage change from the beginning of the year to the end of the

            If (Opening_Price = 0 And Closing_Price = 0) Then
                Percent_Change = 0
            
            ElseIf (Opening_Price = 0 And Closing_Price <> 0) Then
                Percent_Change = 1
            
            Else
                Percent_Change = Yearly_Change / Opening_Price
                Cells(Row, "K").Value = Percent_Change
                Cells(Row, "K").NumberFormat = "0.00%"
            End If

' Set the variable for total volume
            Volume = Volume + Cells(i, "G").Value


' Print in the column
            Cells(Row, "L").Value = Volume

' Add to the summary table rows and reset the opening price
            Row = Row + 1
            Opening_Price = Cells(i + 1, Column + 2)


' Reset the volume to 0
            Volume = 0


        Else
            Volume = Volume + Cells(i, "G").Value
        End If
    
    Next i
        
        
' Add "Yearly Change" to each worksheet and add the colors

YearChangeLastRow = ws.Cells(Rows.Count, "J").End(xlUp).Row
        
    For j = 2 To YearChangeLastRow
        If (ws.Cells(j, "J").Value > 0 Or ws.Cells(j, "J").Value = 0) Then
            ws.Cells(j, "J").Interior.ColorIndex = 4
        
        ElseIf ws.Cells(j, "J").Value < 0 Then
            ws.Cells(j, "J").Interior.ColorIndex = 3
        End If
        
    Next j
        
        
    Next ws
        
End Sub