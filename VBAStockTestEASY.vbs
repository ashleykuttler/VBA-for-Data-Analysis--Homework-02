Attribute VB_Name = "Module1"
Sub StockTestEasy()
    'define worksheets as variables
    Dim ws As Worksheet
    Dim startingws As Worksheet
    Set startingws = ActiveSheet
    
    'Loop through worksheets
    For Each ws In Worksheets
    ws.Activate
    
    'title summary columns
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Total Stock Volume"
    
    'define data set variables
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    Dim ticker As String
    Dim totalvol As Double
    Dim summaryrow As Integer
    totalvol = 0
    summaryrow = 2
    
    'Loop through each row to find distinct ticker symbols
    For i = 2 To lastrow
    
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        
            'place unique symbol in summary column
            ticker = Cells(i, 1).Value
            Range("I" & summaryrow).Value = ticker
                               
           
           'aggregate total vol for symbol and place in summary column
            totalvol = totalvol + Cells(i, 7).Value
            Range("J" & summaryrow).Value = totalvol
            
                                
            'begin new summary row and reset total volumn
            summaryrow = summaryrow + 1
            totalvol = 0
            
        Else
            'accumulate volumn for the same ticker
            totalvol = totalvol + Cells(i, 7).Value
            
        End If
        
        
    Next i
    Next ws
        
    MsgBox ("Aggregation Complete")
End Sub
