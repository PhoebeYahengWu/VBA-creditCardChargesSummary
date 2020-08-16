# VBA-creditCardChargesSummary

## Start File
![image](https://user-images.githubusercontent.com/52837649/90323177-9c793b80-df2b-11ea-8bd9-775cd46b4e2d.png)

## Finished File
![image](https://user-images.githubusercontent.com/52837649/90323676-9cc90500-df32-11ea-8ea3-5b0c4f38f8d9.png)

## Code
```
Sub summarize()

    Dim Card_Name As String
    Dim Card_Total As Double
    
    Card_Total = 0
    
    'keep track of the location for each credit card in the summary table
    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2
    
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    For i = 2 To lastRow
        'check if we are still within the same credit card brand
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            Card_Name = Cells(i, 1).Value
            
            'add to the card total
            Card_Total = Card_Total + Cells(i, 3).Value
            
            'print the credit card name in the summary table
            Range("G" & Summary_Table_Row).Value = Card_Name
            
            'print the card total in the summary table
            Range("H" & Summary_Table_Row).Value = Card_Total
            
            'add 1 to the summary table row
            Summary_Table_Row = Summary_Table_Row + 1
            
            'reset the card total
            Card_Total = 0
        Else
            Card_Total = Card_Total + Cells(i, 3).Value
        End If
        
    Next i
    
End Sub
```
