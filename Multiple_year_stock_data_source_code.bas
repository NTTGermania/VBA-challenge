Attribute VB_Name = "Module1"
Sub BudgetChecker():
    
    'Running my codes through all the worksheets
    Dim ws As Worksheet
    For Each ws In Worksheets
        
        
        Dim Ticker As String
        Dim Counter As Integer
        Dim TotalStockVolume As Double
        Dim YearlyChange As Double
        Dim PercentChange As Double
        PercentChange = 0
        TotalStockVolume = 0
        YearlyChange = 0
        Counter = 1
        Dim i As Long
        Dim openingprice As Double
        Dim closingprice As Double
        
        
        ' Counts the number of rows
        LastRow = Cells(Rows.Count, 1).End(xlUp).Row
        
        ' Labelling headers
        Range("I1").Value = "Ticker"
        Range("J1").Value = "Yearly Change"
        Range("K1").Value = "Percent Change"
        Range("L1").Value = "Total Stock Volume"
        
        
        'Capturing the first opening stock value
        openingprice = Cells(2, 3).Value
        
        ' Loop through each row
        For i = 2 To LastRow
            
            'Adding volumes of each stock together
            TotalStockVolume = TotalStockVolume + Cells(i, 7).Value
            
            'Keeping track of Ticker changes and displaying them
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
                Ticker = Cells(i, 1).Value
                Counter = Counter + 1
                Cells(Counter, 9).Value = Ticker
                closingprice = Cells(i, 6).Value
                
                'Calculating the Yearly Changes
                YearlyChange = closingprice - openingprice
                Cells(Counter, 10).Value = YearlyChange
                
                'Keeping track of negative and positive stock changes and color them accordingly
                    If (YearlyChange > 0) Then
                        Cells(Counter, 10).Interior.ColorIndex = 4
                    ElseIf (YearlyChange < 0) Then
                        Cells(Counter, 10).Interior.ColorIndex = 3
                    Else
                         Cells(Counter, 10).Value = YearlyChange
                    End If
                'Calculating Percent Changes and coverting it into percent format
                PercentChange = YearlyChange / openingprice
                openingprice = Cells(i + 1, 3).Value
                Cells(Counter, 11).Value = PercentChange
                Cells(Counter, 11).NumberFormat = "0.00%"
                'Code to write TotalStockVolume volume to the cell
                Cells(Counter, 12) = TotalStockVolume
                'Reseting TotalStockVolume to 0
                TotalStockVolume = 0
                
              
            End If
            
        Next i

    ws.Activate
    Debug.Print ws.Name
    Next
End Sub

