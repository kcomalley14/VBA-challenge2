Attribute VB_Name = "Module1"
Sub tickervolumes()
'Define Variable for Tickers
Dim TickerName As String

' OPen price found needs to be for next ticker
Dim OpenPrice As Double
Dim ClosePrice As Variant

' Set initial variable for holding the total per ticker
Dim TotalVolume As Variant
TotalVolume = 0

'Keep track of the locations for each ticker in summary
Dim SummaryTable As Variant
SummaryTable = 2



' Loop through all the tickers
    For i = 2 To Cells(Rows.Count, 1).End(xlUp).Row
    
If Cells(i - 1, 1).Value <> Cells(i, 1).Value Then
    ' Find open value for each ticker
    OpenPrice = Cells(i, 3).Value
    
    ' Open price range printed for analysis
    Range("N" & SummaryTable).Value = OpenPrice
    End If
    If OpenPrice = 0 Then
        Cells(i, 3).Value = 1E-09
    End If
    
    
' Check if we are still within the same ticker value
If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    
    ' Set the ticker name
    TickerName = Cells(i, 1).Value
    
    ' Add to total volume
    TotalVolume = TotalVolume + Cells(i, 7).Value
    
    ' Range for closed price analysis
    Range("O" & SummaryTable).Value = ClosePrice
    
   ' Print ticker value in summary table
    Range("I" & SummaryTable).Value = TickerName
    
    ' Print total volume in summary table
    Range("L" & SummaryTable).Value = TotalVolume
    
    ' Add one to the summary row to go to next row
    SummaryTable = SummaryTable + 1
    
    ' Reset Value of summary table
    TotalVolume = 0

'If the immediate row is the same value then..
Else
    
    ' Add to the Volume total
    TotalVolume = TotalVolume + Cells(i, 7).Value
    
    ' Find close price from each ticker
    ClosePrice = Cells(i + 1, 6).Value
       
   End If
  
    
    Next i

    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"
    Cells(i, 14).Value = "Open Price"
    Cells(i, 15).Value = "Close Price"

    Dim percentchange As Variant

    Dim openingstock As Variant
    Dim closingstock As Variant
    Dim YearlyChange As Variant
    
    For i = 2 To Cells(Rows.Count, 1).End(xlUp).Row
    
    
    openingstock = Cells(i, 14).Value
    closingstock = Cells(i, 15).Value
    
    YearlyChange = closingstock - openingstock
    Cells(i, 10).Value = YearlyChange
    
    percentchange = ((Cells(i, 10).Value / Cells(i, 14).Value) * 100)
    
    Cells(i, 11).Value = percentchange
    
        
    Next i
End Sub

    Sub conditionals()
    
    For i = 2 To Cells(Rows.Count, 1).End(xlUp).Row
    
    If Cells(i, 10).Value < 0 Then
        Cells(i, 10).Interior.ColorIndex = 3
    
    Else
        Cells(i, 10).Interior.ColorIndex = 4
    End If
    Next i
    

End Sub


