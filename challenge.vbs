Sub Reset():

' Define variables

    Dim i As Integer
    Dim ws As Worksheet
    For i = 1 To 6

    ThisWorkbook.Worksheets(i).Activate

' Reset the data inputs

    Range("$I:$Z").Value = ""

' Reset the color formatting

    Range("$A:$Z").Interior.Color = -4142
    
    Next i
    
    MsgBox ("Reset Complete")
    
End Sub



Sub ColumnFormat():

' Set the Variables

    Dim i As Integer
    Dim ws As Worksheet
    For i = 1 To 6
    
    ThisWorkbook.Worksheets(i).Activate

' Create New Columns

    Range("I1:P1").Value = Array("Ticker", "Yearly Change", "Percent Change", "Total Stock Volume", "", "", "Ticker", "Value")
    Range("I1:L1,O1:P1").Interior.ColorIndex = 17
    Range("N2:N4").Interior.ColorIndex = 24
    Columns("I:P").ColumnWidth = 16
    Rows("1").RowHeight = 16

' Set New Rows

    Range("N2:N4").Value = Array("Greatest % Increase", "Greatest % Decrease", "Greatest Total Volume")
    
' Run through all worksheets
    
    Next i

End Sub



Sub LastRow():

Dim LR As Long

LR = Cells(Rows.Count, 1).End(xlUp).Row

MsgBox LR

End Sub





Sub CalculateDataFields():

' Identify variables
Dim Ticker As String
Dim TotalVol As String
Dim OpenPrice As Double
Dim OpenPrice2 As Double
Dim ClosePrice As Double
Dim YearlyChange As Double
Dim PercentChange As Double
Dim SummaryTableRow As Integer
Dim LastRow As Long
Dim i As Long

' Where do the variables start
TotalVol = 0
SummaryTableRow = 2
LastRow = Cells(Rows.Count, 1).End(xlUp).Row
OpenPrice = Range("$C$2").Value

' Set data range
For i = 2 To LastRow
        
    If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then

        ' Volume
        TotalVol = TotalVol + Cells(i, 7).Value
            Range("L" & SummaryTableRow).Value = TotalVol
        
        ' Ticker
        Ticker = Cells(i, 1).Value
            Range("I" & SummaryTableRow).Value = Ticker
        
        ' Close Price
        ClosePrice = Cells(i, 6).Value
        
        ' Percent Change
        PercentChange = ((ClosePrice - OpenPrice) / OpenPrice)
            Range("K" & SummaryTableRow).Value = PercentChange
            
         ' Yearly Change
        YearlyChange = (ClosePrice - OpenPrice)
            Range("J" & SummaryTableRow).Value = YearlyChange

        ' Reset Variables
        TotalVol = 0
        SummaryTableRow = SummaryTableRow + 1
        OpenPrice = Cells(i + 1, 3).Value

    Else
        
        ' Add to total volume
        TotalVol = TotalVol + Cells(i, 7).Value
    
    End If

Next i


End Sub


Sub Notes():
            
        ' reset the row/column sizing
        ' add interior color to a:g
        ' calculate winners summary
        ' create buttons?
        ' can you format percentage
        

End Sub






















