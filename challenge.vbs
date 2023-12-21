Sub ResetWorkbook():

' run through the entire workbook
    
    For Each ws In Worksheets
        ws.Activate

' reset data inputs

    Range("$I:$Z").Value = ""

' reset the color formatting

    Range("$A:$Z").Interior.Color = -4142
    
' reset the column and row width

    Columns("A:Z").ColumnWidth = 8
    Rows("1").RowHeight = 15
    
' close the loop
    
    Next ws

' you did it!
    MsgBox ("the workbook has been reset!")

End Sub


Sub AnalyzeStocks():

' APPLY TO THE WORKBOOK

    For Each ws In Worksheets
        ws.Activate
    
' THE VARIABLES
    Dim Ticker As String
    Dim TotalVol As Double
    Dim OpenPrice As Double
    Dim ClosePrice As Double
    Dim YearlyChange As Double
    Dim PercentChange As Double
    Dim SummaryTableRow As Integer
    Dim LastRow As Long
    Dim i As Long
    Dim PercentIncrease As Double
    Dim PercentDecrease As Double
    Dim GPITicker As String
    Dim GPDTicker As String
    Dim GTVTicker As String

' WORKBOOK FORMATTING
 
    ' set new columns for summary tables

        Range("I1:Q1").Value = Array("Ticker", "Yearly Change", "Percent Change", "Total Stock Volume", "", "", "", "Ticker", "Value")
        Range("I1:L1,P1:Q1,A1:G1").Interior.ColorIndex = 17
        Range("O2:O4").Interior.ColorIndex = 24
        Columns("I:L").ColumnWidth = 18
        Columns("O:Q").ColumnWidth = 18
        Rows("1").RowHeight = 16

    ' set new rows for summary table

        Range("O2").Value = ("Greatest % Increase")
        Range("O3").Value = ("Greatest % Decrease")
        Range("O4").Value = ("Greatest Total Volume")
    
    ' set certain values as a %

        Range("K2:K999").NumberFormat = "0.00%"
        Range("Q2:Q3").NumberFormat = "0.00%"
        
' SUMMARY TABLE - 1

    ' set the variables
        TotalVol = 0
        SummaryTableRow = 2
        LastRow = Cells(Rows.Count, "A").End(xlUp).Row
        OpenPrice = Range("$C$2").Value

    ' what range are we looping
        For i = 2 To LastRow
    
            If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then

                ' volume outcome
                    TotalVol = TotalVol + Cells(i, 7).Value
                        Range("L" & SummaryTableRow).Value = TotalVol
    
                ' ticker outcome
                    Ticker = Cells(i, 1).Value
                        Range("I" & SummaryTableRow).Value = Ticker
    
                ' close price outcome
                    ClosePrice = Cells(i, 6).Value
    
                ' percent change outcome
                    PercentChange = ((ClosePrice - OpenPrice) / OpenPrice)
                        Range("K" & SummaryTableRow).Value = PercentChange
        
                ' yearly change outcome
                    YearlyChange = (ClosePrice - OpenPrice)
                        Range("J" & SummaryTableRow).Value = YearlyChange

                ' reset the variables
                    TotalVol = 0
                    SummaryTableRow = SummaryTableRow + 1
                    OpenPrice = Cells(i + 1, 3).Value

            Else
    
            ' add to cumulative volume
                TotalVol = (TotalVol + Cells(i, 7).Value)

            End If

        Next i

' SUMMARY TABLE - 2

    ' set the variables
        LastRow = Cells(Rows.Count, "I").End(xlUp).Row
        PercentDecrease = 0
        PercentIncrease = 0
        TotalVol = 0

        For i = 2 To LastRow

            ' greatest percent increase
        
                If Cells(i, "K").Value > PercentIncrease Then
            
                    PercentIncrease = Cells(i, "K").Value
                        Range("Q2") = PercentIncrease
            
                    GPITicker = Cells(i, "I").Value
                        Range("P2") = GPITicker
        
            ' greatest percent decrease
        
                ElseIf Cells(i, "K").Value < PercentDecrease Then
            
                    PercentDecrease = Cells(i, "K").Value
                        Range("Q3") = PercentDecrease
            
                    GPDTicker = Cells(i, "I").Value
                        Range("P3") = GPDTicker
        
                End If
    
            ' greatest total volume
        
                If Cells(i, "L").Value > TotalVol Then
            
                    TotalVol = Cells(i, "L").Value
                        Range("Q4") = TotalVol
            
                    GTVTicker = Cells(i, "I").Value
                        Range("P4") = GTVTicker
        
                End If
    
        Next i

' CONDITIONAL FORMATTING
    
    ' set the variables
    
        LastRow = Cells(Rows.Count, "I").End(xlUp).Row
        
        ' first summary table formatting
        
            For i = 2 To LastRow

                If Cells(i, 10) > 0 Then
                    Cells(i, 10).Interior.ColorIndex = 43
    
                ElseIf Cells(i, 10) < 0 Then
                    Cells(i, 10).Interior.ColorIndex = 3
    
                End If

                If Cells(i, 11) > 0 Then
                    Cells(i, 11).Interior.ColorIndex = 43
    
                ElseIf Cells(i, 11) < 0 Then
                    Cells(i, 11).Interior.ColorIndex = 3
    
                End If

            Next i

        ' second summary table formatting
        
            For i = 2 To 3

                If Cells(i, "Q") > 0 Then
                    Cells(i, "Q").Interior.ColorIndex = 43
    
                ElseIf Cells(i, "Q") < 0 Then
                    Cells(i, "Q").Interior.ColorIndex = 3
    
                End If

            Next i

' LOOP THROUGH

Next ws

' the end :))

    MsgBox ("analysis complete! :)")

End Sub