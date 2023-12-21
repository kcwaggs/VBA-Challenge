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
    
' column and row width

    Columns("A:Z").ColumnWidth = 8
    Rows("1").RowHeight = 15
    
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

    Range("I1:Q1").Value = Array("Ticker", "Yearly Change", "Percent Change", "Total Stock Volume", "", "", "", "Ticker", "Value")
    Range("I1:L1,P1:Q1,A1:G1").Interior.ColorIndex = 17
    Range("O2:O4").Interior.ColorIndex = 24
    Columns("I:L").ColumnWidth = 18
    Columns("O:Q").ColumnWidth = 18
    Rows("1").RowHeight = 16

' Set New Rows

    Range("O2").Value = ("Greatest % Increase")
    Range("O3").Value = ("Greatest % Decrease")
    Range("O4").Value = ("Greatest Total Volume")
    
' set Column J as %

    Range("K2:K999").NumberFormat = "0.00%"
    Range("Q2:Q3").NumberFormat = "0.00%"

' Run through all worksheets
    
    Next i

End Sub



Sub CalculateDataFields():

' run through all tabs


' Identify variables
Dim Ticker As String
Dim TotalVol As Long
Dim OpenPrice As Double
Dim ClosePrice As Double
Dim YearlyChange As Double
Dim PercentChange As Double
Dim SummaryTableRow As Integer
Dim LastRow As Long
Dim i As Long

' Where do the variables start
TotalVol = 0
SummaryTableRow = 2
LastRow = Cells(Rows.Count, "A").End(xlUp).Row
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





Sub ConditionalFormatting():

' set your variables
    
    Dim i As Double
    Dim LastRow As Integer
    
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
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
    
    For i = 2 To 3
    
        If Cells(i, "Q") > 0 Then
            Cells(i, "Q").Interior.ColorIndex = 43
            
        ElseIf Cells(i, "Q") < 0 Then
            Cells(i, "Q").Interior.ColorIndex = 3
            
        End If
        
    Next i
    
    



End Sub



Sub SecondTable():

' set variables
    Dim PercentIncrease As Double
    Dim PercentDecrease As Double
    Dim TotalVol As String
    Dim LastRow As Integer
    Dim GPITicker As String
    Dim GPDTicker As String
    Dim GTVTicker As String
    
    LastRow = Cells(Rows.Count, "I").End(xlUp).Row
    PercentDecrease = 0
    PercentIncrease = 0
    TotalVol = 0
    
    For i = 2 To LastRow
    
        If Cells(i, "K").Value > PercentIncrease Then
            PercentIncrease = Cells(i, "K").Value
            Range("Q2") = PercentIncrease
            GPITicker = Cells(i, "I").Value
            Range("P2") = GPITicker
            
        ElseIf Cells(i, "K").Value < PercentDecrease Then
            PercentDecrease = Cells(i, "K").Value
            Range("Q3") = PercentDecrease
            GPDTicker = Cells(i, "I").Value
            Range("P3") = GPDTicker
            
        End If
        
        If Cells(i, "L").Value > TotalVol Then
            TotalVol = Cells(i, "L").Value
            Range("Q4") = TotalVol
            GTVTicker = Cells(i, "I").Value
            Range("P4") = GTVTicker
            
        End If
        
        
    Next i



End Sub






Sub Stock():

' run through the entire workbook

    For Each ws In Worksheets
        ws.Activate
    
' what kind of variables do we have
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


    ' workbook formatting
 
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
        

    ' first summary table

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


    ' second summary table
    
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


    ' conditional formatting
        
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

Next ws


End Sub


