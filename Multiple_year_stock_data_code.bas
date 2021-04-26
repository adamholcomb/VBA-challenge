Attribute VB_Name = "Module1"
Sub Analysis()
    
    'Declare variables
    Dim Ticker As String
    Dim YearlyChange As Double
    Dim PercentChange As Double
    Dim TotalVolume As Double
        TotalVolume = 0
    Dim OutputRow As Integer
        OutputRow = 2
    Dim FirstOpen As Double
    Dim LastClose As Double
        
    'Set-up for multiple worksheets
    For Each ws In Worksheets
        ws.Activate
        LastRow = Cells(Rows.Count, 1).End(xlUp).Row
            
        'Find variables and output into new table
        For x = 2 To LastRow
            If Cells(x - 1, 1).Value <> Cells(x, 1).Value Then
                FirstOpen = Cells(x, 3).Value
            ElseIf Cells(x + 1, 1).Value <> Cells(x, 1).Value Then
                Ticker = Cells(x, 1).Value
                TotalVolume = TotalVolume + Cells(x, 7).Value
                Range("I" & OutputRow) = Ticker
                Range("L" & OutputRow) = TotalVolume
                TotalVolume = 0
                LastClose = Cells(x, 6).Value
                Range("J" & OutputRow) = LastClose - FirstOpen
                Range("K" & OutputRow) = ((LastClose - FirstOpen) / FirstOpen)
                OutputRow = OutputRow + 1
            ElseIf FirstOpen = 0 Then
                FirstOpen = FirstOpen + 1
            Else
                TotalVolume = TotalVolume + Cells(x, 7).Value
            End If
        Next x
        
        'Headings for output table
        OutputRow = 2
        Range("I1") = "Ticker"
        Range("J1") = "Yearly Change"
        Range("K1") = "Percent Change"
        Range("L1") = "Total Volume"
        
        'Conditional formatting for Yearly Change
        LastRow2 = Cells(Rows.Count, 10).End(xlUp).Row
        Range("K2:K" & LastRow2).Style = "percent"
        For y = 2 To LastRow2
            If Cells(y, 10) > 0 Then
                Cells(y, 10).Interior.ColorIndex = 4
            Else
                Cells(y, 10).Interior.ColorIndex = 3
            End If
            
        'Challenge: % Increase and Decrease Table
            MaxChange = Application.WorksheetFunction.Max(Range("K:K"))
            If Cells(y, 11) = MaxChange Then
                Range("O2") = Cells(y, 9)
                Range("P2") = Cells(y, 11)
            End If
            
            MinChange = Application.WorksheetFunction.Min(Range("K:K"))
            If Cells(y, 11) = MinChange Then
                Range("O3") = Cells(y, 9)
                Range("P3") = Cells(y, 11)
            End If
            
            BigVol = Application.WorksheetFunction.Max(Range("L:L"))
            If Cells(y, 12) = BigVol Then
                Range("O4") = Cells(y, 9)
                Range("P4") = Cells(y, 12)
            End If
            
        Next y
        
        'Challenge: max and min table headings and formatting
        Range("N2") = "Greatest % Increase"
        Range("N3") = "Greatest % Decrease"
        Range("N4") = "Greatest Total Volume"
        Range("P2:P3").Style = "percent"
        Columns(14).AutoFit

    Next ws
        
End Sub


