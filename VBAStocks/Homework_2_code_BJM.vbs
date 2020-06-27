Attribute VB_Name = "Module1"
'REFERENCES
'https://www.excelcampus.com/vba/find-last-row-column-cell/
'https://www.mrexcel.com/board/threads/add-column-headers-in-a-worksheet-using-vba.1078803/
'https://www.bluepecantraining.com/portfolio/excel-vba-macro-to-apply-conditional-formatting-based-on-value/
'https://docs.microsoft.com/en-us/office/vba/api/excel.range.autofit
'https://stackoverflow.com/questions/6854764/formatting-cells-as-percentage
'https://www.extendoffice.com/documents/excel/5333-excel-run-macro-multiple-sheets.html

Sub Move_through_sheets()
    Dim xSh As Worksheet
    Application.ScreenUpdating = False
    For Each xSh In Worksheets
        xSh.Select
        Call WorksheetLoop
    Next
    Application.ScreenUpdating = True
End Sub

Sub WorksheetLoop()

    Dim ws As Integer
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim m As Long
    Dim ticker_1 As String
    Dim ticker_2 As String
    Dim lastrow As Double
    Dim first_open As Double
    Dim last_close As Double
    Dim yearly_change As Double
    Dim proportion_change As Double
    Dim total_volume As Double
    Dim rg As Range
    Dim cond1 As FormatCondition
    Dim proportion_change_2 As Double
    Dim highest_change As Double
    Dim highest_ticker As String
    Dim lowest_change As Double
    Dim lowest_ticker As String
    Dim highest_volume As Double
    Dim highest_volume_2 As Double
    Dim highest_volume_ticker As String
    Dim highest_change_position As String
    Dim lowest_change_position As String
    Dim highest_volume_position As String
    Dim ties As Integer
        
    ticker_1 = A
    ticker_2 = A
    k = 2
    j = 1
        
    [I1:L1] = Split("Ticker, Yearly change, Percent change, Total stock volume", ",")
    [O2] = "Greatest % increase"
    [O3] = "Greatest %  decrease"
    [O4] = "Greatest total volume"
    [P1:Q1] = Split("Ticker, Value", ",")
        
    'Collect and print out data
        lastrow = Cells(Rows.Count, 1).End(xlUp).Row
        first_open = Cells(2, 3)
            For j = 2 To lastrow
                ticker_1 = Cells(j, 1)
                ticker_2 = Cells(j + 1, 1)
         
         'Add up volume, calculate yearly change
                If (ticker_2 = ticker_1) Then
                    total_volume = total_volume + Cells(j, 7)
          
                ElseIf (ticker_2 <> ticker_1) Then
                    total_volume = total_volume + Cells(j, 7)
                    Cells(k, 9).Value = ticker_1
                    last_close = Cells(j, 6)
                    yearly_change = last_close - first_open
           
            'From class discussion: zero starting values represent issues new this year, find first non-zero
                    If (first_open = 0) Then
                        m = j
                        If Cells(m, 1).Value = ticker_1 Then
                            If Cells(m, 3).Value > 0 Then
                                  first_open = Cells(m, 3)
                            m = m + 1
                            End If
                        Else: first_open = 0
                        End If
                    End If
                    
                'Calculate proportion change
                    If (first_open <> 0) Then
                        proportion_change = yearly_change / first_open
                    ElseIf (first_open = 0) And (last_close = 0) Then
                        proportion_change = 0
                    End If
                 
                 'Format results, get ready for next round
                    Cells(k, 10).Value = Round(yearly_change, 2)
                    Cells(k, 11).Value = proportion_change
                    Cells(k, 11).NumberFormat = "0.00%"
                    Cells(k, 12).Value = total_volume
                    total_volume = 0
                End If
        
                If (ticker_2 <> ticker_1) And (j < lastrow) Then
                    first_open = Cells(j + 1, 3)
                    k = k + 1
                End If
            Next j
        
    'Color cells in eye-searing contrast
        Set rg = Range("J2", Range("J2").End(xlDown))
        rg.FormatConditions.Delete
        Set cond1 = rg.FormatConditions.Add(xlCellValue, xlGreater, 0)
        Set cond2 = rg.FormatConditions.Add(xlCellValue, xlLess, 0)
        With cond1
            .Interior.Color = vbGreen
        End With
        With cond2
            .Interior.Color = vbRed
        End With
         
    'Find extremes
        lastrow = Cells(Rows.Count, 9).End(xlUp).Row
        highest_change = Cells(2, 11).Value
        lowest_change = Cells(2, 11).Value
        highest_volume = Cells(2, 12).Value
        
        For j = 2 To (lastrow - 1)
            proportion_change_2 = Cells(j + 1, 11).Value
                If (proportion_change_2 > highest_change) Then
                    highest_change = proportion_change_2
                    highest_ticker = Cells(j + 1, 9).Value
                    highest_change_position = Cells(j + 1, 9)
                ElseIf (proportion_change_2 < lowest_change) Then
                    lowest_change = proportion_change_2
                    lowest_ticker = Cells(j + 1, 9).Value
                    lowest_change_position = Cells(j + 1, 9)
                End If
                highest_volume_2 = Cells(j + 1, 12).Value
                If (highest_volume_2 > highest_volume) Then
                    highest_volume = highest_volume_2
                    highest_volume_ticker = Cells(j + 1, 9).Value
                    highest_volume_position = Cells(j + 1, 9)
                End If
        Next j
         
    'Check for and print out any ties (none found, alas)
        ties = 1
        For j = 2 To lastrow
            If (Cells(j, 11).Value > highest_change) And (Cells(j, 9).Value <> highest_ticker) Then
                Cells(2, 16 + (2 * ties)).Value = Cells(j, 1).Value
                Cells(2, 17 + (2 * ties)).Value = Cells(j, 3).Value
                Cells(2, 17 + (2 * ties)).NumberFormat = "0.00%"
                Cells(1, 16 + (2 * ties)).Value = "Ticker"
                Cells(1, 17 + (2 * ties)).Value = "Value"
                ties = ties + 1
            End If
        Next j
        
        ties = 1
        For j = 2 To lastrow
            If (Cells(j, 11).Value < lowest_change) And (Cells(j, 9).Value <> lowest_ticker) Then
                Cells(3, 16 + (2 * ties)).Value = Cells(j, 1).Value
                Cells(3, 17 + (2 * ties)).Value = Cells(j, 3).Value
                Cells(3, 17 + (2 * ties)).NumberFormat = "0.00%"
                Cells(1, 16 + (2 * ties)).Value = "Ticker"
                Cells(1, 17 + (2 * ties)).Value = "Value"
                ties = ties + 1
            End If
        Next j
        
        ties = 1
        For j = 2 To lastrow
            If (Cells(j, 12).Value > highest_volume) And (Cells(j, 9).Value <> highest_volume_ticker) Then
                Cells(4, 16 + (2 * ties)).Value = Cells(j, 1).Value
                Cells(4, 17 + (2 * ties)).Value = Cells(j, 3).Value
                Cells(4, 17 + (2 * ties)).NumberFormat = "0.00%"
                Cells(1, 16 + (2 * ties)).Value = "Ticker"
                Cells(1, 17 + (2 * ties)).Value = "Value"
                ties = ties + 1
                End If
        Next j
        
'Print out results
        Cells(2, 16).Value = highest_ticker
        Cells(2, 17).Value = highest_change
        Cells(2, 17).NumberFormat = "0.00%"
        Cells(3, 16).Value = lowest_ticker
        Cells(3, 17).Value = lowest_change
        Cells(3, 17).NumberFormat = "0.00%"
        Cells(4, 16).Value = highest_volume_ticker
        Cells(4, 17).Value = highest_volume
        
        Columns("A:Z").AutoFit
    'Next i
End Sub



