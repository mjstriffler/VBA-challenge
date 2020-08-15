Attribute VB_Name = "Module1"
Sub stockanalysis_test()

' 1. Create a script that will loop through all the stocks for one year and output the following information.
'       The ticker symbol.

'       Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.

'       The percent change from opening price at the beginning of a given year to the closing price at the end of that year.

'       The total stock volume of the stock.

' 2. You should also have conditional formatting that will highlight positive change in green and negative change in red.

'Define variables
      
        Dim ticker_symbol As String
        Dim total_volume As Double
         ticker_volume = 0
        Dim summary_table_row As Integer
        summary_table_row = 2
        Dim date_open As Double
        Dim date_close As Double
        Dim yearly_change As Double
        Dim percent_change As Double
        Dim ws As Worksheet
        For Each ws In Worksheets
        
'Create column labels for the analysis

        Cells(1, 9).Value = "Ticker"
        Cells(1, 10).Value = "Yearly Change"
        Cells(1, 11).Value = "Percent Change"
        Cells(1, 12).Value = "Total Stock Volume"
        
'Create a best/worst performance table

        Cells(2, 15).Value = "Greatest % Increase"
        Cells(3, 15).Value = "Greatest % Decrease"
        Cells(4, 15).Value = "Greatest Total Volume"
        Cells(1, 16).Value = "Ticker"
        Cells(1, 17).Value = "Value"
 
' Determine the Last Row
        
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        For i = 2 To lastrow

' Add To Ticker Total Volume
            
            total_volume = total_volume + ws.Cells(i, 7).Value
            
' Check If We Are Still Within The Same Ticker Name If It Is Not...
            
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

' Set Ticker Name
                
                ticker_symbol = ws.Cells(i, 1).Value
                
 ' Print The Ticker Name In The Summary Table
                
                ws.Range("I" & summary_table_row).Value = ticker_symbol
                
' Print The Ticker Total Amount To The Summary Table
                
                ws.Range("L" & summary_table_row).Value = total_volume
                
' Reset Ticker Total
                total_volume = 0
 
' Set Yearly Open, Yearly Close and Yearly Change Name
                
                Dim YearlyOpen As Double
                Dim YearlyClose As Double
                Dim YearlyChange As Double
                Dim PreviousAmount As Long
                     PreviousAmount = 2
                
                YearlyOpen = ws.Range("C" & PreviousAmount)
                YearlyClose = ws.Range("F" & i)
                YearlyChange = YearlyClose - YearlyOpen
                ws.Range("J" & summary_table_row).Value = YearlyChange
 
 ' Determine Percent Change
                
                If YearlyOpen = 0 Then
                    
                    PercentChange = 0
                
                Else
                    
                    YearlyOpen = ws.Range("C" & PreviousAmount)
                    PercentChange = YearlyChange / YearlyOpen
                
                End If
                
 ' Format Double To Include % Symbol And Two Decimal Places
                
                ws.Range("K" & summary_table_row).NumberFormat = "0.00%"
                ws.Range("K" & summary_table_row).Value = PercentChange

 ' Conditional Formatting Highlight Positive (Green) / Negative (Red)
                
                If ws.Range("J" & summary_table_row).Value >= 0 Then
                    ws.Range("J" & summary_table_row).Interior.ColorIndex = 4
                
                Else
                    
                    ws.Range("J" & summary_table_row).Interior.ColorIndex = 3
                
                End If
            
 ' Add One To The Summary Table Row
                summary_table_row = summary_table_row + 1
                PreviousAmount = i + 1
                
                End If
            
            Next i


' Greatest % Increase, Greatest % Decrease and Greatest Total Volume
            
            lastrow = ws.Cells(Rows.Count, 11).End(xlUp).Row
        
' Start Loop For Final Results
            
            For i = 2 To lastrow
                If ws.Range("K" & i).Value > ws.Range("Q2").Value Then
                    ws.Range("Q2").Value = ws.Range("K" & i).Value
                    ws.Range("P2").Value = ws.Range("I" & i).Value
                
                End If

                If ws.Range("K" & i).Value < ws.Range("Q3").Value Then
                    ws.Range("Q3").Value = ws.Range("K" & i).Value
                    ws.Range("P3").Value = ws.Range("I" & i).Value
                
                End If

                If ws.Range("L" & i).Value > ws.Range("Q4").Value Then
                    ws.Range("Q4").Value = ws.Range("L" & i).Value
                    ws.Range("P4").Value = ws.Range("I" & i).Value
                End If

            Next i
            
' Format Double To Include % Symbol And Two Decimal Places
            
            ws.Range("Q2").NumberFormat = "0.00%"
            ws.Range("Q3").NumberFormat = "0.00%"
            
 ' Format Table Columns To Auto Fit
 
        ws.Columns("I:Q").AutoFit

    Next ws

End Sub

