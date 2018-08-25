Attribute VB_Name = "Module111"


Sub TSV_Ticker()

Dim WS_Count As Integer
WS_Count = ThisWorkbook.Worksheets.Count
Dim uTickerCount As Integer
Dim rTickerCount As Long

Dim rTicker As String
Dim uTicker As String

Dim Volume As Long
Dim Sum As LongLong
Dim OpenDaily As Double
Dim CloseDaily As Double
Dim openVolume As Long
Dim closeVolume As Long
Dim YearlyVChange As Double




Dim Data As Range
Dim Dates As Range
Dim Opens As Range
Dim Closes As Range
Dim TopResults As Range

Dim PercentChange As Range
Dim TotalVolume As Range
Dim YearlyChange As Range
Dim UniqueTickers As Range
Dim RawTickers As Range

Dim Pmax As Double
Dim Pmin As Double
Dim Vmax As LongLong
Dim Pagebreak As Boolean
Pagebreak = True


''Worksheet Loop
For J = 1 To WS_Count
    ThisWorkbook.Worksheets(J).Activate

        ''Headings
    Range("I1").Value = "Ticker"
    Range("L1").Value = "Total Stock Volume"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("O1").Value = "Ticker"
    Range("P1").Value = "Value"
    Range("N2").Value = "Greatest % Increase"
    Range("N3").Value = "Greatest % Decrease"
    Range("N4").Value = "Greatest Total Volume"
    Set RawTickers = Worksheets(J).Range("a2", Range("a2").End(xlDown))
    Set Dates = Worksheets(J).Range("B2", Range("B2").End(xlDown))
    Set Opens = Worksheets(J).Range("C2", Range("C2").End(xlDown))
    Set Closes = Worksheets(J).Range("F2", Range("F2").End(xlDown))
    

''Populates all Unique Tickers
''Finds Earliest Volume Reading and Last Volume Reading
''Finds Total Volume for that period
''Finds The Change in Volume, and Finds that Change as Percentage of Total Volume
    rTickerCount = Cells(Rows.Count, 1).End(xlUp).Row
    uTickerCount = 1
    For i = 2 To rTickerCount + 2

    
        rTicker = Cells(i, 1).Value
        Volume = Cells(i, 7).Value
        
        If uTicker <> rTicker Then
            
            If uTicker <> "" Then
                YearlyVChange = CloseDaily - OpenDaily
                Cells(uTickerCount, 12).Value = Sum
                Cells(uTickerCount, 10).Value = YearlyVChange
                If OpenDaily <> 0 Then
                    Cells(uTickerCount, 11).Value = (OpenDaily - CloseDaily) / OpenDaily
                End If
        
            End If
           
       '     Range("").Value =
            ''Reset Sum, Ticker
            OpenDaily = Cells(i, 3).Value
            Sum = Volume
            uTicker = rTicker
            uTickerCount = uTickerCount + 1
            uTicker = rTicker
            Cells(uTickerCount, 9).Value = uTicker
            
            Else
                closeVolume = Volume
                CloseDaily = Cells(i, 6).Value
                Sum = Volume + Sum
            
            
            End If
           
        Next i



    Set UniqueTickers = Worksheets(J).Range("I2", Range("I2").End(xlDown))
    Set TotalVolume = Worksheets(J).Range("L2", Range("L2").End(xlDown))
    Set PercentChange = Worksheets(J).Range("k2", Range("k2").End(xlDown))
    Set YearlyChange = Worksheets(J).Range("j2", Range("j2").End(xlDown))
    

    
''Populating Top Change Summary
    Pmax = Application.WorksheetFunction.Max(PercentChange)
    Range("p2").Value = Pmax
    Range("O2").Value = Application.WorksheetFunction.Index(UniqueTickers, Application.WorksheetFunction.Match(Pmax, PercentChange, 0), 1)

    Pmin = Application.WorksheetFunction.Min(PercentChange)
    Range("p3").Value = Pmin
    Range("O3").Value = Application.WorksheetFunction.Index(UniqueTickers, Application.WorksheetFunction.Match(Pmin, PercentChange, 0), 1)
  
    Vmax = WorksheetFunction.Max(TotalVolume)
    Range("p4").Value = Vmax
    Range("O4").Value = Application.WorksheetFunction.Index(UniqueTickers, Application.WorksheetFunction.Match(Vmax, TotalVolume, 0), 1)
        
   

'' Applying Formatting
'.FormatConditions.Delete
PercentChange.NumberFormat = "0.00%"
YearlyChange.NumberFormat = "0.0000" 
Worksheets(J).Range("P2:P3").NumberFormat = "0.00%"
Worksheets(J).Columns("A:P").AutoFit

''Conditional Formatting Yearly Change, If less than 0, Interior Red
    With YearlyChange.FormatConditions _
        .Add(xlCellValue, xlGreater, "0")
        .Interior.ColorIndex = 10

    End With
''Conditional Formatting Yearly Change, If Greater than 0, Interior Green
    With YearlyChange.FormatConditions _
        .Add(xlCellValue, xlLess, "0")
        .Interior.ColorIndex = 9
        
    End With

    If WS_Count - J <> 0 Then
        Pagecount = MsgBox(Str(J) + " Pages complete" + Str(WS_Count - J) + " To go!", vbOKOnly, "Page Complete")
        If Pagebreak = True Then
            Pagebreak = MsgBox("Do you want to break the loop", vbYesNo + vbQuestion, "Page Break")
            If Pagebreak = vbYes Then
            Exit For
            Else
            Pagebreak = False
         End If
    End If
    Else: Pagecount = MsgBox("All sheets completed!")
    End If
     Next J
End Sub
