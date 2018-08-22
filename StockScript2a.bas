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
Dim sVolume As Long
Dim sDate As Long
Dim eDate As Long
Dim eVolume As Long
Dim YearlyVChange As Long
Dim Pchange As Long
Dim UtickerOpen As Double
Dim UtickerClass As Double



Dim Data As Range
Dim Opens As Range
Dim Closes As Range
Dim TopResults As Range
Dim Results As Range
Dim PercentChange As Range
Dim TotalVolume As Range
Dim YearlyChange As Range
Dim UniqueTickers As Range
Dim RawTickers As Range

Dim Pmax As Double
Dim Pmin As Double
Dim Vmax As LongLong



''Worksheet Loop
For j = 1 To WS_Count
    ThisWorkbook.Worksheets(j).Activate

        ''Headings
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Total Stock Volume"
    Range("K1").Value = "Yearly Change"
    Range("L1").Value = "Percent Change"
    Range("O1").Value = "Ticker"
    Range("P1").Value = "Value"
    Range("N2").Value = "Greatest % Increase"
    Range("N3").Value = "Greatest % Decrease"
    Range("N4").Value = "Greatest Total Volume"
    Set Dates = Worksheets(j).Range("B2", Range("B2").End(xlDown))
    Set Opens = Worksheets(j).Range("C2", Range("C2").End(xlDown))
    Set Closes = Worksheets(j).Range("F2", Range("F2").End(xlDown))
    Set RawTickers = Worksheets(j).Range("a2", Range("a2").End(xlDown))

''Populates all Unique Tickers
''Finds Earliest Volume Reading and Last Volume Reading
''Finds Total Volume for that period
''Finds The Change in Volume, and Finds that Change as Percentage of Total Volume
    rTickerCount = Cells(Rows.Count, 1).End(xlUp).Row
    uTickerCount = 2
    For i = 2 To rTickerCount + 1

    
        rTicker = Cells(i, 1).Value
        Volume = Cells(i, 7).Value
        
        If uTicker <> rTicker And uTicker <> "" Then
            sDate = Cells(i, 2)
            UtickerOpen = Application.WorksheetFunction.Index(Opens, Application.WorksheetFunction.Match(sDate & uTicker, Dates & RawTickers, 0), 1)
            UtickerClose = Application.WorksheetFunction.Index(Closes, Application.WorksheetFunction.Match(eDate & uTicker, Dates & RawTickers, 0), 1)
            YearlyVChange = eVolume - sVolume
            sVolume = Volume
            Cells(uTickerCount, 10).Value = Sum
            Cells(uTickerCount, 11).Value = YearlyVChange
            Cells(uTickerCount, 12).Value = UtickerOpen / UtickerClose
           
            
           
       '     Range("").Value =
            ''Reset Sum, Ticker
            
            Sum = Volume
            uTicker = rTicker
            uTickerCount = uTickerCount + 1
            Cells(uTickerCount, 9).Value = uTicker
            
            ElseIf uTicker <> rTicker Then
            sVolume = Volume
            uTicker = rTicker
            Cells(uTickerCount, 9).Value = uTicker
            
            Else
            eVolume = Volume
            eDate = Cells(i, 2)
            Sum = Volume + Sum
            
            
        End If
           
        Next i



    Set UniqueTickers = Worksheets(j).Range("I2", Range("I2").End(xlDown))
    Set TotalVolume = Worksheets(j).Range("J2", Range("J2").End(xlDown))
    Set PercentChange = Worksheets(j).Range("L2", Range("L2").End(xlDown))
    Set YearlyChange = Worksheets(j).Range("K2", Range("K2").End(xlDown))
    Set Results = Application.Union(UniqueTickers, TotalVolume, PercentChange, YearlyChange)

    
''Populating Top Change Summary
    Pmax = Application.WorksheetFunction.Max(PercentChange)
    Range("p2").Value = Pmax
    Range("O2").Value = Application.WorksheetFunction.Index(UniqueTickers, Application.WorksheetFunction.Match(Pmax, PercentChange, 0), 1)

    Pmin = WorksheetFunction.Min(PercentChange)
    Range("p3").Value = Pmin
    Range("O3").Value = Application.WorksheetFunction.Index(UniqueTickers, Application.WorksheetFunction.Match(Pmin, PercentChange, 0), 1)
  
    Vmax = WorksheetFunction.Max(TotalVolume)
    Range("p4").Value = Vmax
    Range("O4").Value = Application.WorksheetFunction.Index(UniqueTickers, Application.WorksheetFunction.Match(Vmax, TotalVolume, 0), 1)
        
   

'' Applying Formatting
Results.FormatConditions.Delete
PercentChange.NumberFormat = "0.00%"
Worksheets(j).Columns("A:P").AutoFit

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


   





   
    PageBreak = MsgBox(Str(j) + " Pages complete" + Str(WS_Count - j) + " To go!", vbOKOnly, "Page Complete")
    PageBreak = MsgBox("Do you want to break the loop", vbYesNo + vbQuestion, "Page Break")
    If PageBreak = vbYes Then
        Exit For
    End If
    Next j
End Sub
