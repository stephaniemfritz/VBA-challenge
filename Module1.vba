Attribute VB_Name = "Module1"
Sub GetQuarterlyData()
    'Declare variables'
    Dim Wb As Workbook
    Dim CurrentTicker As String
    Dim NextTicker As String
    Dim Volume As LongLong
    Dim CurrentVolume As LongLong
    Dim PercentChange As Double
    Dim QuarterlyChange As Double
    Dim RecordingLine As Long
    Dim OpeningPrice As Double
    Dim ClosingPrice As Double
    Dim i As Long
    Dim Rowcount As Long
    Dim CountTicker As Integer
    Dim GIncrease As Double
    Dim GDecrease As Double
    Dim GVolume As LongLong
    Dim GInreaseTicker As String
    Dim GDecreaseTicker As String
    Dim GVolumeTicker As String
    Dim QuarterlyChangeRange As Range
    'Set Wb equal to the active workbook so that its components can be read and manipulated. This assumes that the active workbook is the one with the stock information to be gathered.'
    Set Wb = Application.ActiveWorkbook
    'Loop through each sheet in the workbook'
    For Each Ws In Wb.Worksheets()
        'Set headers for the summary by ticker.'
        Ws.Range("I1").Value = "Ticker"
        Ws.Range("J1").Value = "Quarterly Change"
        Ws.Range("K1").Value = "Percent Change"
        Ws.Range("L1").Value = "Total Stock Volume"
        'Get the number of rows of daily data for each ticker'
        Rowcount = Ws.Cells(Ws.Rows.Count, "A").End(xlUp).Row
        'Get volume to start with for the first ticker because the volume is set to 0 before starting to collect data for a new ticker.'
        Volume = 0
        'Get the opening price for the first ticker because the opening price for a ticker is stored for later use before starting to collect the rest of the data for a new ticker.'
        OpeningPrice = Cells(2, 3).Value
        'Set the value of the row in which to write summary for each ticker to 2.'
        RecordingLine = 2
        'Loop through the rows that contain the daily data for each ticker.'
        For i = 2 To Rowcount
            CurrentTicker = Ws.Cells(i, 1).Value
            NextTicker = Ws.Cells(i + 1, 1).Value
            CurrentVolume = Ws.Cells(i, 7).Value
            Volume = Volume + CurrentVolume
            If CurrentTicker <> NextTicker Then
                ClosingPrice = Ws.Cells(i, 6).Value
                Ws.Cells(RecordingLine, 9).Value = CurrentTicker
                QuarterlyChange = ClosingPrice - OpeningPrice
                Ws.Cells(RecordingLine, 10).Value = ClosingPrice - OpeningPrice
                Ws.Cells(RecordingLine, 11).Value = (ClosingPrice - OpeningPrice) / OpeningPrice
                Ws.Cells(RecordingLine, 11).NumberFormat = "0.00%"
                Ws.Cells(RecordingLine, 12).Value = Volume
                Volume = 0
                OpeningPrice = Ws.Cells(i + 1, 3).Value
                RecordingLine = RecordingLine + 1
            End If
        Next i
        Set QuarterlyChangeRange = Ws.Range("J2:K" & CStr(Ws.Cells(Ws.Rows.Count, "J").End(xlUp).Row))
        QuarterlyChangeRange.FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, _
            Formula1:="=0"
        QuarterlyChangeRange.FormatConditions(1).Interior.ColorIndex = 3
        QuarterlyChangeRange.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, _
            Formula1:="=0"
        QuarterlyChangeRange.FormatConditions(2).Interior.ColorIndex = 4
        Ws.Cells(2, 15).Value = "Greatest % Increase"
        Ws.Cells(3, 15).Value = "Greatest % Decrease"
        Ws.Cells(4, 15).Value = "Greatest Total Volume"
        GVolume = Ws.Cells(2, 12).Value
        GIncrease = Ws.Cells(2, 11).Value
        GDecrease = Ws.Cells(2, 11).Value
        GVolumeTicker = Ws.Cells(2, 9).Value
        GIncreaseTicker = Ws.Cells(2, 9).Value
        GDecreaseTicker = Ws.Cells(2, 9).Value
        Ws.Cells(1, 16).Value = "Ticker"
        Ws.Cells(1, 17).Value = "Value"
        For i = 2 To RecordingLine
            If Ws.Cells(i, 11) > GIncrease Then
                GIncreaseTicker = Ws.Cells(i, 9).Value
                GIncrease = Ws.Cells(i, 11).Value
            End If
            If Ws.Cells(i, 11) < GDecrease Then
                GDecreaseTicker = Ws.Cells(i, 9).Value
                GDecrease = Ws.Cells(i, 11).Value
            End If
            If Ws.Cells(i, 12) > GVolume Then
                GVolumeTicker = Ws.Cells(i, 9).Value
                GVolume = Ws.Cells(i, 12).Value
            End If
        Next i
        Ws.Cells(2, 16).Value = GIncreaseTicker
        Ws.Cells(2, 17).Value = GIncrease
        Ws.Cells(2, 17).NumberFormat = "0.00%"
        Ws.Cells(3, 16).Value = GDecreaseTicker
        Ws.Cells(3, 17).Value = GDecrease
        Ws.Cells(3, 17).NumberFormat = "0.00%"
        Ws.Cells(4, 16).Value = GVolumeTicker
        Ws.Cells(4, 17).Value = GVolume
    Next Ws
End Sub


