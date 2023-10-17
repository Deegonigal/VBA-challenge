Attribute VB_Name = "Module1"
Sub ticker_code():

'Declare worksheet variab;e
Dim ws As Worksheet
For Each ws In Worksheets

'Setdimensions
Dim stock_date As Date
Dim open_price As Double
Dim closing_price As Double
Dim LR As Double
Dim min_date As Long
Dim volume As Double
Dim tkr_index As Double
Dim tkrRng As String
Dim offsettkrRng As String
Dim max_increase As Double
Dim max_decrease As Double
Dim max_volume As Long


'Assign columns for ticker summary
ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Annual Change"
ws.Cells(1, 11).Value = "Annual Percent Change"
ws.Cells(1, 12).Value = "Annual Total Volume"
ws.Cells(1, 16).Value = "Ticker"
ws.Cells(1, 17).Value = "Value"

'Assign header for min/max summary
ws.Cells(2, 15).Value = "Greatest % Increase"
ws.Cells(3, 15).Value = "Greatest % Decrease"
ws.Cells(4, 15).Value = "Greatest Total Volume"

'Set first row
FR = 2

'Set last row
LR = Cells(Rows.Count, 1).End(xlUp).Row
    'MsgBox LR

'Set data row
data_row = 2

'Set volume
volume = 0

'Set index
tkr_index = 2

        
'Loop through Names and find unique stock name and list in cloumn I
For i = FR + 1 To LR
    tkrRng = ws.Cells(i, 1).Value
    offsettkrRng = ws.Cells(i + 1, 1).Value
    If tkrRng <> offsettkrRng Then
        ws.Cells(data_row, 9).Value = tkrRng
        alphabet = tkrRng
        data_row = data_row + 1
    End If
Next i


'Identify opening and closing price of stock
'Calculate the difference between opening and closing price as dollar amt and percentage change
tkr_index = 2
For i = FR To LR
        tkrRng = ws.Cells(i, 1).Value
        offsettkrRng = ws.Cells(i - 1, 1).Value
        offsetrngclose = ws.Cells(i + 1, 1).Value
        If tkrRng <> offsettkrRng Then
            open_price = ws.Cells(i, 3).Value
        ElseIf tkrRng <> offsetrngclose Then
        closing_price = ws.Cells(i, 6).Value
            Yearly_Change = closing_price - open_price
            ws.Cells(tkr_index, 10) = Yearly_Change
            Percent_change = Yearly_Change / open_price
            ws.Cells(tkr_index, 11) = FormatPercent(Percent_change)
            tkr_index = tkr_index + 1
        End If
Next i

'Calculate total volume
data_row = 2
For i = FR To LR
        tkrRng = ws.Cells(i, 1).Value
        If tkrRng = ws.Cells(i + 1, 1).Value And i > 1 Then
            volume = volume + ws.Cells(i, 7).Value
        ElseIf tkrRng <> ws.Cells(i + 1, 1).Value And i > 1 Then
            volume = volume + ws.Cells(i, 7).Value
            ws.Cells(data_row, 12).Value = volume
            data_row = data_row + 1
            volume = 0
        Else
            volume = volume + Cells(i, 7).Value
        End If
Next i

'Format Annual Price Variance column based on increase or decrease to value
For i = FR To ws.Cells(Rows.Count, 10).End(xlUp).Row
If ws.Cells(i, 10).Value > 0 Then
    ws.Cells(i, 10).Interior.ColorIndex = 4
Else: ws.Cells(i, 10).Interior.ColorIndex = 3
End If
Next i

'Identify greatest % increase, greaetst % decrease and largest volume
maxinc = ws.Application.WorksheetFunction.Max(Range("k:k"))
maxdec = ws.Application.WorksheetFunction.Min(Range("k:k"))
maxvol = ws.Application.WorksheetFunction.Max(Range("l:l"))

'Find last row of ticker summary table
LR_sum = Cells(Rows.Count, 9).End(xlUp).Row

'Loop through ticker and assign associated max and min ticker to summary
For i = FR To LR_sum
If maxinc = Cells(i, 11).Value Then
    ws.Cells(2, 16).Value = Cells(i, 9).Value
ElseIf maxdec = Cells(i, 11).Value Then
    ws.Cells(3, 16).Value = Cells(i, 9).Value
ElseIf maxvol = Cells(i, 12).Value Then
    ws.Cells(4, 16).Value = Cells(i, 9).Value
End If
Next i

'Populate min and max values
ws.Cells(2, 17) = FormatPercent(maxinc)
ws.Cells(3, 17) = FormatPercent(maxdec)
ws.Cells(4, 17) = maxvol

'Execute to next worksheet
Next ws

End Sub
