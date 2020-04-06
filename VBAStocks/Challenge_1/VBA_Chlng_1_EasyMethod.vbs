Sub ShowTicker()

    ' --------------------------------------
    'Better Method to handle ALL Worksheets
    Dim ws As Worksheet
    
    ' --------------------------------------
    'Basic Declaration
    'Dim lastRow As Integer
    'lastRow = 10
    'MsgBox (lastRow)
    
    ' --------------------------------------
    'HW2 Declarations
    'Dim i As Integer  'Row Counter for all Tickets in Column A; used with lastRow
    'Dim j As Integer  'Row Counter to write Tickter in Column I and Column J
    j = 0
    
    'Need to show below for HW2 per image
    Dim Ticker As String
    
    Dim Yearly_Change As Integer
    Yearly_Change = 0
    
    Dim Percentage_Change As Integer
    Percentage_Change = 0
    
    Dim TotalVol As Double
    
    ' --------------------------------------
    'Declared for my Calc
    Dim OpenPrice As Integer
    OpenPrice = 0
    
    Dim ClosePrice As Integer
    ClosePrice = 0
    
    Dim Greatest_PctIncr As Integer
    Greatest_PctIncr = 0
    
    Dim Greatest_PctDecr As Integer
    Greatest_PctDecr = 0
     
    ' --------------------------------------
    ' My Tests...
    'Ticker = Range("I2").Value
    'Total_Stock_Volume = Range("J2").Value
    
    'Check Ticker Column'For i = 2 To 263
    'i = 4
    'MsgBox i
    
    ' --------------------------------------
    ' Build Loop to populate for Column I & Column J
    
    
    For Each ws In Worksheets
        TotalVol = 0
        j = 0
    
        ws.Range("I1").Value = "ShowTicker"
        ws.Range("J1").Value = "Total Stock Volume"

        lastRow = Cells(Rows.Count, "A").End(xlUp).Row
    
        For i = 2 To lastRow

            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

                'Start as new Ticker for calc to write Ticker+Volume
                TotalVol = TotalVol + ws.Cells(i, 7).Value

                ws.Range("I" & 2 + j).Value = ws.Cells(i, 1).Value
                ws.Range("J" & 2 + j).Value = TotalVol

                TotalVol = 0
    
                'Increments the Column I and Column J counter (j) for new Ticker
                j = j + 1

            Else
                'Keep adding Volume for matched Ticker symbol
                TotalVol = TotalVol + ws.Cells(i, 7).Value

            End If

        Next i

        'Reset Total Stock Volume Counter to 0
        TotalVol = 0
        j = 0

    Next ws

End Sub
