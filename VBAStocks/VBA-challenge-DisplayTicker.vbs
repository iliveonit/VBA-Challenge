Private Sub ShowTicker()
Dim WhichYear As Integer
Dim ShowTicker As String
Dim TotalSheets As Integer
Dim rng As Range, s As String, x As String

TotalSheets = Worksheets.Count
 ShowTicker = Worksheets("MasterSheet").Cells(2, 3).Value

' Loop to Check in all Sheets
For i = 1 to TotalSheets
  If Worksheets(i).Name <> "MasterSheet" Then
     lastrow=Worksheets(i).Cells(Rows.Count, 1).End(xlUp).Row

' Start from 2nd Row in each sheet
     For j = 2 To lastrow
         Set rng = Range("A1")
         ShowTicker = "The Ticker symbol is "
         x = ShowTicker & rng.Value
         MsgBox x
         Range("C2") = x
     Next j
   End If
   Next i
End Sub
