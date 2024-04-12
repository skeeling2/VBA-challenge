Sub AlphabeticalListing()

' Make It Apply To Data in Whole Workbook

For Each ws In Worksheets
LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

' Define All the Variables

Dim Ticker As String
Dim YearlyChange As Double
Dim PercentChange As Double
Dim Volume As Long
Dim TotalStockVolume As Long
Dim OutputTableRow As Integer
Dim OpenPrice As Double
Dim FirstVolume As Long

'Label New Columns

ws.Range("I1") = "Ticker"
ws.Range("J1") = "Yearly Change"
ws.Range("K1") = "Percent Change"
ws.Range("L1") = "Total Stock Volume"
ws.Range("O2") = "Greatest % Increase"
ws.Range("O3") = "Greatest % Decrease"
ws.Range("O4") = "Greatest Total Volume"
ws.Range("P1") = "Ticker"
ws.Range("Q1") = "Value"

'Tell It Where Results Go

OutputTableRow = 2

'Start Values

TotalStockVolume = 0
YearlyChange = 0
PercentChange = 0
FirstVolume = 0
OpenPrice = ws.Cells(2, 3).Value


'Start Looping
    'If this worked properly it would loop "For i = 2 To LastRow" but that breaks my computer

For i = 2 To 10
 
    
'Check Ticker Values Against Each Other
 
         If ws.Cells(i + 1, 1).Value = ws.Cells(i, 1).Value Then
         
'Direct to Numbers
          
        Ticker = ws.Cells(i, 1).Value
        Volume = ws.Cells(i, 7).Value
        ClosePrice = ws.Cells(i, 6).Value

'Volume Math
   
        TotalStockVolume = TotalStockVolume + Volume
        

'Price Math: These formulas would correctly calcaulte the desired numbers if I could figure out how to pull in OpenPrice

        'YearlyChange = ClosePrice - OpenPrice
        'PercentChange = YearlyChange / OpenPrice
        
'Output
            ws.Range("I" & OutputTableRow).Value = Ticker
            ws.Range("J" & OutputTableRow).Value = YearlyChange
            ws.Range("J" & OutputTableRow).Style = "Currency"
            ws.Range("K" & OutputTableRow).Value = PercentChange
            ws.Range("K" & OutputTableRow).Style = "Percent"
            ws.Range("L" & OutputTableRow).Value = TotalStockVolume
            
        
 'When Ticker Values Are Not The Same
                 
        ElseIf ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
'Go To The Next Row In The Table
        
         OutputTableRow = OutputTableRow + 1
         Ticker = ws.Cells(i + 1, 1).Value
         
         
'Make Sure to Reset Amounts
        
         Volume = 0
         YearlyChange = 0
         PercentChange = 0
         TotalStockVolume = 0

'Pull  in Values From Row Where Ticker Changes
          'OpenPrice = ws.Cells(i - 1, 3).Value
          'FirstVolume = ws.Cells(i - 1, 7).Value
          
          
                       
        
'Output

         ws.Range("I" & OutputTableRow).Value = Ticker
         ws.Range("J" & OutputTableRow).Value = YearlyChange
         ws.Range("J" & OutputTableRow).Style = "Currency"
         ws.Range("K" & OutputTableRow).Value = PercentChange
         ws.Range("K" & OutputTableRow).Style = "Percent"
         ws.Range("L" & OutputTableRow).Value = TotalStockVolume
         
       
        
        End If
           
    Next i
    
'Formatting
For j = 2 To LastRow
    If ws.Cells(j, 10).Value > 0 Then
        ws.Cells(j, 10).Interior.ColorIndex = 4
    ElseIf ws.Cells(j, 10).Value < 0 Then
        ws.Cells(j, 10).Interior.ColorIndex = 3
    Else
        ws.Cells(j, 10).Interior.ColorIndex = 0
    End If

Next j

For k = 2 To LastRow
    If ws.Cells(k, 11).Value > 0 Then
        ws.Cells(k, 11).Interior.ColorIndex = 4
    ElseIf ws.Cells(k, 11).Value < 0 Then
        ws.Cells(k, 11).Interior.ColorIndex = 3
    Else
        ws.Cells(k, 11).Interior.ColorIndex = 0
    End If

Next k

'Greatest Change Table

ws.Range("Q2") = WorksheetFunction.Max(ws.Range("K2:K" & LastRow))
ws.Range("Q2").Style = "Percent"
ws.Range("Q3") = WorksheetFunction.Min(ws.Range("K2:K" & LastRow))
ws.Range("Q3").Style = "Percent"
ws.Range("Q4") = WorksheetFunction.Max(ws.Range("L2:L" & LastRow))



Next ws

End Sub

'If Cells(i + 1, 1).Value = CurrentTicker Then (Started with =. Was always off after the first ticker by 1 row. Try starting with not equal?)
'Open and Close Price Based on Date Choice 1 (just working on first 3 lines to see if it works)- Is pulling both from B3
            'Dim DateRange As Range
            'Set DateRange = Range("B2:B4")
            'Dim EarliestDate As Long
            'Dim LatestDate As Long
            'Dim MinValue As Range
            'Dim MaxValue As Range
            'Dim OpenPrice As Double
            'Dim ClosePrice As Double
            
            'EarliestDate = WorksheetFunction.Min(Range("B2:B4"))
            'LatestDate = WorksheetFunction.Max(Range("B2:B4"))
            'Set MinValue = Range("B2:B4").Find(What:=EarliestDate)
            'Set MaxValue = Range("B2:B4").Find(What:=LatestDate)
            'OpenPrice = MinValue.Offset(, 1).Value
            'ClosePrice = MaxValue.Offset(, 4).Value
'Open and Close Price Based on Date Choice 2 - Keeps going to the row prior
            'If Cells(i, 2).Value < Cells(i + 1, 2) Then
            'OpenPrice = Cells(i, 3).Value
            'End If
            
            'If Cells(i + 1, 2).Value > Cells(i, 2) Then
            'ClosePrice = Cells(i + 1, 6).Value
            'End If

'Open and Close Price Based on Date Choice 3 - Finds earliest date but doesn't pull correct open price
            'For j = 2 To 4
            'MinValue = Cells(i, 2).Value
            'If Cells(j, 2).Value < MinValue Then
            'MinValue = Cells(j, 2).Value
            'End If
            'Next j
            
            
            'If Cells(i + 1, 2).Value > Cells(i, 2) Then
            'ClosePrice = Cells(i + 1, 6).Value
            'End If
         
            'MsgBox (MinValue)
            'MsgBox (OpenPrice)
            

