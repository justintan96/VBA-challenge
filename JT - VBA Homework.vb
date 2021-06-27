Sub Homework2():

' set initial variable to hold ticker number
Dim ticker As String
' set initial variable to hold total stock volume, change percentage change and worksheets
Dim total As LongLong
Dim ws As Worksheet
Dim Lastrow As Long
Dim change As Double
Dim percent_change As Double
Dim Summary_Table_Row As Integer
Dim year_open As Double
Dim year_close As Double
Dim Start As Integer

'Run through each worksheet andreetting my variables to 0

For Each ws In Worksheets
ws.Activate

MsgBox (ws.Name)

Summary_Table_Row = 2

' making sure total is initially 0

total = 0
Start = 2

'change = 0 no need to set to 0 as it cause error


' putting header in each worksehet

    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"

' Setting last row

Lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

' Loop through all ticker numbers

For i = Start To Lastrow

' Check if we are still within the same credit card brand, if it is not...

If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

'  set the Ticker number

ticker = ws.Cells(i, 1).Value

' add to ticker volume

total = total + ws.Cells(i, 7).Value

' set change

year_open = ws.Cells(Start, 3).Value
year_close = ws.Cells(i, 6).Value
change = year_close - year_open
percent_change = change / ws.Cells(Start, 3)

' Print the ticker number onto the summary

ws.Cells(Summary_Table_Row, 9).Value = ticker

ws.Cells(Summary_Table_Row, 12).Value = total

ws.Cells(Summary_Table_Row, 10).Value = change

ws.Cells(Summary_Table_Row, 11).Value = percent_change

'Add one to the summary table row

Summary_Table_Row = Summary_Table_Row + 1

'Reset total volume

    total = 0
    change = 0
    percent_change = 0

' If the cell immediately following a row is the sam brand

    Else

total = total + ws.Cells(i, 7).Value

    End If

Next i

' Conditional formatting the change cells for colours

For i = Start To Lastrow

If ws.Cells(i, 11).Value > 0 Then
    ws.Cells(i, 11).Interior.ColorIndex = 4
    
ElseIf ws.Cells(i, 11).Value < 0 Then
    ws.Cells(i, 11).Interior.ColorIndex = 3

End If

Next i

For i = Start To Lastrow

If ws.Cells(i, 10).Value > 0 Then
    ws.Cells(i, 10).Interior.ColorIndex = 4
    
ElseIf ws.Cells(i, 10).Value < 0 Then
    ws.Cells(i, 10).Interior.ColorIndex = 3
    
End If

Next i

Next ws

End Sub