VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub multiple_year()

'Creating Dims

Dim ws As String
Dim ticker As String
Dim vol As Integer
Dim year_open As Double
Dim year_close As Double
Dim yearly_change As Double
Dim percent_change As Double
Dim Summary_Table_Row As Integer
Dim Top_Ticker As String
Dim Top_yearly As Integer
Dim Top_percent_change As Double

'Step 1. Ticker Set
'Stop at 0

On Error Resume Next

'run through each worksheet
'Worksheets

For Each ws In Worksheets

'set headers
ws.Cells(1, 9).Value = "Ticker"
    
    Summary_Table_Row = 2

    'loop
        For I = 2 To ws.UsedRange.Rows.Count
             If ws.Cells(I + 1, 1).Value <> ws.Cells(I, 1).Value Then
                
    'find all the values
            ticker = ws.Cells(I, 1).Value
            
    'insert values into summary
            ws.Cells(Summary_Table_Row, 9).Value = ticker
            
            vol = 0
        End If

'Step 2. Yearly Change

'Stop at 0

On Error Resume Next

'run through each worksheet
For Each ws In Worksheets

    'set headers
    ws.Cells(1, 10).Value = "Yearly Change"
    
    Summary_Table_Row = 2

    'loop
        For I = 2 To ws.UsedRange.Rows.Count
             If ws.Cells(I + 1, 1).Value <> ws.Cells(I, 1).Value Then
                
    'find all the values
            year_open = ws.Cells(I, 3).Value
            year_close = ws.Cells(I, 6).Value

            yearly_change = year_close - year_open
            
    'insert values into summary
            ws.Cells(Summary_Table_Row, 10).Value = yearly_change
            
            vol = 0
        End If

'Step 3 Percent Change

'stop at 0
On Error Resume Next

'run through each worksheet

For Each ws In Worksheets

    'set headers
    ws.Cells(1, 11).Value = "Percent Change"

    'setup integers for loop
    Summary_Table_Row = 2

    'loop
        For I = 2 To ws.UsedRange.Rows.Count
             If ws.Cells(I + 1, 1).Value <> ws.Cells(I, 1).Value Then
            
            'find all the values
            year_open = ws.Cells(I, 3).Value
            year_close = ws.Cells(I, 6).Value

            percent_change = (year_close - year_open) / year_close

            'insert values into summary
            ws.Cells(Summary_Table_Row, 11).Value = percent_change
            Summary_Table_Row = Summary_Table_Row + 1

             vol = 0
            
        End If

'Step 4. Loop Volume

'Stop at 0

On Error Resume Next

'run through each worksheet
For Each ws In Worksheets

    'set headers
    ws.Cells(1, 9).Value = "Volume"
    
    Summary_Table_Row = 2

    'loop
        For I = 2 To ws.UsedRange.Rows.Count
             If ws.Cells(I + 1, 1).Value <> ws.Cells(I, 1).Value Then
                
    'find all the values
            vol = ws.Cells(I, 1).Value
            
    'insert values into summary
            ws.Cells(Summary_Table_Row, 7).Value = vol
            
            vol = 0
        End If


ws.Columns("K").NumberFormat = "0.00%"

'Step 5 Conditonal

'stop at 0
On Error Resume Next

'run through each worksheet

For Each ws In Worksheets

   Summary_Table_Row = 2
   
    If ws.Cells(I + 1, 10).Value > 0 Then
                ws.Cells(I + 1, 10).Interior.ColorIndex = 4
        Else
                ws.Cells((I+1, 10)).Interior.ColorIndex = 3
        End If
                ticker = 0
        
'move to next worksheet
Next ws

End Sub

'Step 6 Bonus

Dim Top_Ticker As String
Dim Top_yearly As Integer
Dim Top_percent_change As Double

'stop at 0
On Error Resume Next

'run through each worksheet

For Each ws In Worksheets

    'set headers
        ws.Cells(1, 15).Value = "Top_Ticker"
        ws.Cells(1, 16).Value = "Top_yearly"
        ws.Cells(1, 17).Value = "Top_percent_change"

    Summary_Table_Row = 2

    'loop
        For I = 2 To ws.UsedRange.Rows.Count
             If ws.Cells(I + 1, 1).Value > ws.Cells(I, 1).Value Then
                
    'find all the values
            Top_Ticker = ws.Cells(I, 1).Value
            
    'insert values into summary
            ws.Cells(Summary_Table_Row, 15).Value = Top_Ticker
            ws.Cells(Summary_Table_Row, 15).Value = Top_yearly
            ws.Cells(Summary_Table_Row, 15).Value = Top_percent_change
            vol = 0
        End If

ws.Columns("K").NumberFormat = "0.00%"

Next I

End Sub







