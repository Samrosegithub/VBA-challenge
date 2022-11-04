VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

'Other Model

Sub Module2()

'define everything

        Dim ticker As String
        Dim vol As Double
        vol = 0
        Dim Summary_Table_Row As Integer
        Dim year_open As Double
        Dim year_close As Double

'Step 1
        Cells(1, 9).Value = "Ticker"
        Summary_Table_Row = 2

    For I = 2 To 797712

                If year_open = 0 Then
        
                year_open = Cells(I, 3).Value
    End If

        Cells(1, 10).Value = "Yearly Change"
        Cells(1, 12).Value = "Total Stock Vol:"
        Cells(1, 11).Value = "Yearly_percentage"

        Summary_Table_Row = 2

        For I = 2 To 797712

            If year_open = 0 Then
        
               year_open = Cells(I, 3).Value
            End If
    
            If Cells(I - 1, 1) = Cells(I, 1) And Cells(I + 1, 1).Value <> Cells(I, 1).Value Then
                year_close = Cells(I, 6).Value
                
                yearly_change = year_close - year_open
            
                ticker = Cells(I, 1).Value
        
                vol = vol + Cells(I, 7).Value
        
                Range("j" & Summary_Table_Row).Value = yearly_change
        
                Range("I" & Summary_Table_Row).Value = ticker
        
                Range("k" & Summary_Table_Row).Value = year_percent
        
                Range("L" & Sumary_Table_Row).Value = vol
        
                Summary_Table_Row = Summary_Table_Row + 1
                      
                vol = 0
        
          Else
        
                vol = vol + Cells(I, 7).Value
    
          End If
    
        Next I
        
    End Sub
    
        

