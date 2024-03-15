Attribute VB_Name = "Module1"
Sub stock()
' Set an initial variable for holding the brand name
  


  'loop through each worksheet
  For Each ws In Worksheets
    Dim ticker As String

    Dim total_volume As Double
    total_volume = 0

    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2
  
    Dim yearly_change As Double
  
    Dim percent_change As Double
  
    Dim open_stock As Double
  
    Dim close_stock As Double
  
    last_row = ws.Cells(Rows.Count, 1).End(xlUp).Row
    open_stock = ws.Cells(2, 3).Value
  
    'loop through each row to last row of each sheet
    For i = 2 To last_row
  

    ' Check if we are still within the same ticker type, if it is not...
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            ' Set the ticker type
            ticker = ws.Cells(i, 1).Value

      ' Add to the total volume
            total_volume = total_volume + ws.Cells(i, 7).Value
        
        ' yearly change
        
            close_stock = ws.Cells(i, 6).Value
        
            yearly_change = close_stock - open_stock
            
            'loop for yearly change
            If yearly_change <> 0 Then
                
                percent_change = (yearly_change / open_stock) * 100
                            
            Else
                percent_change = 0
            End If
        
        ' Print the ticker type in the Summary Table
            ws.Range("I" & Summary_Table_Row).Value = ticker
        ' Print the total volume to the Summary Table
            ws.Range("L" & Summary_Table_Row).Value = total_volume
        
        'yearly change column in summary table
            ws.Range("J" & Summary_Table_Row).Value = yearly_change
        
        'percent change column in summary table
            ws.Range("K" & Summary_Table_Row).Value = percent_change
        

      ' Add one to the summary table row
            Summary_Table_Row = Summary_Table_Row + 1
      
      ' Reset the Total volume
            total_volume = 0
        
     'reset opening stock price
            open_stock = ws.Cells(i + 1, 3)
        
    ' If the cell immediately following a row is the same ticker..
        Else
            total_volume = total_volume + ws.Cells(i, 7).Value
            
        End If
    
    Next i

  
   
   Dim last_sum_row As Integer
   last_sum_row = Cells(Rows.Count, 11).End(xlUp).Row
  
  For i = 2 To last_sum_row
  'conditional formatting
    If ws.Cells(i, 10).Value And ws.Cells(i, 11).Value > 0 Then
        ws.Cells(i, 10).Interior.ColorIndex = 4
        ws.Cells(i, 11).Interior.ColorIndex = 4
    Else
        ws.Cells(i, 10).Interior.ColorIndex = 3
        ws.Cells(i, 11).Interior.ColorIndex = 3
    End If

  Dim max_per_inc As Double
  Dim min_per_dec As Double
  
  max_per_inc = Application.WorksheetFunction.Max(Range("K2:K" & last_sum_row))
  Cells(2, 17).Values = max_per_inc
  
  max_total_volume = Application.WorksheetFunction.Max(Range("L2:L" & last_sum_row))
  Cells(4, 17).Values = max_total_volume
  Cells(4, 16).Values = Cells(, 9).Values
  
  max_per_dec = Application.WorksheetFunction.Min(Range("K2:K" & last_sum_row))
  Cells(3, 17).Values = max_per_dec
  
  
  Next i
Next ws

  
End Sub
