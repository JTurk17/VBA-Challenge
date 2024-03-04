Sub Challenge2_Final_Code()
    For Each ws In Worksheets
    
        ' variable to hold the stock name
        Dim stock_name As String
        
        ' variable to hold the total per stock brand
        Dim stock_total As Double
        stock_total = 0
        
        ' variable to hold the location for the stock name and
        ' stock total volume in the summary table
        Dim Sum_Table_Row As Integer
        Sum_Table_Row = 2
        
        ' variable to interate to the last row in the for loop
        lastrow = Cells(Rows.Count, 1).End(xlUp).Row
        
        ' loop through all the stock values
        
        For i = 2 To lastrow
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
                
                ' change the stock name
                stock_name = Cells(i, 1).Value
                
                ' Add to stock total
                stock_total = stock_total + Cells(i, 7).Value
                
                'Put stock name into summary table
                Range("I" & Sum_Table_Row).Value = stock_name
                
                'Put stock total into summary table
                Range("L" & Sum_Table_Row).Value = stock_total
                
                ' move summary table to next row
                Sum_Table_Row = Sum_Table_Row + 1
                
                ' Reset stock total
                stock_total = 0
                
            Else
            
                ' Add to brand total if the stock names are the same
                stock_total = stock_total + Cells(i, 7).Value
                
            End If
            
        Next i
           
        ' Reset the summary table row
        Sum_Table_Row = 2
        
        ' Variable to hold the first date value of the stock value
        Dim first_date_value As Double
        
        ' Variable to hold the end date value of the stock value
        Dim end_date_value As Double
        
        ' initialize the first date value to be the value of the first row
        first_date_value = Cells(2, 3).Value
        
        ' Variable to hold the yearly change of a stock
        Dim year_change As Double
        
        ' Variable to hold the percent change
        Dim percent_change As Double
        
        ' loop through the stock names to see if they are the same
        For i = 2 To lastrow
            
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
                
                ' Set the end date value
                end_date_value = Cells(i, 6).Value
                
                ' calculate the year change
                year_change = end_date_value - first_date_value
                
                ' input the year change
                Range("J" & Sum_Table_Row).Value = year_change
                
                ' Find the percent change of the first date and end date value
                percent_change = (end_date_value - first_date_value) / first_date_value
                
                ' input percent change
                Range("K" & Sum_Table_Row).Value = percent_change
                
                ' Move summary table to next row
                Sum_Table_Row = Sum_Table_Row + 1
                
                ' Change the first date value to the new stock
                first_date_value = Cells(i + 1, 3).Value
                
            End If
            
        Next i
        
        ' look for the last row in column J
        lastrow = Cells(Rows.Count, 10).End(xlUp).Row
        
        For i = 2 To lastrow
            
            If Cells(i, 10).Value < 0 Then
                
                ' Turn the cell red if the value is negative
                Cells(i, 10).Interior.ColorIndex = 3
                
            ElseIf Cells(i, 10).Value > 0 Then
                ' Turn the cell green if positive
                Cells(i, 10).Interior.ColorIndex = 4
                
            
            Else
                
                ' Turn the cell white if equal to 0
                Cells(i, 10).Interior.ColorIndex = 0
            
            End If
            
        Next i
                
        ' look for the last row in column K
        lastrow = Cells(Rows.Count, 11).End(xlUp).Row
        
        For i = 2 To lastrow
            
            Cells(i, 11).NumberFormat = "0.00%"
    
        Next i
               
        ' variables to hold max volume, max positive change, and max negative change
        Dim max_volume, max_pos_change, max_neg_change As Double
        
        max_volume = Application.WorksheetFunction.Max(Range("L2:L3001"))
        Range("Q4").Value = max_volume
        
        For i = 2 To 3001
            If Cells(i, 12).Value = max_volume Then
                stock_name = Cells(i, 9).Value
            End If
        Next i
        
        Range("P4").Value = stock_name
        
        max_pos_change = Application.WorksheetFunction.Max(Range("K2:K3001"))
        Range("Q2").Value = max_pos_change
        Range("Q2").NumberFormat = "0.00%"
        
        For i = 2 To 3001
            If Cells(i, 11).Value = max_pos_change Then
                stock_name = Cells(i, 9).Value
            End If
        Next i
        
        Range("P2").Value = stock_name
        
        max_neg_change = Application.WorksheetFunction.Min(Range("K2:K3001"))
        Range("Q3").Value = max_neg_change
        Range("Q3").NumberFormat = "0.00%"
        
        For i = 2 To 3001
            If Cells(i, 11).Value = max_neg_change Then
                stock_name = Cells(i, 9).Value
            End If
        Next i
        
        Range("P3").Value = stock_name
        MsgBox ("Just finished one")
    Next ws
    MsgBox ("All Complete")
End Sub