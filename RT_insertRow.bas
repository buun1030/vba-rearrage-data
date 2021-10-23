Sub insertRow()
    
    For sh = 1 To Sheets.Count
    
    With Sheets(sh)
        
        num_id = 2465
        
        ' Find missing number and insert row
        For i = 1 To num_id
            If .Cells(i + 1, 1) <> i Then
                .Rows(i + 1).Insert Shift:=xlShiftDown
                .Cells(i + 1, 1) = i
            End If
        Next i
        
        row_add = 0
        
        'Number of column set to add
        a = Application.Index(.Range(.Cells(2, 1), .Cells(num_id + 1, 1)), _
            Application.Match(1, .Range(.Cells(2, 9), .Cells(num_id + 1, 9)), 0)).Row
        num_col = .Cells(a, Columns.Count).End(xlToLeft).Column - 9
    
    For i = 1 To num_id
        If .Cells(i + row_add + 1, 9).Value > 1 Then
            '---------------------Must choose first!!---------------------
            num = .Cells(i + row_add + 1, 9).Value
            'num = (.Cells(i + row_add + 1, Columns.Count).End(xlToLeft).Column - 9) / num_col
            '---------------------Must choose first!!---------------------
            
            'add row(s) depend on number of appliances minus one
            For j = 1 To num - 1
                .Rows(i + j + row_add + 1).Insert Shift:=xlShiftDown
                
                For k = 1 To num_col
                    .Cells(i + j + row_add + 1, 9 + k).Value _
                        = .Cells(i + row_add + 1, 9 + j * num_col + k).Value
                Next k
                
                .Range(.Cells(i + j + row_add + 1, 1), .Cells(i + j + row_add + 1, 8)).Value _
                        = .Range(.Cells(i + row_add + 1, 1), .Cells(i + row_add + 1, 8)).Value
                
                .Range(.Cells(i + row_add + 1, 10 + j * num_col), _
                       .Cells(i + row_add + 1, 9 + (j + 1) * num_col)).ClearContents
                        
            Next j
            
            row_add = row_add + num - 1
            
        End If
    Next i
    
    End With
    Next sh
    
End Sub