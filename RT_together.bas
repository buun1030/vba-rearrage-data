Sub together()

    num_id = 2465
    num_sh = Sheets.Count
    
    ReDim num_col_each_sh(1 To num_sh) As Integer
    ReDim num_rep_each_id(1 To num_id, 1 To num_sh) As Integer
    ReDim max_num_rep_each_id(1 To num_id) As Integer
    
    ' Determine number of columns each sheet
    For sh = 1 To num_sh
        With Sheets(sh)
            a = Application.Index(.Range(.Cells(2, 1), .Cells(num_id + 1, 1)), _
                Application.Match(1, .Range(.Cells(2, 9), .Cells(num_id + 1, 9)), 0)).Row
            num_col_each_sh(sh) = .Cells(a, Columns.Count).End(xlToLeft).Column - 9
        End With
    Next sh
    
    ' Determine maximum number of repeated id of all sheets 2570
    For i = 1 To num_id
        For sh = 1 To num_sh
        With Sheets(sh)
            num_rep_each_id(i, sh) = Application.CountIf(.Range("A:A"), i)
        End With
        Next sh
        max_num_rep_each_id(i) = Application.Max(Application.Index(num_rep_each_id, i, 0))
    Next i
    
    ' New sheet
    Sheets.Add(After:=Sheets(num_sh)).Name = "Result"
    With Sheets(1)
        .Range(.Cells(1, 1), .Cells(.Cells(Rows.Count, 1).End(xlUp).Row, 8)).Copy Sheets(num_sh + 1).Range("A1")
    End With
    Sheets(num_sh + 1).Range("A:H").RemoveDuplicates Columns:=1, Header:=xlNo
    
    ' Start New sheet operations
    With Sheets(num_sh + 1)
    
    ' Find missing number and insert row
    For i = 1 To num_id
        If .Cells(i + 1, 1) <> i Then
            .Rows(i + 1).Insert Shift:=xlShiftDown
            .Cells(i + 1, 1) = i
        End If
    Next i
    
    ' Insert rows
    row_add = 0
    For i = 1 To num_id
        If max_num_rep_each_id(i) > 1 Then
            For j = 1 To max_num_rep_each_id(i) - 1
                .Rows(i + j + row_add + 1).Insert Shift:=xlShiftDown
                .Range(.Cells(i + j + row_add + 1, 1), .Cells(i + j + row_add + 1, 8)).Value _
                    = .Range(.Cells(i + row_add + 1, 1), .Cells(i + row_add + 1, 8)).Value
            Next j
            row_add = row_add + max_num_rep_each_id(i) - 1
        End If
    Next i
    
    ' Fill in data
    col_add = 0
    For sh = 1 To num_sh
        row_add_result = 0
        row_add_source = 0
        For i = 1 To num_id
            For j = 1 To num_rep_each_id(i, sh)
                .Range(.Cells(i + j + row_add_result, 9 + col_add), _
                       .Cells(i + j + row_add_result, 9 + col_add + num_col_each_sh(sh))).Value _
                     = Sheets(sh).Range(Sheets(sh).Cells(i + j + row_add_source, 9), _
                       Sheets(sh).Cells(i + j + row_add_source, 9 + num_col_each_sh(sh))).Value
            Next j
            row_add_result = row_add_result + max_num_rep_each_id(i) - 1
            row_add_source = row_add_source + num_rep_each_id(i, sh) - 1
        Next i
        Range(Cells(1, 9 + col_add), Cells(Cells(Rows.Count, 9).End(xlUp).Row, 9 + col_add)).Interior.Color = RGB(255, 255, 60)
        col_add = col_add + num_col_each_sh(sh) + 1
    Next sh
   
    End With
    
End Sub