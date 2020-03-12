Attribute VB_Name = "ArrayFunctions"
' These functions are designed to perform advanced operations on VBA Arrays

Function DISPLAY_ARRAY(ByVal array_data, Optional delimiter = " ", Optional prefix = "", Optional suffix = "")
    If VarType(array_data) <> 8204 Then
        DISPLAY_ARRAY = array_data
        Exit Function
    End If
    
    DISPLAY_ARRAY = prefix
    arr_count = 0

    If ARRAY_DIMENSIONS(array_data) = 1 Then
        For a = LBound(array_data) To UBound(array_data)
            If array_data(a) <> "" Then
                If arr_count > 0 Then DISPLAY_ARRAY = DISPLAY_ARRAY & delimiter
                DISPLAY_ARRAY = DISPLAY_ARRAY & array_data(a)
                arr_count = arr_count + 1
            End If
        Next
    End If
    
    If ARRAY_DIMENSIONS(array_data) > 2 Then
        temp_array = array_data
        array_data = ARRAY_MAKEMATRIX(temp_array)
    End If
    
    If ARRAY_DIMENSIONS(array_data) = 2 Then
        For a = LBound(array_data, 1) To UBound(array_data, 1)
            For b = LBound(array_data, 2) To UBound(array_data, 2)
                If array_data(a, b) <> "" Then
                    If arr_count > 0 Then DISPLAY_ARRAY = DISPLAY_ARRAY & delimiter
                    DISPLAY_ARRAY = DISPLAY_ARRAY & array_data(a, b)
                    arr_count = arr_count + 1
                End If
            Next
            If a < UBound(array_data, 1) Then DISPLAY_ARRAY = DISPLAY_ARRAY & Chr(10)
        Next
    End If
        
    DISPLAY_ARRAY = DISPLAY_ARRAY & suffix
End Function

Sub testdisp()
    Dim test_array()
    ReDim test_array(1 To 4, 1 To 3, 1 To 2)
    test_array(1, 1, 1) = "A"
    test_array(2, 1, 1) = "B"
    test_array(3, 1, 1) = "C"
    test_array(4, 1, 1) = "D"
    
    test_array(1, 2, 1) = 1
    test_array(2, 2, 1) = 2
    test_array(3, 2, 1) = 3
    test_array(4, 2, 1) = 4
    
    test_array(1, 3, 1) = " One"
    test_array(2, 3, 1) = " Two"
    test_array(3, 3, 1) = " Three"
    test_array(4, 3, 1) = " Four"
    
    test_array(1, 1, 2) = "B"
    test_array(2, 1, 2) = "A"
    test_array(3, 1, 2) = "D"
    test_array(4, 1, 2) = "C"
    
    test_array(1, 2, 2) = 2
    test_array(2, 2, 2) = 1
    test_array(3, 2, 2) = 4
    test_array(4, 2, 2) = 3
    
    test_array(1, 3, 2) = " Two"
    test_array(2, 3, 2) = " One"
    test_array(3, 3, 2) = " Four"
    test_array(4, 3, 2) = " Three"
    
    MsgBox DISPLAY_ARRAY(test_array)
End Sub

Function ARRAY_MIN(ByRef temp_array)
    ARRAY_MIN = temp_array(1)
    For a = 1 To UBound(temp_array)
        If temp_array(a) < ARRAY_MIN Then ARRAY_MIN = temp_array(a)
    Next
End Function

Function ARRAY_MAX(ByRef temp_array)
    ARRAY_MAX = temp_array(1)
    For a = 1 To UBound(temp_array)
        If temp_array(a) > ARRAY_MAX Then ARRAY_MAX = temp_array(a)
    Next
End Function

Function ARRAY_POS(ByRef temp_array, n, target)
' Looks in the array (temp_array), for instance n of the target
' It increases the target by 1 if it runs out of things to look for

    instances = 0
    Do While target <= ARRAY_MAX(temp_array)
        For a = LBound(temp_array) To UBound(temp_array)
            ''Debug.Print "Looking for " & target & ", loop " & a & ", value in array: " & temp_array(a)
            If temp_array(a) = target Then
                instances = instances + 1
                If instances = n Then
                    ARRAY_POS = a
                    Exit Function
                End If
            End If
        Next
        target = target + 1
    Loop
        
End Function

Function ARRAY_INSTANCES(ByRef temp_array, target)
' Looks in the array, and returns the number of instances of a target

    ARRAY_INSTANCES = 0
    For a = LBound(temp_array) To UBound(temp_array)
        If temp_array(a) = target Then ARRAY_INSTANCES = ARRAY_INSTANCES + 1
    Next
End Function


Function ARRAY_RND(ByRef temp_array, target, ByRef temp_array2, wnum)
' Returns a random instance of the target
' It increments the target by 1 if an instance is not found

    num_instances = 0
    Do While num_instances = 0
        For a = 1 To UBound(temp_array, 1)
            If temp_array(a) = target And temp_array2(a) <> wnum Then num_instances = num_instances + 1
        Next
        If num_instances = 0 Then target = target + 1
    Loop
    
    n = rand_between(1, num_instances)
       
    Ctr = 0
    For a = 1 To UBound(temp_array, 1)
        If temp_array(a) = target And temp_array2(a) <> wnum Then Ctr = Ctr + 1
        If Ctr = n Then
            ARRAY_RND = a
            Exit Function
        End If
    Next

    ' ARRAY_RND = ARRAY_POS(temp_array, n, target)
End Function
Function ARRAY_RND_PREV(ByRef temp_array1, target, ByRef temp_array2, wnum)
'ARRAY_RND_PREV(chosen_array, 0, lastwave_array, wave_num - 1)

' Returns a random instance of the target
' Only returns people who participated in the previous wave

    num_instances = 0
    Do While num_instances = 0 And target <= ARRAY_MAX(temp_array1)
        For a = 1 To UBound(temp_array1, 1)
            If temp_array2(a) = wnum And temp_array1(a) = target Then num_instances = num_instances + 1
        Next
        If num_instances = 0 Then target = target + 1
    Loop
    
    If num_instances = 0 Then
        Debug.Print "No instances found"
        Exit Function
    End If
    
    n = rand_between(1, num_instances)
    countr = 0
    
    For a = 1 To UBound(temp_array1, 1)
        If temp_array2(a) = wnum And temp_array1(a) = target Then countr = countr + 1
        If countr = n Then
            Debug.Print "Instance found at " & a
            ARRAY_RND_PREV = a
            Exit Function
        End If
    Next
End Function

Sub show_array(temp_array)
    Debugging.TextBox1 = ""
    
    If ARRAY_DIMENSIONS(temp_array) = 2 Then
        For a = LBound(temp_array, 1) To UBound(temp_array, 1)
            For b = LBound(temp_array, 2) To UBound(temp_array, 2)
                Debugging.TextBox1 = Debugging.TextBox1 & " " & temp_array(a, b)
            Next
            Debugging.TextBox1 = Debugging.TextBox1 & " " & Chr(10)
        Next
    ElseIf ARRAY_DIMENSIONS(temp_array) = 1 Then
        For a = LBound(temp_array, 1) To UBound(temp_array, 1)
            Debugging.TextBox1 = Debugging.TextBox1 & " " & temp_array(a)
        Next
    End If
    Debugging.Show
End Sub

Function ARRAY_DIMENSIONS(temp_array)
    On Error GoTo leavehere
    For a = 1 To 10
        b = UBound(temp_array, a)
    Next
    
leavehere:
    ARRAY_DIMENSIONS = a - 1
    On Error Resume Next
    
End Function

Function ARRAY_ROW(temp_array, row_num)
    'Debug.Print "Devising temporary array, elements: " & UBound(temp_array, 1) - LBound(temp_array, 1) + 1
    Dim temp_array2()
    ReDim temp_array2(LBound(temp_array, 2) To UBound(temp_array, 2))

    For a = LBound(temp_array, 2) To UBound(temp_array, 2)
        'Debug.Print "Loop: " & a & " of " & UBound(temp_array, 2)
        temp_array2(a) = temp_array(row_num, a)
    Next
    
    ARRAY_ROW = temp_array2
End Function

Function ARRAY_COLUMN(temp_array, column_num)
    'Debug.Print "Devising temporary array, elements: " & UBound(temp_array, 1) - LBound(temp_array, 1) + 1
    Dim temp_array2()
    ReDim temp_array2(LBound(temp_array, 1) To UBound(temp_array, 1))

    For a = LBound(temp_array, 1) To UBound(temp_array, 1)
        'Debug.Print "Loop: " & a & " of " & UBound(temp_array, 1)
        temp_array2(a) = temp_array(a, column_num)
    Next
    
    ARRAY_COLUMN = temp_array2
End Function

Function ARRAY_SUBSET(temp_array, y1, x1, y2, x2, Optional cols_rows = True)
    ' Returns a subset of an array, denoted by x1 x2 y1 y2
    ' Default is to go across columns and then rows
    
    If y1 > y2 Then
        temp = y1
        y1 = y2
        y2 = temp
    End If
    If x1 > x2 Then
        temp = x1
        x1 = x2
        x2 = temp
    End If
    
    If y1 < LBound(temp_array, 1) Then y1 = LBound(temp_array, 1)
    If x1 < LBound(temp_array, 2) Then x1 = LBound(temp_array, 2)
    If y2 > UBound(temp_array, 1) Then y2 = UBound(temp_array, 1)
    If x2 > UBound(temp_array, 2) Then x2 = UBound(temp_array, 2)

    'Debug.Print "Devising temporary array, elements: " & UBound(temp_array, 1) - LBound(temp_array, 1) + 1
    Dim temp_array2()
    ReDim temp_array2(1 To (y2 - y1 + 1) * (x2 - x1 + 1))

    a = 1
    For y = y1 To y2
        For x = x1 To x2
            temp_array2(a) = temp_array(y, x)
            a = a + 1
        Next
    Next
    
    ARRAY_SUBSET = temp_array2
End Function

Function ARRAY_SUBXY(temp_array, x_vector, y_vector)
    Dim new_array()
    x_len = COUNT_STRINGVECTOR(x_vector)
    y_len = COUNT_STRINGVECTOR(y_vector)
    ReDim new_array(1 To y_len, 1 To x_len)

    For y = 1 To y_len
        y_ref = CInt(GET_STRINGVECTOR(y_vector, y))
        For x = 1 To x_len
            x_ref = CInt(GET_STRINGVECTOR(x_vector, x))
            new_array(y, x) = temp_array(y_ref, x_ref)
        Next
    Next
    
    ARRAY_SUBXY = new_array
End Function

Function ARRAY_FLATTEN(temp_array, Optional miss_value = "")
' Flattens the array (i.e., puts all data into a single vector)
    
    ' Exit the function if the array is already flat
    ' Debug.Print "Array dimensions: " & ARRAY_DIMENSIONS(temp_array)
    If ARRAY_DIMENSIONS(temp_array) < 2 Then
        ARRAY_FLATTEN = temp_array
        Exit Function
    End If
    
    Dim temp_array2()
    ReDim temp_array2(1 To (UBound(temp_array, 1) - LBound(temp_array, 1) + 1) * (UBound(temp_array, 2) - LBound(temp_array, 2) + 1))

    cnt = 1
    miss_cnt = 0
    For a = LBound(temp_array, 1) To UBound(temp_array, 1)
        For b = LBound(temp_array, 2) To UBound(temp_array, 2)
            If temp_array(a, b) <> miss_value Then
                temp_array2(cnt) = temp_array(a, b)
                cnt = cnt + 1
            Else
                miss_cnt = miss_cnt + 1
            End If
        Next
    Next
    
    If miss_cnt > 0 Then ReDim Preserve temp_array2(1 To UBound(temp_array2) - miss_cnt)
    ARRAY_FLATTEN = temp_array2
End Function

Function ARRAY_MAKEMATRIX(temp_array)
' Turns the provided array (whether a vector or a 3+ dimensional array, into a 2-dimensional matrix)

    Dim new_array()
    Debug.Print "Converting array to matrix..."
    
    num_dim = ARRAY_DIMENSIONS(temp_array)
    Debug.Print "Array has"; num_dim; "dimensions"
    
    If num_dim = 0 Then
        ReDim new_array(1 To 1, 1 To 1)
        new_array(1, 1) = temp_array
    ElseIf num_dim = 1 Then
        ReDim new_array(1 To UBound(temp_array), 1)
        For a = 1 To UBound(temp_array)
            new_array(a, 1) = temp_array(a)
        Next
    ElseIf num_dim = 2 Then
        new_array = temp_array
    Else
        ReDim new_array(1 To UBound(temp_array, 1), 1 To UBound(temp_array, 2))
        For y = 1 To UBound(temp_array, 1)
            For x = 1 To UBound(temp_array, 2)
                new_array(y, x) = temp_array(y, x, 1)
            Next
        Next
    End If
    
    ARRAY_MAKEMATRIX = new_array
End Function

Function ARRAY_COUNT(temp_array, Optional miss_value = "")
' Returns the number of values in the array; skips the missing value
    temp_array = ARRAY_FLATTEN(temp_array, miss_value)
   
    Dim cnt: cnt = 0

    'Debug.Print "Variable type: " & VarType(temp_array)
    'Debug.Print "Number of dimensions: " & ARRAY_DIMENSIONS(temp_array)
    For a = LBound(temp_array) To UBound(temp_array)
        b = temp_array(a)
        If IsNumeric(b) And b <> "" Then
            If b <> miss_value Then cnt = cnt + 1
        End If
    Next
    
    ARRAY_COUNT = cnt
End Function

Function ARRAY_MEAN(temp_array, Optional miss_value = "")
' Returns the mean of the array; skips the missing value
    If IsNumeric(miss_value = False) Then miss_value = CInt(miss_value)
    temp_array = ARRAY_FLATTEN(temp_array, miss_value)
    
    Dim sum: sum = 0
    Dim cnt: cnt = 0

    For a = LBound(temp_array) To UBound(temp_array)
        b = temp_array(a)
        If IsNumeric(b) And b <> "" Then
            If b <> miss_value Then
                sum = sum + b
                cnt = cnt + 1
            End If
        End If
    Next
    
    If cnt = 0 Then Exit Function
    ARRAY_MEAN = sum / cnt
End Function

Function ARRAY_VARIANCE(temp_array, Optional miss_value = "")
' Returns the variance of the array; skips the missing value
    temp_array = ARRAY_FLATTEN(temp_array, miss_value)
    
    Dim sos: sos = 0
    Dim cnt: cnt = 0
    Dim Mean: Mean = ARRAY_MEAN(temp_array, miss_value)
    
    For a = LBound(temp_array) To UBound(temp_array)
        b = temp_array(a)
        If IsNumeric(b) And b <> "" Then
            If b <> miss_value Then
                sos = sos + (b - Mean) ^ 2
                cnt = cnt + 1
            End If
        End If
    Next
    
    ARRAY_VARIANCE = sos / (cnt - 1)
End Function

Function ARRAY_STDEV(temp_array, Optional miss_value = "")
    ARRAY_STDEV = Sqr(ARRAY_VARIANCE(temp_array, miss_value))
End Function

Function ARRAY_DELETEROW(temp_array, row_num)
    Dim temp_array2()
    ReDim temp_array2(LBound(temp_array, 1) To UBound(temp_array, 1) - 1, LBound(temp_array, 2) To UBound(temp_array, 2))

    y2 = LBound(temp_array, 1)
    For y = LBound(temp_array, 1) To UBound(temp_array, 1)
        If y <> row_num Then
            For x = LBound(temp_array, 2) To UBound(temp_array, 2)
                temp_array2(y2, x) = temp_array(y, x)
            Next
            y2 = y2 + 1
        End If
    Next

    ARRAY_DELETEROW = temp_array2
End Function
Function ARRAY_DELETECOLUMN(temp_array, column_num)
    Dim temp_array2()
    ReDim temp_array2(LBound(temp_array, 1) To UBound(temp_array, 1), LBound(temp_array, 2) To UBound(temp_array, 2) - 1)

    x2 = LBound(temp_array, 2)
    For x = LBound(temp_array, 2) To UBound(temp_array, 2)
        If x <> column_num Then
            For y = LBound(temp_array, 1) To UBound(temp_array, 1)
                temp_array2(y, x2) = temp_array(y, x)
            Next
            x2 = x2 + 1
        End If
    Next

    ARRAY_DELETECOLUMN = temp_array2
End Function

Function ARRAY_DELIMIT(temp_array, Optional miss_value = "", Optional delimiter = ",")
    temp_array = ARRAY_FLATTEN(temp_array, miss_value)
    Dim temp_string As String: temp_string = ""
    
    For a = LBound(temp_array) To UBound(temp_array)
        If a > LBound(temp_array) Then temp_string = temp_string & delimiter
        temp_string = temp_string & temp_array(a)
    Next
    
    ARRAY_DELIMIT = temp_string
End Function

Function ARRAY_CLEAN(temp_array, Optional miss_value = "")
    d = ARRAY_DIMENSIONS(temp_array)
    
    If d = 1 Then
        For x = LBound(temp_array, 1) To UBound(temp_array, 1)
            If temp_array(x) = "" Or IsNumeric(temp_array(x)) = False Then temp_array(x) = miss_value
        Next
    End If
    
    If d = 2 Then
        For y = LBound(temp_array, 1) To UBound(temp_array, 1)
            For x = LBound(temp_array, 2) To UBound(temp_array, 2)
                If temp_array(y, x) = "" Or IsNumeric(temp_array(y, x)) = False Then temp_array(y, x) = miss_value
            Next
        Next
    End If
    ARRAY_CLEAN = temp_array
End Function

Function ARRAY_POPULATE(ByVal temp_array, Optional pop_value = "")
    ' Populates an entire array with pop_value
    
    ' Check to see how many dimensions the array has
    d = ARRAY_DIMENSIONS(temp_array)
    
    If d = 1 Then
        For x = LBound(temp_array, 1) To UBound(temp_array, 1)
            temp_array(x) = pop_value
        Next
    End If
    
    If d = 2 Then
        For y = LBound(temp_array, 1) To UBound(temp_array, 1)
            For x = LBound(temp_array, 2) To UBound(temp_array, 2)
                temp_array(y, x) = pop_value
            Next
        Next
    End If
  
    If d = 3 Then
        For y = LBound(temp_array, 1) To UBound(temp_array, 1)
            For x = LBound(temp_array, 2) To UBound(temp_array, 2)
                For w = LBound(temp_array, 3) To UBound(temp_array, 3)
                    temp_array(y, x, w) = pop_value
                Next
            Next
        Next
    End If
    
    If d = 4 Then
        For y = LBound(temp_array, 1) To UBound(temp_array, 1)
            For x = LBound(temp_array, 2) To UBound(temp_array, 2)
                For w = LBound(temp_array, 3) To UBound(temp_array, 3)
                    For v = LBound(temp_array, 4) To UBound(temp_array, 4)
                        temp_array(y, x, w, v) = pop_value
                    Next
                Next
            Next
        Next
    End If
    
    ARRAY_POPULATE = temp_array
End Function

Sub ARRAY_FILTER(ByRef temp_array, mdata)
    For y = LBound(temp_array, 1) To UBound(temp_array, 1)
        For x = LBound(temp_array, 2) To UBound(temp_array, 2)
            
        Next
    Next
End Sub

Sub test_asc()
    Dim test_array()
    ReDim test_array(1 To 4, 1 To 3)
    test_array(1, 1) = "B"
    test_array(2, 1) = "A"
    test_array(3, 1) = "D"
    test_array(4, 1) = "C"
    
    test_array(1, 2) = 2
    test_array(2, 2) = 1
    test_array(3, 2) = 4
    test_array(4, 2) = 3
    
    test_array(1, 3) = " Two"
    test_array(2, 3) = " One"
    test_array(3, 3) = " Four"
    test_array(4, 3) = " Three"
    
    test_array2 = ARRAY_SORT_COLUMN(test_array, 2, True, False)
    
    MsgBox test_array2(1, 1) & test_array2(1, 2) & test_array2(1, 3)
    MsgBox test_array2(2, 1) & test_array2(2, 2) & test_array2(2, 3)
    MsgBox test_array2(3, 1) & test_array2(3, 2) & test_array2(3, 3)
    MsgBox test_array2(4, 1) & test_array2(4, 2) & test_array2(4, 3)
    
End Sub

Function ARRAY_SORT_COLUMN(temp_array, col_num, Optional Ascending = False, Optional absolute = False)
    n_row = UBound(temp_array, 1) - LBound(temp_array, 1) + 1
    n_col = UBound(temp_array, 2) - LBound(temp_array, 2) + 1
    
    Dim count_array(): ReDim count_array(LBound(temp_array, 1) To UBound(temp_array, 1))
    Dim target_array(): ReDim target_array(LBound(temp_array, 1) To UBound(temp_array, 1), LBound(temp_array, 2) To UBound(temp_array, 2))
    
    count_array = ARRAY_POPULATE(count_array, 0)
    
    temp_array2 = ARRAY_COLUMN(temp_array, col_num)
    
    ' Create pivot
    For a = 1 To UBound(count_array, 1)
        highest_value = ""
        highest_entry = ""
        For b = LBound(temp_array2) To UBound(temp_array2)
            current_value = temp_array2(b)
            If highest_value = "" Then
                If current_value <> "NA" Then
                    If absolute = True Then current_value = Abs(current_value)
                    highest_value = current_value
                    highest_entry = b
                End If
            Else
                If current_value <> "NA" Then
                    If absolute = True Then current_value = Abs(current_value)
                    If current_value > highest_value Then
                        highest_value = current_value
                        highest_entry = b
                    End If
                End If
            End If
        Next
        count_array(a) = highest_entry
        temp_array2(highest_entry) = "NA"
    Next
    
    ' Populate target array
    For a = 1 To UBound(count_array)
        For b = LBound(target_array, 2) To UBound(target_array, 2)
            target = count_array(a)
            If Ascending = True Then target = n_row - target + 1
            target_array(a, b) = temp_array(target, b)
        Next
    Next
        
    ARRAY_SORT_COLUMN = target_array
End Function

Function COLUMN_COUNT(ByRef temp_array, col_num, Optional target = 1)
' Searches the nominated column for a target, and returns the number of instances
    COLUMN_COUNT = 0

    For a = 1 To UBound(temp_array, 1)
        If temp_array(a, col_num) = target Then COLUMN_COUNT = COLUMN_COUNT + 1
    Next
End Function

Function COLUMN_COUNT_OVERLAP(ByRef temp_array, col_num1, col_num2, Optional target = 1)
' Searches the nominated two columns for a target, and returns the number of instances that it appears in both columns
    COLUMN_COUNT_OVERLAP = 0

    For a = 1 To UBound(temp_array, 1)
        If temp_array(a, col_num1) = target And temp_array(a, col_num2) = target Then COLUMN_COUNT_OVERLAP = COLUMN_COUNT_OVERLAP + 1
    Next
End Function

Function ARRAY_TRANSPOSE(ByRef temp_array)
    ' Transposes an array
    
    num_dim = ARRAY_DIMENSIONS(temp_array)
    
    If num_dim = 1 Then
        x_length = 1
    Else
        x_length = UBound(temp_array, 2)
    End If
    y_length = UBound(temp_array, 1)
    
    Dim new_array()
    ReDim new_array(1 To x_length, 1 To y_length)
    
    For y = 1 To x_length
        For x = 1 To y_length
            If num_dim > 1 Then
                new_array(y, x) = temp_array(x, y)
            Else
                new_array(y, x) = temp_array(x)
            End If
        Next
    Next
    
    ARRAY_TRANSPOSE = new_array
End Function

Function ARRAY_ALLONES(y_elements, x_elements)
    Dim temp_array()
    ReDim temp_array(1 To y_elements, 1 To x_elements)
    
    ARRAY_ALLONES = ARRAY_POPULATE(temp_array, 1)
End Function

Function ARRAY_MULTIPLY(ByVal arr1, ByVal arr2, Optional formula = True)
    Dim new_array()
    
    Debug.Print "Converting matrices"
    array1 = ARRAY_MAKEMATRIX(arr1)
    array2 = ARRAY_MAKEMATRIX(arr2)

    Debug.Print "Determining dimensions"
    a1_y_len = UBound(array1, 1)
    a1_x_len = UBound(array1, 2)
    
    a2_y_len = UBound(array2, 1)
    a2_x_len = UBound(array2, 2)
    
    If a1_x_len <> a2_y_len Then
        ARRAY_MULTIPLY = "NA"
        Exit Function
    End If

    ReDim new_array(1 To a1_y_len, 1 To a2_x_len)

    For x = 1 To a2_x_len ' Move across new array column
        For y = 1 To a1_y_len ' Move across new array rows
            For C = 1 To a1_x_len ' Populate the cell
                Debug.Print "X:"; x; "of"; a2_x_len; "Y:"; y; "of"; a1_y_len; "Cell entry: "; C; "of"; a1_x_len
                If formula = True Then
                    Debug.Print "Providing formula as output"
                    If C = 1 Then new_array(y, x) = "SUM("
                    new_array(y, x) = new_array(y, x) & CStr(array1(y, C)) & "*" & CStr(array2(C, x))
                    If C < a1_x_len Then
                        new_array(y, x) = new_array(y, x) & ","
                    Else
                        new_array(y, x) = new_array(y, x) & ")"
                    End If
                Else
                    If C = 1 Then new_array(y, x) = 0
                    new_array(y, x) = new_array(y, x) + array1(y, C) * array2(C, x)
                End If
            Next
        Next
    Next
    
    If UBound(new_array, 1) = 1 And UBound(new_array, 2) = 1 Then
        ARRAY_MULTIPLY = new_array(1, 1)
    Else
        ARRAY_MULTIPLY = new_array
    End If
End Function
Sub test_arraymul()
    Dim temp0
    ReDim temp0(1 To 1, 1 To 2)
    temp0(1, 1) = 1
    temp0(1, 2) = 1
    
    MsgBox VarType(temp0)

    Dim temp1
    ReDim temp1(1 To 2, 1 To 3)
    temp1(1, 1) = 0.246
    temp1(1, 2) = 0.246
    temp1(1, 3) = 0.246
    temp1(2, 1) = 0.252
    temp1(2, 2) = 0.252
    temp1(2, 3) = 0.252
    
    Dim temp2
    ReDim temp2(1 To 3, 1 To 1)
    temp2(1, 1) = 1
    temp2(2, 1) = 1
    temp2(3, 1) = 1
    
    temp1 = ARRAY_MULTIPLY(temp0, temp1, False)
    temp2 = ARRAY_MULTIPLY(temp1, temp2, True)
    MsgBox DISPLAY_ARRAY(temp1)
    MsgBox DISPLAY_ARRAY(temp2)
End Sub


Sub test_trans()
    Dim temp()
    ReDim temp(1 To 3, 1 To 3)
    temp(1, 1) = 1
    temp(2, 1) = 2
    temp(3, 1) = 3
    temp(1, 2) = 4
    temp(2, 2) = 5
    temp(3, 2) = 6
    temp(1, 3) = 7
    temp(2, 3) = 8
    temp(3, 3) = 9
    
    Dim temp2
    ReDim temp2(1 To 3)
    temp2(1) = 1
    temp2(2) = 2
    temp2(3) = 3
    
    temp4 = ARRAY_TRANSPOSE(temp2)
        
    temp3 = ARRAY_MULTIPLY(temp4, temp, False)

    MsgBox DISPLAY_ARRAY(temp3)
End Sub

' Required functions
' Array_row - returns all the values in the row of an array
' Array_column - returns all of the values in the column of an array
' Array_mean - returns the mean of the array, not counting a missing value
' Array_SD - returns the variance of the array, not counting a missing value
' Array_count - returns the number of values in the array, not counting a missing value

Sub testnewarray()
    Dim a()
    ReDim a(1 To 3, 1 To 3)
    a(1, 1) = 0
    a(1, 2) = 1
    a(1, 3) = 2
    a(2, 1) = 3
    a(2, 2) = 4
    a(2, 3) = 5
    a(3, 1) = 6
    a(3, 2) = 7
    a(3, 3) = 8

    'Call show_array(array_row(a, 2))
    'Call show_array(array_column(a, 1))
    'Call show_array(array_subset(a, 2, 2, 3, 3))
    
    'Call show_array(a)
    Call show_array(ARRAY_DELETECOLUMN(a, 3))
    MsgBox "Mean: " & ARRAY_MEAN(a, 999) & ", SD: " & ARRAY_SD(a, 999) & ", Count: " & ARRAY_COUNT(a, 999)
End Sub

Sub testxy()
    Dim a()
    ReDim a(1 To 3, 1 To 3)
    a(1, 1) = 0
    a(1, 2) = 1
    a(1, 3) = 2
    a(2, 1) = 3
    a(2, 2) = 4
    a(2, 3) = 5
    a(3, 1) = 6
    a(3, 2) = 7
    a(3, 3) = 8
    
    new_ar = ARRAY_SUBXY(a, SET_STRINGVECTOR(1, 2, 3), SET_STRINGVECTOR(1, 2, 3))
    MsgBox DISPLAY_ARRAY(new_ar)
End Sub


Function UNIQUE(data_range, Optional unique_case = 0)
' Identifies a unique case from a list
' Leaving "unique_case" blank shows the number of unique cases
' Works either from a worksheet range, or 1-dimensional VBA Array

    Dim temp_array()
    If TypeName(data_range) = "Range" Then
        temp_data = data_range.Value
        n_rows = UBound(temp_data, 1)
        n_cols = UBound(temp_data, 2)
        ReDim temp_array(1 To n_rows * n_cols)
    ElseIf TypeName(data_range) = "Variant()" Then
        temp_data = ARRAY_FLATTEN(data_range)
        n_rows = UBound(temp_data, 1)
        n_cols = 0
        ReDim temp_array(1 To n_rows)
    Else
        temp_data = data_range
    End If
    
    C = 0
    
    If n_cols > 0 Then
        For a = 1 To n_rows
            For b = 1 To n_cols
                If temp_data(a, b) <> "" Then
                    If C > 0 Then
                        found_match = False
                        For d = 1 To C
                            If temp_array(d) = temp_data(a, b) Then
                                found_match = True
                                Exit For
                            End If
                        Next
                        If found_match = False Then
                            temp_array(C + 1) = temp_data(a, b)
                            C = C + 1
                        End If
                    Else
                        temp_array(1) = temp_data(a, b)
                        C = 1
                    End If
                End If
            Next
        Next
    Else
        For a = 1 To n_rows
            If temp_data(a) <> "" Then
                If C > 0 Then
                    found_match = False
                    For d = 1 To C
                        If temp_array(d) = temp_data(a) Then
                            found_match = True
                            Exit For
                        End If
                    Next
                    If found_match = False Then
                        temp_array(C + 1) = temp_data(a)
                        C = C + 1
                    End If
                Else
                    temp_array(1) = temp_data(a)
                    C = 1
                End If
            End If
        Next
    End If
    
    If unique_case = 0 Then
        UNIQUE = C
    Else
        UNIQUE = temp_array(unique_case)
    End If
End Function


Function SET_STRINGVECTOR(ParamArray vals())
' Creates a vector stored as a string
    SET_STRINGVECTOR = "{"
    For a = LBound(vals) To UBound(vals)
        SET_STRINGVECTOR = SET_STRINGVECTOR & vals(a)
        If a < UBound(vals) Then SET_STRINGVECTOR = SET_STRINGVECTOR & ","
    Next
    SET_STRINGVECTOR = SET_STRINGVECTOR & "}"
End Function

Function GET_STRINGVECTOR(ByVal temp_array, Optional element = -1)
' Gets an element from a string vector (or returns the whole thing as an array)
' If element is a number, it returns the numbered element
' If element is a string, then it returns the value based on a dictionary
    Dim temp_val

    ' Set the value of GET_STRINGVECTOR to NA by default
    temp_val = "NA"

    temp_array = Trim(temp_array)
    temp_array = Replace(temp_array, "{", "")
    temp_array = Replace(temp_array, "}", "")
    temp_array = Split(temp_array, ",")
    
    If IsNumeric(element) = True Then
        If element > -1 Then
            If element - 1 <= UBound(temp_array) Then
                temp_str = temp_array(element - 1)
                If InStr(1, temp_str, ":") > 0 Then
                    temp_str = Split(temp_str, ":")
                    temp_str = temp_str(1)
                End If
                temp_val = temp_str
            End If
        Else
            temp_val = temp_array
        End If
    Else
        For a = LBound(temp_array) To UBound(temp_array)
            temp_dict = Split(temp_array(a), ":")
            If UBound(temp_dict) > 0 And element = temp_dict(0) Then temp_val = temp_dict(1)
        Next
    End If
    
    On Error GoTo catcherror
    GET_STRINGVECTOR = CDbl(temp_val)
    Exit Function
    
catcherror:
    GET_STRINGVECTOR = temp_val
    On Error GoTo 0
End Function

Function ADD_STRINGVECTOR(ByVal temp_array, val, Optional element = -1, Optional overwrite = False)
' Adds an element to a string vector (by default at the end, or as specified in another place)
' The function creates blank entries if the element is specified beyond the end of the vector
    'MsgBox temp_array & " " & val & " " & element

    temp_array2 = GET_STRINGVECTOR(temp_array)

    ' If the element is a dictionary entry, find if a previous instance exists
    If IsNumeric(element) = False Then
        overwrite = True
        val = element & ":" & val
        For a = LBound(temp_array2) To UBound(temp_array2)
            temp_dict = Split(temp_array2(a), ":")
            If UBound(temp_dict) > 0 And element = temp_dict(0) Then
                element = a + 1
                Exit For
            End If
        Next
        If IsNumeric(element) = False Then element = UBound(temp_array2) + 2
    End If

    UB = UBound(temp_array2)
    If UB < 0 Then UB = 0
    commas = ""
    If element - 1 > UB Then
        For a = 1 To element - 1 - UB
            commas = commas & ","
        Next
    End If
    If Len(temp_array) < 1 And element <= 1 Then
        ADD_STRINGVECTOR = val ' "{" & val & "}"
        Exit Function
    End If
    temp_array = "{"
    For a = LBound(temp_array2) To UBound(temp_array2)
        If a + 1 = element Then
            temp_array = temp_array & val
            If overwrite = False Then temp_array = temp_array & "," & temp_array2(a)
        Else
            temp_array = temp_array & temp_array2(a)
        End If
        If a < UBound(temp_array2) Then temp_array = temp_array & ","
    Next
    If element = -1 Then temp_array = temp_array & "," & val
    If Len(commas) > 0 Then temp_array = temp_array & commas & val
    ADD_STRINGVECTOR = temp_array & "}"
End Function

Function JOIN_STRINGVECTOR(vector1, vector2)
    temp_array = Trim(vector2)
    temp_array = Replace(temp_array, "{", "")
    temp_array = Replace(temp_array, "}", "")
    temp_array = Split(temp_array, ",")
    
    new_vector = vector1
    
    For a = LBound(temp_array) To UBound(temp_array)
        temp_text = temp_array(a)
        'MsgBox a & " " & temp_text
        If InStr(1, temp_text, ":") > 0 Then
            temp_text = Split(temp_text, ":")
            new_vector = ADD_STRINGVECTOR(new_vector, temp_text(1), temp_text(0))
        Else
            new_vector = ADD_STRINGVECTOR(new_vector, temp_text)
        End If
    Next
    
    JOIN_STRINGVECTOR = new_vector
End Function

Function COUNT_STRINGVECTOR(ByVal temp_array)
    COUNT_STRINGVECTOR = UBound(GET_STRINGVECTOR(temp_array)) + 1
End Function

Function EXPAND_STRINGVECTOR(temp_array, Optional sep = "", Optional prefix = "", Optional suffix = "")
    EXPAND_STRINGVECTOR = ""
    For a = 1 To COUNT_STRINGVECTOR(temp_array)
        EXPAND_STRINGVECTOR = EXPAND_STRINGVECTOR & prefix & GET_STRINGVECTOR(temp_array, a) & suffix
        If a < COUNT_STRINGVECTOR(temp_array) Then EXPAND_STRINGVECTOR = EXPAND_STRINGVECTOR & sep
    Next
End Function

Function POSVAL_STRINGVECTOR(ByVal temp_array, val)
' Searches the array for a value (val), and then returns its position

    temp_array2 = GET_STRINGVECTOR(temp_array)
    
    POSVAL_STRINGVECTOR = 0
    For a = 1 To UBound(temp_array2)
        If temp_array2(a) = val Then
            POSVAL_STRINGVECTOR = a
            Exit Function
        End If
    Next
End Function

Function EMPTY_STRINGVECTOR(Optional length = 1, Optional val = "NA")
' Creates an empty string vector of "length" entries
    temp = SET_STRINGVECTOR(val)
    If length > 2 Then
        For a = 2 To length
            temp = ADD_STRINGVECTOR(temp, val)
        Next
    End If
    EMPTY_STRINGVECTOR = temp
End Function
