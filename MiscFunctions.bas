Attribute VB_Name = "MiscFunctions"
Public num_obs, z_crit, r_crit
Public use_z, use_r As Boolean
'Public MPlusOutput As Object

Function format_mplus(ByVal orig_string)
' This function is designed to convert a variable name into Mplus format (8 characters or less)
    
    ' Clean all non-alphabetic, non-numeric characters, except for underscore
    a = 1
    Do While a <= Len(orig_string)
        b = Asc(Mid(orig_string, a, 1))
        If b <= 31 Or (b >= 33 And b <= 47) Or (b >= 58 And b <= 64) Or (b >= 91 And b <= 94) Or (b = 96) Or (b >= 123) Then
            orig_string = Replace(orig_string, Chr(b), "")
            a = 0
        End If
        a = a + 1
    Loop
    
    ' Mode 1: Simply return the variable name without modification
    If Len(orig_string) <= 8 Then
        orig_string = Replace(orig_string, " ", "_")
        format_mplus = orig_string
        Exit Function
    End If
    
    orig_string = auto_capital(orig_string)
    orig_string = Replace(orig_string, " ", "_")
    
    ' Mode 2: If there 3 or more words, then create a shortened name based on all words
    temp_array = Split(orig_string, "_")
    If UBound(temp_array) >= 2 Then
        Dim temp_array2()
        ReDim temp_array2(LBound(temp_array) To UBound(temp_array))
        cntr = 0
        temp_array2 = ARRAY_POPULATE(temp_array2, "")
        For a = 1 To 8
            For b = LBound(temp_array) To UBound(temp_array)
                temp = Mid(temp_array(b), a, 1)
                If a = 1 Then
                    temp = UCase(temp)
                Else
                    temp = LCase(temp)
                End If
                If cntr < 8 Then
                    temp_array2(b) = temp_array2(b) & temp
                    cntr = cntr + 1
                End If
            Next
        Next
        
        format_mplus = ""
        For a = LBound(temp_array2) To UBound(temp_array2)
            format_mplus = format_mplus & temp_array2(a)
        Next
        Exit Function
    End If
    
    ' Mode 3: Try to reduce the string
    ' Define a variable to count the length
    Dim len_cntr: len_cntr = 0
    Dim end_str
    Dim item_text: item_text = ""

    ' Get the wave number
    For d = 1 To Len(orig_string) - 1
        If Mid(orig_string, Len(orig_string) - d, 1) = "W" And IsNumeric(Mid(orig_string, Len(orig_string) + 1 - d, 1)) = True Then
            item_id = Mid(orig_string, Len(orig_string) - d, 2)
        End If
    Next
    
    ' Get the item number
    If item_id <> "" Then
        a = InStr(1, orig_string, item_id)
    Else
        a = Len(orig_string) + 1
    End If
    temp = ""
    For b = 1 To a - 1
        d = Mid(orig_string, a - b, 1)
        If IsNumeric(d) Then
            temp = d & temp
        Else
            Exit For
        End If
    Next
    item_id = temp & item_id
    item_id = CStr(item_id)
    
    'Get the number of capitalised words
    Dim junk: junk = 0
    If item_id <> "" Then orig_string = Replace(orig_string, item_id, "")
    b = Len(orig_string)
    
    ' Capitalise the first letter if it's not a capital. This is needed to parse the number of capitalised words.
    
    If Left(orig_string, 1) <> UCase(Left(orig_string, 1)) Then orig_string = UCase(Left(orig_string, 1)) & Right(orig_string, Len(orig_string) - 1)
        
    For C = 1 To Len(orig_string) + 1
        If Len(orig_string) > 2 Then
            If Mid(orig_string, Len(orig_string) - C, 1) = "_" Then
                junk = junk + 1
            Else
                Exit For
            End If
        End If
    Next
    
    end_str = Len(orig_string) - junk
    num_wrds = num_caps(orig_string, end_str)
    'Num words based on capitalisation; if 0 then make 1
    If num_wrds = 0 Then num_wrds = 1
    'Debug.Print num_wrds
    num_lr = 8 - Len(item_id)
    'Debug.Print 3 & " " & num_wrds
    num_lpw = Int(num_lr / num_wrds)
    'Debug.Print 4

    If num_wrds > num_lr Then
        ' If the number of words more than the remaining number of letters, just use the first letter from each word
        For a = 1 To num_lr
            item_text = item_text & trunc(word_num(orig_string, a, end_str), 1)
        Next
    Else
        For a = 1 To num_wrds
            If a < num_wrds Then
                item_text = item_text & trunc(word_num(orig_string, a, end_str), num_lpw)
            Else
                item_text = item_text & trunc(word_num(orig_string, a, end_str), num_lr - Len(item_text))
            End If
        Next
    End If
    
    format_mplus = item_text & item_id

End Function

Function num_caps(orig_string, end_str)
' Provides the number of capital letters in a string
    For a = 1 To end_str
        If Mid(orig_string, a, 1) = UCase(Mid(orig_string, a, 1)) Then
            num_caps = num_caps + 1
        ElseIf end_str - a > 1 And Mid(orig_string, a, 1) = "_" Then
            num_caps = num_caps + 1
        End If
    Next
End Function

Function word_num(orig_string, w_num, end_str)
' Returns the w_num'th word in a string.

    word_num = Mid(orig_string, word_pos(orig_string, w_num, end_str), word_pos(orig_string, w_num + 1, end_str) - word_pos(orig_string, w_num, end_str))
End Function

Function word_pos(orig_string, w_num, end_str)
    temp = 0
    For a = 1 To end_str
        If Mid(orig_string, a, 1) = UCase(Mid(orig_string, a, 1)) Then
            temp = temp + 1
        ElseIf end_str - a > 1 And Mid(orig_string, a, 1) = "_" Then
            temp = temp + 1
        End If
        If w_num = temp Then
            word_pos = a
            If Mid(orig_string, a, 1) = "_" Then word_pos = word_pos + 1
            Exit Function
        End If
    Next
    word_pos = end_str + 1
End Function

Function trunc(long_string, num_chars)
    If num_chars < Len(long_string) Then
        trunc = Left(long_string, num_chars)
    Else
        trunc = long_string
    End If
End Function

Function write_comment(long_string)
' Parses a comment; makes it appear over multiple lines if over 90 characters
    write_comment = ""
    
    ' Replace line breaks and carriage feeds with spaces
    long_string = Replace(long_string, Chr(10), " ")
    long_string = Replace(long_string, Chr(13), "")

    rem_length = Len(long_string)
    
    For a = 0 To Int(Len(long_string) / 88)
        write_comment = write_comment & "!"
        If rem_length > 89 Then
            str_length = 88
        Else
            str_length = rem_length
        End If
                
        write_comment = write_comment & Mid(long_string, 1 + a * 88, str_length)
        rem_length = rem_length - str_length
        If rem_length > 0 Then write_comment = write_comment & "+" & Chr(10)
        'If a < Int(Len(long_string) / 88) Then write_comment = write_comment & "+" & Chr(10)
    Next
End Function

Function show_ascii(long_string)
    show_ascii = ""
    For a = 1 To Len(long_string)
        If a > 1 Then show_ascii = show_ascii & " "
        show_ascii = show_ascii & Asc(Mid(long_string, a, 1))
    Next
End Function


Function auto_capital(var_name)
' This function capitalises the first letter of the string
    If var_name = "" Then
        auto_capital = ""
        Exit Function
    End If
    auto_capital = UCase(Left(var_name, 1)) & Right(var_name, Len(var_name) - 1)
End Function

Function clean_correlation(corr As String) As String
    If corr = "" Then Exit Function
    
    temp_corr = Replace(corr, "*", "")
    stars = Replace(corr, temp_corr, "")
    
    temp_corr = WorksheetFunction.Fixed(temp_corr, 2)
    
    clean_correlation = temp_corr & stars
    clean_correlation = Replace(clean_correlation, "0.", ".")
End Function

Public Sub flip_correlations_below()
    cell_range = Selection.Value

    If UBound(cell_range, 1) <> UBound(cell_range, 2) Then
        MsgBox "The number of rows must be the same as the number of columns."
        Exit Sub
    End If

    For a = 1 To UBound(cell_range, 1)
        For b = 1 To UBound(cell_range, 2)
            If a <> b Then
                temp1 = cell_range(a, b)
                temp2 = cell_range(b, a)
                
                If temp1 <> "" And temp2 = "" Then
                    temp = temp1
                ElseIf temp1 = "" And temp2 <> "" Then
                    temp = temp2
                ElseIf temp1 = temp2 Then
                    temp = temp1
                ElseIf (temp1 <> "" And temp2 <> "") And (temp1 <> temp2) Then
                    temp = "Ambiguous: " & temp1 & ", " & temp2
                End If
                
                cell_range(a, b) = temp
                cell_range(b, a) = ""
            End If
        Next
    Next
    
    With Selection
        .Value = cell_range
        .HorizontalAlignment = xlRight
        .NumberFormat = ".00"
    End With
    
End Sub


Public Sub flip_correlations_above()
    cell_range = Selection.Value

    If UBound(cell_range, 1) <> UBound(cell_range, 2) Then
        MsgBox "The number of rows must be the same as the number of columns."
        Exit Sub
    End If

    For a = 1 To UBound(cell_range, 1)
        For b = 1 To UBound(cell_range, 2)
            If a <> b Then
                temp1 = cell_range(a, b)
                temp2 = cell_range(b, a)
                
                If temp1 <> "" And temp2 = "" Then
                    temp = temp1
                ElseIf temp1 = "" And temp2 <> "" Then
                    temp = temp2
                ElseIf temp1 = temp2 Then
                    temp = temp1
                ElseIf (temp1 <> "" And temp2 <> "") And (temp1 <> temp2) Then
                    temp = "Ambiguous: " & temp1 & ", " & temp2
                End If
                
                cell_range(b, a) = temp
                cell_range(a, b) = ""
            End If
        Next
    Next
    
    With Selection
        .Value = cell_range
        .HorizontalAlignment = xlRight
        .NumberFormat = ".00"
    End With
        
End Sub

Function RLZ(num, Optional decimals = 3) As String
    RLZ = Right(WorksheetFunction.Fixed(num, decimals), decimals + 1)
End Function

Sub CollectionFormat(myColl As VBA.Collection, format As String)
    For Each a In myColl
        Range.NumberFormat = format
    Next
End Sub

Function STEM_STRING(Label)
' Removes numbers at the end of a string
    STEM_STRING = Label

    For a = 1 To Len(Label)
        If IsNumeric(Right(Label, a)) Then
            STEM_STRING = Left(STEM_STRING, Len(Label) - a)
        Else
            Exit For
        End If
    Next
End Function

Function STEM_FILE(ByVal file_name)
' Removes the extension (e.g., ".csv", ".txt" from the end of a file name)
    STEM_FILE = file_name
    If InStrRev(STEM_FILE, ".") = 0 Then
        Exit Function
    Else
        STEM_FILE = Left(STEM_FILE, InStrRev(STEM_FILE, ".") - 1)
    End If
End Function


Function NumFormat(Optional decimals = 2, Optional PVal = 1, Optional p10 = False, Optional lead0 = False)
' Returns a text string indicating formatting, including decimals and asterisks representing p-values
' decimals is the number of decimal places desired
' pval is the p value indicating significance
' p10: if this is true, then it will return (*) for p < .10
' lead0: should there be a leading zero?

    NumFormat = "."
    If lead0 = True Then NumFormat = "0" & NumFormat
    NumFormat = NumFormat & String$(decimals, "0")
    If PVal < 0.1 And p10 = True Then NumFormat = NumFormat & """(*)"""
    If PVal < 0.05 And PVal >= 0.01 Then NumFormat = NumFormat & """*"""
    If PVal < 0.01 And PVal >= 0.001 Then NumFormat = NumFormat & """**"""
    If PVal < 0.001 Then NumFormat = NumFormat & """***"""
End Function

Function asterisk_pval(p_val, Optional pLessTen = False)
    asterisk_pval = ""
    If p_val < 0.1 And p_val >= 0.05 And pLessTen = True Then asterisk_pval = "(*)"
    If p_val < 0.05 And p_val >= 0.01 Then asterisk_pval = "*"
    If p_val < 0.01 And p_val >= 0.001 Then asterisk_pval = "**"
    If p_val < 0.001 Then asterisk_pval = "***"
End Function

Function STRING_RIGHT(text, phrase, Optional ReturnBlank = True)
' Returns the text to the right of a target phrase
' text is the string to be searched
' phrase is the string to be searched for
' If ReturnBlank is set to True, then the function will return an empty string if the phrase is not found. If it is false, the original phrase will be returned.
    STRING_RIGHT = ""
    If ReturnBlank = False Then STRING_RIGHT = text
    If InStr(1, text, phrase) = 0 Then Exit Function
    If Len(text) - InStr(1, text, phrase) + Len(phrase) < 1 Then Exit Function
    STRING_RIGHT = Right(text, Len(text) - InStr(1, text, phrase) - Len(phrase) + 1)
End Function

Function STRING_LEFT(text, phrase, Optional ReturnBlank = True)
' Returns the text to the left of a target phrase
' text is the string to be searched
' phrase is the string to be searched for
' If ReturnBlank is set to True, then the function will return an empty string if the phrase is not found. If it is false, the original phrase will be returned.
    STRING_LEFT = ""
    If ReturnBlank = False Then STRING_LEFT = text
    If InStr(1, text, phrase) <= 1 Then Exit Function
    STRING_LEFT = Left(text, InStr(1, text, phrase) - 1)
End Function

Function CONTAINS(text, phrase, Optional case_sens = False)
' Returns True if the target phrase is contained within the text
    If case_sens = False Then
        text = UCase(text)
        phrase = UCase(phrase)
    End If

    If InStr(1, text, phrase) > 0 Then
        CONTAINS = True
    Else
        CONTAINS = False
    End If
End Function

Function CONCAT_SEP(sep, ParamArray items())
    CONCAT_SEP = ""
    For a = LBound(items) To UBound(items)
        If a > LBound(items) And Len(items(a)) > 0 Then CONCAT_SEP = CONCAT_SEP & sep
        CONCAT_SEP = CONCAT_SEP & items(a)
    Next
End Function

Function Cronbach(CellsArea As Range, Optional format = True, Optional AboveDiag = False)
' Above diagonal means it looks for the values above the diagonal

    cellscopy = CellsArea.Value
    Cronbach = CRONBACH_array(cellscopy)
End Function


Function CRONBACH_array(cellscopy, Optional format = True, Optional AboveDiag = False)
    ' Check that the number of cells is equal to the number of columns
    
    If UBound(cellscopy, 1) <> UBound(cellscopy, 2) Then
        CRONBACH_array = "The number of selected rows must be the same as the number of columns."
        Exit Function
    End If
    
    Dim num_v: num_v = UBound(cellscopy, 1)
    
    ' Get variances
    Dim vars()
    ReDim vars(1 To num_v)
    For a = 1 To UBound(vars)
        vars(a) = cellscopy(a, a)
    Next
    
    ' Get covariances / correlations
    Dim covars()
    C = 1
    
    ReDim covars(1 To (num_v ^ 2 - num_v) / 2)
    num_cv = UBound(covars)
    For a = 1 To num_v
        For b = 1 To a
            If AboveDiag = False Then
                If a <> b Then
                    covars(C) = cellscopy(a, b)
                    C = C + 1
                End If
            Else
                If a <> b Then
                    covars(C) = cellscopy(b, a)
                    C = C + 1
                End If
            End If
        Next
    Next
    
    Dim temp_text: temp_text = ""
    For a = 1 To UBound(covars)
        temp_text = temp_text & covars(a) & ", "
    Next
    
    'MsgBox (temp_text)
    
    ' Get variance mean
    Dim vars_mean: vars_mean = 0
    For a = 1 To num_v
        vars_mean = vars_mean + vars(a) / num_v
    Next
    
    ' Get covariance mean
    Dim covars_mean: covars_mean = 0
    For a = 1 To num_cv
        covars_mean = covars_mean + covars(a) / num_cv
    Next
    
    ' Return alpha
    CRONBACH_array = (num_v * covars_mean) / (vars_mean + (num_v - 1) * covars_mean)
    
    If format = True Then
        'CRONBACH = WorksheetFunction.text(CRONBACH, "(.00)")
        ActiveCell.NumberFormat = "00%"
    End If
    
End Function




Sub LoadResults(OUTPUT)
' Load the Mplus Output into an Mplus Object
    Set MplusOutput = New cMplusOutput
    MplusOutput.ParseOutput = OUTPUT
    
    Call CreateCorrelationTable
    
    For a = 1 To MplusOutput.ObsMatrix_n()
        MsgBox MplusOutput.ObsMatrixName(a)
    Next
End Sub

Sub CreateCorrelationTable(Optional below_matrix_num = 1, Optional above_matrix_num = 2)
    Dim x_start: x_start = ActiveCell.Column
    Dim y_start: y_start = ActiveCell.Row
    Dim y_offset: y_offset = 1
    Dim x_offset: x_offset = 1

    For y = 1 To MplusOutput.ObsVarNum()
        y_var = MplusOutput.ObsVarNum(y)
        Cells(y_start + y, x_start) = y & ". " & MplusOutput.VarName(y_var)
        Cells(y_start, x_start + y) = y
        For x = 1 To y
            x_var = MplusOutput.ObsVarNum(x)
            If x = y Then
                Cells(y_start + y, x_start + x) = "'--"
            Else
                If below_matrix_num > 0 Then Cells(y_start + y, x_start + x) = MplusOutput.Sample_Covariance(y_var, x_var, below_matrix_num)
                If above_matrix_num > 0 Then Cells(y_start + x, x_start + y) = MplusOutput.Sample_Covariance(y_var, x_var, above_matrix_num)
            End If
        Next
    Next
End Sub


Public Sub FormatMPlusOutput()
    OutputForm.Show
End Sub

Public Sub MsgboxA(text)
    'MsgBox text
    'Debugging.TextBox1.text = Debugging.TextBox1.text & Chr(13) & Chr(10) & text
End Sub



Function t_to_r(t, n)
    t_to_r = (Abs(t) / t) * Sqr((t ^ 2) / ((t ^ 2) + n - 2))
End Function

Function crit_val_asterisk(R, n)
    If n = 0 Then
        crit_val_asterisk = ""
        Exit Function
    End If
    t05 = WorksheetFunction.TInv(0.05, n - 1)
    t01 = WorksheetFunction.TInv(0.01, n - 1)
    t001 = WorksheetFunction.TInv(0.001, n - 1)

    r05 = t_to_r(t05, n)
    r01 = t_to_r(t01, n)
    r001 = t_to_r(t001, n)

    If Abs(R) > r05 Then crit_val_asterisk = "*"
    If Abs(R) > r01 Then crit_val_asterisk = "**"
    If Abs(R) > r001 Then crit_val_asterisk = "***"
End Function

Function IsArrayAllocated(arr As Variant) As Boolean
    On Error Resume Next
    IsArrayAllocated = IsArray(arr) And Not IsError(LBound(arr, 1)) And LBound(arr, 1) <= UBound(arr, 1)
    On Error GoTo 0
End Function

Sub testword()
    Debug.Print GetWord(" TL_PGO_M       0.000         0.000", 2)
End Sub

Function GetWord(temp_word, num_word)

    temp_word = Trim(temp_word)
    Do While InStr(1, temp_word, "  ") > 0
        temp_word = Replace(temp_word, "  ", " ")
    Loop

    Dim start_pos: start_pos = 1
    Dim end_pos: end_pos = Len(temp_word)

    If num_word > 1 Then start_pos = instr_instance(temp_word, " ", num_word - 1) + 1
    If num_word < Numb_Words(temp_word) Then end_pos = instr_instance(temp_word, " ", num_word)
    
    'Debug.Print end_pos
    GetWord = Mid(temp_word, start_pos, end_pos - start_pos + 1)
End Function

Function instr_instance(string1, string2, instance)
    Dim start_pos: start_pos = 0
    
    'Debug.Print instance

    For a = 1 To instance
        start_pos = InStr(start_pos + 1, string1, string2)
        'Debug.Print a & " " & start_pos
    Next

    instr_inst = start_pos
End Function

Function Numb_Words(text)
    text = remove_whitespace(text)
    
    If InStr(1, text, " ") = 0 Then
        If Len(text) > 0 Then
            Numb_Words = 1
        Else
            Numb_Words = 0
        End If
        Exit Function
    Else
        Numb_Words = 1
    End If

    start = 1
    Do While InStr(start, text, " ") > 0
        start = InStr(start, text, " ") + 1
        Numb_Words = Numb_Words + 1
    Loop
End Function

Function remove_whitespace(orig_string)
    remove_whitespace = Trim(orig_string)
    
    Do While InStr(1, remove_whitespace, "  ") > 0
        remove_whitespace = Replace(remove_whitespace, "  ", " ")
    Loop
End Function

Function format_FL(FL, p, Optional decimals = 2)
    If FL = "" Or FL = 1 Then
        format_FL = FL
        Exit Function
    End If
    
    FL = WorksheetFunction.Fixed(FL, decimals) & ""
    FL = Replace(FL, "0.", ".")
    
    Dim stars
    If p < 0.05 Then stars = "*"
    If p < 0.01 Then stars = "**"
    If p < 0.001 Then stars = "***"
    
    format_FL = FL & stars
End Function


Public Function MyDocsPath() As String
    ' Get the "My Documents" path
    MyDocsPath = Environ$("USERPROFILE") & "\My Documents"
End Function


Function SUBFOLDER_EXISTS(subfolder)
    Dim file_path: file_path = ActiveWorkbook.Path
    If file_path = "" Then file_path = MyDocsPath()

    If Right(file_path, 1) <> "\" Then
        file_path = file_path & "\"
    End If
    
    If Left(subfolder, 1) = "\" Then
        subfolder = Right(subfolder, Len(subfolder) - 1)
    End If
    file_path = file_path & subfolder
        
    If Right(file_path, 1) <> "\" Then
        file_path = file_path & "\"
    End If
    
    If VBA.FileSystem.Dir(file_path) <> vbNullString Then
        SUBFOLDER_EXISTS = True
    Else
        SUBFOLDER_EXISTS = False
    End If
    
    filename = VBA.FileSystem.Dir("your folder name\your file name")
End Function

Function as_formula(ByVal formula_text)
    If formula_text = "" Then
        as_formula = "="""""
        Exit Function
    ElseIf InStr(formula_text, "NA") > 0 Then
        as_formula = "=""NA"""
        Exit Function
    End If
    as_formula = "=" & formula_text
End Function


Sub dump_matrix()
' Used for debugging. Displays content of a matrix into Excel.

    LoadMplusOutput.Show
    If LoadMplusOutput.execute = False Then Exit Sub
    Set MplusOutput = New cMplusOutput
    MplusOutput.ParseOutput = LoadMplusOutput.MPlusInput.text
    Unload LoadMplusOutput
    
    y_start = ActiveCell.Row
    x_start = ActiveCell.Column
    
    For y = 1 To MplusOutput.ModelVariable_n
        For x = 1 To MplusOutput.ModelVariable_n
            Cells(y_start + y - 1, x_start + x - 1) = MplusOutput.pMod_Matrix(y, x, 2, 1)
        Next
    Next
End Sub

