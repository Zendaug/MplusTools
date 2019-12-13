Attribute VB_Name = "Output_ObsMatrix"
Sub Start_ObsMatrix()
    LoadMplusOutput.Show
    
    ' Halt operations if the user exits the form without clicking on the "Proceed" button
    If LoadMplusOutput.execute = False Then Exit Sub
    
    Set MplusOutput = New cMplusOutput
    MplusOutput.ParseOutput = LoadMplusOutput.MPlusInput.text
    
    Unload LoadMplusOutput
    
    If MplusOutput.IsSAMPSTAT = False Then
        MsgBox "Sample statistics not found."
        Exit Sub
    End If
    
' Reset defaults
    Call ResetDefaults
    
    Form_ObsMatrices.matrix2.AddItem ("None")
    If MplusOutput.ObsMatrix_n = 1 Then
        Form_ObsMatrices.matrix1.AddItem ("Entire Dataset")
    Else
        For a = 1 To MplusOutput.ObsMatrix_n
            Form_ObsMatrices.matrix1.AddItem (MplusOutput.ObsMatrixName(a))
            Form_ObsMatrices.matrix2.AddItem (MplusOutput.ObsMatrixName(a))
        Next
    End If
    Form_ObsMatrices.matrix1.ListIndex = 0
    Form_ObsMatrices.matrix2.ListIndex = 0
        
    If MplusOutput.IsICC = True Then
        Form_ObsMatrices.ICC1.Enabled = True
        Form_ObsMatrices.ICC2.Enabled = True
    End If
    
    Form_ObsMatrices.Show
    
    ' Halt operations if the user exits the form without clicking on the "Generate" button
    If Form_ObsMatrices.execute = False Then Exit Sub
    
    matrix1 = Form_ObsMatrices.matrix1.ListIndex + 1
    matrix2 = Form_ObsMatrices.matrix2.ListIndex
    
    If Form_ObsMatrices.Correlation = True Then output_type = "Correlation"
    If Form_ObsMatrices.Covariance = True Then output_type = "Covariance"
    
    If Form_ObsMatrices.Mean = True Then Means = True
    If Form_ObsMatrices.SD = True Then SDs = True
    
    head1 = Form_ObsMatrices.Heading1.text
    head2 = Form_ObsMatrices.Heading2.text
    
    note_text = Form_ObsMatrices.Note.text
    
    If Form_ObsMatrices.Option_APA = True Then sig = True
    If Form_ObsMatrices.Option_None = True Then sig = False
    
    ICC1 = Form_ObsMatrices.ICC1
    ICC2 = Form_ObsMatrices.ICC2
    
    Unload Form_ObsMatrices
    
    Call CreateCorrelationTable(MplusOutput, matrix1, matrix2, output_type, Means, SDs, False, head1, head2, note_text, sig, n_decimals, ICC1, ICC2)
End Sub

Sub CreateCorrelationTable(MplusOutput, Optional below_matrix_num = 1, Optional above_matrix_num = 0, Optional output_type = "Correlation", Optional Means = True, Optional SDs = True, Optional no_label = False, Optional Heading1 = "", Optional Heading2 = "", Optional note_text = "", Optional sig = False, Optional ByVal decimals = 2, Optional ICC1 = False, Optional ICC2 = False)
    Dim x_start: x_start = ActiveCell.Column
    Dim y_start: y_start = ActiveCell.Row
    Dim y_offset: y_offset = 1
    Dim x_offset: x_offset = 1
    
    If no_label = True Then
        no_label = 0
    Else
        no_label = 1
    End If

    ' Leave heading2 blank for now; insert it after resizing the row
    If Heading1 <> "" Then y_start = y_start + 1
    If Heading2 <> "" Then y_start = y_start + 1
        
    For y = 1 To MplusOutput.ObsVarNum()
        y_var = MplusOutput.ObsVarNum(y)
        Cells(y_start + y, x_start) = y & ". " & MplusOutput.VarName(y_var, no_label)
        With Cells(y_start, x_start + y)
            .Value = y
            .HorizontalAlignment = xlCenter
        End With
        For x = 1 To y
            x_var = MplusOutput.ObsVarNum(x)
            If x = y Then
                With Cells(y_start + y, x_start + x)
                    .Value = "'--"
                    .HorizontalAlignment = xlRight
                End With
            Else
                p_below = 1: p_above = 1
                If below_matrix_num > 0 Then
                    If sig = True Then p_below = MplusOutput.Sample_P(y_var, x_var, below_matrix_num)
                    If output_type = "Correlation" Then res_below = MplusOutput.Sample_Correlation(y_var, x_var, below_matrix_num)
                    If output_type = "Covariance" Then res_below = MplusOutput.Sample_Covariance(y_var, x_var, below_matrix_num)
                    With Cells(y_start + y, x_start + x)
                        .Value = res_below
                        .NumberFormat = NumFormat(decimals, p_below)
                    End With
                    'MsgBox p_below & " " & NumFormat(decimals, p_below)
                End If
                If above_matrix_num > 0 Then
                    If sig = True Then p_above = MplusOutput.Sample_P(y_var, x_var, above_matrix_num)
                    If output_type = "Correlation" Then res_above = MplusOutput.Sample_Correlation(y_var, x_var, above_matrix_num)
                    If output_type = "Covariance" Then res_above = MplusOutput.Sample_Covariance(y_var, x_var, above_matrix_num)
                    With Cells(y_start + x, x_start + y)
                        .Value = res_above
                        .NumberFormat = NumFormat(decimals, p_above)
                    End With
                    'MsgBox p_above & " " & NumFormat(decimals, p_above)
                End If
            End If
        Next
    Next
    
    ' Insert the means
    If Means = True Then
        If (below_matrix_num > 0 And above_matrix_num > 0) Then
            Cells(y_start + y, x_start) = "Mean (below diagonal)"
            Cells(y_start + y + 1, x_start) = "Mean (above diagonal)"
            For x = 1 To MplusOutput.ObsVarNum()
                x_var = MplusOutput.ObsVarNum(x)
                With Cells(y_start + y + 0, x_start + x)
                    .Value = MplusOutput.Sample_Mean(x_var, below_matrix_num)
                    .NumberFormat = NumFormat(decimals)
                End With
                With Cells(y_start + y + 1, x_start + x)
                    .Value = MplusOutput.Sample_Mean(x_var, above_matrix_num)
                    .NumberFormat = NumFormat(decimals)
                End With
            Next
            y = y + 2
        Else
            Cells(y_start + y, x_start) = "Mean"
            For x = 1 To MplusOutput.ObsVarNum()
                x_var = MplusOutput.ObsVarNum(x)
                If below_matrix_num > 0 Then
                    With Cells(y_start + y, x_start + x)
                        .Value = MplusOutput.Sample_Mean(x_var, below_matrix_num)
                        .NumberFormat = NumFormat(decimals)
                    End With
                End If
                If above_matrix_num > 0 Then
                    With Cells(y_start + y, x_start + x)
                        .Value = MplusOutput.Sample_Mean(x_var, above_matrix_num)
                        .NumberFormat = NumFormat(decimals)
                    End With
                End If
            Next
            y = y + 1
        End If
    End If
    
    ' Insert the SDs
    If SDs = True Then
        If (below_matrix_num > 0 And above_matrix_num > 0) Then
            Cells(y_start + y, x_start) = "SD (below diagonal)"
            Cells(y_start + y + 1, x_start) = "SD (above diagonal)"
            For x = 1 To MplusOutput.ObsVarNum()
                x_var = MplusOutput.ObsVarNum(x)
                var_a = MplusOutput.Sample_Variance(x_var, above_matrix_num)
                var_b = MplusOutput.Sample_Variance(x_var, below_matrix_num)
                With Cells(y_start + y + 0, x_start + x)
                    If var_b = "NA" Then
                        .Value = "NA"
                    Else
                        .formula = as_formula("SQRT(" & var_b & ")")
                    End If
                    .NumberFormat = NumFormat(decimals)
                End With
                With Cells(y_start + y + 1, x_start + x)
                    If var_a = "NA" Then
                        .Value = "NA"
                    Else
                        .formula = as_formula("SQRT(" & var_a & ")")
                    End If
                    .NumberFormat = NumFormat(decimals)
                End With
            Next
            y = y + 2
        Else
            Cells(y_start + y, x_start) = "SD"
            For x = 1 To MplusOutput.ObsVarNum()
                x_var = MplusOutput.ObsVarNum(x)
                If below_matrix_num > 0 Then
                    var_b = MplusOutput.Sample_Variance(x_var, below_matrix_num)
                    With Cells(y_start + y, x_start + x)
                        If var_b = "NA" Then
                            .Value = "NA"
                        Else
                            .formula = as_formula("SQRT(" & var_b & ")")
                        End If
                        .NumberFormat = NumFormat(decimals)
                    End With
                End If
                If above_matrix_num > 0 Then
                    var_a = MplusOutput.Sample_Variance(x_var, above_matrix_num)
                    With Cells(y_start + y, x_start + x)
                        If var_b = "NA" Then
                            .Value = "NA"
                        Else
                            .formula = as_formula("SQRT(" & var_a & ")")
                        End If
                        .NumberFormat = NumFormat(decimals)
                    End With
                End If
            Next
            y = y + 1
        End If
    End If
    
    ' Insert the ICC1S
    If ICC1 = True Then
        If (below_matrix_num > 0 And above_matrix_num > 0) Then
            Cells(y_start + y, x_start) = "ICC1 (below diagonal)"
            Cells(y_start + y + 1, x_start) = "ICC1 (above diagonal)"
            For x = 1 To MplusOutput.ObsVarNum()
                x_var = MplusOutput.ObsVarNum(x)
                With Cells(y_start + y + 0, x_start + x)
                    .formula = as_formula(MplusOutput.Sample_ICC1(x_var, below_matrix_num))
                    .NumberFormat = NumFormat(decimals)
                End With
                With Cells(y_start + y + 1, x_start + x)
                    .formula = as_formula(MplusOutput.Sample_ICC1(x_var, above_matrix_num))
                    .NumberFormat = NumFormat(decimals)
                End With
            Next
            y = y + 2
        Else
            Cells(y_start + y, x_start) = "ICC1"
            For x = 1 To MplusOutput.ObsVarNum()
                x_var = MplusOutput.ObsVarNum(x)
                If below_matrix_num > 0 Then
                    With Cells(y_start + y, x_start + x)
                        .formula = as_formula(MplusOutput.Sample_ICC1(x_var, below_matrix_num))
                        .NumberFormat = NumFormat(decimals)
                    End With
                End If
                If above_matrix_num > 0 Then
                    With Cells(y_start + y, x_start + x)
                        .formula = as_formula(MplusOutput.Sample_ICC1(x_var, above_matrix_num))
                        .NumberFormat = NumFormat(decimals)
                    End With
                End If
            Next
            y = y + 1
        End If
    End If
    
    ' Insert the ICC2S
    If ICC2 = True Then
        If (below_matrix_num > 0 And above_matrix_num > 0) Then
            Cells(y_start + y, x_start) = "ICC2 (below diagonal)"
            Cells(y_start + y + 1, x_start) = "ICC2 (above diagonal)"
            For x = 1 To MplusOutput.ObsVarNum()
                x_var = MplusOutput.ObsVarNum(x)
                With Cells(y_start + y + 0, x_start + x)
                    .formula = as_formula(MplusOutput.Sample_ICC2(x_var, below_matrix_num, use_formula))
                    .NumberFormat = NumFormat(decimals)
                End With
                With Cells(y_start + y + 1, x_start + x)
                    .formula = as_formula(MplusOutput.Sample_ICC2(x_var, above_matrix_num, use_formula))
                    .NumberFormat = NumFormat(decimals)
                End With
            Next
            y = y + 2
        Else
            Cells(y_start + y, x_start) = "ICC2"
            For x = 1 To MplusOutput.ObsVarNum()
                x_var = MplusOutput.ObsVarNum(x)
                If below_matrix_num > 0 Then
                    With Cells(y_start + y, x_start + x)
                        .formula = as_formula(MplusOutput.Sample_ICC2(x_var, below_matrix_num))
                        .NumberFormat = NumFormat(decimals)
                    End With
                End If
                If above_matrix_num > 0 Then
                    With Cells(y_start + y, x_start + x)
                        .formula = as_formula(MplusOutput.Sample_ICC2(x_var, above_matrix_num))
                        .NumberFormat = NumFormat(decimals)
                    End With
                End If
            Next
            y = y + 1
        End If
    End If
    
    With Range(Cells(ActiveCell.Row, ActiveCell.Column), Cells(y_start + y, x_start + x - 1))
        .Font.Name = "Times New Roman"
        .Font.Size = 12
        .VerticalAlignment = xlTop
    End With
    
    ' Insert borders
    ' Top of table
    With Range(Cells(y_start, x_start), Cells(y_start, x_start + x - 1))
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeTop).weight = xlThin
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).weight = xlThin
    End With
    ' Below table
    With Range(Cells(y_start + y - 1, x_start), Cells(y_start + y - 1, x_start + x - 1))
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).weight = xlThin
    End With
    
    ' Resize the column
    Columns(x_start).EntireColumn.AutoFit
        
    ' Insert headings finally
    y_offset = 1
    If Heading2 <> "" Then
        With Cells(y_start - y_offset, x_start)
            .Value = Heading2
            .Font.Italic = True
        End With
        y_offset = y_offset + 1
    End If
    If Heading1 <> "" Then Cells(y_start - y_offset, x_start) = Heading1

    ' Insert note
    Cells(y_start + y, x_start) = note_text

End Sub


Sub FormatTNR()
'
' FormatTNR Macro
'

'
    With Selection.Font
        .Name = "Times New Roman"
        .Size = 11
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
    With Selection.Font
        .Name = "Times New Roman"
        .Size = 12
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
    Range("G35").Select
    Selection.NumberFormat = ".00"
    Range("F35").Select
    Selection.NumberFormat = ".00**"
    Selection.NumberFormat = ".00""*"""
    Range("F38").Select
End Sub
