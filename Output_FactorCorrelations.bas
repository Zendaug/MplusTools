Attribute VB_Name = "Output_FactorCorrelations"
Sub Start_FactorCorrelations()
    LoadMplusOutput.Show
    
    ' Halt operations if the user exits the form without clicking on the "Proceed" button
    If LoadMplusOutput.execute = False Then Exit Sub
    
    Set MplusOutput = New cMplusOutput
    MplusOutput.ParseOutput = LoadMplusOutput.MPlusInput.text
    
    Unload LoadMplusOutput
    
    If MplusOutput.IsModel = False Then
        MsgBox "Model not found."
        Exit Sub
    End If
    
' Reset defaults
    Call ResetDefaults
    
    Form_FactorCorrelations.matrix1.AddItem ("None")
    Form_FactorCorrelations.matrix2.AddItem ("None")

    If MplusOutput.Model_n = 1 Then
        Form_FactorCorrelations.matrix1.AddItem ("Entire Dataset")
        Form_FactorCorrelations.matrix2.AddItem ("Entire Dataset")
    Else
        For a = 1 To MplusOutput.Model_n
            Form_FactorCorrelations.matrix1.AddItem (MplusOutput.ModelName(a))
            Form_FactorCorrelations.matrix2.AddItem (MplusOutput.ModelName(a))
        Next
    End If
    Form_FactorCorrelations.matrix1.ListIndex = 1
    Form_FactorCorrelations.matrix2.ListIndex = 0
    
    If MplusOutput.IsSAMPSTAT = True Then
        Form_FactorCorrelations.matrix2.Enabled = True
        If MplusOutput.Model_n = 1 Then
            Form_FactorCorrelations.matrix1.AddItem ("Entire Dataset (composite sample correlations)")
            Form_FactorCorrelations.matrix2.AddItem ("Entire Dataset (composite sample correlations)")
        Else
            For a = 1 To MplusOutput.ObsMatrix_n
                Form_FactorCorrelations.matrix1.AddItem (MplusOutput.ObsMatrixName(a) & "(composite sample correlations)")
                Form_FactorCorrelations.matrix2.AddItem (MplusOutput.ObsMatrixName(a) & "(composite sample correlations)")
            Next
        End If
    End If

    Form_FactorCorrelations.StandNum.AddItem ("Unstandardized")
    Form_FactorCorrelations.StandNum.ListIndex = 0
    If MplusOutput.Std_n > 1 Then
        For a = 2 To MplusOutput.Std_n
            Form_FactorCorrelations.StandNum.AddItem (MplusOutput.StdName(a))
        Next
        Form_FactorCorrelations.StandNum.ListIndex = 1
    End If
               
    Form_FactorCorrelations.Show
    If Form_FactorCorrelations.execute = False Then Exit Sub
    
    diagonal = 0
    If Form_FactorCorrelations.Alpha = True Then diagonal = 1
    If Form_FactorCorrelations.Omega = True Then diagonal = 2
    If Form_FactorCorrelations.AVE = True Then diagonal = 3
    
    Call CreateFactorCorrelations(MplusOutput, Form_FactorCorrelations.matrix1.ListIndex, _
    Form_FactorCorrelations.matrix2.ListIndex, _
    Form_FactorCorrelations.StandNum.ListIndex + 1, _
    Form_FactorCorrelations.Mean, _
    Form_FactorCorrelations.SD, _
    Form_FactorCorrelations.Heading1, _
    Form_FactorCorrelations.Note.text, _
    Form_FactorCorrelations.PVal, _
    diagonal, _
    Form_FactorCorrelations.VarDisplay.ListIndex, _
    Form_FactorCorrelations.Heading2, _
    n_decimals)
    
    Unload Form_FactorCorrelations
End Sub

Sub CreateFactorCorrelations(MplusOutput, Optional below_matrix_num = 1, Optional above_matrix_num = 0, Optional std_num = 1, Optional Means = True, Optional SDs = True, Optional Heading1 = "", Optional note_text = "", Optional sig = True, Optional diagonal = 0, Optional format_option = 1, Optional Heading2 = "", Optional ByVal decimals = 2)

    Dim x_start: x_start = ActiveCell.Column
    Dim y_start: y_start = ActiveCell.Row
    Dim y_offset: y_offset = 1
    Dim x_offset: x_offset = 1
           
    Dim below_composite: below_composite = False
    Dim above_composite: above_composite = False
    
    ' Set up diagonal format
    Dim diag_format: diag_format = "(.0"
    If decimals >= 2 Then diag_format = diag_format & "0"
    If decimals >= 3 Then diag_format = diag_format & "0"
    diag_format = diag_format & ")"
           
    ' Leave heading2 blank for now; insert it after resizing the row
    If Heading1 <> "" Then y_start = y_start + 1
    If Heading2 <> "" Then y_start = y_start + 1
    
    ' Determine the model number for below/above diagonal, and also whether it is a composite
    If below_matrix_num > MplusOutput.Model_n Then
        below_matrix_num = below_matrix_num - MplusOutput.Model_n
        below_composite = True
    End If
    If above_matrix_num > MplusOutput.Model_n Then
        above_matrix_num = above_matrix_num - MplusOutput.Model_n
        above_composite = True
    End If
    
    'MsgBox "Below model " & below_matrix_num & ", Above model " & above_matrix_num
        
    For y = 1 To MplusOutput.Composite()
        y_var = MplusOutput.Composite(y)
        Cells(y_start + y, x_start) = y & ". " & MplusOutput.VarName(y_var, format_option)
        With Cells(y_start, x_start + y)
            .Value = y
            .HorizontalAlignment = xlCenter
        End With
        For x = 1 To y
            x_var = MplusOutput.Composite(x)
            If x = y Then  ' Add Cronbach's alpha, CR, AVE
                f_num = MplusOutput.FactorFind(y_var)
                'MsgBox f_num
                If f_num > 0 Then
                    If diagonal = 1 Then d_text = MplusOutput.Alpha(f_num, 1, below_matrix_num, True, use_formula)
                    If diagonal = 2 Then d_text = MplusOutput.Omega(f_num, 1, below_matrix_num, True, use_formula)
                    If diagonal = 3 Then d_text = "SQRT(" & MplusOutput.AVE(f_num, 1, below_matrix_num, True, use_formula) & ")"
                End If
                With Cells(y_start + y, x_start + x)
                    If diagonal = 0 Or f_num = 0 Then .Value = "'--"
                    If diagonal > 0 And f_num > 0 Then
                        .formula = as_formula(d_text)
                        .NumberFormat = diag_format
                    End If
                    .HorizontalAlignment = xlRight
                End With
            Else
                p_below = 1: p_above = 1 ' Set p to 1 by default
                If below_matrix_num > 0 Then
                    If below_composite = True Then
                        res_below = MplusOutput.Composite_Correlation(y_var, x_var, 1, below_matrix_num, use_formula)
                        If sig = True Then p_below = MplusOutput.Composite_P(y_var, x_var, below_matrix_num)
                    Else
                        res_below = MplusOutput.NonDirectionalPath(y_var, x_var, std_num, below_matrix_num)
                        If sig = True Then p_below = MplusOutput.NonDirectionalPathP(y_var, x_var, std_num, below_matrix_num)
                    End If
                    With Cells(y_start + y, x_start + x)
                        .formula = as_formula(res_below)
                        .NumberFormat = NumFormat(decimals, p_below)
                    End With
                End If
                If above_matrix_num > 0 Then
                    If above_composite = True Then
                        res_above = MplusOutput.Composite_Correlation(x_var, y_var, 2, above_matrix_num, use_formula)
                        If sig = True Then p_above = MplusOutput.Composite_P(x_var, y_var, above_matrix_num)
                    Else
                        res_above = MplusOutput.NonDirectionalPath(x_var, y_var, std_num, above_matrix_num)
                        If sig = True Then p_above = MplusOutput.NonDirectionalPathP(x_var, y_var, std_num, above_matrix_num)
                    End If
                    With Cells(y_start + x, x_start + y)
                        .formula = as_formula(res_above)
                        .NumberFormat = NumFormat(decimals, p_above)
                    End With
                End If
            End If
        Next
    Next
    
    ' Insert the means
    If Means = True Then
        If (below_matrix_num > 0 And above_matrix_num > 0) Then
            Cells(y_start + y, x_start) = "Mean (above diagonal)"
            Cells(y_start + y + 1, x_start) = "Mean (below diagonal)"
            For x = 1 To MplusOutput.Composite()
                x_var = MplusOutput.Composite(x)
                
                If below_composite = True Then
                    mean_1 = MplusOutput.Composite_Mean(x_var, below_matrix_num, use_formula)
                Else
                    mean_1 = MplusOutput.Mean(x_var, std_num, below_matrix_num)
                End If
                
                If above_composite = True Then
                    mean_2 = MplusOutput.Composite_Mean(x_var, above_matrix_num, use_formula)
                Else
                    mean_2 = MplusOutput.Mean(x_var, std_num, above_matrix_num)
                End If
                
                With Cells(y_start + y + 0, x_start + x)
                    .formula = as_formula(mean_2)
                    .NumberFormat = NumFormat(decimals)
                End With
                With Cells(y_start + y + 1, x_start + x)
                    .formula = as_formula(mean_1)
                    .NumberFormat = NumFormat(decimals)
                End With
            Next
            y = y + 2
        Else
            Cells(y_start + y, x_start) = "Mean"
            For x = 1 To MplusOutput.Composite()
                x_var = MplusOutput.Composite(x)
                If below_matrix_num > 0 Then
                    If below_composite = True Then
                        mean_1 = MplusOutput.Composite_Mean(x_var, below_matrix_num, use_formula)
                    Else
                        mean_1 = MplusOutput.Mean(x_var, std_num, below_matrix_num)
                    End If
                    With Cells(y_start + y, x_start + x)
                        .formula = as_formula(mean_1)
                        .NumberFormat = NumFormat(decimals)
                    End With
                End If
                If above_matrix_num > 0 Then
                    If above_composite = True Then
                        mean_2 = MplusOutput.Composite_Mean(x_var, above_matrix_num, use_formula)
                    Else
                        mean_2 = MplusOutput.Mean(x_var, std_num, above_matrix_num)
                    End If
                    With Cells(y_start + y, x_start + x)
                        .formula = as_formula(mean_2)
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
            Cells(y_start + y, x_start) = "SD (above diagonal)"
            Cells(y_start + y + 1, x_start) = "SD (below diagonal)"
            For x = 1 To MplusOutput.Composite()
                x_var = MplusOutput.Composite(x)
                
                If below_composite = True Then
                    SD_1 = MplusOutput.Composite_SD(x_var, below_matrix_num, use_formula)
                Else
                    SD_1 = MplusOutput.SD(x_var, 1, below_matrix_num, use_formula)
                End If
                
                If above_composite = True Then
                    SD_2 = MplusOutput.Composite_SD(x_var, above_matrix_num, use_formula)
                Else
                    SD_2 = MplusOutput.SD(x_var, 1, above_matrix_num, use_formula)
                End If
                
                With Cells(y_start + y + 0, x_start + x)
                    .formula = as_formula(SD_2)
                    .NumberFormat = NumFormat(decimals)
                End With
                With Cells(y_start + y + 1, x_start + x)
                    .formula = as_formula(SD_1)
                    .NumberFormat = NumFormat(decimals)
                End With
            Next
            y = y + 2
        Else
            Cells(y_start + y, x_start) = "SD"
            For x = 1 To MplusOutput.Composite()
                x_var = MplusOutput.Composite(x)
                If below_matrix_num > 0 Then
                    If below_composite = True Then
                        SD_1 = MplusOutput.Composite_SD(x_var, below_matrix_num, use_formula)
                    Else
                        SD_1 = MplusOutput.SD(x_var, 1, below_matrix_num, use_formula)
                    End If
                    With Cells(y_start + y, x_start + x)
                        .formula = as_formula(SD_1)
                        .NumberFormat = NumFormat(decimals)
                    End With
                End If
                If above_matrix_num > 0 Then
                    If above_composite = True Then
                        SD_2 = MplusOutput.Composite_SD(x_var, above_matrix_num, use_formula)
                    Else
                        SD_2 = MplusOutput.SD(x_var, 1, above_matrix_num, use_formula)
                    End If
                    With Cells(y_start + y, x_start + x)
                        .formula = as_formula(SD_2)
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
