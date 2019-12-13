Attribute VB_Name = "Output_CFA"
Sub Start_ObsCFATable()
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
    
    If MplusOutput.IsSAMPSTAT = False Then
        Form_CFATable.Alpha.Enabled = False
        Form_CFATable.Alpha.Value = False
        Form_CFATable.Alpha.ControlTipText = "SAMPSTAT is required."
    End If
    
    If MplusOutput.Model_n = 1 Then
        Form_CFATable.ModelNum.AddItem ("Entire Dataset")
    Else
        For a = 1 To MplusOutput.Model_n
            Form_CFATable.ModelNum.AddItem (MplusOutput.ModelName(a))
        Next
    End If
    Form_CFATable.ModelNum.ListIndex = 0
    
    Form_CFATable.StandNum.AddItem ("Unstandardized")
    If MplusOutput.Std_n > 1 Then
        For a = 2 To MplusOutput.Std_n
            Form_CFATable.StandNum.AddItem (MplusOutput.StdName(a))
        Next
    End If
    Form_CFATable.StandNum.ListIndex = 0
    
    Form_CFATable.FitStats = ModelFitAuto(MplusOutput)
    
    Form_CFATable.Show
    If Form_CFATable.execute = False Then Exit Sub
    Call CreateCFATable(MplusOutput, Form_CFATable.ModelNum.ListIndex + 1, Form_CFATable.StandNum.ListIndex + 1, Form_CFATable.SESD.Value, Form_CFATable.PVal.Value, Form_CFATable.Means.Value, Form_CFATable.SDs.Value, Form_CFATable.CR.Value, Form_CFATable.CR_Total.Value, Form_CFATable.AVE.Value, Form_CFATable.Alpha.Value, Form_CFATable.ModelFit.Value, Form_CFATable.Intercepts.Value, Form_CFATable.Heading1.Value, Form_CFATable.Heading2.Value, Form_CFATable.Note.Value, n_decimals, Form_CFATable.SortBySize.Value, Form_CFATable.HideBelow, Form_CFATable.CoefAction.ListIndex, var_disp_mode, Form_CFATable.SingleColumn, Form_CFATable.obs_only)

    Unload Form_CFATable
End Sub

Sub CreateCFATable(MplusOutput, _
                    Optional ModelNum = 1, _
                    Optional StandNum = 1, _
                    Optional SESD = False, _
                    Optional PV = False, _
                    Optional Means = False, _
                    Optional SDs = False, _
                    Optional CR = False, _
                    Optional CR_Total = False, _
                    Optional AVE = False, _
                    Optional Alpha = False, _
                    Optional ModelFit = False, _
                    Optional Intercepts = False, _
                    Optional Heading1 = "Table X.", _
                    Optional Heading2 = "", _
                    Optional note_text = "", _
                    Optional ByVal DecimalPlaces = 2, _
                    Optional SortBySize = False, _
                    Optional HideBelow = "", _
                    Optional CoefAction = 0, _
                    Optional VarDisplay = 1, _
                    Optional SingleColumn = False, _
                    Optional obs_only = False)

    If DecimalPlaces = 0 Then ftext = "0"
    If DecimalPlaces = 1 Then ftext = ".0"
    If DecimalPlaces = 2 Then ftext = ".00"
    If DecimalPlaces = 3 Then ftext = ".000"
    If HideBelow = "" Then HideBelow = -1
    HideBelow = CDbl(HideBelow)

    start_row = ActiveCell.Row
    start_col = ActiveCell.Column
    
    ' Add Headings
    Cells(start_row, start_col).Value = Heading1
    If Heading2 <> "" Then
        start_row = start_row + 1
        With Cells(start_row, start_col)
            .Value = Heading2
            .Font.Italic = True
        End With
    End If
    start_row = start_row + 1
       
    ' Create an order of factors
    factor_order = ""
    For a = 1 To MplusOutput.Factor()
        If MplusOutput.FactorIndicator(a, 0, StandNum, ModelNum, obs_only) > 0 Then factor_order = ADD_STRINGVECTOR(factor_order, a)
    Next

    ' Multiple columns
    If SingleColumn = False Then
        
        ' Label the factor columns
        For a = 1 To COUNT_STRINGVECTOR(factor_order)
            With Cells(start_row, start_col + a)
                .Value = MplusOutput.FactorName(CInt(GET_STRINGVECTOR(factor_order, a)))
                .HorizontalAlignment = xlCenter
            End With
        Next
        If Intercepts = True Then
            With Cells(start_row, start_col + a)
                .Value = "Intercepts"
                .HorizontalAlignment = xlCenter
            End With
        End If

        ' Create an order of indicators
        indicator_order = ""
        indicator_remaining = ""
        For a = 1 To COUNT_STRINGVECTOR(factor_order)
            ' Only proceed if the number of indicators is more than zero
            temp_array = MplusOutput.FactorIndicatorArray(CInt(GET_STRINGVECTOR(factor_order, a)), StandNum, ModelNum, obs_only, SortBySize)
            For b = 1 To UBound(temp_array)
                ind_num = temp_array(b, 1)
                ind_name = MplusOutput.VarName(ind_num, 0)
                coef = Abs(temp_array(b, 2))
                If coef >= HideBelow Then
                    indicator_order = ADD_STRINGVECTOR(indicator_order, ind_num, ind_name)
                Else
                    indicator_remaining = ADD_STRINGVECTOR(indicator_remaining, ind_num, ind_name)
                End If
            Next
        Next
        indicator_order = JOIN_STRINGVECTOR(indicator_order, indicator_remaining)
    
        With Cells(start_row, start_col)
            .Value = "Indicator"
            .HorizontalAlignment = xlCenter
        End With
        For a = 1 To COUNT_STRINGVECTOR(indicator_order)
            With Cells(start_row + a, start_col)
                .Value = MplusOutput.VarName(GET_STRINGVECTOR(indicator_order, a), VarDisplay)
                .HorizontalAlignment = xlLeft
            End With
        Next
        
        For a = 1 To COUNT_STRINGVECTOR(factor_order)
            For b = 1 To COUNT_STRINGVECTOR(indicator_order)
                fac_num = MplusOutput.Factor(CInt(GET_STRINGVECTOR(factor_order, a)))
                ind_num = CInt(GET_STRINGVECTOR(indicator_order, b))
                temp_text = MplusOutput.Path(fac_num, ind_num, StandNum, ModelNum)
                'MsgBox "Factor: " & fac_num & ", Indicator number: " & ind_num & ", Path coefficient: " & temp_text
                If temp_text <> "NA" Then
                    temp_text2 = format(temp_text, ftext)
                    If PV = True Then temp_text2 = temp_text2 & asterisk_pval(MplusOutput.PathP(fac_num, ind_num, StandNum, ModelNum))
                    If SESD = True Then temp_text2 = temp_text2 & " (" & format(MplusOutput.PathSE(fac_num, ind_num, StandNum, ModelNum), ftext) & ")"
                    If CoefAction = 1 And Abs(temp_text) < Abs(HideBelow) Then
                        ' Hidden!
                    Else
                        With Cells(start_row + b, start_col + a)
                            .Value = temp_text2
                            .HorizontalAlignment = xlRight
                            .NumberFormat = ftext
                        End With
                    End If
                    If CoefAction = 2 And Abs(temp_text) > Abs(HideBelow) Then Cells(start_row + b, start_col + a).Font.Bold = True
                End If
                
                ' Add the intercepts on the last loop
                If a = COUNT_STRINGVECTOR(factor_order) And Intercepts = True Then
                    temp_text = "="
                    If MplusOutput.Intercept(ind_num, StandNum, ModelNum) <> "NA" Then
                        For u = 1 To MplusOutput.PathNCategories(fac_num, ind_num, StandNum, ModelNum)
                            If u > 1 Then temp_text = temp_text & " & CHAR(10) & "
                            temp_text = temp_text & "text(" & MplusOutput.Intercept(ind_num, StandNum, ModelNum, u) & ", """ & ftext & """)"
                            If PV = True Then temp_text = temp_text & " & """ & asterisk_pval(MplusOutput.PathP(fac_num, ind_num, StandNum, ModelNum, u)) & """"
                            If SESD = True Then temp_text = temp_text & " & "" ("" & text(" & MplusOutput.InterceptSE(ind_num, StandNum, ModelNum, u) & ",""" & ftext & """) & "")"""
                        Next
                    Else
                        temp_text = temp_text & """NA"""
                    End If
                    With Cells(start_row + b, start_col + a + 1)
                        .formula = temp_text
                        .HorizontalAlignment = xlRight
                    End With
                End If
            Next
        Next
        
        ' Mid-table line
        row_midtable = start_row + b
        
        ' Insert Factor Means (if available)
        If Means = True Then
            Cells(start_row + b, start_col) = "Means"
            For a = 1 To COUNT_STRINGVECTOR(factor_order)
                var_num = MplusOutput.Factor(CInt(GET_STRINGVECTOR(factor_order, a)))
                With Cells(start_row + b, start_col + a)
                    .Value = MplusOutput.Mean(var_num, StandNum, ModelNum)
                    .NumberFormat = ftext
                End With
            Next
            b = b + 1
        End If
    
        ' Insert Factor SDs (if available)
        If SDs = True Then
            Cells(start_row + b, start_col) = "SDs"
            For a = 1 To COUNT_STRINGVECTOR(factor_order)
                var_num = MplusOutput.Factor(CInt(GET_STRINGVECTOR(factor_order, a)))
                get_var = MplusOutput.Variance(var_num, StandNum, ModelNum, True)
                If var_num <> "Residual" Then
                    With Cells(start_row + b, start_col + a)
                        .formula = as_formula("SQRT(" & MplusOutput.Variance(var_num, StandNum, ModelNum, True) & ")")
                        .NumberFormat = ftext
                    End With
                Else
                    Cells(start_row + b, start_col + a) = "NA"
                End If
            Next
            b = b + 1
        End If
        
        ' Insert Coefficient Alpha / Cronbach's
        If Alpha = True Then
            Cells(start_row + b, start_col) = "Cronbach's Alpha"
            For a = 1 To COUNT_STRINGVECTOR(factor_order)
                With Cells(start_row + b, start_col + a)
                    .formula = as_formula(MplusOutput.Alpha(CInt(GET_STRINGVECTOR(factor_order, a)), StandNum, ModelNum, obs_only, use_formula))
                    .NumberFormat = ftext
                End With
            Next
            b = b + 1
        End If
        
        ' Insert Coefficient Omega / Composite Reliability
        If CR = True Then
            Cells(start_row + b, start_col) = "Omega (single factor)"
            For a = 1 To COUNT_STRINGVECTOR(factor_order)
                temp_val = MplusOutput.Omega(CInt(GET_STRINGVECTOR(factor_order, a)), StandNum, ModelNum, obs_only, use_formula)
                With Cells(start_row + b, start_col + a)
                    If temp_val = "NA" Then
                        .Value = "NA"
                    Else
                        '.Value = temp_val
                        .formula = as_formula(temp_val)
                    End If
                    .NumberFormat = ftext
                End With
            Next
            b = b + 1
        End If
    
        ' Insert Coefficient Omega Total / Composite Reliability
        If CR_Total = True Then
            Cells(start_row + b, start_col) = "Omega Total (% variance explained)"
            For a = 1 To COUNT_STRINGVECTOR(factor_order)
                temp_val = MplusOutput.OmegaTotal(CInt(GET_STRINGVECTOR(factor_order, a)), StandNum, ModelNum, obs_only, use_formula)
                With Cells(start_row + b, start_col + a)
                    If temp_val = "NA" Then
                        .Value = "NA"
                    Else
                        .formula = as_formula(temp_val)
                    End If
                    .NumberFormat = ftext
                End With
            Next
            b = b + 1
        End If
    
        ' Insert Average Variance Extracted
        If AVE = True Then
            Cells(start_row + b, start_col) = "Average Variance Extracted"
            For a = 1 To COUNT_STRINGVECTOR(factor_order)
                With Cells(start_row + b, start_col + a)
                    .formula = as_formula(MplusOutput.AVE(CInt(GET_STRINGVECTOR(factor_order, a)), StandNum, ModelNum, obs_only, use_formula))
                    .NumberFormat = ftext
                End With
            Next
            b = b + 1
        End If
        
    Else
    
    ' Format output in a single column
        With Cells(start_row, start_col)
            .Value = "Indicator"
            .HorizontalAlignment = xlCenter
        End With
        With Cells(start_row, start_col + 1)
            .Value = "Loading"
            .HorizontalAlignment = xlCenter
        End With
        If Intercepts = True Then
            With Cells(start_row, start_col + 2)
                .Value = "Intercepts"
                .HorizontalAlignment = xlCenter
            End With
        End If
        
        b = 0
        
        Dim out_text As Collection
        Set out_text = New Collection
        
        For d = 1 To COUNT_STRINGVECTOR(factor_order)
            Set out_text = Nothing
            Set out_text = New Collection
        
            fac_num = MplusOutput.Factor(CInt(GET_STRINGVECTOR(factor_order, d)))
            b = b + 1
            
            ' Additional Output
            If Means = True Then
                temp_val = MplusOutput.Mean(fac_num, StandNum, ModelNum)
                If temp_val <> "NA" Then out_text.Add """M = "" & text(" & temp_val & ",""" & ftext & """)"
            End If
            If SDs = True Then
                get_var = MplusOutput.Variance(fac_num, StandNum, ModelNum, True)
                If get_var <> "Residual" And get_var <> "NA" Then out_text.Add """SD = "" & text(SQRT(" & get_var & "),""" & ftext & """)"
            End If
            If Alpha = True Then
                temp_var = MplusOutput.Alpha(d, 1, ModelNum, obs_only, True)
                If temp_var <> "NA" Then out_text.Add """Alpha = "" & text(" & temp_var & ",""" & ftext & """)"
            End If
            'ADD_STRINGVECTOR(out_text, "Alpha = " & format(MPlusOutput.Alpha(d, 1, ModelNum, obs_only, False), ftext))
            If CR = True Then
                temp_var = MplusOutput.Omega(d, StandNum, ModelNum, obs_only, use_formula)
                If temp_var <> "NA" Then out_text.Add """Omega = "" & text(" & temp_var & ",""" & ftext & """)"
            End If
            'ADD_STRINGVECTOR(out_text, "Omega = " & format(MPlusOutput.Omega(d, StandNum, ModelNum, obs_only, False), ftext))
            If CR_Total = True Then
                temp_var = MplusOutput.OmegaTotal(d, StandNum, ModelNum, obs_only, use_formula)
                If temp_var <> "NA" Then out_text.Add """Omega (% var explained) = "" & text(" & temp_var & ",""" & ftext & """)"
            End If
            'ADD_STRINGVECTOR(out_text, "Omega (% var explained) = " & format(MPlusOutput.OmegaTotal(d, StandNum, ModelNum, obs_only, False), ftext))
            If AVE = True Then
                temp_var = MplusOutput.AVE(d, StandNum, ModelNum, obs_only, use_formula)
                If temp_var <> "NA" Then out_text.Add """AVE = "" & text(" & temp_var & ",""" & ftext & """)"
            End If
            'ADD_STRINGVECTOR(out_text, "AVE = " & format(MPlusOutput.AVE(d, StandNum, ModelNum, obs_only, False), ftext))
            
            temp_text = ""
            If out_text.Count > 0 Then
                temp_text = " & "" ("" & "
                For e = 1 To out_text.Count
                    If e > 1 Then temp_text = temp_text & " & "", "" & "
                    temp_text = temp_text & out_text(e)
                Next
                temp_text = temp_text & "& "")"""
            End If
                                              
            With Cells(start_row + b, start_col)
                .formula = "=""" & MplusOutput.VarName(fac_num) & """" & temp_text
                '.Value = MPlusOutput.VarName(fac_num) & temp_text
                '.formula = "=" & MPlusOutput.VarName(fac_num) & " & " & out_text
                .HorizontalAlignment = xlCenter
            End With
                        
            ' Create an order of indicators
            indicator_order = ""
            indicator_remaining = ""
            temp_array = MplusOutput.FactorIndicatorArray(d, StandNum, ModelNum, obs_only, SortBySize)
            For f = 1 To UBound(temp_array)
                ind_num = temp_array(f, 1)
                ind_name = MplusOutput.VarName(ind_num, 0)
                coef = Abs(temp_array(f, 2))
                If coef >= HideBelow Then
                    indicator_order = ADD_STRINGVECTOR(indicator_order, ind_num, ind_name)
                Else
                    indicator_remaining = ADD_STRINGVECTOR(indicator_remaining, ind_num, ind_name)
                End If
            Next
            indicator_order = JOIN_STRINGVECTOR(indicator_order, indicator_remaining)
            
            For e = 1 To COUNT_STRINGVECTOR(indicator_order)
                b = b + 1
                ind_num = GET_STRINGVECTOR(indicator_order, e)
                With Cells(start_row + b, start_col)
                    .Value = MplusOutput.VarName(ind_num)
                    .HorizontalAlignment = xlLeft
                End With
                temp_text = MplusOutput.Path(fac_num, ind_num, StandNum, ModelNum)
                If temp_text <> "NA" Then
                    temp_text2 = format(temp_text, ftext)
                    If SESD = True Then temp_text2 = temp_text2 & " (" & format(MplusOutput.PathSE(fac_num, ind_num, StandNum, ModelNum), ftext) & ")"
                    If PV = True Then temp_text2 = temp_text2 & asterisk_pval(MplusOutput.PathP(fac_num, ind_num, StandNum, ModelNum))
                    If Not (CoefAction = 1 And Abs(temp_text) < Abs(HideBelow)) Then
                        With Cells(start_row + b, start_col + 1)
                            .Value = temp_text2
                            .HorizontalAlignment = xlRight
                            .NumberFormat = ftext
                        End With
                    End If
                    If CoefAction = 2 And Abs(temp_text) > Abs(HideBelow) Then Cells(start_row + b, start_col + 1).Font.Bold = True
                End If
                
                If Intercepts = True Then
                    temp_text = "="
                    If MplusOutput.Intercept(ind_num, StandNum, ModelNum) <> "NA" Then
                        For u = 1 To MplusOutput.PathNCategories(fac_num, ind_num, StandNum, ModelNum)
                            If u > 1 Then temp_text = temp_text & " & CHAR(10) & "
                            temp_text = temp_text & "text(" & MplusOutput.Intercept(ind_num, StandNum, ModelNum, u) & ", """ & ftext & """)"
                            If PV = True Then temp_text = temp_text & " & """ & asterisk_pval(MplusOutput.PathP(fac_num, ind_num, StandNum, ModelNum, u)) & """"
                            If SESD = True Then temp_text = temp_text & " & "" ("" & text(" & MplusOutput.InterceptSE(ind_num, StandNum, ModelNum, u) & ",""" & ftext & """) & "")"""
                        Next
                    Else
                        temp_text = temp_text & """NA"""
                    End If
                    With Cells(start_row + b, start_col + 2)
                        .formula = temp_text
                        .HorizontalAlignment = xlRight
                    End With
                End If
                
            Next
            b = b + 1
        Next
    End If
    
    
    ' Determine where the end row is
    If SingleColumn = False Then
        end_col = COUNT_STRINGVECTOR(factor_order)
    Else
        end_col = 1
        row_midtable = start_row + b
    End If
    If Intercepts = True Then end_col = end_col + 1
    
    
    ' Format the entire table
    With Range(Cells(ActiveCell.Row, ActiveCell.Column), Cells(start_row + b, start_col + end_col))
        .Font.Name = "Times New Roman"
        .Font.Size = 12
        .VerticalAlignment = xlTop
    End With

    
    ' Top of table
    With Range(Cells(start_row, start_col), Cells(start_row, start_col + end_col))
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeTop).weight = xlThin
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).weight = xlThin
    End With
    
    ' Mid table
    With Range(Cells(row_midtable, start_col), Cells(row_midtable, start_col + end_col))
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeTop).weight = xlThin
    End With
    
    ' Below table
    With Range(Cells(start_row + b, start_col), Cells(start_row + b, start_col + end_col))
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeTop).weight = xlThin
    End With
    
    ' Resize the column
    Columns(start_col).EntireColumn.AutoFit
        
    ' Insert note
    Cells(start_row + b, start_col) = note_text
    
    ' Move cursor to the next cell
    Cells(start_row + b + 1, start_col).Select
    ActiveSheet.Select

End Sub



Sub arrtype()
    Dim a()
    ReDim a(1 To 2)
    MsgBox VarType(a)
End Sub
