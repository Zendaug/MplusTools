Attribute VB_Name = "Output_SEM"

Sub Start_SEMTable()
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
    
    If MplusOutput.IsRSquare = False Then
        Form_PathSEM.RSquare.Value = False
        Form_PathSEM.RSquare.Enabled = False
    End If
    
    ' Populate combo boxes
    control_index = 1
    Form_PathSEM.PopCol.AddItem ("None")
    Form_PathSEM.PopRow.AddItem ("None")
    Form_PathSEM.PopAdjRow.AddItem ("None")
    Form_PathSEM.PopAdjCol.AddItem ("None")
    
    Form_PathSEM.PopCol.ListIndex = 0
    Form_PathSEM.PopRow.ListIndex = 0
    Form_PathSEM.PopAdjRow.ListIndex = 0
    Form_PathSEM.PopAdjCol.ListIndex = 0
    
    Form_PathSEM.group_index = -99
    Form_PathSEM.level_index = -99
    Form_PathSEM.mixture_index = -99
    Form_PathSEM.stand_index = -99
    Form_PathSEM.ci_index = -99
    
    If MplusOutput.nGroups = 1 And MplusOutput.nLevels = 1 And MplusOutput.nMixtures = 1 And MplusOutput.nStandards = 1 Then
        Form_PathSEM.PopCol.Enabled = False
        Form_PathSEM.PopRow.Enabled = False
        Form_PathSEM.PopAdjRow.Enabled = False
        Form_PathSEM.Label8.Enabled = False
        Form_PathSEM.Label9.Enabled = False
        Form_PathSEM.Label11.Enabled = False
        Form_PathSEM.Label12.Enabled = False
    Else
        If MplusOutput.nGroups > 1 Then
            Form_PathSEM.PopCol.AddItem ("Groups")
            Form_PathSEM.PopRow.AddItem ("Groups")
            Form_PathSEM.PopAdjRow.AddItem ("Groups")
            Form_PathSEM.PopAdjCol.AddItem ("Groups")
            Form_PathSEM.group_index = control_index
            control_index = control_index + 1
        End If
        If MplusOutput.nLevels > 1 Then
            Form_PathSEM.PopCol.AddItem ("Levels")
            Form_PathSEM.PopRow.AddItem ("Levels")
            Form_PathSEM.PopAdjRow.AddItem ("Levels")
            Form_PathSEM.PopAdjCol.AddItem ("Levels")
            Form_PathSEM.level_index = control_index
            control_index = control_index + 1
        End If
        If MplusOutput.nMixtures > 1 Then
            Form_PathSEM.PopCol.AddItem ("Mixtures")
            Form_PathSEM.PopRow.AddItem ("Mixtures")
            Form_PathSEM.PopAdjRow.AddItem ("Mixtures")
            Form_PathSEM.PopAdjCol.AddItem ("Mixtures")
            Form_PathSEM.mixture_index = control_index
            control_index = control_index + 1
        End If
        If MplusOutput.nStandards > 1 Then
            Form_PathSEM.PopCol.AddItem ("Standardizations")
            Form_PathSEM.PopRow.AddItem ("Standardizations")
            Form_PathSEM.PopAdjRow.AddItem ("Standardizations")
            Form_PathSEM.PopAdjCol.AddItem ("Standardizations")
            Form_PathSEM.stand_index = control_index
            control_index = control_index + 1
        End If
        If MplusOutput.IsCInterval = True Then
            cname = "Confidence "
            If MplusOutput.Estimator = "BAYES" Then cname = "Credibility "
            Form_PathSEM.PopCol.AddItem (cname & "Intervals")
            Form_PathSEM.PopRow.AddItem (cname & "Intervals")
            Form_PathSEM.PopAdjRow.AddItem (cname & "Intervals")
            Form_PathSEM.PopAdjCol.AddItem (cname & "Intervals")
            Form_PathSEM.ci_index = control_index
            control_index = control_index + 1
        End If
    End If
    
    Form_PathSEM.PopCol.AddItem ("Relative Importance %")
    Form_PathSEM.PopRow.AddItem ("Relative Importance %")
    Form_PathSEM.PopAdjRow.AddItem ("Relative Importance %")
    Form_PathSEM.PopAdjCol.AddItem ("Relative Importance %")
    Form_PathSEM.ri_index = control_index
    control_index = control_index + 1
    
    ' Share number of groups, levels, mixtures with form
    Form_PathSEM.nGroups = MplusOutput.nGroups
    Form_PathSEM.nLevels = MplusOutput.nLevels
    Form_PathSEM.nMixtures = MplusOutput.nMixtures
    
    Form_PathSEM.CLevel.AddItem ("None")
    Form_PathSEM.CLevel.ListIndex = 0
    If MplusOutput.IsCInterval = True Then
        Form_PathSEM.ci90 = 1
        Form_PathSEM.ci95 = 2
        Form_PathSEM.ci99 = 3
        Form_PathSEM.CLevel.AddItem ("90%")
        Form_PathSEM.CLevel.AddItem ("95%")
        Form_PathSEM.CLevel.AddItem ("99%")
        Form_PathSEM.CLevel.Enabled = True
        Form_PathSEM.CLevel.ListIndex = 2
    ElseIf MplusOutput.Estimator = "BAYES" Then
        Form_PathSEM.ci95 = 1
        Form_PathSEM.CLevel.AddItem ("95%")
        Form_PathSEM.CLevel.Enabled = True
        Form_PathSEM.CLevel.ListIndex = 1
    Else
        Form_PathSEM.CLevel.Enabled = False
        Form_PathSEM.Label10.Enabled = False
    End If
        
    If MplusOutput.Model_n = 1 Then
        Form_PathSEM.ModelNum.AddItem ("Entire Dataset")
    Else
        Form_PathSEM.ModelNum.AddItem ("Multiple groups / levels / mixtures")
        For a = 1 To MplusOutput.Model_n
            Form_PathSEM.ModelNum.AddItem (MplusOutput.ModelName(a))
        Next
    End If
    Form_PathSEM.ModelNum.ListIndex = 0
    
    Form_PathSEM.StandNum.AddItem ("Unstandardized")
    If MplusOutput.Std_n > 1 Then
        For a = 2 To MplusOutput.Std_n
            Form_PathSEM.StandNum.AddItem (MplusOutput.StdName(a))
        Next
    End If
    Form_PathSEM.StandNum.ListIndex = 0
    
    Form_PathSEM.FitStats = ModelFitAuto(MplusOutput)
    
    Form_PathSEM.Show
    
    If Form_PathSEM.execute = False Then
        Unload Form_PathSEM
        Exit Sub
    End If
    
    Call CreateSEMTable(MplusOutput, Form_PathSEM.ModelNum.ListIndex + 1, Form_PathSEM.group_output, Form_PathSEM.level_output, Form_PathSEM.mixture_output, Form_PathSEM.stand_output, Form_PathSEM.ci_output, Form_PathSEM.ri_output, Form_PathSEM.StandNum.ListIndex + 1, Form_PathSEM.SESD, Form_PathSEM.PVal, Form_PathSEM.CI, Form_PathSEM.Intercept, Form_PathSEM.RSquare, Form_PathSEM.Heading1, Form_PathSEM.Heading2, Form_PathSEM.Note, n_decimals)
    
    Unload Form_PathSEM
'    If Form_CFATable.execute = False Then Exit Sub
'    Call CreateCFATable(MPlusOutput, Form_CFATable.ModelNum.ListIndex + 1, Form_CFATable.StandNum.ListIndex + 1, Form_CFATable.SESD.Value, Form_CFATable.PVal.Value, Form_CFATable.Means.Value, Form_CFATable.SDs.Value, Form_CFATable.CR.Value, Form_CFATable.AVE.Value, Form_CFATable.Alpha.Value, Form_CFATable.ModelFit.Value, Form_CFATable.Heading1.Value, Form_CFATable.Heading2.Value, Form_CFATable.Note.Value, Form_CFATable.DecimalPlaces.Value, Form_CFATable.SortBySize.Value, Form_CFATable.HideBelow, Form_CFATable.CoefAction.ListIndex, Form_CFATable.VarDisplay.ListIndex, Form_CFATable.SingleColumn, Form_CFATable.obs_only)

'    Unload Form_CFATable
End Sub

Sub CreateSEMTable(MplusOutput, Optional ModelNum = 1, Optional group_output = 0, Optional level_output = 0, Optional mixture_output = 0, Optional stand_output = 0, Optional ci_output = 0, Optional ri_output = 0, Optional StandNum = 1, Optional SESD = True, Optional PV = True, Optional CIs = 95, Optional Intercepts = True, Optional RSquare = True, Optional Heading1 = "Table X.", Optional Heading2 = "", Optional note_text = "", Optional ByVal DecimalPlaces = 2, Optional include_indicator = False)
' Creates the SEM table
    Dim x_start: x_start = ActiveCell.Column
    Dim y_start: y_start = ActiveCell.Row
    Dim y_offset: y_offset = 0
    Dim x_offset: x_offset = 0
    
    Dim nrow: nrow = 1 ' Number of rows in the grid
    Dim ncol: ncol = 1 ' Number of columns in the grid
    Dim nadr: nadr = 1 ' Number of adjacent rows in the grid
    Dim nadc: nadc = 1 ' Number of adjacent columns in the grid
    
    Dim model_grid()   ' The model grid
    Dim stand_grid()   ' A list of standardisations for the grid
    Dim ci_grid()      ' A list of whether the grid shows estimates or CIs
    Dim ri_grid()      ' A list of whether the grid shows estimates or relative importance
    
    Dim col_headings() ' Column headings, based on a list of DVs
    Dim row_headings() ' Row headings, based on a list of IVs

' Create formatting text
    If DecimalPlaces = 0 Then ftext = "0"
    If DecimalPlaces = 1 Then ftext = ".0"
    If DecimalPlaces = 2 Then ftext = ".00"
    If DecimalPlaces = 3 Then ftext = ".000"
    
' Define confidence limits
    If CIs = 90 Then
        ci_lower = "050"
        ci_upper = "950"
    End If
    If CIs = 95 Then
        ci_lower = "025"
        ci_upper = "975"
    End If
    If CIs = 99 Then
        ci_lower = "005"
        ci_upper = "995"
    End If

' Create the model grid
    If group_output = 0 And level_output = 0 And mixture_output = 0 And stand_output = 0 And ci_output = 0 And ri_output = 0 Then
        ReDim model_grid(1, 1, 1, 1)
        ReDim stand_grid(1, 1, 1, 1)
        ReDim ci_grid(1, 1, 1, 1)
        ReDim ri_grid(1, 1, 1, 1)
        
        model_grid(1, 1, 1, 1) = ModelNum
        stand_grid(1, 1, 1, 1) = StandNum
        ci_grid(1, 1, 1, 1) = 1
        ri_grid(1, 1, 1, 1) = 1
    Else
        ' Assign specific model matrix numbers to each part of the grid, as well as standardizatins and CIs
        If group_output = 1 Then nrow = MplusOutput.nGroups
        If level_output = 1 Then nrow = MplusOutput.nLevels
        If mixture_output = 1 Then nrow = MplusOutput.nMixtures
        If stand_output = 1 Then nrow = MplusOutput.nStandards
        If ci_output = 1 Or ri_output = 1 Then nrow = 2
    
        If group_output = 2 Then ncol = MplusOutput.nGroups
        If level_output = 2 Then ncol = MplusOutput.nLevels
        If mixture_output = 2 Then ncol = MplusOutput.nMixtures
        If stand_output = 2 Then ncol = MplusOutput.nStandards
        If ci_output = 2 Or ri_output = 2 Then ncol = 2
    
        If group_output = 3 Then nadr = MplusOutput.nGroups
        If level_output = 3 Then nadr = MplusOutput.nLevels
        If mixture_output = 3 Then nadr = MplusOutput.nMixtures
        If stand_output = 3 Then nadr = MplusOutput.nStandards
        If ci_output = 3 Or ri_output = 3 Then nadr = 2

        If group_output = 4 Then nadc = MplusOutput.nGroups
        If level_output = 4 Then nadc = MplusOutput.nLevels
        If mixture_output = 4 Then nadc = MplusOutput.nMixtures
        If stand_output = 4 Then nadc = MplusOutput.nStandards
        If ci_output = 4 Or ri_output = 4 Then nadc = 2

        ReDim model_grid(1 To nrow, 1 To ncol, 1 To nadr, 1 To nadc)
        ReDim stand_grid(1 To nrow, 1 To ncol, 1 To nadr, 1 To nadc)
        ReDim ci_grid(1 To nrow, 1 To ncol, 1 To nadr, 1 To nadc)
        ReDim ri_grid(1 To nrow, 1 To ncol, 1 To nadr, 1 To nadc)
                
        temp_group = 1
        temp_level = 1
        temp_mixture = 1
        temp_stand = 1
        temp_ci = 1
        temp_ri = 1
                
        For y = 1 To nrow
            If group_output = 1 Then temp_group = y
            If level_output = 1 Then temp_level = y
            If mixture_output = 1 Then temp_mixture = y
            If stand_output = 1 Then temp_stand = y
            If ci_output = 1 Then temp_ci = y
            If ri_output = 1 Then temp_ri = y
            For x = 1 To ncol
                If group_output = 2 Then temp_group = x
                If level_output = 2 Then temp_level = x
                If mixture_output = 2 Then temp_mixture = x
                If stand_output = 2 Then temp_stand = x
                If ci_output = 2 Then temp_ci = x
                If ri_output = 2 Then temp_ri = x
                For w = 1 To nadr
                    If group_output = 3 Then temp_group = w
                    If level_output = 3 Then temp_level = w
                    If mixture_output = 3 Then temp_mixture = w
                    If stand_output = 3 Then temp_stand = w
                    If ci_output = 3 Then temp_ci = w
                    If ri_output = 3 Then temp_ri = w
                    For v = 1 To nadc
                        If group_output = 4 Then temp_group = v
                        If level_output = 4 Then temp_level = v
                        If mixture_output = 4 Then temp_mixture = v
                        If stand_output = 4 Then temp_stand = v
                        If ci_output = 4 Then temp_ci = v
                        If ri_output = 4 Then temp_ri = v
                        model_grid(y, x, w, v) = MplusOutput.ModelNum(temp_group, temp_level, temp_mixture)
                        stand_grid(y, x, w, v) = temp_stand
                        ci_grid(y, x, w, v) = temp_ci
                        ri_grid(y, x, w, v) = temp_ri
                        Debug.Print "Model number: " & model_grid(y, x, w, v) & ", group:" & temp_group & " level:" & temp_level & " mixture:" & temp_mixture & " standardization:" & temp_stand & " CI:" & temp_ci & " RI:" & temp_ri
                    Next
                Next
            Next
        Next
    End If

    ' Make a list of dependent variables
    ' First, populate the column headings
    ReDim col_headings(1 To nrow, 1 To ncol, 1 To nadr, 1 To nadc)
    col_headings = ARRAY_POPULATE(col_headings, "")
    
    ' Populate the first row of column headings
    For x = 1 To ncol
        For v = 1 To MplusOutput.DV(0, model_grid(1, x, 1, 1))
            var_num1 = MplusOutput.DV(v, model_grid(1, x, 1, 1))
            col_headings(1, x, 1, 1) = ADD_STRINGVECTOR(col_headings(1, x, 1, 1), var_num1, MplusOutput.VarName(var_num1))
        Next
        Debug.Print "Original column headings: " & col_headings(1, x, 1, 1)
    Next
    
    ' Duplicate the first row, but replace entries with duplicate latent variables
    For x = 1 To ncol
        For y = 1 To nrow
            For w = 1 To nadr
                For v = 1 To nadc
                    If y > 1 Or w > 1 Or v > 1 Then
                        col_headings(y, x, w, v) = EMPTY_STRINGVECTOR(COUNT_STRINGVECTOR(col_headings(1, x, 1, 1)), "")
                        ' If there any latent variables in the first set of headings, replace with matching ones
                        For u = 1 To COUNT_STRINGVECTOR(col_headings(1, x, 1, 1))
                            var_num1 = GET_STRINGVECTOR(col_headings(1, x, 1, 1), u)
                            If MplusOutput.IsFactor(var_num1) = True Then
                                var_num2 = MplusOutput.FactorMatch(var_num1, model_grid(1, x, 1, 1), model_grid(y, x, w, v))
                                Debug.Print "Matching: " & var_num1 & " (" & MplusOutput.VarName(var_num1) & ") " & var_num2 & " (" & MplusOutput.VarName(var_num2) & ")"
                                If var_num2 > 0 Then
                                    Debug.Print col_headings(1, x, 1, 1); "(original)"
                                    col_headings(y, x, w, v) = ADD_STRINGVECTOR(col_headings(y, x, w, v), MplusOutput.VarName(var_num2) & ":" & var_num2, u, True)
                                    Debug.Print col_headings(y, x, w, v); "(variant)"
                                End If
                            End If
                        Next
                        
                        ' Add extra variables that were not in the first row
                        Debug.Print model_grid(y, x, w, v)
                        For u = 1 To MplusOutput.DV(0, model_grid(y, x, w, v))
                            var_num1 = MplusOutput.DV(u, model_grid(y, x, w, v))
                            
                            ' Check whether it's already on our master list. If not, add it to the list:
                            If GET_STRINGVECTOR(col_headings(1, x, 1, 1), MplusOutput.VarName(var_num1)) = "NA" Then
                            
                                ' Check if it's a factor
                                If MplusOutput.IsFactor(var_num1) = False Then
                                    'MsgBox "Adding new observed variable: " & var_num1 & " " & MPlusOutput.VarName(var_num1)
                                    col_headings(1, x, 1, 1) = ADD_STRINGVECTOR(col_headings(1, x, 1, 1), var_num1, MplusOutput.VarName(var_num1))
                                    ' MsgBox col_headings(1, x, 1) & " " & col_headings(y, x, w) & " after change to secondary row"
                                Else
                                    var_num2 = MplusOutput.FactorMatch(var_num1, model_grid(y, x, w, v), model_grid(1, x, 1, 1))
                                    
                                    'temp_text = "Trying to match factors: " & var_num1 & " (" & MPlusOutput.VarName(var_num1) & ") with " & var_num2
                                    'If var_num2 > 0 Then temp_text = temp_text & " (" & MPlusOutput.VarName(var_num2) & ")"
                                    'temp_text = temp_text & " from models " & model_grid(y, x, w, v) & " and " & model_grid(1, x, 1, 1)
                                    'MsgBox temp_text
                                    
                                    If var_num2 = 0 Then ' No match with an existing factor, so add the new latent variable to the master set of headings
                                        Debug.Print "Adding new latent variable: " & var_num1 & " " & MplusOutput.VarName(var_num1)
                                        col_headings(1, x, 1, 1) = ADD_STRINGVECTOR(col_headings(1, x, 1, 1), var_num1, MplusOutput.VarName(var_num1))
                                    Else ' There is a match with an existing factor, so replace the old name with the new.
                                        Debug.Print "Found match!"
                                        Debug.Print "Replacing latent variable in the modified list."
                                        var_pos = POSVAL_STRINGVECTOR(col_headings(1, x, 1, 1), var_num2)
                                        col_headings(y, x, w, v) = ADD_STRINGVECTOR(col_headings(y, x, w, v), MplusOutput.VarName(var_num1) & ":" & var_num1, var_pos, True)
                                    End If
                                End If
                            End If
                            Debug.Print col_headings(1, x, 1, 1); "(row 1)"
                            Debug.Print col_headings(y, x, w, v); "(row " & y & ", variant " & w & " " & v & ")"
                        Next
                    End If
                Next
            Next
        Next
    Next

    ' Make a list of independent variables
    ' First, populate the column headings
    ReDim row_headings(1 To nrow, 1 To ncol, 1 To nadr, 1 To nadc)
    row_headings = ARRAY_POPULATE(row_headings, "")
    
    ' Populate the first row of column headings
    For y = 1 To nrow
        If Intercepts = True Then row_headings(y, 1, 1, 1) = ADD_STRINGVECTOR(row_headings(y, 1, 1, 1), 0, "Constant", True)
        For v = 1 To MplusOutput.IV(0, model_grid(y, 1, 1, 1))
            var_num1 = MplusOutput.IV(v, model_grid(y, 1, 1, 1))
            row_headings(y, 1, 1, 1) = ADD_STRINGVECTOR(row_headings(y, 1, 1, 1), var_num1, MplusOutput.VarName(var_num1))
        Next
        Debug.Print "Original row headings: " & row_headings(y, 1, 1, 1)
    Next
    
    ' Duplicate the first row, but replace entries with duplicate latent variables
    For y = 1 To nrow
        For x = 1 To ncol
            For w = 1 To nadr
                For v = 1 To nadc
                    If x > 1 Or w > 1 Or v > 1 Then
                        row_headings(y, x, w, v) = EMPTY_STRINGVECTOR(COUNT_STRINGVECTOR(row_headings(1, x, 1, 1)), "")
                        start_n = 1
                        If Intercepts = True Then start_n = 2

                        ' If there any latent variables in the first set of headings, replace with matching ones
                        For u = start_n To COUNT_STRINGVECTOR(row_headings(y, 1, 1, 1))
                            var_num1 = GET_STRINGVECTOR(row_headings(y, 1, 1, 1), u)
                            If MplusOutput.IsFactor(var_num1) = True Then
                                var_num2 = MplusOutput.FactorMatch(var_num1, model_grid(y, 1, 1, 1), model_grid(y, x, w, v))
                                Debug.Print "Matching: " & var_num1 & " (" & MplusOutput.VarName(var_num1) & ") " & var_num2 & " (" & MplusOutput.VarName(var_num2) & ")"
                                If var_num2 > 0 Then
                                    Debug.Print row_headings(y, 1, 1, 1); "(original)"
                                    row_headings(y, x, w, v) = ADD_STRINGVECTOR(row_headings(y, x, w, v), MplusOutput.VarName(var_num2) & ":" & var_num2, u, True)
                                    Debug.Print row_headings(y, x, w, v); "(variant)"
                                End If
                            End If
                        Next
                        
                        ' Add extra variables that were not in the first row
                        Debug.Print model_grid(y, x, w, v)
                        For u = 1 To MplusOutput.IV(0, model_grid(y, x, w, v))
                            var_num1 = MplusOutput.IV(u, model_grid(y, x, w, v))
                            
                            'MsgBox "Number of IVs: " & MplusOutput.IV(0, model_grid(y, x, w, v)) & ", IV name: " & MplusOutput.VarName(var_num1) & " " & start_n
                            
                            ' Check whether it's already on our master list. If not, add it to the list:
                            If GET_STRINGVECTOR(row_headings(y, 1, 1, 1), MplusOutput.VarName(var_num1)) = "NA" Then
                            
                                ' Check if it's a factor
                                If MplusOutput.IsFactor(var_num1) = False Then
                                    'MsgBox "Adding new observed variable: " & var_num1 & " " & MplusOutput.VarName(var_num1)
                                    row_headings(y, 1, 1, 1) = ADD_STRINGVECTOR(row_headings(y, 1, 1, 1), var_num1, MplusOutput.VarName(var_num1))
                                    ' MsgBox row_headings(1, x, 1) & " " & row_headings(y, x, w) & " after change to secondary row"
                                Else
                                    var_num2 = MplusOutput.FactorMatch(var_num1, model_grid(y, x, w, v), model_grid(y, 1, 1, 1))
                                    If var_num2 = 0 Then ' No match with an existing factor, so add the new latent variable to the master set of headings
                                        Debug.Print "Adding new latent variable: " & var_num1 & " " & MplusOutput.VarName(var_num1)
                                        row_headings(y, 1, 1, 1) = ADD_STRINGVECTOR(row_headings(1, x, 1, 1), var_num1, MplusOutput.VarName(var_num1))
                                    Else ' There is a match with an existing factor, so replace the old name with the new.
                                        Debug.Print "Replacing latent variable in the modified list."
                                        var_pos = POSVAL_STRINGVECTOR(row_headings(y, 1, 1, 1), var_num2)
                                        row_headings(y, x, w, v) = ADD_STRINGVECTOR(row_headings(y, x, w, v), MplusOutput.VarName(var_num1) & ":" & var_num1, var_pos, True)
                                    End If
                                    Debug.Print row_headings(y, 1, 1, 1); "(row 1)"
                                    Debug.Print row_headings(y, x, w, v); "(row " & y & ", variant " & w & " " & v & ")"
                                End If
                            End If
                        Next
                    End If
                Next
            Next
        Next
    Next
        
    ' Populate the table data
    ' First, do the headings

    ' Insert Table Headings
    If Heading1 <> "" Then
        With Cells(y_start + y_offset, x_start)
            .Value = Heading1
            .Font.Name = "Times New Roman"
            .Font.Size = 12
        End With
        y_offset = y_offset + 1
    End If
    
    If Heading2 <> "" Then
        With Cells(y_start + y_offset, x_start)
            .Value = Heading2
            .Font.Italic = True
            .Font.Name = "Times New Roman"
            .Font.Size = 12
        End With
        y_offset = y_offset + 1
    End If
    
    ' Insert top table border
    x_end = -1
    For x1 = 1 To ncol
        x_end = x_end + COUNT_STRINGVECTOR(col_headings(1, x1, 1, 1)) * nadc + 1
    Next
    With Range(Cells(y_start + y_offset, x_start), Cells(y_start + y_offset, x_start + x_end))
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeTop).weight = xlThin
    End With
    
    ' Insert column headings
    x_offset2 = x_offset + 1
    For x = 1 To ncol
        c_off = 0
        For u = 1 To COUNT_STRINGVECTOR(col_headings(1, x, 1, 1))
            With Cells(y_start + y_offset, x_start + x_offset2 + c_off)
                .Value = MplusOutput.VarName(GET_STRINGVECTOR(col_headings(1, x, 1, 1), u))
                .Font.Name = "Times New Roman"
                .Font.Size = 12
            End With
            With Range(Cells(y_start + y_offset, x_start + x_offset2 + c_off), Cells(y_start + y_offset, x_start + x_offset2 + c_off + nadc - 1))
                .HorizontalAlignment = xlCenterAcrossSelection
                .Borders(xlEdgeTop).LineStyle = xlContinuous
                .Borders(xlEdgeTop).weight = xlThin
                .Borders(xlEdgeBottom).LineStyle = xlContinuous
                .Borders(xlEdgeBottom).weight = xlThin
            End With
            If nadc > 1 Then
                For t = 1 To nadc
                    If group_output = 4 Then temp_val = MplusOutput.GroupName(t)
                    If level_output = 4 Then temp_val = MplusOutput.LevelName(t)
                    If mixture_output = 4 Then temp_val = MplusOutput.MixtureName(t)
                    If stand_output = 4 Then temp_val = MplusOutput.StdName(t)
                    If (ci_output = 4 Or ri_output = 4) And t = 1 Then temp_val = "B"
                    If ci_output = 4 And t = 2 Then temp_val = CIs & "% CI"
                    If ri_output = 4 And t = 2 Then temp_val = "Relative Importance"
                    With Cells(y_start + y_offset + 1, x_start + x_offset2 + c_off + t - 1)
                        .Value = temp_val
                        .HorizontalAlignment = xlCenter
                        .Font.Name = "Times New Roman"
                        .Font.Size = 12
                        .Borders(xlEdgeTop).LineStyle = xlContinuous
                        .Borders(xlEdgeTop).weight = xlThin
                        .Borders(xlEdgeBottom).LineStyle = xlContinuous
                        .Borders(xlEdgeBottom).weight = xlThin
                    End With
                Next
            End If
            c_off = c_off + nadc
        Next
        x_offset2 = x_offset2 + COUNT_STRINGVECTOR(col_headings(1, x, 1, 1)) * nadc + 1
    Next
    
    y_offset2 = y_offset + 1
    If nadc > 1 Then y_offset2 = y_offset2 + 1
    For y = 1 To nrow
        ' Add the name of the model if there is more than 1
        If nrow > 1 Or ncol > 1 Then
            x_offset = 1
            For x = 1 To ncol
                temp_name = ""
                If group_output = 1 Then temp_name = temp_name & MplusOutput.GroupName(y)
                If level_output = 1 Then temp_name = temp_name & MplusOutput.LevelName(y)
                If mixture_output = 1 Then temp_name = temp_name & MplusOutput.MixtureName(y)
                If stand_output = 1 Then temp_name = temp_name & MplusOutput.StdName(y)
                If (ci_output = 1 Or ri_output = 1) And y = 1 Then temp_name = temp_name & "Estimate"
                If ci_output = 1 And y = 2 Then temp_name = temp_name & CIs & "% CI"
                If ri_output = 1 And y = 2 Then temp_name = temp_name & "Relative Importance"
                If Len(temp_name) > 0 And ncol > 1 Then temp_name = temp_name & " "
                
                If group_output = 2 Then temp_name = temp_name & MplusOutput.GroupName(x)
                If level_output = 2 Then temp_name = temp_name & MplusOutput.LevelName(x)
                If mixture_output = 2 Then temp_name = temp_name & MplusOutput.MixtureName(x)
                If stand_output = 2 Then temp_name = temp_name & MplusOutput.StdName(x)
                If (ci_output = 2 Or ri_output = 2) And y = 1 Then temp_name = temp_name & "Estimate"
                If ci_output = 2 And y = 2 Then temp_name = temp_name & CIs & "% CI"
                If ri_output = 2 And y = 2 Then temp_name = temp_name & "Relative Importance"
                
                With Cells(y_start + y_offset2, x_start + x_offset)
                    .Value = temp_name
                    .Font.Name = "Times New Roman"
                    .Font.Size = 12
                End With
                Range(Cells(y_start + y_offset2, x_start + x_offset), Cells(y_start + y_offset2, x_start + x_offset + COUNT_STRINGVECTOR(col_headings(1, x, 1, 1)) * nadc - 1)).HorizontalAlignment = xlCenterAcrossSelection
                If y > 1 Then
                    With Range(Cells(y_start + y_offset2, x_start), Cells(y_start + y_offset2, x_start + x_end))
                        .Borders(xlEdgeTop).LineStyle = xlContinuous
                        .Borders(xlEdgeTop).weight = xlThin
                    End With
                End If
                x_offset = x_offset + COUNT_STRINGVECTOR(col_headings(1, x, 1, 1)) * nadc + 1
            Next
            y_offset2 = y_offset2 + 1
        End If
        
        ' Add the independent variables
        For u = 1 To COUNT_STRINGVECTOR(row_headings(y, 1, 1, 1))
            x_offset = 0
            var_num = GET_STRINGVECTOR(row_headings(y, 1, 1, 1), u)
            If var_num = 0 Then
                With Cells(y_start + y_offset2, x_start + x_offset)
                    .Value = "Constant"
                    .Font.Name = "Times New Roman"
                    .Font.Size = 12
                    .VerticalAlignment = xlTop
                End With
            Else
                With Cells(y_start + y_offset2, x_start + x_offset)
                    .Value = MplusOutput.VarName(var_num)
                    .Font.Name = "Times New Roman"
                    .Font.Size = 12
                    .VerticalAlignment = xlTop
                End With
            End If
            'If nrow > 1 Or ncol > 1 Then y_offset2 = y_offset2 + 1
            y_offset2 = y_offset2 + 1
            If nadr > 1 Then
                For t = 1 To nadr
                    If group_output = 3 Then temp_val = MplusOutput.GroupName(t)
                    If level_output = 3 Then temp_val = MplusOutput.LevelName(t)
                    If mixture_output = 3 Then temp_val = MplusOutput.MixtureName(t)
                    If stand_output = 3 Then temp_val = MplusOutput.StdName(t)
                    If (ci_output = 3 Or ri_output = 3) And t = 1 Then temp_val = "B"
                    If ci_output = 3 And t = 2 Then temp_val = CIs & "% CI"
                    If ri_output = 3 And t = 2 Then temp_val = "Relative Importance"
                    With Cells(y_start + y_offset2, x_start + x_offset)
                        .Value = temp_val
                        .HorizontalAlignment = xlLeft
                        .VerticalAlignment = xlTop
                        .IndentLevel = 1
                        .Font.Name = "Times New Roman"
                        .Font.Size = 12
                    End With
                    y_offset2 = y_offset2 + 1
                Next
            End If
        Next
    Next
    Columns(x_start + x_offset).ColumnWidth = 20
    Columns(x_start + x_offset).EntireColumn.AutoFit
    
    ' Add the model results
    y_offset = y_offset + 1
    If ncol > 1 Or nrow > 1 Then y_offset = y_offset + 1
    If nadc > 1 Then y_offset = y_offset + 1
    For y = 1 To nrow
        'y_offset2 = y_offset
        x_offset = 0
        Debug.Print "Row"; y; "of"; nrow
        For x = 1 To ncol
            x_offset2 = x_offset
            Debug.Print "Column"; x; "of"; ncol
            Debug.Print "Column headings:"; COUNT_STRINGVECTOR(col_headings(1, x, 1, 1))
            For t = 1 To COUNT_STRINGVECTOR(col_headings(1, x, 1, 1))
                y_offset2 = y_offset
                'Cells(y_start + y_offset, x_start + x_offset - 1).Value = "Start from here..."
                Debug.Print "Row headings:"; COUNT_STRINGVECTOR(row_headings(y, 1, 1, 1))
                For s = 1 To COUNT_STRINGVECTOR(row_headings(y, 1, 1, 1))
                    For w = 1 To nadr
                        If nadr > 1 And w = 1 Then y_offset2 = y_offset2 + 1
                        For v = 1 To nadc
                            model_num = model_grid(y, x, w, v)
                            Stand_Num = stand_grid(y, x, w, v)
                            ci_num = ci_grid(y, x, w, v)
                            ri_num = ri_grid(y, x, w, v)
                            dv_num = GET_STRINGVECTOR(col_headings(y, x, w, v), t)
                            If dv_num = "" Or dv_num = "NA" Then dv_num = GET_STRINGVECTOR(col_headings(1, x, 1, 1), t)
                            iv_num = GET_STRINGVECTOR(row_headings(y, x, w, v), s)
                            If iv_num = "" Or iv_num = "NA" Then iv_num = GET_STRINGVECTOR(row_headings(y, 1, 1, 1), s)
                            If ci_num = 1 And ri_num = 1 Then ' Add estimate
                                temp_text = "="
                                If iv_num = 0 Then
                                    For u = 1 To MplusOutput.InterceptNCategories(dv_num, Stand_Num, model_num)
                                        temp_val = MplusOutput.Intercept(dv_num, Stand_Num, model_num, u)
                                        If u > 1 Then temp_text = temp_text & " & CHAR(10) & "
                                        temp_text = temp_text & "text(" & temp_val & ", """ & ftext & """)"
                                        If PV = True Then
                                            temp_p = MplusOutput.InterceptP(dv_num, Stand_Num, model_num, u)
                                            If temp_p <> "NA" Then temp_text = temp_text & " & """ & asterisk_pval(temp_p) & """"
                                        End If
                                        If SESD = True Then
                                           temp_SE = MplusOutput.InterceptSE(dv_num, Stand_Num, model_num, u)
                                            If temp_SE <> "NA" Then temp_text = temp_text & "& "" ("" & text(" & temp_SE & ",""" & ftext & """) & "")"""
                                        End If
                                    Next
                                Else
                                    For u = 1 To MplusOutput.PathNCategories(iv_num, dv_num, Stand_Num, model_num)
                                        temp_val = MplusOutput.Path(iv_num, dv_num, Stand_Num, model_num, u)
                                        If u > 1 Then temp_text = temp_text & " & CHAR(10) & "
                                        temp_text = temp_text & "text(" & temp_val & ", """ & ftext & """)"
                                        If PV = True Then
                                            temp_p = MplusOutput.PathP(iv_num, dv_num, Stand_Num, model_num, u)
                                            If temp_p <> "NA" Then temp_text = temp_text & " & """ & asterisk_pval(temp_p) & """"
                                        End If
                                        If SESD = True Then
                                            temp_SE = MplusOutput.PathSE(iv_num, dv_num, Stand_Num, model_num, u)
                                            If temp_SE <> "NA" Then temp_text = temp_text & "& "" ("" & text(" & temp_SE & ",""" & ftext & """) & "")"""
                                        End If
                                    Next
                                End If
                                With Cells(y_start + y_offset2, x_start + x_offset2 + v)
                                    If temp_val = "NA" Or CStr(temp_val) = "999" Then
                                        .formula = "–"
                                    Else
                                    .formula = temp_text
                                    End If
                                    .VerticalAlignment = xlTop
                                    .HorizontalAlignment = xlRight
                                    .Font.Name = "Times New Roman"
                                    .Font.Size = 12
                                    .WrapText = True
                                End With
                                Debug.Print "Adding estimate: " & temp_text
                            End If
                            If ci_num > 1 Then ' Add confidence intervals
                                temp_text = "="
                                    If iv_num = 0 Then
                                        For u = 1 To MplusOutput.InterceptNCategories(dv_num, Stand_Num, model_num)
                                            If u > 1 Then temp_text = temp_text & " & CHAR(10) & "
                                            If MplusOutput.CI_Intercept(dv_num, Stand_Num, model_num, ci_lower) <> "NA" And CStr(MplusOutput.CI_Intercept(dv_num, Stand_Num, model_num, ci_lower)) <> "999" Then
                                                temp_text = temp_text & """["" & text(" & MplusOutput.CI_Intercept(dv_num, Stand_Num, model_num, ci_lower, u) & ",""" & ftext & """) & "", "" & text(" & MplusOutput.CI_Intercept(dv_num, Stand_Num, model_num, ci_upper, u) & ",""" & ftext & """) & ""]"""
                                            Else
                                                temp_text = temp_text & """[–, –]"""
                                            End If
                                        Next
                                    Else
                                        For u = 1 To MplusOutput.PathNCategories(iv_num, dv_num, Stand_Num, model_num)
                                            If u > 1 Then temp_text = temp_text & " & CHAR(10) & "
                                            If MplusOutput.CI(iv_num, dv_num, Stand_Num, model_num, ci_lower) <> "NA" And CStr(MplusOutput.CI(iv_num, dv_num, Stand_Num, model_num, ci_lower)) <> "999" Then
                                                temp_text = temp_text & """["" & text(" & MplusOutput.CI(iv_num, dv_num, Stand_Num, model_num, ci_lower, u) & ",""" & ftext & """) & "", "" & text(" & MplusOutput.CI(iv_num, dv_num, Stand_Num, model_num, ci_upper, u) & ",""" & ftext & """) & ""]"""
                                            Else
                                                temp_text = temp_text & """[–, –]"""
                                            End If
                                        Next
                                    End If
                                MsgBox temp_text
                                With Cells(y_start + y_offset2, x_start + x_offset2 + v)
                                    .formula = temp_text
                                    .VerticalAlignment = xlTop
                                    .HorizontalAlignment = xlRight
                                    .Font.Name = "Times New Roman"
                                    .Font.Size = 12
                                    .WrapText = True
                                End With
                                Debug.Print "Adding CIs: " & temp_text
                            End If
                            If ri_num > 1 Then ' Add Relative Importance
                                If iv_num = 0 Then ' Add nothing for the intercept
                                    temp_text = ""
                                Else
                                    ' Only perform relative importance analysis if the information can be retrieved, and there is only a single category
                                    temp_text = """NA"""
                                    If MplusOutput.PathNCategories(iv_num, dv_num, Stand_Num, model_num) = 1 Then
                                        rel_imp = MplusOutput.RelativeImportance(dv_num, iv_num, Stand_Num, model_num, use_formula)
                                        If rel_imp <> "NA" Then
                                            temp_text = rel_imp
                                        End If
                                    End If
                                End If
                                With Cells(y_start + y_offset2, x_start + x_offset2 + v)
                                    .formula = as_formula(temp_text)
                                    .VerticalAlignment = xlTop
                                    .HorizontalAlignment = xlRight
                                    .Font.Name = "Times New Roman"
                                    .Font.Size = 12
                                    .NumberFormat = "0" & ftext & "%"
                                End With
                                Debug.Print "Adding RIs: " & temp_text
                            End If
                            
                            Columns(x_start + x_offset2 + v).ColumnWidth = 20
                            Columns(x_start + x_offset2 + v).EntireColumn.AutoFit
                        Next
                        y_offset2 = y_offset2 + 1
                    Next
                Next
                x_offset2 = x_offset2 + nadc
            Next
            x_offset = x_offset + COUNT_STRINGVECTOR(col_headings(1, x, 1, 1)) * nadc + 1
            If x < ncol Then Columns(x_start + x_offset).ColumnWidth = 1
        Next
        If nadr > 1 Then
            y_offset = y_offset + COUNT_STRINGVECTOR(row_headings(y, 1, 1, 1)) * (nadr + 1)
        Else
            y_offset = y_offset + COUNT_STRINGVECTOR(row_headings(y, 1, 1, 1))
        End If
        If (ncol > 1 Or nrow > 1) And y < nrow Then y_offset = y_offset + 1 'Make room for heading
    Next
    
    ' Draw line under the table
    With Range(Cells(y_start + y_offset, x_start), Cells(y_start + y_offset, x_start + x_end))
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeTop).weight = xlThin
    End With
    
    ' Create R-Square
    If RSquare = True Then
        row_block_n = 1
        col_block_n = 1
        row_adj_n = 1
        col_adj_n = 1
        
        y_offset2 = y_offset
        x_offset = 0
        
        If nrow > 1 And (group_output = 1 Or level_output = 1) Then
            row_block_n = nrow
        End If
        If ncol > 1 And (group_output = 2 Or level_output = 2) Then
            col_block_n = ncol
        End If
        If nadr > 1 And (group_output = 3 Or level_output = 3) Then
            row_adj_n = nadr
        End If
        If nadc > 1 And (group_output = 4 Or level_output = 4) Then
            col_adj_n = nadc
        End If
        
        ' First come up with a list of row names
        If row_block_n = 1 And row_adj_n = 1 Then
            With Cells(y_start + y_offset, x_start + x_offset)
                .Value = "R²"
                .HorizontalAlignment = xlLeft
                .Font.Name = "Times New Roman"
                .Font.Size = 12
            End With
        Else
            For a = 1 To row_block_n
                For b = 1 To row_adj_n
                    temp_text = "R² ("
                    If group_output = 1 Then temp_text = temp_text & MplusOutput.GroupName(a)
                    If level_output = 1 Then temp_text = temp_text & MplusOutput.LevelName(a)
                    If Len(temp_text) > 4 And (group_output = 3 Or level_output = 3) Then temp_text = temp_text & " "
                    If group_output = 3 Then temp_text = temp_text & MplusOutput.GroupName(b)
                    If level_output = 3 Then temp_text = temp_text & MplusOutput.LevelName(b)
                    temp_text = temp_text & ")"
                    With Cells(y_start + y_offset2, x_start + x_offset)
                        .Value = temp_text
                        .HorizontalAlignment = xlLeft
                        .Font.Name = "Times New Roman"
                        .Font.Size = 12
                    End With
                    y_offset2 = y_offset2 + 1
                Next
            Next
        End If
        
        y_offset2 = y_offset
        x_offset2 = x_offset + 1
        For x1 = 1 To ncol ' Number of column blocks
            For dv_cnt = 1 To COUNT_STRINGVECTOR(col_headings(1, x1, 1, 1)) ' Number of DVs
                For x2 = 1 To nadc ' Number of adjacent columns
                    For y1 = 1 To nrow ' Number of rows
                        For y2 = 1 To nadr ' Number of adjacent rows
                            model_num = model_grid(y1, x1, y2, x2)
                            
                            dv_num = GET_STRINGVECTOR(col_headings(y1, x1, y2, x2), dv_cnt)
                            If dv_num = "NA" Or dv_num = "" Then dv_num = GET_STRINGVECTOR(col_headings(1, x1, 1, 1), dv_cnt)
                            
                            'MsgBox col_headings(y1, x1, y2, x2) & Chr(10) & col_headings(1, x1, 1, 1) & Chr(10) & dv_cnt & " " & dv_num
                            
                            ' Only print the R-Square if particular circumstances are met
                            print_rsq = False
                            If y1 = 1 Or (y1 > 1 And row_block_n > 1) Then
                                If y2 = 1 Or (y2 > 1 And row_adj_n > 1) Then
                                    'If x1 = 1 Or (x1 > 1 And col_block_n > 1) Then
                                        If x2 = 1 Or (x2 > 1 And col_adj_n > 1) Then
                                            print_rsq = True
                                        End If
                                    'End If
                                End If
                            End If
                                                        
                            If print_rsq = True Then
                                temp_text = MplusOutput.RSquare(dv_num, model_num)
                                With Cells(y_start + y_offset2, x_start + x_offset2)
                                    If temp_text = "NA" Or temp_text = "Undefined" Then
                                        .Value = "–"
                                    Else
                                    .formula = "=text(" & temp_text & ",""" & ftext & """)"
                                    End If
                                    .HorizontalAlignment = xlRight
                                    .Font.Name = "Times New Roman"
                                    .Font.Size = 12
                                End With
                                y_offset2 = y_offset2 + 1
                            End If
                        Next
                    Next
                    x_offset2 = x_offset2 + 1
                    y_offset2 = y_offset
                Next
            Next
            x_offset2 = x_offset2 + 1
        Next
        
        ' Draw line under the table
        With Range(Cells(y_start + y_offset + row_block_n * row_adj_n, x_start), Cells(y_start + y_offset + row_block_n * row_adj_n, x_start + x_end))
            .Borders(xlEdgeTop).LineStyle = xlContinuous
            .Borders(xlEdgeTop).weight = xlThin
        End With
    End If
        
End Sub
