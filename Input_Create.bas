Attribute VB_Name = "Input_Create"
Dim model_syntax
Dim model_define
Dim model_type
Dim model_constraint
Dim model_priors
Dim usevariables
Dim model_data
Dim lb

' This module is designed to create a syntax file and data file for use with Mplus
Sub PrepareMPlus()
    ' Set defaults
    model_syntax = ""
    usevariables = ""
    model_priors = ""
    model_constraint = ""
    model_define = ""
    model_data = ""
    model_type = "TYPE IS BASIC;"
    nreps = 1

    Set DataStructure = New cDataStructure
    DataStructure.CreateDataStructure
    
    InputForm.Show
    If InputForm.execute = False Then End
    
    ' Call the CFA table if CFA syntax has been requested
    If InputForm.CFASyntax = True Then
        InputForm_CFA.Show
        If InputForm_CFA.execute = True Then CFA_model_syntax
        Unload InputForm_CFA
    End If
    
    ' Set up Monte Carlo syntax if requested
    If InputForm.MCSyntax = True Then
        model_data = "TYPE = MONTECARLO;" & lb
        data_filename = STEM_FILE(InputForm.data_filename) & ".dat"
        nreps = InputForm.MCDatasets
        model_type = ""
        model_syntax = "!Insert model here"
    Else
        data_filename = InputForm.data_filename
    End If
        
    If InputForm.create_datafile = True Then Call CreateDataFile(DataStructure, data_filename, InputForm.export_to.text, nreps)
    If InputForm.create_syntax = True Then Call CreateInputSyntax(DataStructure, _
                                                                  data_filename, _
                                                                  InputForm.input_filename, _
                                                                  InputForm.var_labels, _
                                                                  InputForm.scale_names, _
                                                                  InputForm.CFASyntax, _
                                                                  InputForm.export_to.text, _
                                                                  DataStructure.MissingValue, _
                                                                  "", _
                                                                  model_data, _
                                                                  model_define, _
                                                                  "", _
                                                                  usevariables, _
                                                                  model_type, _
                                                                  model_syntax, _
                                                                  model_constraint, _
                                                                  model_priors, _
                                                                  "", _
                                                                  "", _
                                                                  "", _
                                                                  "", _
                                                                  "")
    Unload InputForm
    
    If InputForm.create_datafile = True And input_create_syntax = True Then MsgBox "Data file and syntax file created successfully."
    If InputForm.create_datafile = True And input_create_syntax = False Then MsgBox "Data file created successfully."
    If InputForm.create_datafile = False And input_create_syntax = True Then MsgBox "Syntax file created successfully."
End Sub

Sub CreateDataFile(data_object, ByVal data_filename, Optional export_to = "Mplus", Optional nreps = 1)
    ' Exports a data file to Mplus or R
    ' If it's Mplus, it saves the file as a CSV with no header
    ' If it's R, it saves the file as a CSV with a variable name header
    ' If nreps > 1 then it generates a Monte Carlo dataset

    Dim file_path: file_path = ActiveWorkbook.Path
    If file_path = "" Then file_path = MyDocsPath()
    
    If nreps > 1 Then ' Set up Monte Carlo file
        Open file_path & "\" & data_filename For Output As 1
        For a = 1 To nreps
            Print #1, STEM_FILE(data_filename) & a & ".csv"
        Next
        Close #1
    End If
        
    For a = 1 To nreps
        If nreps > 1 Then ' Monte Carlo; open a numbered file and refresh random numbers
            Open file_path & "\" & STEM_FILE(data_filename) & a & ".csv" For Output As 1
            ActiveSheet.EnableCalculation = False
            ActiveSheet.EnableCalculation = True
        Else
            Open file_path & "\" & STEM_FILE(data_filename) & ".csv" For Output As 1
        End If
        
        If export_to = "R" Then
        ' If we are creating an R file then write the first row of the file as the variable names
            temp = ""
            For b = 1 To data_object.Variables_N
                If b > 1 Then temp = temp & ","
                temp = temp & data_object.VarName(b)
            Next
            Print #1, temp
        End If
        
        temp = ""
        For b = 1 To data_object.Cases_N
            For c = 1 To data_object.Variables_N
                temp_val = data_object.Dataset(b, c)
                If temp_val = data_object.MissingValue And export_to = "R" Then
                    ' If there is missing data and you are exporting to R, then write NA instead
                    temp = temp & "NA"
                Else
                    temp = temp & temp_val
                End If
                If c < data_object.Variables_N Then temp = temp & ","
            Next
            Print #1, temp
            temp = ""
        Next
    Close #1
    Next
    
End Sub

Sub CreateInputSyntax(data_object, _
                      ByVal data_filename, _
                      input_filename, _
                      Optional var_labels = False, _
                      Optional scale_names = False, _
                      Optional CFA_syntax = False, _
                      Optional export_to = "Mplus", _
                      Optional missing_value = "", _
                      Optional TITLE = "", _
                      Optional DATA = "", _
                      Optional DEFINE = "", _
                      Optional VARIABLE = "", _
                      Optional usevariables = "", _
                      Optional ANALYSIS = "TYPE = BASIC;", _
                      Optional MODEL = "", _
                      Optional MODELCONSTRAINT = "", _
                      Optional MODELPRIORS = "", _
                      Optional OUTPUT = "!sampstat tech4 standardized residual modindices;", _
                      Optional SAVEDATA = "", _
                      Optional PLOT = "", _
                      Optional MONTECARLO = "", _
                      Optional MODELPOPULATION = "")
' Creates input syntax; saves it as a file
   
    If data_object.var_names = False Then
        MsgBox "Error: Variable names must be provided."
        Exit Sub
    End If
    
    ' Locate the file path; if this fails, save to "My Documents"
    Dim file_path: file_path = ActiveWorkbook.Path
    If file_path = "" Then file_path = MyDocsPath()

    Dim y2: y2 = 0
    lb = Chr(10)
    Dim temp_array()
        
    ' Set up defaults
    Dim line: line = ""
    If TITLE = "" Then TITLE = ActiveSheet.Name
    If VARIABLE = "" Then
        line = "NAMES ARE"
        VARIABLE = line
        For a = 1 To data_object.VarName
            'MsgBox data_object.VarName
            line = line & " " & data_object.VarName(a)
            If a = data_object.VarName Then line = line & ";"
            If Len(line) > 89 Then
                VARIABLE = VARIABLE & lb & data_object.VarName(a)
                line = data_object.VarName(a)
            Else
                VARIABLE = VARIABLE & " " & data_object.VarName(a)
            End If
            If a = data_object.VarName Then VARIABLE = VARIABLE & ";"
        Next
    End If
    
' Phase #2 - Generate MPlus / R syntax

' 2.3 Generate MPlus syntax
    If export_to = "Mplus" Then
        DATA = DATA & "FILE IS " & data_filename & ";"
    
        Open file_path & "\" & input_filename For Output As 1
        
        Print #1, "TITLE: " & TITLE
        Print #1, ""
        Print #1, "DATA: " & DATA
        Print #1, ""
        Print #1, "VARIABLE:"
        Print #1, VARIABLE
        If usevariables <> "" Then Print #1, lb & usevariables
        Print #1, ""
        
    ' 2.4 Report missing values
        If missing_value <> "" Then
            Print #1, "MISSING ARE ALL(" & missing_value & ");"
            Print #1, ""
        End If
    
    ' 2.5. Generate defined variables
        If DEFINE <> "" Then
            Print #1, "DEFINE:"
            Print #1, DEFINE
            Print #1, ""
        End If
    
        If ANALYSIS <> "" Then
            Print #1, "ANALYSIS:"
            Print #1, ANALYSIS ' model_type
            Print #1, ""
        End If
        
        If MONTECARLO <> "" Then
            Print #1, "MONTECARLO:"
            Print #1, MONTECARLO
            Print #1, ""
        End If

        If MODELPOPULATION <> "" Then
            Print #1, "MODEL POPULATION:"
            Print #1, MODELPOPULATION
            Print #1, ""
        End If
        
        If MODEL <> "" Then
            Print #1, "MODEL:"
            Print #1, MODEL
            If MODELCONSTRAINT <> "" Then
                Print #1, ""
                Print #1, "MODEL CONSTRAINT:"
                Print #1, MODELCONSTRAINT
            End If
            If MODELPRIORS <> "" Then
                Print #1, ""
                Print #1, "MODEL PRIORS:"
                Print #1, MODELPRIORS
            End If
            Print #1, ""
    '       Print #1, model_syntax & model_priors
        End If
        
        If OUTPUT <> "" Then
            Print #1, "OUTPUT:"
            Print #1, OUTPUT
            Print #1, ""
        End If
        
        If SAVEDATA <> "" Then
            Print #1, "SAVEDATA:"
            Print #1, SAVEDATA
            Print #1, ""
        End If
        
        If PLOT <> "" Then
            Print #1, "PLOT:"
            Print #1, PLOT
            Print #1, ""
        End If
    
        ' Insert factor/scale names
        If scale_names = True Then
            Print #1, ""
            Print #1, "!SCALES:"
                
            For a = 1 To data_object.ScaleName
                temp_text = data_object.ScaleName(a) & ":"
                For b = 1 To data_object.ScaleIndicator(a, 0)
                    temp_text = temp_text & data_object.ScaleIndicator(a, b) & " "
                Next
                Print #1, write_comment(temp_text)
            Next
        End If
    
        If var_labels = True Then
            Print #1, ""
            Print #1, "!LABELS:"
            For a = 1 To data_object.VarName
                If data_object.VarLabel(a) <> "" Then
                    'MsgBox data_object.VarName(34)
                    Print #1, write_comment(data_object.VarName(a) & ":" & data_object.VarLabel(a))
                End If
            Next
            If row_scales > 0 Then
                For a = 1 To data_object.ScaleName
                    'MsgBox data_object.ScaleName(a)
                    'MsgBox data_object.ScaleLabel(a)
                    Print #1, write_comment(data_object.ScaleName(a) & ":" & data_object.ScaleLabel(a))
                Next
            End If
        End If
    
        Close #1
    
    ElseIf export_to = "R" Then
        ' Reformat "usevariables"
        usevariables = Replace(usevariables, "USEVARIABLES ARE", "")
        ' Get rid of line breaks
        usevariables = Replace(usevariables, lb, " ")
        ' Get rid of double spaces
        Do
            temp_len = Len(usevariables)
            usevariables = Replace(usevariables, "  ", " ")
        Loop Until Len(usevariables) = temp_len
        usevariables = Replace(usevariables, ";", "")
        usevariables = Trim(usevariables)
        usevariables = """" & usevariables
        usevariables = Replace(usevariables, " ", """, """)
        usevariables = usevariables & """"
    
        Open file_path & "\" & input_filename For Output As 1
        Print #1, "load.package <- function(x)"
        Print #1, "{"
        Print #1, "  if (!require(x,character.only = TRUE))"
        Print #1, "  {"
        Print #1, "    install.packages(x,dep=TRUE)"
        Print #1, "    if(!require(x,character.only = TRUE)) stop(""Package not found"")"
        Print #1, "  }"
        Print #1, "}"
        Print #1, ""
        Print #1, "load.package (""MplusAutomation"")"
        Print #1, "load.package (""tidyverse"")"
        Print #1, ""
        Print #1, "# Function to obtain the mode"
        Print #1, "getmode <- function(v) {"
        Print #1, "  v <- v[!is.na(v)]"
        Print #1, "  uniqv <- unique(v)"
        Print #1, "  tablv <- tabulate(match(v, uniqv))"
        Print #1, "  if (length(tablv[tablv==max(tablv)]) > 1) return(NA)"
        Print #1, "  else uniqv [which.max(tabulate(match(v, uniqv)))]"
        Print #1, "}"
        Print #1, ""
        Print #1, "# Load dataset"
        Print #1, "dataset <- read.csv(""" & data_filename & """, header=TRUE)"
        Print #1, ""
        Print #1, "# Filter the dataset (e.g., remove observations where group membership is unknown)"
        Print #1, "#dataset <- dataset[!is.na(dataset$group),]"
        Print #1, ""
        Print #1, "# Data preparation functions----"
        Print #1, "# Calculate mean of multiple variables----"
        Print #1, "#dataset$xmean <- rowMeans(subset(dataset, select=x1:x5), na.rm = TRUE)"
        Print #1, ""
        Print #1, "# Aggregation functions----"
        Print #1, "# Create a Level 2 data frame"
        Print #1, "#dataset.level2 <- dataset %>% group_by(group) %>% summarise("
        Print #1, "#  group2 = unique(group2),"
        Print #1, "#  x_grp_mean = mean(x, na.rm = TRUE),      # Group mean of x"
        Print #1, "#  x_grp_sd = sd(x, na.rm = TRUE),          # Group standard deviation of x"
        Print #1, "#  x_grp_median = median(x, na.rm = TRUE),  # Group median of x"
        Print #1, "#  x_grp_mode = getmode(x),                 # Group mode of x"
        Print #1, "#  x_grp_max = max(x, na.rm = TRUE),        # Group maxima (largest value of x in the group)"
        Print #1, "#  x_grp_min = min(x, na.rm = TRUE),        # Group minima (smallest value of x in the group)"
        Print #1, "#  x_grp_sum = sum(x, na.rm = TRUE),        # Group sum of x"
        Print #1, "#  grp_n = length(group),                   # Group size / number of cases in group"
        Print #1, "#  x_missing = sum(is.na(x))                # Number of missing values of x"
        Print #1, "#)"
        Print #1, "# Merge the Level 1 and Level 2 data frames"
        Print #1, "#dataset <- merge(x = dataset, y = dataset.level2, by = c(""group"", ""group2""), all.x = TRUE)"
        Print #1, ""
        Print #1, "# Create a Level 3 data frame from Level 2 data frame"
        Print #1, "#dataset.level3 <- dataset.level2 %>% group_by(group2) %>% summarise("
        Print #1, "#  x_grp_mean2 = mean(x_grp_mean, na.rm = TRUE)    # Group mean of x"
        Print #1, "#)"
        Print #1, "# Merge the Level 1 and Level 3 data frames"
        Print #1, "#dataset <- merge(x = dataset, y = dataset.level3, by = ""group2"", all.x = TRUE)"
        Print #1, ""
        Print #1, "# Grand mean centre----"
        Print #1, "#dataset$x_gmc <- scale(dataset$x, center = TRUE, scale = FALSE) #set scale = TRUE to create Z scores"
        Print #1, ""
        Print #1, "# Group mean center"
        Print #1, "#dataset$x_grpc <- dataset$x - dataset$x_grp_mean       # Mean centre by group (subtract group mean from each case)"
        Print #1, ""
        Print #1, "# Group grand mean centre, subtracting the Level 2 mean of all groups----"
        Print #1, "#dataset$x_grp_mean_cent <- dataset$x_grp_mean - mean(dataset.level2$x_grp_mean, na.rm = TRUE)"
        Print #1, ""
        Print #1, "# Mplus syntax----"
        Print #1, "mplus_object <- mplusObject("
        Print #1, "TITLE = """ & TITLE & ""","
        If DATA <> "" Then
            Print #1, "DATA = """ & DATA & """, "
        Else
            Print #1, "#DATA = "";"", "
        End If
        Print #1, "rdata = dataset,"
        If usevariables <> "" Then
            Print #1, "usevariables = c(" & usevariables & "),"
        Else
            Print #1, "usevariables = c(""""), # Insert variables here"
        End If
        Print #1, "#VARIABLE = "";"", "
        If ANALYSIS <> "" Then
            Print #1, "ANALYSIS = """ & ANALYSIS & """, "
        Else
            Print #1, "#ANALYSIS = "";"", "
        End If
        If DEFINE <> "" Then
            Print #1, "DEFINE = """ & DEFINE & """, "
        Else
            Print #1, "#DEFINE = "";"", "
        End If
        If MODEL <> "" Then
            Print #1, "MODEL = """ & MODEL & """, "
            If MODELCONSTRAINT <> "" Then Print #1, "MODELCONSTRAINT = """ & MODELCONSTRAINT & """, "
            If MODELPRIORS <> "" Then Print #1, "MODELPRIORS = """ & MODELPRIORS & """, "
            Print #1, "#MODELTEST = "";"","
        Else
            Print #1, "#MODEL = "";"", "
            Print #1, "#MODELCONSTRAINT = "";"", "
            Print #1, "#MODELTEST = "";"","
        End If
        If SAVEDATA <> "" Then
            Print #1, "SAVEDATA = """ & SAVEDATA & """, "
        Else
            Print #1, "#SAVEDATA = "";"","
        End If
        If PLOT <> "" Then
            Print #1, "PLOT = """ & PLOT & """, "
        Else
            Print #1, "#PLOT = "";"","
        End If
        Print #1, "OUTPUT = """ & OUTPUT & """"
        Print #1, ")"
        Print #1, ""
        Print #1, "res <- mplusModeler(mplus_object, """ & ActiveSheet.Name & ".dat"", modelout = """ & ActiveSheet.Name & ".inp"", run = 1L)"
        Print #1, ""
        Print #1, "res$results$covariance.coverage"
        Print #1, "res$results$sampstat$means"
        Print #1, "res$results$sampstat$covariances"
        Print #1, "res$results$sampstat$correlations"
        Print #1, "res$results$summaries"
        Print #1, "#res$results$parameters"
        Print #1, ""
        Print #1, "#List of all variables:"
        
        ' Produce the list of all of the variables
        tempvarlist = "varlist <- c("
        tempaliaslist = "labellist <- c("
        For a = 1 To data_object.VarName
            tempvarlist = tempvarlist & """" & data_object.VarName(a) & """"
            tempaliaslist = tempaliaslist & """" & data_object.VarLabel(a) & """"
            If a < data_object.VarName Then
                tempvarlist = tempvarlist & ", "
                tempaliaslist = tempaliaslist & ", "
                If a Mod 8 = 0 Then
                    tempvarlist = tempvarlist & Chr(10)
                    tempaliaslist = tempaliaslist & Chr(10)
                End If
            Else
                tempvarlist = tempvarlist & ")"
                tempaliaslist = tempaliaslist & ")"
            End If
        Next
        Print #1, tempvarlist
        If row_varlabels > 0 Then Print #1, tempaliaslist
        Print #1, ""
    
        Close #1
    End If

    Application.ScreenUpdating = True
    
End Sub

Private Sub CFA_model_syntax()
' This subroutine creates the CFA model syntax

Dim temp_line, n_ind, n_ind_line
Dim fl_count: fl_count = 0

lb = Chr(10)

' Create list of USEVARIABLES
    usevariables = lb & "USEVARIABLES ARE" & lb
    temp_line = ""
    For a = 1 To DataStructure.ScaleInclude
        If InputForm_CFA.LV4 = True Then
            temp_ind = format_mplus(DataStructure.ScaleName(a, True))
            If Len(temp_line) + Len(temp_ind) > 88 Then
                usevariables = usevariables & temp_line & lb
                temp_line = temp_ind
            Else
                If Len(temp_line) > 0 Then temp_line = temp_line & " "
                temp_line = temp_line & temp_ind
            End If
        Else
            For b = 1 To DataStructure.ScaleIndicator(a, 0, True)
                temp_ind = format_mplus(DataStructure.ScaleIndicator(a, b, True))
                If Len(temp_line) + Len(temp_ind) > 88 Then
                    usevariables = usevariables & temp_line & lb
                    temp_line = temp_ind
                Else
                    If Len(temp_line) > 0 Then temp_line = temp_line & " "
                    temp_line = temp_line & temp_ind
                End If
            Next
            usevariables = usevariables & temp_line & lb
            temp_line = ""
        End If
    Next
    If InputForm_CFA.LV4 = True Then usevariables = usevariables & temp_line
    usevariables = usevariables & ";" & lb

' Case #1, normal CFA with no Bayesian small prior factor loadings
    If (InputForm_CFA.LV1 = True Or InputForm_CFA.LV2 = True Or InputForm_CFA.LV3 = True) And InputForm_CFA.BCFA_CL = False Then
        model_type = ""
        n_ind = 0
        For a = 1 To DataStructure.ScaleInclude
            temp_line = DataStructure.ScaleName(a, True) & " BY"
            n_ind_line = 0
            For b = 1 To DataStructure.ScaleIndicator(a, 0, True)
                n_ind_line = n_ind_line + 1
                temp_ind = format_mplus(DataStructure.ScaleIndicator(a, b, True))
                If InputForm_CFA.LV2 = True And b = 1 Then
                    temp_ind = temp_ind & "*"
                ElseIf InputForm_CFA.LV3 = True And DataStructure.ScaleIndicator(a, 0, True) > 1 Then
                    temp_ind = temp_ind & "*1"
                End If
                If Len(temp_line) + Len(temp_ind) < 75 Then
                    temp_line = temp_line & " " & temp_ind
                Else
                    If InputForm_CFA.LV3 = True And DataStructure.ScaleIndicator(a, 0, True) > 1 Then
                        temp_line = temp_line & " (L" & n_ind + 1
                        If n_ind_line > 1 Then
                            temp_line = temp_line & "-L" & n_ind + n_ind_line - 1
                        End If
                        temp_line = temp_line & ")"
                        n_ind = n_ind + n_ind_line - 1
                        n_ind_line = 1
                    End If
                
                    model_syntax = model_syntax & temp_line & lb
                    temp_line = temp_ind
                End If
            Next
                        
            If InputForm_CFA.LV3 = True And DataStructure.ScaleIndicator(a, 0, True) > 1 Then
                temp_line = temp_line & " (L" & n_ind + 1
                If n_ind_line > 1 Then
                    temp_line = temp_line & "-L" & n_ind + n_ind_line
                End If
                temp_line = temp_line & ")"
                n_ind = n_ind + n_ind_line
            End If
            
            model_syntax = model_syntax & temp_line & ";" & lb
        Next
        
        ' Fix variances to 1
        If InputForm_CFA.LV2 = True Then
            model_syntax = model_syntax & DataStructure.ScaleName(1, True) & "-" & DataStructure.ScaleName(DataStructure.ScaleInclude, True) & "@1;" & lb
        End If
        
        ' Set up intercept labelling and factor mean estimation
        If InputForm_CFA.LV3b = True Then
            int_num = 0
            For a = 1 To DataStructure.ScaleInclude
                n_ind = DataStructure.ScaleIndicator(a, 0, True)
                If n_ind > 1 Then
                    model_syntax = model_syntax & "[" & format_mplus(DataStructure.ScaleIndicator(a, 1, True)) & "-" & format_mplus(DataStructure.ScaleIndicator(a, n_ind, True)) & "*0] (N" & int_num + 1 & "-N" & int_num + n_ind & ");" & lb
                    int_num = int_num + n_ind
                End If
            Next
            model_syntax = model_syntax & "[" & DataStructure.ScaleName(1, True) & "-" & DataStructure.ScaleName(DataStructure.ScaleInclude, True) & "];" & lb
        End If
        
        ' Set up factor loadings constraint
        If InputForm_CFA.LV3 = True Then
            ind_num = 0
            For a = 1 To DataStructure.ScaleInclude
                n_ind = DataStructure.ScaleIndicator(a, 0, True)
                If n_ind > 1 Then
                    model_constraint = model_constraint & "0="
                    For b = 1 To n_ind
                        ind_num = ind_num + 1
                        If b > 1 Then model_constraint = model_constraint & "+"
                        model_constraint = model_constraint & "L" & ind_num
                    Next
                    model_constraint = model_constraint & "-" & DataStructure.ScaleIndicator(a, 0, True) & ";" & lb
                End If
            Next
        End If
        
        ' Set up intercepts constraint
        If InputForm_CFA.LV3b = True Then
            int_num = 0
            For a = 1 To DataStructure.ScaleInclude
                n_ind = DataStructure.ScaleIndicator(a, 0, True)
                If n_ind > 1 Then
                    model_constraint = model_constraint & "0="
                    For b = 1 To n_ind
                        int_num = int_num + 1
                        If b > 1 Then model_constraint = model_constraint & "+"
                        model_constraint = model_constraint & "N" & int_num
                    Next
                    model_constraint = model_constraint & ";" & lb
                End If
            Next
        End If
    End If

' Add information related to Bayes
    If InputForm_CFA.BCFA_CL = True Or InputForm_CFA.BCFA_RC = True Then
        model_type = "ESTIMATOR = BAYES;" & lb & "PROCESSORS = 2;" & lb
    End If

' Case #2, CFA with Bayesian small priors for factor cross-loadings
    If InputForm_CFA.BCFA_CL = True Then
        n_ind = 0
        For a = 1 To DataStructure.ScaleInclude
            For c = 1 To DataStructure.ScaleInclude
                If c = 1 Then
                    temp_line = DataStructure.ScaleName(a, True) & " BY"
                Else
                    temp_line = ""
                End If
                n_ind_line = 0
                For b = 1 To DataStructure.ScaleIndicator(c, 0, True)
                    If a <> c Then
                        n_ind_line = n_ind_line + 1
                    End If
                    temp_ind = format_mplus(DataStructure.ScaleIndicator(c, b, True))
                    If Len(temp_line) + Len(temp_ind) < 75 Then
                        If Len(temp_line) > 0 Then temp_line = temp_line & " "
                        temp_line = temp_line & temp_ind
                        If a <> c And b = 1 Then temp_line = temp_line & "*"
                        If a = c And c > 1 And b = 1 Then temp_line = temp_line & "@1"
                    Else
                        If a <> c Then
                            temp_line = temp_line & " (L" & n_ind + 1
                            If n_ind_line > 1 Then
                                temp_line = temp_line & "-L" & n_ind + n_ind_line - 1
                            End If
                            temp_line = temp_line & ")"
                            n_ind = n_ind + n_ind_line - 1
                            n_ind_line = 1
                        End If
        
                        model_syntax = model_syntax & temp_line & lb
                        temp_line = temp_ind
                    End If
                Next
                
                model_syntax = model_syntax & temp_line
                
                If a <> c Then
                    model_syntax = model_syntax & " (L" & n_ind + 1
                    If n_ind_line > 1 Then
                        model_syntax = model_syntax & "-L" & n_ind + n_ind_line
                    End If
                    model_syntax = model_syntax & ")"
                    n_ind = n_ind + n_ind_line
                End If
                If c = DataStructure.ScaleInclude Then model_syntax = model_syntax & ";"
                model_syntax = model_syntax & lb
            Next
            model_syntax = model_syntax & lb
        Next
    model_priors = model_priors & "L1-L" & n_ind & "~N(0,0.01);" & lb
    End If
    
' Case #3, CFA with Bayesian small priors for all residual covariances
    n_ind = 0
    If InputForm_CFA.BCFA_RC = True Then
        For a = 1 To DataStructure.ScaleInclude
            n_ind = n_ind + DataStructure.ScaleIndicator(a, 0, True)
        Next
        ind_first = format_mplus(DataStructure.ScaleIndicator(1, 1, True))
        ind_last = format_mplus(DataStructure.ScaleIndicator(DataStructure.ScaleInclude, DataStructure.ScaleIndicator(DataStructure.ScaleInclude, 0, True), True))
        diag_n = ((n_ind * (n_ind - 1)) / 2)
        d = WorksheetFunction.Round(DataStructure.Cases_N / 5, 0)
        
        model_syntax = model_syntax & ind_first & "-" & ind_last & " (T1-T" & n_ind & ");" & lb
        model_syntax = model_syntax & ind_first & "-" & ind_last & " WITH " & ind_first & "-" & ind_last & " (T" & n_ind + 1 & "-T" & diag_n + n_ind & ");" & lb
        
        i_num = 1
        For a = 1 To DataStructure.ScaleInclude
            For b = 1 To DataStructure.ScaleIndicator(a, 0, True)
                ind_name = format_mplus(DataStructure.ScaleIndicator(a, b, True))
                model_priors = model_priors & "T" & i_num & "~IW(0.5*" & WorksheetFunction.Round(DataStructure.Variance(ind_name), 3) & "*" & d & "," & d & ");" & lb
                i_num = i_num + 1
            Next
        Next
        
        model_priors = model_priors & "T" & n_ind + 1 & "-T" & diag_n + n_ind & "~IW(0," & d & ");" & lb
        model_priors = model_priors & "! Note: For the Inverse Wishart residual variance priors, IW(Dd, d), D is set to half of the observed variance in each indicator, which is then multiplied by d." & lb
        model_priors = model_priors & "! This assumes the latent variables account for half of the variance in each observed indicator." & lb
        model_priors = model_priors & "! Run a CFA without residual covariances to obtain exact residual variances to be used as priors." & lb
        model_priors = model_priors & "! d set to number of cases (" & DataStructure.Cases_N & ") divided by 5." & lb
        model_priors = model_priors & "! Increase d if model does not converge. Decrease d if it converges too quickly with poor model fit." & lb
    End If

' Case #4: Observed variable means using the DEFINE command
    If InputForm_CFA.LV4 = True Then
        For a = 1 To DataStructure.ScaleInclude
            temp_line = DataStructure.ScaleName(a, True) & " = mean("
            For b = 1 To DataStructure.ScaleIndicator(a, 0, True)
                temp_ind = format_mplus(DataStructure.ScaleIndicator(a, b, True))
                If Len(temp_line) + Len(temp_ind) > 88 Then
                    model_define = model_define & temp_line & lb
                    temp_line = temp_ind
                Else
                    If b < DataStructure.ScaleIndicator(a, 0, True) Then temp_ind = temp_ind & " "
                    temp_line = temp_line & temp_ind
                End If
            Next
            model_define = model_define & temp_line & ");" & lb
        Next
    End If
End Sub



