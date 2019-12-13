VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cMplusOutput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Wrapper goes: (1) Multiple groups, (2) WITHIN-BETWEEN, (3) Mixtures

' Scale list
' - Scale short name
' - Scale alias

' Collection: List of observed matrices
' Collection: List of model matrices

' Covariance coverage matrix
' - Covariance coverage

' Array for observed variables + vector for means
' - Covariances (variances in diagonal)s
' - Correlations
' - Model implied covariances
' - Residual covariances
' - Model implied correlations
' - Residual correlations

' Model array
' - Non-directional array (theta and psi matrix)
' - Directional array (beta, lambda, gamma)
' Each element: effect, SE/SD, t, p, CI_005 CI_025 CI_050 CI_950 CI_975 CI_995
' Each with layers for (1) unstandardized, stydx, std, (2) groups / mixtures / levels

' Model array:


Dim out_text ' Contains the raw text from the output file
Dim pNumLines ' Number of lines in the output file
Dim pLinePos() ' The location in "out_text" of the begininng of each line

Dim var_cntr ' Number of variables
Dim var_list() ' List of variables, labels, scales, and sample-level information
Dim scale_list() ' List of scales and their aliases

Dim obs_matrix_list As New Collection ' A list of the matrices of observed variables
Dim obs_matrix_lookup() ' An array containing the lookup coordinates for different models. Dimension 1: (1) Group, (2) Level. Dimension 2: Model number.
Dim obs_matrix() ' An array containing data from the covariances of the observed variables (y,x,std,grp)
Dim obs_vector() ' A vector containing the means of each variable (y, std, grp)
Dim obs_n() ' An array containing the overall sample size of each group / level / mixture
Dim avg_csize() ' An array containing the average cluster size of each group / level
Dim cov_coverage() ' A matrix containing information about covariance coverage

Dim mod_matrix_list As New Collection ' A list of the matrices from the model
Dim mod_matrix_lookup() ' An array containing the lookup coordinates for different models. Dimension 1: (1) Group, (2) Level, (3) Mixture. Dimension 2: Model number
Dim mod_matrix_directional() ' An array containing directional results from the model (BETA / LAMBDA / GAMMA / KAPPA). Dimension 1: IV, 2: DV, 3: Standardization, 4: Group/Level/Mixture
Dim mod_matrix_nondirectional() ' An array containing non-directional results from the mode (PSI / THETA)
Dim mod_vector() ' A vector cotaining the means/intercepts of each variable
Dim mod_rsquare() ' A vector containing information about the R-Square of each variable
'1: Coefficient / Bayes Coefficient / Monte Carlo Average
'2: Standard error / Bayes SD / Monte Carlo SD
'3: t statistic
'4: P value / Bayes P-value
'5: CI: 0.01
'6: CI: 0.05
'7: CI: 0.10
'8: CI: 0.90
'9: CI: 0.95
'10: CI: 0.99
'11: Monte Carlo Population
'12: Monte Carlo SE Average
'13: Monte Carlo MSE
'14: Monte Carlo 95% Cover
'15: Monte Carlo % Sig Coef
'16: Multiple imputation (rate of missing)

Dim grouplist As New Collection ' A list of all of the groups for multiple group analyses
Dim levellist As New Collection ' A list of all of the clusters for multilevel analysis
Dim mixturelist As New Collection ' A list of all of the mixtures within mixture models
Dim standardlist As New Collection ' A list of the types of standardization used

'Dim SampleSize As New Collection ' Total sample size for each group / cluster / mixture

Dim stdyx_num ' The standardisation number of STDYX

Public IsSAMPSTAT As Boolean ' Whether sample statistics are available
Public IsCovCoverage As Boolean ' Whether information about covariance coverage is available
Public IsICC As Boolean ' Whether sample ICCs are available

Public IsSTDYX As Boolean ' Whether STDYX output is available
Public IsSTDY As Boolean ' Whether STY output is available
Public IsSTD As Boolean ' Whether STD output is available
Public IsOddsRatio As Boolean ' Whether odds ratio output is available
Public IsRSquare As Boolean ' Whether R-Square is available
Public IsCInterval As Boolean ' Whether confidence/credibility intervals are available

Public IsBasic As Boolean ' Whether the analysis is TYPE = BASIC
Public IsResidual As Boolean ' Whether residual output is available
Public IsModel As Boolean ' Whether model results are available
Public IsImputation As Boolean ' Whether the analysis is based on imputation
Public IsMonteCarlo As Boolean ' Whether the analysis is a Monte Carlo simulation

Public IsTECH4 As Boolean ' Does Technical 4 output exist?

Public nGroups As Integer ' Number of groups
Public nLevels As Integer ' Number of levels in the analysis (default is 1)
Public nMixtures As Integer ' Number of mixtures
Public nStandards As Integer ' Number of formats for standardised results
Public Estimator As String ' The estimator used by the analysis

Public nDV, nIV, nCoLV, nCaLV, nTotalV As Integer ' Number of DVs, IVs, Continuous Latents, Categorical Latents, Total Variables

Public ChiSq
Public DF
Public ChiSqP
Public CFI, TLI
Public RMSEA
Public SRMR, SRMR_W, SRMR_B
Public AIC
Public BIC
Public BICssa

Public BayesLower
Public BayesUpper
Public PPP
Public PPPP
Public DIC
Public PD

Dim SAMPSTAT_vList As New Collection ' An ordered list of variables contained within "Sample Statistics"

Dim VarList() ' Information about the variables
' Variable list
' 1 Variable short name
' 2 Variable label
' 3 Scale
' 4 Variable type (1 = continuous, 2 = binary / ordered categorical, 3 = nominal, 4 = count, 5 = observed IV, 6 = continuous latent, 7 = categorical latent, 8 = EFA/ESEM factor)
' 5 Observed (False) or latent (True)
' 6 Mean
' 7 Variance
' 8 Sample size
' 9 Skew
' 10 Kurtosis
' 11 Minimum
' 12 20%
' 13 40%
' 14 50% (median)
' 15 60%
' 16 80%
' 17 Maximum
' 18 % Minimum
' 19 % Maximum
' 20 % Zero (for COUNT variables)
' 21 Observed matrix numbered variable
' 22 ICC
' 23 DV order (if a dependent variable)
' 24 Antecedent order (if an antecedent in an ON command)
' 25 Composite (is this an observed variable that is correlated with a latent variable TRUE/FALSE)

Public Property Get pVarList(y, x)
    pVarList = VarList(y, x)
End Property

Public Property Get pMod_Matrix(y, x, Optional Stand_Num = 1, Optional model_num = 1)
    pMod_Matrix = mod_matrix_directional(y, x, Stand_Num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num))
End Property
Public Property Let ParseOutputFile(filename As String)
    Dim text: text = ""
    On Error GoTo errorhandler
    
    Open filename For Input As #1
    
    Do Until EOF(1)
        Line Input #1, textline
        text = text & textline & Chr(10)
    Loop
    Close #1
    
    ParseOutput text
    Exit Property

errorhandler:
    MsgBox "Error opening Mplus output file."
    End
End Property

Public Property Let ParseOutput(text As String)
    VarList = Array()
    
    ' This method parses an Mplus output file and transforms it into usable data
    Debug.Print "Parsing output..."

    ' Populate the private variable, out_text, and get rid of carriage returns (to make it compatible with Macs)
    out_text = Replace(text, Chr(13), "")

    ' Set up the defaults
    IsSAMPSTAT = False
    IsSTDYX = False
    IsSTDY = False
    IsSTD = False
    IsOddsRatio = False
    IsRSquare = False
    IsCInterval = False
    
    IsResidual = False
    IsModel = False
    nLevels = 1
    nGroups = 1
    nMixtures = 1
    stdyx_num = 0
    ReDim mod_matrix_lookup(1 To 3, 1 To 1)
      
    ParseLines ' Call parse the lines sub
    AnalysisInfo ' Generate analysis information
                        
    ' Based on the above information, produce list of matrices (sample statistics and models)
    ' 1) Create list of observed matrices
    Debug.Print "Creating list of observed matrices"
    ReDim obs_matrix_lookup(1 To 2, 1 To nGroups * nLevels)
    If nLevels < 3 Then ' Sample statistics are usually not provided for THREELEVEL (using Bayesian estimation)
        For a = 1 To nGroups
            temp_name = ""
            If nGroups > 1 Then temp_name = grouplist(a)
            For b = 1 To nLevels
                Debug.Print "Group: "; temp_name
                temp_name2 = temp_name
                If nLevels > 1 Then
                    Debug.Print "Level: "; levellist(b)
                    temp_name2 = temp_name2 & " " & levellist(b)
                End If
                obs_matrix_list.Add Trim(temp_name2)
                obs_matrix_lookup(1, obs_matrix_list.Count) = a
                obs_matrix_lookup(2, obs_matrix_list.Count) = b
                Debug.Print a; b; obs_matrix_list(obs_matrix_list.Count)
            Next
        Next
    End If
    
    ' 2) Create list of sample sizes
    Debug.Print "Creating list of sample sizes"
    ReDim obs_n(1 To nGroups, 1 To nLevels, 1 To nMixtures)
    ReDim avg_csize(1 To nGroups, 1 To nLevels)
    If nGroups = 1 And nLevels = 1 Then
        obs_n(1, 1, 1) = Word(FindLine("Number of observations"), 4)
        'SampleSize.Add Word(FindLine("Number of observations"), 4)
        Debug.Print "Number of cases: " & Word(FindLine("Number of observations"), 4)
    ElseIf nGroups = 1 And nLevels > 1 Then
        obs_n(1, 1, 1) = Word(FindLine("Number of observations"), 4)
        obs_n(1, 2, 1) = Word(FindLine("Number of clusters"), 4)
        avg_csize(1, 2) = Word(FindLine("Average cluster size"), 4)
        Debug.Print "Number of cases: " & obs_n(1, 1, 1)
        Debug.Print "Number of clusters: " & obs_n(1, 2, 1)
        Debug.Print "Average cluster size: " & avg_csize(1, 2)
    ElseIf nGroups > 1 Then
        temp_line = FindLine("Number of groups")
        For a = 1 To nGroups
            obs_n(a, 1, 1) = Word(temp_line + a + 1, 3)
            'SampleSize.Add Word(temp_line + a + 1, 3)
            Debug.Print "-Group " & a & ": " & grouplist.Item(grouplist.Count) & " (n = " & obs_n(a, 1, 1) & ")"
            If nLevels > 1 Then
                temp_line2 = FindLine("SUMMARY OF DATA")
                temp_line2 = FindLine("Group " & grouplist(a), temp_line2)
                obs_n(a, 2, 1) = Word(FindLine("Number of clusters", temp_line2), 4)
                avg_csize(a, 2) = Word(FindLine("Average cluster size", temp_line2), 4)
            End If
        Next
    End If
    
    ' 2.5) Find ICCs
    If nLevels > 1 And FindLine("Estimated Intraclass Correlations") > 0 Then
        Debug.Print "Finding sample ICCs"
        IsICC = True
        start_line = FindLine("SUMMARY OF DATA")
        For a = 1 To nGroups
            If nGroups > 1 Then start_line = FindLine("Group " & grouplist(a), start_line)
            start_line2 = FindLine("Estimated Intraclass Correlations", start_line) + 5
            Do While Num_Words_on_Line(start_line2) > 0
                For b = 1 To Num_Words_on_Line(start_line2) / 2
                    v_num = VarNum(Word(start_line2, (b * 2) - 1))
                    VarList(v_num, 22) = ADD_STRINGVECTOR(VarList(v_num, 22), Word(start_line2, b * 2))
                Next
                start_line2 = start_line2 + 1
            Loop
        Next
    End If
    
    ' 3) Create a list of models
    ' Add a single group if there's only unstandardised output
    ' If grouplist.Count = 0 And levellist.Count = 0 And mixturelist.Count = 0 And standardlist.Count = 1 Then mod_matrix_list.Add SET_STRINGVECTOR("", "", "", "")
    Debug.Print nGroups, nLevels, nMixtures, nStandards
    ReDim mod_matrix_lookup(1 To 3, 1 To nGroups * nLevels * nMixtures)
    For a = 1 To nGroups
        For b = 1 To nLevels
            For c = 1 To nMixtures
                grp_name = ""
                lev_name = ""
                mix_name = ""
                If grouplist.Count >= a Then grp_name = grouplist(a)
                If levellist.Count >= b Then lev_name = levellist(b)
                If mixturelist.Count >= c Then mix_name = mixturelist(c)
                mod_matrix_list.Add SET_STRINGVECTOR(grp_name, lev_name, mix_name) ', std_name)
                mod_matrix_lookup(1, mod_matrix_list.Count) = a ' Set it to the group
                mod_matrix_lookup(2, mod_matrix_list.Count) = b ' Set it to the level
                mod_matrix_lookup(3, mod_matrix_list.Count) = c ' Set it to the mixture
                Debug.Print "Adding model matrix: " & mod_matrix_list(mod_matrix_list.Count)
            Next
        Next
    Next
    
    ' Add sample statistics
    ' Are there any sample statistics?
    start_line = FindLine("SAMPLE STATISTICS")
    
    ' Check to see if it's an exact match
    If line(start_line) <> "SAMPLE STATISTICS" Then start_line = 0
    
    If start_line = 0 Then start_line = FindLine("RESULTS FOR BASIC ANALYSIS")
    
    If start_line > 0 Then
        Debug.Print "Found sample statistics (SAMPSTAT)"
        IsSAMPSTAT = True
        
        ' Count the number of variables
        start_line2 = FindLine("Covariances", start_line) + 3
        If start_line2 = 3 Then start_line2 = FindLine("CORRELATION MATRIX", start_line) + 3
        
        var_cntr = 0
        For a = start_line2 To Next_Blank_Line(start_line2) - 1
            Debug.Print "Line ", a, "to line", Next_Blank_Line(start_line2) - 1
            temp_var = VarNum(Word(a, 1))
            Debug.Print temp_var & ": " & Word(a, 1)
            'If temp_var > 0 Then VarList(temp_var, 21) = var_cntr + 1 ' Numbering the variable's position in the matrix. Deleted this.
            var_cntr = var_cntr + 1
        Next
                
        ' Resize the observed matrix array based on the number of variables and the number of groups
        ReDim obs_matrix(1 To var_cntr, 1 To var_cntr, 1 To 2, 1 To nGroups, 1 To nLevels)
        ReDim obs_vector(1 To var_cntr, 1 To nGroups, 1 To nLevels)
        'MsgBox "Dimensions are " & var_cntr & " * " & var_cntr & " * " & nGroups & " groups * " & nLevels & " levels."
        
        ' Loop through each of the observed matrices and add them to the arrays
        For a = 1 To obs_matrix_list.Count
            temp_line = "SAMPLE STATISTICS"
            If obs_matrix_list.Count > 1 Then temp_line = temp_line & " FOR " & obs_matrix_list(a)
            start_line = FindLine(temp_line)
            
            ' Add Means
            start_line2 = FindLine("Means", start_line, False)
            
            Do While Word(start_line2, 1) = "Means" Or Word(start_line2, 1) = "MEANS/INTERCEPTS/THRESHOLDS"
                For b = 1 To Num_Words_on_Line(start_line2 + 1)
                    Debug.Print "Adding means " & b & " of " & Num_Words_on_Line(start_line2 + 1)
                    var_name = Word(start_line2 + 1, b)
                    dol_sign = InStr(1, var_name, "$")
                    category = 1
                    If dol_sign > 0 Then
                        category = CInt(Mid(var_name, dol_sign + 1, Len(var_name) - dol_sign))
                        var_name = Left(var_name, dol_sign - 1)
                    End If
                    
                    ' Check to see whether the variable is categorical
                    If dol_sign > 0 Then
                        obs_vector(VarNum(var_name), obs_Group(a), obs_Level(a)) = ADD_STRINGVECTOR(obs_vector(VarNum(Word(start_line2 + 1, b)), obs_Group(a), obs_Level(a)), Word(start_line2 + 3, b))
                    Else
                        obs_vector(VarNum(var_name), obs_Group(a), obs_Level(a)) = Word(start_line2 + 3, b)
                    End If
                Next
                start_line2 = Next_NonBlank_Line(start_line2 + 4)
            Loop
            
            ' Add Covariances/Variances
            Debug.Print "Adding Covariances for: " & obs_matrix_list(a)
            start_line2 = FindLine("Covariances", start_line)
            
            Do While Word(start_line2, 1) = "Covariances"
                start_line3 = start_line2 + 2
                For y = 1 To Num_NonBlankLines(start_line3 + 1)
                    Debug.Print y; ": searching for Y variable"
                    y_num = VarNum(Word(start_line3 + y, 1))
                    If y_num = 0 Then MsgBox "Cannot find Y variable: " & Word(start_line3 + y, 1)
                    temp_text = ""
                    Debug.Print "Y number is: " & y_num
                    If Num_Words_on_Line(start_line3 + y) > 1 Then
                        For x = 2 To Num_Words_on_Line(start_line3 + y)
                            x_num = VarNum(Word(start_line3 - 1, x - 1))
                            Debug.Print "X number is: " & x_num
                            If x_num = 0 Then MsgBox "Cannot find X variable: " & Word(start_line3 - 1, x)
                            obs_matrix(y_num, x_num, 1, obs_Group(a), obs_Level(a)) = Word(start_line3 + y, x)
                            obs_matrix(x_num, y_num, 1, obs_Group(a), obs_Level(a)) = obs_matrix(y_num, x_num, 1, obs_Group(a), obs_Level(a))
                            temp_text = temp_text & " " & Word(start_line3 + y, x)
                        Next
                    End If
                    end_line = start_line3 + y
                    Debug.Print "Found line of covariances: " & temp_text
                Next
                start_line2 = Next_NonBlank_Line(end_line)
            Loop
            
            ' Add Correlations
            Debug.Print "Adding Correlations for: " & obs_matrix_list(a)
            start_line2 = FindLine("Correlations", start_line)
            If start_line2 = 0 Then start_line2 = FindLine("CORRELATION MATRIX", start_line)
                        
            Do While Word(start_line2, 1) = "Correlations" Or (Word(start_line2, 1) = "CORRELATION" And Word(start_line2, 2) = "MATRIX")
                start_line3 = start_line2 + 2
                For y = 1 To Num_NonBlankLines(start_line3 + 1)
                    Debug.Print y; ": searching for Y variable"
                    y_num = VarNum(Word(start_line3 + y, 1))
                    If y_num = 0 Then MsgBox "Cannot find Y variable: " & Word(start_line3 + y, 1)
                    temp_text = ""
                    Debug.Print "Y number is: " & y_num
                    If Num_Words_on_Line(start_line3 + y) > 1 Then
                        For x = 2 To Num_Words_on_Line(start_line3 + y)
                            x_num = VarNum(Word(start_line3 - 1, x - 1))
                            'MsgBox "Y:" & Word(start_line3 + y, 1) & "(" & y_num & "), X:" & Word(start_line3 - 1, x - 1) & " (" & x_num & ")"
                            Debug.Print "X number is: " & x_num
                            If x_num = 0 Then MsgBox "Cannot find X variable: " & Word(start_line3 - 1, x)
                            obs_matrix(y_num, x_num, 2, obs_Group(a), obs_Level(a)) = Word(start_line3 + y, x)
                            obs_matrix(x_num, y_num, 2, obs_Group(a), obs_Level(a)) = obs_matrix(y_num, x_num, 2, obs_Group(a), obs_Level(a))
                            temp_text = temp_text & " " & Word(start_line3 + y, x)
                        Next
                    End If
                    end_line = start_line3 + y
                    Debug.Print "Found line of correlations: " & temp_text
                Next
                start_line2 = Next_NonBlank_Line(end_line)
            Loop
            
            ' Produce covariance table if categorical data has been requested
            If FindLine("CORRELATION MATRIX (WITH VARIANCES ON THE DIAGONAL)", start_line) > 0 Then
                ' Do the diagonal first
                For y = LBound(obs_matrix, 1) To UBound(obs_matrix, 1)
                    Debug.Print "Diagonal..."
                    If VarList(y, 4) = 2 Or VarList(y, 4) = 3 Then
                        Debug.Print "Categorical..."
                        ' Set the variance to 1 for categorical variables
                        obs_matrix(y, y, 1, obs_Group(a), obs_Level(a)) = 1
                        obs_matrix(y, y, 2, obs_Group(a), obs_Level(a)) = 1
                    Else
                        Debug.Print "Non-categorical"
                        ' Move variance to covariance table
                        obs_matrix(y, y, 1, obs_Group(a), obs_Level(a)) = obs_matrix(y, y, 2, obs_Group(a), obs_Level(a))
                        obs_matrix(y, y, 2, obs_Group(a), obs_Level(a)) = 1
                    End If
                Next
                
                ' Calculate the covariances
                Debug.Print "Covariances"
                For y = 2 To UBound(obs_matrix, 1)
                    For x = 1 To y - 1
                        Debug.Print y & " " & x & " 1 " & obs_Group(a) & " " & obs_Level(a)
                        obs_matrix(y, x, 1, obs_Group(a), obs_Level(a)) = obs_matrix(y, x, 2, obs_Group(a), obs_Level(a)) * Sqr(obs_matrix(y, y, 1, obs_Group(a), obs_Level(a))) * Sqr(obs_matrix(x, x, 1, obs_Group(a), obs_Level(a)))
                        obs_matrix(x, y, 1, obs_Group(a), obs_Level(a)) = obs_matrix(y, x, 1, obs_Group(a), obs_Level(a))
                    Next
                Next
            End If
            
        Next
    End If
    
    ' Add covariance coverage
    start_line = FindLine("COVARIANCE COVERAGE OF DATA")
    If start_line > 0 Then
        IsCovCoverage = True
        'MsgBox "Variables: " & var_cntr & ", Groups: " & nGroups
        ReDim cov_coverage(1 To ObservedVariable, 1 To ObservedVariable, 1 To nGroups)
        For a = 1 To nGroups
            Debug.Print "Covariance coverage "; a; " of "; nGroups
            temp_line = "PROPORTION OF DATA PRESENT"
            If nGroups > 1 Then temp_line = temp_line & " FOR " & grouplist(a)
            start_line = FindLine(temp_line)
            start_line2 = FindLine("Covariance Coverage", start_line)
            
            Do While CONTAINS(line(start_line2), "Covariance Coverage")
                start_line3 = start_line2 + 2
                For y = 1 To Num_NonBlankLines(start_line3 + 1)
                    y_var = VarNum(Word(start_line3 + y, 1))
                    temp_text = ""
                    For x = 2 To Num_Words_on_Line(start_line3 + y)
                        x_var = VarNum(Word(start_line3 - 1, x - 1))
                        Debug.Print y_var, x_var, a, Word(start_line3 + y, x)
                        Debug.Print LBound(cov_coverage, 1), UBound(cov_coverage, 1)
                        Debug.Print LBound(cov_coverage, 2), UBound(cov_coverage, 2)
                        Debug.Print LBound(cov_coverage, 3), UBound(cov_coverage, 3)
                        'MsgBox y_var & " " & x_var & " " & a & " " & UBound(cov_coverage, 1) & " " & UBound(cov_coverage, 2) & " " & UBound(cov_coverage, 3)
                        cov_coverage(y_var, x_var, a) = Word(start_line3 + y, x)
                        cov_coverage(x_var, y_var, a) = cov_coverage(y_var, x_var, a)
                        Debug.Print "Complete 2"
                        temp_text = temp_text & " " & Word(start_line3 + y, x)
                    Next
                    end_line = start_line3 + y
                    Debug.Print "Found line of covariance coverages: " & temp_text
                Next
                start_line2 = Next_NonBlank_Line(end_line)
            Loop
        Next
    End If
    
    AddModelResults
    
    ' Find model fit statistics
    If IsModel = True Then

        temp_line = FindLine("Information Criteria")
        If temp_line > 0 Then
            AIC = Word(FindLine("Akaike (AIC)", temp_line), 3)
            BIC = Word(FindLine("Bayesian (BIC)", temp_line), 3)
            BICssa = Word(FindLine("Sample-Size Adjusted BIC", temp_line), 4)
        End If
        
        If Estimator = "BAYES" Then
            temp_line = FindLine("Bayesian Posterior Predictive Checking using Chi-Square")
            If temp_line > 0 Then
                BayesLower = Word(FindLine("the Observed and the Replicated Chi-Square Values", temp_line, 2) + 2, 1)
                BayesUpper = Word(FindLine("the Observed and the Replicated Chi-Square Values", temp_line, 2) + 2, 2)
                PPP = Word(FindLine("Posterior Predictive P-Value", temp_line), 4)
                If FindLine("Prior Posterior Predictive P-Value") > 0 Then PPPP = Word(FindLine("Prior Posterior Predictive P-Value", temp_line), 5)
                DIC = Word(FindLine("Deviance (DIC)", temp_line), 3)
                PD = Word(FindLine("Estimated Number of Parameters (pD)", temp_line), 6)
                BIC = Word(FindLine("Bayesian (BIC)", temp_line), 3)
            End If
        Else
            temp_line = FindLine("Chi-Square Test of Model Fit")
            If temp_line > 0 Then
                ChiSq = Word(FindLine("Value", temp_line), 2)
                DF = Word(FindLine("Degrees of Freedom", temp_line), 4)
                ChiSqP = Word(FindLine("P-Value", temp_line), 2)
                If Right(ChiSq, 1) = "*" Then ChiSq = Left(ChiSq, Len(ChiSq) - 1)
            End If
            
            temp_line = FindLine("RMSEA (Root Mean Square Error Of Approximation)")
            If temp_line > 0 Then RMSEA = Word(FindLine("Estimate", temp_line), 2)
            
            temp_line = FindLine("CFI/TLI")
            If temp_line > 0 Then CFI = Word(FindLine("CFI", temp_line + 1), 2)
            If temp_line > 0 Then TLI = Word(FindLine("TLI", temp_line + 1), 2)
            
            temp_line = FindLine("SRMR (Standardized Root Mean Square Residual)")
            If temp_line > 0 Then
                If nLevels > 1 Then
                    SRMR_W = Word(FindLine("Value for Within", temp_line), 4)
                    SRMR_B = Word(FindLine("Value for Between", temp_line), 4)
                Else
                    SRMR = Word(FindLine("Value", temp_line), 2)
                End If
            End If
        End If
    End If
    
    ' Tag variables as composites or not
    FindComposites
    
    ' Find TECH4 output
    start_line = FindLine("TECHNICAL 4 OUTPUT")
    If start_line > 0 Then
        IsTECH4 = True
        
        start_line = FindLine("ESTIMATES DERIVED FROM THE MODEL", start_line)
        
        Do While start_line > 0
            ' Set defaults for group, level and mixture
            grp_num = 1
            lev_num = 1
            mix_num = 1
            
            ' Find the group, level and mixture
            temp_line = Trim(line(start_line))
            temp_line = STRING_RIGHT(temp_line, "ESTIMATES DERIVED FROM THE MODEL")
            If Len(temp_line) > 0 Then
                ' Find groups
                If nGroups > 1 Then
                    For a = 1 To nGroups
                        If InStr(1, UCase(temp_line), UCase(grouplist(a))) > 0 Then
                            grp_num = a
                            Debug.Print "Found group "; a, grouplist(a)
                        End If
                    Next
                End If
                
                ' Find levels
                If nLevels > 1 Then
                    For a = 1 To nLevels
                        If InStr(1, UCase(temp_line), UCase(levellist(a))) > 0 Then
                            lev_num = a
                            Debug.Print "Found level "; a, levellist(a)
                        End If
                    Next
                End If
                
                ' Find mixtures
                If nMixtures > 1 Then
                    temp_line2 = STRING_RIGHT(temp_line, "PATTERN")
                    If temp_line2 = "" Then temp_line2 = STRING_RIGHT(temp_line, "CLASS")
                    For a = 1 To nLevels
                        If InStr(1, UCase(temp_line2), UCase(mixturelist(a))) > 0 Then
                            mix_num = a
                            Debug.Print "Found mixture "; a, mixturelist(a)
                        End If
                    Next
                End If
            End If
            
            ' Find estimated means
            start_line2 = FindLine("ESTIMATED MEANS", start_line)
            Do While CONTAINS(line(start_line2), "ESTIMATED MEANS")
                For b = 1 To Num_Words_on_Line(start_line2 + 1)
                    Debug.Print "Adding estimated means " & b & " of " & Num_Words_on_Line(start_line2 + 1)
                    var_num = VarNum(Word(start_line2 + 1, b))
                    mod_vector(var_num, 1, grp_num, lev_num, mix_num) = ADD_STRINGVECTOR(mod_vector(var_num, 1, grp_num, lev_num, mix_num), Word(start_line2 + 3, b), "TECH4_Estimate")
                Next
                start_line2 = Next_NonBlank_Line(start_line2 + 4)
            Loop
            
            ' Find SEs for estimated means
            start_line2 = FindLine("S.E. FOR ESTIMATED MEANS", start_line)
            If start_line2 > 0 Then
                Do While CONTAINS(line(start_line2), "S.E. FOR ESTIMATED MEANS")
                    For b = 1 To Num_Words_on_Line(start_line2 + 1)
                        Debug.Print "Adding estimated SEs " & b & " of " & Num_Words_on_Line(start_line2 + 1)
                        var_num = VarNum(Word(start_line2 + 1, b))
                        mod_vector(var_num, 1, grp_num, lev_num, mix_num) = ADD_STRINGVECTOR(mod_vector(var_num, 1, grp_num, lev_num, mix_num), Word(start_line2 + 3, b), "TECH4_SE")
                    Next
                    start_line2 = Next_NonBlank_Line(start_line2 + 4)
                Loop
            End If

            ' Find ts for estimated means
            start_line2 = FindLine("EST./S.E. FOR ESTIMATED MEANS", start_line)
            If start_line2 > 0 Then
                Do While CONTAINS(line(start_line2), "EST./S.E. FOR ESTIMATED MEANS")
                    For b = 1 To Num_Words_on_Line(start_line2 + 1)
                        Debug.Print "Adding estimated ts " & b & " of " & Num_Words_on_Line(start_line2 + 1)
                        var_num = VarNum(Word(start_line2 + 1, b))
                        mod_vector(var_num, 1, grp_num, lev_num, mix_num) = ADD_STRINGVECTOR(mod_vector(var_num, 1, grp_num, lev_num, mix_num), Word(start_line2 + 3, b), "TECH4_t")
                    Next
                    start_line2 = Next_NonBlank_Line(start_line2 + 4)
                Loop
            End If
            
            ' Find ps for estimated means
            start_line2 = FindLine("P-VALUE FOR ESTIMATED MEANS", start_line)
            If start_line2 > 0 Then
                Do While CONTAINS(line(start_line2), "P-VALUE FOR ESTIMATED MEANS")
                    For b = 1 To Num_Words_on_Line(start_line2 + 1)
                        Debug.Print "Adding p-values for estimated means " & b & " of " & Num_Words_on_Line(start_line2 + 1)
                        var_num = VarNum(Word(start_line2 + 1, b))
                        mod_vector(var_num, 1, grp_num, lev_num, mix_num) = ADD_STRINGVECTOR(mod_vector(var_num, 1, grp_num, lev_num, mix_num), Word(start_line2 + 3, b), "TECH4_p")
                    Next
                    start_line2 = Next_NonBlank_Line(start_line2 + 4)
                Loop
            End If
            
            ' Add Covariances/Variances
            Debug.Print "Adding TECH4 Covariances"
            start_line2 = FindLine("  ESTIMATED COVARIANCE MATRIX", start_line)
            Do While CONTAINS(line(start_line2), "  ESTIMATED COVARIANCE MATRIX")
                start_line3 = start_line2 + 2
                For y = 1 To Num_NonBlankLines(start_line3 + 1)
                    temp_text = ""
                    For x = 2 To Num_Words_on_Line(start_line3 + y)
                        y_num = VarNum(Word(start_line3 + y, 1))
                        x_num = VarNum(Word(start_line3 - 1, x - 1))
                        mod_matrix_nondirectional(y_num, x_num, 1, grp_num, lev_num, mix_num) = ADD_STRINGVECTOR(mod_matrix_nondirectional(y_num, x_num, 1, grp_num, lev_num, mix_num), Word(start_line3 + y, x), "TECH4_Estimate")
                        mod_matrix_nondirectional(x_num, y_num, 1, grp_num, lev_num, mix_num) = mod_matrix_nondirectional(y_num, x_num, 1, grp_num, lev_num, mix_num)
                        temp_text = temp_text & " " & Word(start_line3 + y, x)
                    Next
                    end_line = start_line3 + y
                    Debug.Print "Found line of covariances: " & temp_text
                Next
                start_line2 = Next_NonBlank_Line(end_line)
            Loop
            
            ' Adding SEs for covariances
            Debug.Print "Adding TECH4 Covariances SEs"
            start_line2 = FindLine("S.E. FOR ESTIMATED COVARIANCE MATRIX", start_line)
            If start_line2 > 0 Then
                Do While CONTAINS(line(start_line2), "S.E. FOR ESTIMATED COVARIANCE MATRIX")
                    start_line3 = start_line2 + 2
                    For y = 1 To Num_NonBlankLines(start_line3 + 1)
                        temp_text = ""
                        For x = 2 To Num_Words_on_Line(start_line3 + y)
                            y_num = VarNum(Word(start_line3 + y, 1))
                            x_num = VarNum(Word(start_line3 - 1, x - 1))
                            mod_matrix_nondirectional(y_num, x_num, 1, grp_num, lev_num, mix_num) = ADD_STRINGVECTOR(mod_matrix_nondirectional(y_num, x_num, 1, grp_num, lev_num, mix_num), Word(start_line3 + y, x), "TECH4_SE")
                            mod_matrix_nondirectional(x_num, y_num, 1, grp_num, lev_num, mix_num) = mod_matrix_nondirectional(y_num, x_num, 1, grp_num, lev_num, mix_num)
                            temp_text = temp_text & " " & Word(start_line3 + y, x)
                        Next
                        end_line = start_line3 + y
                        Debug.Print "Found line of SEs: " & temp_text
                    Next
                    start_line2 = Next_NonBlank_Line(end_line)
                Loop
            End If
            
            ' Adding t statistics for covariances
            Debug.Print "Adding TECH4 Covariances ts"
            start_line2 = FindLine("EST./S.E. FOR ESTIMATED COVARIANCE MATRIX", start_line)
            If start_line2 > 0 Then
                Do While CONTAINS(line(start_line2), "EST./S.E. FOR ESTIMATED COVARIANCE MATRIX")
                    start_line3 = start_line2 + 2
                    For y = 1 To Num_NonBlankLines(start_line3 + 1)
                        temp_text = ""
                        For x = 2 To Num_Words_on_Line(start_line3 + y)
                            y_num = VarNum(Word(start_line3 + y, 1))
                            x_num = VarNum(Word(start_line3 - 1, x - 1))
                            mod_matrix_nondirectional(y_num, x_num, 1, grp_num, lev_num, mix_num) = ADD_STRINGVECTOR(mod_matrix_nondirectional(y_num, x_num, 1, grp_num, lev_num, mix_num), Word(start_line3 + y, x), "TECH4_t")
                            mod_matrix_nondirectional(x_num, y_num, 1, grp_num, lev_num, mix_num) = mod_matrix_nondirectional(y_num, x_num, 1, grp_num, lev_num, mix_num)
                            temp_text = temp_text & " " & Word(start_line3 + y, x)
                        Next
                        end_line = start_line3 + y
                        Debug.Print "Found line of ts: " & temp_text
                    Next
                    start_line2 = Next_NonBlank_Line(end_line)
                Loop
            End If
        
            ' Adding p statistics for covariances
            Debug.Print "Adding TECH4 Covariances ps"
            start_line2 = FindLine("P-VALUE FOR ESTIMATED COVARIANCE MATRIX", start_line)
            If start_line2 > 0 Then
                Do While CONTAINS(line(start_line2), "P-VALUE FOR ESTIMATED COVARIANCE MATRIX")
                    start_line3 = start_line2 + 2
                    For y = 1 To Num_NonBlankLines(start_line3 + 1)
                        temp_text = ""
                        For x = 2 To Num_Words_on_Line(start_line3 + y)
                            y_num = VarNum(Word(start_line3 + y, 1))
                            x_num = VarNum(Word(start_line3 - 1, x - 1))
                            mod_matrix_nondirectional(y_num, x_num, 1, grp_num, lev_num, mix_num) = ADD_STRINGVECTOR(mod_matrix_nondirectional(y_num, x_num, 1, grp_num, lev_num, mix_num), Word(start_line3 + y, x), "TECH4_p")
                            mod_matrix_nondirectional(x_num, y_num, 1, grp_num, lev_num, mix_num) = mod_matrix_nondirectional(y_num, x_num, 1, grp_num, lev_num, mix_num)
                            temp_text = temp_text & " " & Word(start_line3 + y, x)
                        Next
                        end_line = start_line3 + y
                        Debug.Print "Found line of ps: " & temp_text
                    Next
                    start_line2 = Next_NonBlank_Line(end_line)
                Loop
            End If
                        
            ' Add additional TECH4 output if STDYX has been requested
            If stdyx_num > 0 Then
                ' Add Correlations
                Debug.Print "Adding TECH4 Correlations"
                start_line2 = FindLine("  ESTIMATED CORRELATION MATRIX", start_line)
                Do While CONTAINS(line(start_line2), "  ESTIMATED CORRELATION MATRIX")
                    start_line3 = start_line2 + 2
                    For y = 1 To Num_NonBlankLines(start_line3 + 1)
                        temp_text = ""
                        For x = 2 To Num_Words_on_Line(start_line3 + y)
                            y_num = VarNum(Word(start_line3 + y, 1))
                            x_num = VarNum(Word(start_line3 - 1, x - 1))
                            mod_matrix_nondirectional(y_num, x_num, stdyx_num, grp_num, lev_num, mix_num) = ADD_STRINGVECTOR(mod_matrix_nondirectional(y_num, x_num, stdyx_num, grp_num, lev_num, mix_num), Word(start_line3 + y, x), "TECH4_Estimate")
                            mod_matrix_nondirectional(x_num, y_num, stdyx_num, grp_num, lev_num, mix_num) = mod_matrix_nondirectional(y_num, x_num, stdyx_num, grp_num, lev_num, mix_num)
                            temp_text = temp_text & " " & Word(start_line3 + y, x)
                        Next
                        end_line = start_line3 + y
                        Debug.Print "Found line of covariances: " & temp_text
                    Next
                    start_line2 = Next_NonBlank_Line(end_line)
                Loop
                
                ' Adding SEs for correlations
                Debug.Print "Adding TECH4 Correlation SEs"
                start_line2 = FindLine("S.E. FOR ESTIMATED CORRELATION MATRIX", start_line)
                If start_line2 > 0 Then
                    Do While CONTAINS(line(start_line2), "S.E. FOR ESTIMATED CORRELATION MATRIX")
                        start_line3 = start_line2 + 2
                        For y = 1 To Num_NonBlankLines(start_line3 + 1)
                            temp_text = ""
                            For x = 2 To Num_Words_on_Line(start_line3 + y)
                                y_num = VarNum(Word(start_line3 + y, 1))
                                x_num = VarNum(Word(start_line3 - 1, x - 1))
                                mod_matrix_nondirectional(y_num, x_num, stdyx_num, grp_num, lev_num, mix_num) = ADD_STRINGVECTOR(mod_matrix_nondirectional(y_num, x_num, stdyx_num, grp_num, lev_num, mix_num), Word(start_line3 + y, x), "TECH4_SE")
                                mod_matrix_nondirectional(x_num, y_num, stdyx_num, grp_num, lev_num, mix_num) = mod_matrix_nondirectional(y_num, x_num, stdyx_num, grp_num, lev_num, mix_num)
                                temp_text = temp_text & " " & Word(start_line3 + y, x)
                            Next
                            end_line = start_line3 + y
                            Debug.Print "Found line of SEs: " & temp_text
                        Next
                        start_line2 = Next_NonBlank_Line(end_line)
                    Loop
                End If
                
                ' Adding t statistics for correlations
                Debug.Print "Adding TECH4 Correlations ts"
                start_line2 = FindLine("EST./S.E. FOR ESTIMATED CORRELATION MATRIX", start_line)
                If start_line2 > 0 Then
                    Do While CONTAINS(line(start_line2), "EST./S.E. FOR ESTIMATED CORRELATION MATRIX")
                        start_line3 = start_line2 + 2
                        For y = 1 To Num_NonBlankLines(start_line3 + 1)
                            temp_text = ""
                            For x = 2 To Num_Words_on_Line(start_line3 + y)
                                y_num = VarNum(Word(start_line3 + y, 1))
                                x_num = VarNum(Word(start_line3 - 1, x - 1))
                                mod_matrix_nondirectional(y_num, x_num, stdyx_num, grp_num, lev_num, mix_num) = ADD_STRINGVECTOR(mod_matrix_nondirectional(y_num, x_num, stdyx_num, grp_num, lev_num, mix_num), Word(start_line3 + y, x), "TECH4_t")
                                mod_matrix_nondirectional(x_num, y_num, stdyx_num, grp_num, lev_num, mix_num) = mod_matrix_nondirectional(y_num, x_num, stdyx_num, grp_num, lev_num, mix_num)
                                temp_text = temp_text & " " & Word(start_line3 + y, x)
                            Next
                            end_line = start_line3 + y
                            Debug.Print "Found line of ts: " & temp_text
                        Next
                        start_line2 = Next_NonBlank_Line(end_line)
                    Loop
                End If
            
                ' Adding p statistics for correlations
                Debug.Print "Adding TECH4 Correlation ps"
                start_line2 = FindLine("P-VALUE FOR ESTIMATED CORRELATION MATRIX", start_line)
                If start_line2 > 0 Then
                    Do While CONTAINS(line(start_line2), "P-VALUE FOR ESTIMATED CORRELATION MATRIX")
                        start_line3 = start_line2 + 2
                        For y = 1 To Num_NonBlankLines(start_line3 + 1)
                            temp_text = ""
                            For x = 2 To Num_Words_on_Line(start_line3 + y)
                                y_num = VarNum(Word(start_line3 + y, 1))
                                x_num = VarNum(Word(start_line3 - 1, x - 1))
                                mod_matrix_nondirectional(y_num, x_num, stdyx_num, grp_num, lev_num, mix_num) = ADD_STRINGVECTOR(mod_matrix_nondirectional(y_num, x_num, stdyx_num, grp_num, lev_num, mix_num), Word(start_line3 + y, x), "TECH4_p")
                                mod_matrix_nondirectional(x_num, y_num, stdyx_num, grp_num, lev_num, mix_num) = mod_matrix_nondirectional(y_num, x_num, stdyx_num, grp_num, lev_num, mix_num)
                                temp_text = temp_text & " " & Word(start_line3 + y, x)
                            Next
                            end_line = start_line3 + y
                            Debug.Print "Found line of ps: " & temp_text
                        Next
                        start_line2 = Next_NonBlank_Line(end_line)
                    Loop
                End If
            End If
        
            start_line = FindLine("ESTIMATES DERIVED FROM THE MODEL", start_line + 1)
        Loop
        
    End If
    
    MsgBox "Mplus output parsed successfully"
End Property

Sub ParseLines()
    ' Work out where each line starts
    Debug.Print "Determining start of each line..."
    pNumLines = NumLines()
    ReDim pLinePos(1 To pNumLines)
    pLinePos(1) = 1
    a = 1
    lnum = 2
    Do
        b = InStr(a, out_text, Chr(10))
        
        ' Connect the lines if a comment runs over multiple lines
        If b <> 0 Then
            'MsgBox Mid(out_text, a, b - a) & Chr(10) & b - a & Chr(10) & Mid(out_text, a, 3) & Chr(10) & Mid(out_text, b - 1, 1)
            If Mid(out_text, a, 3) = "  !" And Mid(out_text, b - 1, 1) = "+" Then
                pNumLines = pNumLines - 1
            Else
                pLinePos(lnum) = b + 1
                lnum = lnum + 1
                'Debug.Print Mid(out_text, a, b - a)
            End If
        End If
    a = b + 1
    Loop Until b = 0
End Sub
Sub AnalysisInfo()
    ' Work out what type of analysis it is
    temp_line = 1
    Do
        temp_line = FindLine("TYPE", temp_line + 1, False)
        If UCase(Word(temp_line, 3)) = "IMPUTATION;" Then IsImputation = True
        If UCase(Word(temp_line, 3)) = "MONTECARLO;" Then IsMonteCarlo = True
        If UCase(Word(temp_line, 3)) = "BASIC;" Then IsBasic = True
        Debug.Print temp_line
    Loop While temp_line > 0
    If FindLine("MONTECARLO:", 1, False) > 0 Then IsMonteCarlo = True
    If FindLine("TYPE=MONTECARLO", 1, False) > 0 Then IsMonteCarlo = True
    If FindLine("TYPE=IMPUTATION", 1, False) > 0 Then IsImputation = True
    If FindLine("TYPE=BASIC", 1, False) > 0 Then IsBasic = True
    
    ' Count the number of variables
    nDV = CInt(Word(FindLine("Number of dependent variables"), 5))
    nIV = CInt(Word(FindLine("Number of independent variables"), 5))
    nCoLV = CInt(Word(FindLine("Number of continuous latent variables"), 6))
    nCaLV = CInt(Word(FindLine("Number of categorical latent variables"), 6))
    nTotalV = nDV + nIV + nCoLV + nCaLV
    ReDim VarList(1 To nTotalV, 1 To 25)
                                
    ' Find the variables
    Debug.Print "Finding variables..."
    var_cntr = 1
    Dim start_line: start_line = FindLine("SUMMARY OF ANALYSIS")
    Dim start_line2
    Dim alias_line: alias_line = FindLine("!LABELS:") + 1
    
    For v_type = 1 To 8
        If v_type = 1 Then v_tt = "  Continuous"
        If v_type = 2 Then v_tt = "  Binary and ordered categorical (ordinal)"
        If v_type = 3 Then v_tt = "  Unordered categorical (nominal)"
        If v_type = 4 Then v_tt = "  Count"
        If v_type = 5 Then v_tt = "Observed independent variables"
        If v_type = 6 Then v_tt = "Continuous latent variables"
        If v_type = 7 Then v_tt = "Categorical latent variables"
        If v_type = 8 Then v_tt = "  EFA factors" ' And ESEM
    
        start_line2 = FindLine(v_tt, start_line + 1) + 1
        
        If start_line2 > 1 Then
            Do While Num_Words_on_Line(start_line2) > 0
                b = Num_Words_on_Line(start_line2)
                Debug.Print "Found " & b & " words"
                For a = 1 To b
                    If Left(Word(start_line2, a), 1) <> "*" Then ' Do not count a variable if it begins with a "*". Needed for EFA and ESEM.
                        VarList(var_cntr, 1) = Word(start_line2, a)
                        VarList(var_cntr, 4) = v_type ' Variable type
                        If v_type = 6 Or v_type = 7 Or v_type = 8 Then ' Set as latent or not latent
                            VarList(var_cntr, 5) = True
                        Else
                            VarList(var_cntr, 5) = False
                        End If
                        Debug.Print "Found " & VarList(var_cntr, 1)
                        
                        ' Find labels
                        If alias_line > 1 Then
                            temp_line = FindLine("!" & VarList(var_cntr, 1) & ":", alias_line, False)
                            If temp_line > 0 Then
                                temp_al = line(temp_line)
                                temp_al = Right(temp_al, Len(temp_al) - InStr(1, temp_al, ":"))
                                VarList(var_cntr, 2) = temp_al
                                Debug.Print "Label: " & VarList(var_cntr, 2)
                            Else
                                VarList(var_cntr, 2) = ""
                            End If
                        End If
                        var_cntr = var_cntr + 1
                    End If
                Next
                start_line2 = start_line2 + 1
            Loop
        End If
    Next


    ' Add scales to the variables
    start_line = FindLine("!SCALES:") + 1
    If start_line > 1 Then
        Do While Num_Words_on_Line(start_line) > 0
            temp_al = Trim(line(start_line))
            scale_name = Mid(temp_al, 2, InStr(1, temp_al, ":") - 2)
            temp_al = Right(temp_al, Len(temp_al) - InStr(1, temp_al, ":")) ': Debug.Print temp_al
            temp_al = Split(temp_al, " ")
        
            For a = LBound(temp_al) To UBound(temp_al)
                temp_num = VarNum(temp_al(a))
                'Debug.Print "Variable number: " & temp_num & " " & temp_al(a)
                If temp_num > 0 Then
                    VarList(temp_num, 3) = scale_name
                    Debug.Print "Assigning " & scale_name & " to " & temp_al(a)
                End If
            Next
            start_line = start_line + 1
        Loop
    End If
    
    ' Add the names of the multilevel clusters
    Debug.Print "Searching for multilevel clusters..."
    temp_line = FindLine("  Cluster variable")
    If temp_line > 0 Then
        temp_words = Num_Words_on_Line(temp_line)
        If temp_words = 3 Then ' TWOLEVEL model
            levellist.Add "WITHIN"
            levellist.Add "BETWEEN"
            nLevels = 2
            Debug.Print "Levels: "; levellist(1) & " " & levellist(2)
        ElseIf temp_words > 3 Then
            levellist.Add "WITHIN"
            For a = 0 To temp_words - 3
                levellist.Add "BETWEEN " & Word(temp_line, temp_words - a)
                nLevels = nLevels + 1
                Debug.Print "Level:"; levellist(levellist.Count)
            Next
        End If
    End If
                
    ' Find the groups
    temp_line = FindLine("Number of groups")
    nGroups = Word(temp_line, 4)
    Debug.Print "Number of groups: " & nGroups
    If nGroups > 1 Then
        For a = 1 To nGroups
            grouplist.Add Word(temp_line + a + 1, 2)
            'SampleSize.Add Word(temp_line + a + 1, 3)
            Debug.Print "-Group " & a & ": " & grouplist.Item(grouplist.Count)
        Next
    'Else
        'SampleSize.Add Word(temp_line + 1, 4)
        'Debug.Print "Sample size: " & SampleSize(1)
    End If
    
    ' Find the mixtures
    Debug.Print "Adding the mixture patterns"
    If nCaLV = 1 Then
        start_line = FindLine("Class Counts and Proportions") + 5
        Do While Num_Words_on_Line(start_line) > 0
            mixturelist.Add Word(start_line, 1)
            start_line = start_line + 1
        Loop
        nMixtures = mixturelist.Count
    ElseIf nCaLV > 1 Then
        start_line = FindLine("Latent Class Variable Patterns") + 5
        Do While Num_Words_on_Line(start_line) > 0
            mix_patt = ""
            For n_word = 1 To Num_Words_on_Line(start_line)
                mix_patt = mix_patt & Word(start_line, n_word) & " "
            Next
            mixturelist.Add Trim(mix_patt)
            start_line = start_line + 1
        Loop
        nMixtures = mixturelist.Count
    End If
       
    ' Find the types of standardizations
    start_line = FindLine("STANDARDIZED MODEL RESULTS")
    nStandards = 1
    standardlist.Add ""
    If start_line > 0 Then
        If FindLine("STDYX Standardization", start_line) Then
            standardlist.Add "STDYX"
            stdyx_num = standardlist.Count
        End If
        If FindLine("STDY Standardization", start_line) Then standardlist.Add "STDY"
        If FindLine("STD Standardization", start_line) Then standardlist.Add "STD"
        nStandards = standardlist.Count
    End If
    If FindLine("LOGISTIC REGRESSION ODDS RATIO RESULTS", start_line) Then
        standardlist.Add "Odds Ratio"
        nStandards = nStandards + 1
    End If
    
    ' Add the other information about the variables from UNIVARIATE HIGHER-ORDER MOMENT DESCRIPTIVE STATISTICS
    Debug.Print "Adding information about the univariate descriptive statistics..."
    For a = 1 To nGroups
        temp_line = "UNIVARIATE HIGHER-ORDER MOMENT DESCRIPTIVE STATISTICS"
        If nGroups > 1 Then temp_line = temp_line & " FOR " & grouplist(a)
        start_line = FindLine(temp_line) + 5
        If start_line > 5 Then
            Do While Num_Words_on_Line(start_line) > 0
                var_n = VarNum(Word(start_line, 1))
                If var_n > 0 Then
                    Debug.Print "Loop", a, "of ", nGroups, "groups"
                    VarList(var_n, 6) = ADD_STRINGVECTOR(VarList(var_n, 6), Word(start_line, 2)) ' Mean
                    VarList(var_n, 7) = ADD_STRINGVECTOR(VarList(var_n, 7), Word(start_line + 1, 2)) ' Variance
                    VarList(var_n, 8) = ADD_STRINGVECTOR(VarList(var_n, 8), Word(start_line + 1, 1)) ' Sample size
                    VarList(var_n, 9) = ADD_STRINGVECTOR(VarList(var_n, 9), Word(start_line, 3)) ' Skew
                    VarList(var_n, 10) = ADD_STRINGVECTOR(VarList(var_n, 10), Word(start_line + 1, 3)) ' Kurtosis
                    VarList(var_n, 11) = ADD_STRINGVECTOR(VarList(var_n, 11), Word(start_line, 4)) ' Minimum
                    VarList(var_n, 12) = ADD_STRINGVECTOR(VarList(var_n, 12), Word(start_line, 6)) ' 20%
                    VarList(var_n, 13) = ADD_STRINGVECTOR(VarList(var_n, 13), Word(start_line, 7)) ' 40%
                    VarList(var_n, 14) = ADD_STRINGVECTOR(VarList(var_n, 14), Word(start_line, 8)) ' Median
                    VarList(var_n, 15) = ADD_STRINGVECTOR(VarList(var_n, 15), Word(start_line + 1, 6)) ' 60%
                    VarList(var_n, 16) = ADD_STRINGVECTOR(VarList(var_n, 16), Word(start_line + 1, 7)) ' 80%
                    VarList(var_n, 17) = ADD_STRINGVECTOR(VarList(var_n, 17), Word(start_line + 1, 4)) ' Maximum
                    VarList(var_n, 18) = ADD_STRINGVECTOR(VarList(var_n, 18), Word(start_line, 5)) ' % minimum
                    VarList(var_n, 19) = ADD_STRINGVECTOR(VarList(var_n, 19), Word(start_line + 1, 5)) ' % maximum
                    Debug.Print "Finished loop"
                End If
                start_line = start_line + 2
            Loop
        End If
    Next
    
    ' Add counts for categorical variables
    Debug.Print "Adding counts for categorical variables"
    For a = 1 To nGroups
        temp_line = "UNIVARIATE PROPORTIONS AND COUNTS FOR CATEGORICAL VARIABLES"
        If nGroups > 1 Then temp_line = temp_line & " FOR " & grouplist(a)
        start_line = FindLine(temp_line) + 2
        If start_line > 2 Then
            Do While Num_Words_on_Line(start_line) > 0
                If Num_Words_on_Line(start_line) = 1 Then
                    v_name = Word(start_line, 1)
                Else
                    v_num = VarNum(v_name)
                    If v_num > 0 Then
                        VarList(v_num, 8) = ADD_STRINGVECTOR(VarList(v_num, 8), Word(start_line, 4))
                    End If
                End If
                start_line = start_line + 1
            Loop
        End If
    Next
    
    ' Add proportions for COUNT variables
    Debug.Print "Adding proportions for COUNT variables"
    For a = 1 To nGroups
        temp_line = "COUNT PROPORTION OF ZERO, MINIMUM AND MAXIMUM VALUES"
        If nGroups > 1 Then temp_line = temp_line & " FOR " & grouplist(a)
        start_line = FindLine(temp_line) + 2
        If start_line > 2 Then
            Do While Num_Words_on_Line(start_line) > 0
                v_num = VarNum(Word(start_line, 1))
                If v_num > 0 Then
                    VarList(v_num, 20) = ADD_STRINGVECTOR(VarList(v_num, 20), Word(start_line, 2))
                    VarList(v_num, 18) = ADD_STRINGVECTOR(VarList(v_num, 18), Word(start_line, 3))
                    VarList(v_num, 19) = ADD_STRINGVECTOR(VarList(v_num, 19), Word(start_line, 4))
                End If
                start_line = start_line + 1
            Loop
        End If
    Next
    
    ' Find the estimator
    temp_line = FindLine("SUMMARY OF ANALYSIS")
    Estimator = Word(FindLine("Estimator", temp_line), 2)
End Sub

Sub AddModelResults()
    ' Add model results
    ' Check to see if a model exists...
    start_line = FindLine("MODEL RESULTS")
    If start_line > 0 Then
        IsModel = True
        start_line2 = start_line + 4
        
        Debug.Print nStandards
        Debug.Print nGroups
        Debug.Print nLevels
        Debug.Print nMixtures
        ' Resize the observed matrix array based on the number of variables and the number of groups
        Debug.Print "0"
        ReDim mod_matrix_directional(1 To nTotalV, 1 To nTotalV, 1 To nStandards, 1 To nGroups, 1 To nLevels, 1 To nMixtures)
        Debug.Print "1"
        ReDim mod_matrix_nondirectional(1 To nTotalV, 1 To nTotalV, 1 To nStandards, 1 To nGroups, 1 To nLevels, 1 To nMixtures)
        Debug.Print "2"
        ReDim mod_vector(1 To nTotalV, 1 To nStandards, 1 To nGroups, 1 To nLevels, 1 To nMixtures)
        ReDim mod_rsquare(1 To nTotalV, 1 To nGroups, 1 To nLevels, 1 To nMixtures)
        Debug.Print "3"

        'ReDim mod_matrix_directional(1 To nTotalV, 1 To nTotalV, 1 To standardlist.Count, 1 To mod_matrix_list.Count)
        'ReDim mod_matrix_nondirectional(1 To nTotalV, 1 To nTotalV, 1 To standardlist.Count, 1 To mod_matrix_list.Count)
        'ReDim mod_vector(1 To nTotalV, 1 To standardlist.Count, 1 To mod_matrix_list.Count)

        Debug.Print "Model Results -- dimensions are " & nTotalV & "X * " & nTotalV & "Y * " & standardlist.Count & " standardizations, * " & nGroups & " groups, " & nLevels & " levels, " & nMixtures & "mixtures."
        
        ' Set up some dummy group names; defaults to the unstandardised model
        Dim results_found: results_found = True
        grp_name = ""
        lev_name = ""
        mix_name = ""
        std_name = ""
        
        ' Set model and std number to 1 as default
        model_num = 1
        std_num = 1
        SMR_line = FindLine("STANDARDIZED MODEL RESULTS")
        CISMR_line = FindLine("INTERVALS OF STANDARDIZED MODEL RESULTS", SMR_line)
        OR_line = FindLine("LOGISTIC REGRESSION ODDS RATIO RESULTS")
        
        ' Set up other counters
        dv_count = 1  ' Counter variable for the number of dependent variables
        IV_count = 1  ' Counter variable for the number of independent variables / predictors
        error_count = 0 'Counter variable for errors; halts looking for output if more than 3 encountered
        
        Do While results_found = True
            model_name = SET_STRINGVECTOR(grp_name, lev_name, mix_name)
            Debug.Print "Model name: " & model_name

            ' Find the model number and standardisation number
            If model_num = 0 Then
                Debug.Print "Trying to match a model..."
                
                IV_count = 1
                dv_count = 1
                
                For a = 1 To mod_matrix_list.Count
                    Debug.Print "Trying to match model "; UCase(model_name); " to "; UCase(mod_matrix_list(a))
                    If UCase(model_name) = UCase(mod_matrix_list(a)) Then
                        model_num = a
                        Debug.Print "Matched model "; a; " of "; mod_matrix_list.Count; ": "; UCase(model_name); ", and "; UCase(mod_matrix_list(a))
                        Exit For
                    End If
                Next

                std_num = 1
                
                For a = 1 To standardlist.Count
                    Debug.Print "Model " & ModelName(model_num) & ", trying to match standardization " & a & " of " & standardlist.Count & ": " & UCase(std_name) & " to " & UCase(standardlist(a))
                    If UCase(std_name) = UCase(standardlist(a)) Then
                        std_num = a
                        Debug.Print "Matched standardization " & a & " of " & standardlist.Count & ": " & UCase(std_name)
                        Exit For
                    End If
                Next
            End If
            start_line2 = Next_NonBlank_Line(start_line2)
        
            ' Try to determine what kind of output is appearing here...
            ' Detect if it's a group name
            Debug.Print "Determining type of output..."
            If Word(start_line2, 1) = "Group" Then
                error_count = 0
                grp_name = Word(start_line2, 2)
                model_num = 0
                Debug.Print "Found group name: " & grp_name
            ' Detect if it's the name of a level
            ElseIf Word_right(start_line2, 1) = "Level" Then
                error_count = 0
                lev_name = Words_exceptfor(start_line2, 1)
                model_num = 0
                Debug.Print "Found level name: " & lev_name
            ' Detect if it's a latent class / mixture pattern (more than one categorical variable)
            ElseIf Word(start_line2, 1) = "Latent" And Word(start_line2, 2) = "Class" And Word(start_line2, 3) = "Pattern" Then
                error_count = 0
                mix_name = ""
                For a = 4 To Num_Words_on_Line(start_line2)
                    mix_name = mix_name & Word(start_line2, a)
                    If a < Num_Words_on_Line(start_line2) Then mix_name = mix_name & " "
                Next
                model_num = 0
                Debug.Print "Found mixture: " & mix_name
            ' Detect if it's a mixture (only one categorical variable)
            ElseIf Word(start_line2, 1) = "Latent" And Word(start_line2, 2) = "Class" And Word(start_line2, 3) <> "Pattern" Then
                error_count = 0
                mix_name = Word(start_line2, 3)
                model_num = 0
                Debug.Print "Found mixture: " & mix_name
            ' Detect if it's the standardized results section
            ElseIf start_line2 = SMR_line Or start_line2 = CISMR_line Then
                error_count = 0
                start_line2 = start_line2 + 1
                Debug.Print "Found standardized model results"
            ' Detect each type of standardization (STDYX, STDY, STD)
            ElseIf Word(start_line2, 2) = "Standardization" Then
                error_count = 0
                std_name = Word(start_line2, 1)
                model_num = 0
                Debug.Print "Found " & std_name & " output"
                start_line2 = start_line2 + 1
            ElseIf (UCase(Word(start_line2, 1)) = "STDYX" Or UCase(Word(start_line2, 1)) = "STDY" Or UCase(Word(start_line2, 1)) = "STD") And Word(start_line2 + 1, 1) = "Estimate" Then
                error_count = 0
                std_name = UCase(Word(start_line2, 1))
                model_num = 0
                Debug.Print "Found " & std_name & " output"
                start_line2 = start_line2 + 1
            ElseIf start_line2 = OR_line Then
                error_count = 0
                std_name = "Odds Ratio"
                model_num = 0
                Debug.Print "Found odds ratio output"
                start_line2 = start_line2 + 4
            ElseIf Word(start_line2, 1) = "R-SQUARE" Then
                error_count = 0
                Debug.Print "Found R-Square"
                IsRSquare = True
                start_line2 = start_line2 + 1
            ElseIf CONTAINS(line(start_line2), "INTERVALS OF MODEL RESULTS") Then
                error_count = 0
                Debug.Print "Found Confidence Intervals"
                start_line2 = start_line2 + 3
                IsCInterval = True
                model_num = 0
                std_name = ""
            ElseIf CONTAINS(line(start_line2), "INTERVALS FOR THE LOGISTIC REGRESSION ODDS RATIO") Then
                error_count = 0
                Debug.Print "Found Confidence Intervals for logistic regression"
                start_line2 = start_line2 + 1
                IsCInterval = True
                model_num = 0
                std_name = "Odds Ratio"
            ' Put in regression coefficients
            ElseIf Word_right(start_line2, 1) = "ON" Then
                error_count = 0
                Debug.Print "Found regression coefficients"
                
                dv_num = VarNum(Word(start_line2, 1))
                dv_cat = CategoryNumber(Word(start_line2, 1))
                dv_suf = ""
                If dv_cat > 1 Then dv_suf = "_" & CStr(dv_cat)

                If GET_STRINGVECTOR(VarList(dv_num, 23), model_num) = "NA" Then ' Label as DV
                    VarList(dv_num, 23) = ADD_STRINGVECTOR(VarList(dv_num, 23), dv_count, model_num)
                    dv_count = dv_count + 1
                End If
                'Debug.Print "Found regression coefficients, trying " & ":" & iv_num & ":"; dv_num & ":"; std_num & ":" & model_num
                For a = start_line2 + 1 To Next_Blank_Line(start_line2) - 1
                    iv_num = VarNum(Word(a, 1))
                    If GET_STRINGVECTOR(VarList(iv_num, 24), model_num) = "NA" Then ' Label as IV
                        VarList(iv_num, 24) = ADD_STRINGVECTOR(VarList(iv_num, 24), IV_count, model_num)
                        IV_count = IV_count + 1
                    End If
                    mod_matrix_directional(iv_num, dv_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)) = ADD_STRINGVECTOR(mod_matrix_directional(iv_num, dv_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)), "ON", "Type") ' Indicates that this is specifically an antecedent
                    If dv_cat > 1 Then
                        mod_matrix_directional(iv_num, dv_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)) = ADD_STRINGVECTOR(mod_matrix_directional(iv_num, dv_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)), CStr(dv_cat), "NCategories") ' Indicates the number of categories the DV has
                    End If
                    If IsCInterval = True Then
                        mod_matrix_directional(iv_num, dv_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)) = ADD_STRINGVECTOR(mod_matrix_directional(iv_num, dv_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)), Word(a, 2), "CI005" & dv_suf)
                        mod_matrix_directional(iv_num, dv_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)) = ADD_STRINGVECTOR(mod_matrix_directional(iv_num, dv_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)), Word(a, 3), "CI025" & dv_suf)
                        mod_matrix_directional(iv_num, dv_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)) = ADD_STRINGVECTOR(mod_matrix_directional(iv_num, dv_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)), Word(a, 4), "CI050" & dv_suf)
                        mod_matrix_directional(iv_num, dv_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)) = ADD_STRINGVECTOR(mod_matrix_directional(iv_num, dv_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)), Word(a, 6), "CI950" & dv_suf)
                        mod_matrix_directional(iv_num, dv_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)) = ADD_STRINGVECTOR(mod_matrix_directional(iv_num, dv_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)), Word(a, 7), "CI975" & dv_suf)
                        mod_matrix_directional(iv_num, dv_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)) = ADD_STRINGVECTOR(mod_matrix_directional(iv_num, dv_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)), Word(a, 8), "CI995" & dv_suf)
                    ElseIf Estimator = "BAYES" Then
                        mod_matrix_directional(iv_num, dv_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)) = ADD_STRINGVECTOR(mod_matrix_directional(iv_num, dv_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)), Word(a, 2), "Estimate" & dv_suf)
                        mod_matrix_directional(iv_num, dv_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)) = ADD_STRINGVECTOR(mod_matrix_directional(iv_num, dv_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)), Word(a, 3), "SD" & dv_suf)
                        mod_matrix_directional(iv_num, dv_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)) = ADD_STRINGVECTOR(mod_matrix_directional(iv_num, dv_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)), Word(a, 4), "p" & dv_suf)
                        mod_matrix_directional(iv_num, dv_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)) = ADD_STRINGVECTOR(mod_matrix_directional(iv_num, dv_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)), Word(a, 5), "CI025" & dv_suf)
                        mod_matrix_directional(iv_num, dv_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)) = ADD_STRINGVECTOR(mod_matrix_directional(iv_num, dv_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)), Word(a, 6), "CI975" & dv_suf)
                    ElseIf IsImputation Then
                        mod_matrix_directional(iv_num, dv_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)) = ADD_STRINGVECTOR(mod_matrix_directional(iv_num, dv_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)), Word(a, 2), "Estimate" & dv_suf)
                        mod_matrix_directional(iv_num, dv_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)) = ADD_STRINGVECTOR(mod_matrix_directional(iv_num, dv_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)), Word(a, 3), "SE" & dv_suf)
                        mod_matrix_directional(iv_num, dv_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)) = ADD_STRINGVECTOR(mod_matrix_directional(iv_num, dv_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)), Word(a, 4), "t" & dv_suf)
                        mod_matrix_directional(iv_num, dv_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)) = ADD_STRINGVECTOR(mod_matrix_directional(iv_num, dv_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)), Word(a, 5), "p" & dv_suf)
                        mod_matrix_directional(iv_num, dv_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)) = ADD_STRINGVECTOR(mod_matrix_directional(iv_num, dv_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)), Word(a, 6), "RateMissing" & dv_suf)
                    ElseIf IsMonteCarlo Then
                        mod_matrix_directional(iv_num, dv_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)) = ADD_STRINGVECTOR(mod_matrix_directional(iv_num, dv_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)), Word(a, 2), "Population" & dv_suf)
                        mod_matrix_directional(iv_num, dv_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)) = ADD_STRINGVECTOR(mod_matrix_directional(iv_num, dv_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)), Word(a, 3), "Estimate" & dv_suf)
                        mod_matrix_directional(iv_num, dv_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)) = ADD_STRINGVECTOR(mod_matrix_directional(iv_num, dv_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)), Word(a, 4), "SD" & dv_suf)
                        mod_matrix_directional(iv_num, dv_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)) = ADD_STRINGVECTOR(mod_matrix_directional(iv_num, dv_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)), Word(a, 5), "SE" & dv_suf)
                        mod_matrix_directional(iv_num, dv_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)) = ADD_STRINGVECTOR(mod_matrix_directional(iv_num, dv_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)), Word(a, 6), "MSE" & dv_suf)
                        mod_matrix_directional(iv_num, dv_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)) = ADD_STRINGVECTOR(mod_matrix_directional(iv_num, dv_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)), Word(a, 7), "95Cov" & dv_suf)
                        mod_matrix_directional(iv_num, dv_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)) = ADD_STRINGVECTOR(mod_matrix_directional(iv_num, dv_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)), Word(a, 8), "%Sig" & dv_suf)
                    Else
                        mod_matrix_directional(iv_num, dv_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)) = ADD_STRINGVECTOR(mod_matrix_directional(iv_num, dv_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)), Word(a, 2), "Estimate" & dv_suf)
                        mod_matrix_directional(iv_num, dv_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)) = ADD_STRINGVECTOR(mod_matrix_directional(iv_num, dv_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)), Word(a, 3), "SE" & dv_suf)
                        mod_matrix_directional(iv_num, dv_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)) = ADD_STRINGVECTOR(mod_matrix_directional(iv_num, dv_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)), Word(a, 4), "t" & dv_suf)
                        mod_matrix_directional(iv_num, dv_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)) = ADD_STRINGVECTOR(mod_matrix_directional(iv_num, dv_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)), Word(a, 5), "p" & dv_suf)
                    End If
                    Debug.Print dv_num & " ON " & iv_num & ":" & 1, model_num, mod_matrix_directional(iv_num, dv_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num))
                Next
                start_line2 = a
            ' Put in factor loadings
            ElseIf Word_right(start_line2, 1) = "BY" Or Word_right(start_line2, 1) = "|" Then
                error_count = 0
                Debug.Print "Found factor loadings"
                iv_num = VarNum(Word(start_line2, 1))
                Debug.Print "IV number", iv_num
                For a = start_line2 + 1 To Next_Blank_Line(start_line2) - 1
                    dv_num = VarNum(Word(a, 1))
                    Debug.Print "DV number", dv_num
                        mod_matrix_directional(iv_num, dv_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)) = ADD_STRINGVECTOR(mod_matrix_directional(iv_num, dv_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)), Word_right(start_line2, 1), "Type") ' Indicates that this is specifically a factor loading
                    If IsCInterval = True Then
                        mod_matrix_directional(iv_num, dv_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)) = ADD_STRINGVECTOR(mod_matrix_directional(iv_num, dv_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)), Word(a, 2), "CI005")
                        mod_matrix_directional(iv_num, dv_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)) = ADD_STRINGVECTOR(mod_matrix_directional(iv_num, dv_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)), Word(a, 3), "CI025")
                        mod_matrix_directional(iv_num, dv_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)) = ADD_STRINGVECTOR(mod_matrix_directional(iv_num, dv_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)), Word(a, 4), "CI050")
                        mod_matrix_directional(iv_num, dv_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)) = ADD_STRINGVECTOR(mod_matrix_directional(iv_num, dv_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)), Word(a, 6), "CI950")
                        mod_matrix_directional(iv_num, dv_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)) = ADD_STRINGVECTOR(mod_matrix_directional(iv_num, dv_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)), Word(a, 7), "CI975")
                        mod_matrix_directional(iv_num, dv_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)) = ADD_STRINGVECTOR(mod_matrix_directional(iv_num, dv_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)), Word(a, 8), "CI995")
                    ElseIf Estimator = "BAYES" Then
                        mod_matrix_directional(iv_num, dv_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)) = ADD_STRINGVECTOR(mod_matrix_directional(iv_num, dv_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)), Word(a, 2), "Estimate")
                        mod_matrix_directional(iv_num, dv_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)) = ADD_STRINGVECTOR(mod_matrix_directional(iv_num, dv_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)), Word(a, 3), "SD")
                        mod_matrix_directional(iv_num, dv_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)) = ADD_STRINGVECTOR(mod_matrix_directional(iv_num, dv_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)), Word(a, 4), "p")
                        mod_matrix_directional(iv_num, dv_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)) = ADD_STRINGVECTOR(mod_matrix_directional(iv_num, dv_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)), Word(a, 5), "CI025")
                        mod_matrix_directional(iv_num, dv_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)) = ADD_STRINGVECTOR(mod_matrix_directional(iv_num, dv_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)), Word(a, 6), "CI975")
                    ElseIf IsImputation Then
                        mod_matrix_directional(iv_num, dv_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)) = ADD_STRINGVECTOR(mod_matrix_directional(iv_num, dv_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)), Word(a, 2), "Estimate")
                        mod_matrix_directional(iv_num, dv_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)) = ADD_STRINGVECTOR(mod_matrix_directional(iv_num, dv_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)), Word(a, 3), "SE")
                        mod_matrix_directional(iv_num, dv_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)) = ADD_STRINGVECTOR(mod_matrix_directional(iv_num, dv_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)), Word(a, 4), "t")
                        mod_matrix_directional(iv_num, dv_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)) = ADD_STRINGVECTOR(mod_matrix_directional(iv_num, dv_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)), Word(a, 5), "p")
                        mod_matrix_directional(iv_num, dv_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)) = ADD_STRINGVECTOR(mod_matrix_directional(iv_num, dv_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)), Word(a, 6), "RateMissing")
                    ElseIf IsMonteCarlo Then
                        mod_matrix_directional(iv_num, dv_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)) = ADD_STRINGVECTOR(mod_matrix_directional(iv_num, dv_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)), Word(a, 2), "Population")
                        mod_matrix_directional(iv_num, dv_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)) = ADD_STRINGVECTOR(mod_matrix_directional(iv_num, dv_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)), Word(a, 3), "Estimate")
                        mod_matrix_directional(iv_num, dv_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)) = ADD_STRINGVECTOR(mod_matrix_directional(iv_num, dv_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)), Word(a, 4), "SD")
                        mod_matrix_directional(iv_num, dv_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)) = ADD_STRINGVECTOR(mod_matrix_directional(iv_num, dv_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)), Word(a, 5), "SE")
                        mod_matrix_directional(iv_num, dv_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)) = ADD_STRINGVECTOR(mod_matrix_directional(iv_num, dv_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)), Word(a, 6), "MSE")
                        mod_matrix_directional(iv_num, dv_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)) = ADD_STRINGVECTOR(mod_matrix_directional(iv_num, dv_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)), Word(a, 7), "95Cov")
                        mod_matrix_directional(iv_num, dv_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)) = ADD_STRINGVECTOR(mod_matrix_directional(iv_num, dv_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)), Word(a, 8), "%Sig")
                    Else
                        mod_matrix_directional(iv_num, dv_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)) = ADD_STRINGVECTOR(mod_matrix_directional(iv_num, dv_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)), Word(a, 2), "Estimate")
                        mod_matrix_directional(iv_num, dv_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)) = ADD_STRINGVECTOR(mod_matrix_directional(iv_num, dv_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)), Word(a, 3), "SE")
                        mod_matrix_directional(iv_num, dv_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)) = ADD_STRINGVECTOR(mod_matrix_directional(iv_num, dv_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)), Word(a, 4), "t")
                        mod_matrix_directional(iv_num, dv_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)) = ADD_STRINGVECTOR(mod_matrix_directional(iv_num, dv_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)), Word(a, 5), "p")
                    End If
                    Debug.Print iv_num & " BY " & dv_num & ":" & ", Model:"; model_num, mod_matrix_directional(iv_num, dv_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num))
                Next
                start_line2 = a
            ' Put in covariance / correlations
            ElseIf Word_right(start_line2, 1) = "WITH" Then
                error_count = 0
                Debug.Print "Found covariances / correlations"
                v_num1 = VarNum(Word(start_line2, 1))
                For a = start_line2 + 1 To Next_Blank_Line(start_line2) - 1
                    v_num2 = VarNum(Word(a, 1))
                    Debug.Print "Trying: "; v_num1; ":" & v_num2; ":"; std_num & ":" & model_num
                    If IsCInterval = True Then
                        mod_matrix_nondirectional(v_num1, v_num2, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)) = ADD_STRINGVECTOR(mod_matrix_nondirectional(v_num1, v_num2, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)), Word(a, 2), "CI005")
                        mod_matrix_nondirectional(v_num1, v_num2, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)) = ADD_STRINGVECTOR(mod_matrix_nondirectional(v_num1, v_num2, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)), Word(a, 3), "CI025")
                        mod_matrix_nondirectional(v_num1, v_num2, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)) = ADD_STRINGVECTOR(mod_matrix_nondirectional(v_num1, v_num2, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)), Word(a, 4), "CI050")
                        mod_matrix_nondirectional(v_num1, v_num2, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)) = ADD_STRINGVECTOR(mod_matrix_nondirectional(v_num1, v_num2, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)), Word(a, 6), "CI950")
                        mod_matrix_nondirectional(v_num1, v_num2, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)) = ADD_STRINGVECTOR(mod_matrix_nondirectional(v_num1, v_num2, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)), Word(a, 7), "CI975")
                        mod_matrix_nondirectional(v_num1, v_num2, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)) = ADD_STRINGVECTOR(mod_matrix_nondirectional(v_num1, v_num2, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)), Word(a, 8), "CI995")
                    ElseIf Estimator = "BAYES" Then
                        mod_matrix_nondirectional(v_num1, v_num2, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)) = ADD_STRINGVECTOR(mod_matrix_nondirectional(v_num1, v_num2, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)), Word(a, 2), "Estimate")
                        mod_matrix_nondirectional(v_num1, v_num2, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)) = ADD_STRINGVECTOR(mod_matrix_nondirectional(v_num1, v_num2, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)), Word(a, 3), "SD")
                        mod_matrix_nondirectional(v_num1, v_num2, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)) = ADD_STRINGVECTOR(mod_matrix_nondirectional(v_num1, v_num2, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)), Word(a, 4), "p")
                        mod_matrix_nondirectional(v_num1, v_num2, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)) = ADD_STRINGVECTOR(mod_matrix_nondirectional(v_num1, v_num2, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)), Word(a, 5), "CI025")
                        mod_matrix_nondirectional(v_num1, v_num2, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)) = ADD_STRINGVECTOR(mod_matrix_nondirectional(v_num1, v_num2, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)), Word(a, 6), "CI975")
                    ElseIf IsImputation Then
                        mod_matrix_nondirectional(v_num1, v_num2, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)) = ADD_STRINGVECTOR(mod_matrix_nondirectional(v_num1, v_num2, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)), Word(a, 2), "Estimate")
                        mod_matrix_nondirectional(v_num1, v_num2, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)) = ADD_STRINGVECTOR(mod_matrix_nondirectional(v_num1, v_num2, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)), Word(a, 3), "SE")
                        mod_matrix_nondirectional(v_num1, v_num2, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)) = ADD_STRINGVECTOR(mod_matrix_nondirectional(v_num1, v_num2, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)), Word(a, 4), "t")
                        mod_matrix_nondirectional(v_num1, v_num2, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)) = ADD_STRINGVECTOR(mod_matrix_nondirectional(v_num1, v_num2, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)), Word(a, 5), "p")
                        mod_matrix_nondirectional(v_num1, v_num2, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)) = ADD_STRINGVECTOR(mod_matrix_nondirectional(v_num1, v_num2, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)), Word(a, 6), "RateMissing")
                    ElseIf IsMonteCarlo Then
                        mod_matrix_nondirectional(v_num1, v_num2, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)) = ADD_STRINGVECTOR(mod_matrix_nondirectional(v_num1, v_num2, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)), Word(a, 2), "Population")
                        mod_matrix_nondirectional(v_num1, v_num2, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)) = ADD_STRINGVECTOR(mod_matrix_nondirectional(v_num1, v_num2, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)), Word(a, 3), "Estimate")
                        mod_matrix_nondirectional(v_num1, v_num2, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)) = ADD_STRINGVECTOR(mod_matrix_nondirectional(v_num1, v_num2, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)), Word(a, 4), "SD")
                        mod_matrix_nondirectional(v_num1, v_num2, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)) = ADD_STRINGVECTOR(mod_matrix_nondirectional(v_num1, v_num2, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)), Word(a, 5), "SE")
                        mod_matrix_nondirectional(v_num1, v_num2, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)) = ADD_STRINGVECTOR(mod_matrix_nondirectional(v_num1, v_num2, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)), Word(a, 6), "MSE")
                        mod_matrix_nondirectional(v_num1, v_num2, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)) = ADD_STRINGVECTOR(mod_matrix_nondirectional(v_num1, v_num2, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)), Word(a, 7), "95Cov")
                        mod_matrix_nondirectional(v_num1, v_num2, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)) = ADD_STRINGVECTOR(mod_matrix_nondirectional(v_num1, v_num2, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)), Word(a, 8), "%Sig")
                    Else
                        mod_matrix_nondirectional(v_num1, v_num2, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)) = ADD_STRINGVECTOR(mod_matrix_nondirectional(v_num1, v_num2, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)), Word(a, 2), "Estimate")
                        mod_matrix_nondirectional(v_num1, v_num2, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)) = ADD_STRINGVECTOR(mod_matrix_nondirectional(v_num1, v_num2, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)), Word(a, 3), "SE")
                        mod_matrix_nondirectional(v_num1, v_num2, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)) = ADD_STRINGVECTOR(mod_matrix_nondirectional(v_num1, v_num2, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)), Word(a, 4), "t")
                        mod_matrix_nondirectional(v_num1, v_num2, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)) = ADD_STRINGVECTOR(mod_matrix_nondirectional(v_num1, v_num2, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)), Word(a, 5), "p")
                    End If
                    mod_matrix_nondirectional(v_num2, v_num1, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)) = mod_matrix_nondirectional(v_num1, v_num2, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num))
                    Debug.Print v_num1 & " WITH " & v_num2 & ":" & 1, model_num, mod_matrix_nondirectional(v_num1, v_num2, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num))
                Next
                start_line2 = a
            ElseIf Word(start_line2, 1) = "Intercepts" Or Word(start_line2, 1) = "Means" Or Word(start_line2, 1) = "Thresholds" Then
                error_count = 0
                Debug.Print "Found means / intercepts / thresholds"
                For a = start_line2 + 1 To Next_Blank_Line(start_line2) - 1
                
                    v_num = VarNum(Word(a, 1))
                    v_cat = CategoryNumber(Word(a, 1))
                    v_suf = ""
                    If v_cat > 1 Then
                        v_suf = "_" & CStr(v_cat)
                        mod_vector(v_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)) = ADD_STRINGVECTOR(mod_vector(v_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)), CStr(v_cat), "NCategories")
                    End If
                
                    mod_vector(v_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)) = ADD_STRINGVECTOR(mod_vector(v_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)), Word(start_line2, 1), "Type")
                    If IsCInterval = True Then
                        mod_vector(v_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)) = ADD_STRINGVECTOR(mod_vector(v_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)), Word(a, 2), "CI005" & v_suf)
                        mod_vector(v_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)) = ADD_STRINGVECTOR(mod_vector(v_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)), Word(a, 3), "CI025" & v_suf)
                        mod_vector(v_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)) = ADD_STRINGVECTOR(mod_vector(v_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)), Word(a, 4), "CI050" & v_suf)
                        mod_vector(v_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)) = ADD_STRINGVECTOR(mod_vector(v_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)), Word(a, 6), "CI950" & v_suf)
                        mod_vector(v_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)) = ADD_STRINGVECTOR(mod_vector(v_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)), Word(a, 7), "CI975" & v_suf)
                        mod_vector(v_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)) = ADD_STRINGVECTOR(mod_vector(v_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)), Word(a, 8), "CI995" & v_suf)
                    ElseIf Estimator = "BAYES" Then
                        mod_vector(v_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)) = ADD_STRINGVECTOR(mod_vector(v_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)), Word(a, 2), "Estimate" & v_suf)
                        mod_vector(v_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)) = ADD_STRINGVECTOR(mod_vector(v_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)), Word(a, 3), "SD" & v_suf)
                        mod_vector(v_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)) = ADD_STRINGVECTOR(mod_vector(v_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)), Word(a, 4), "p" & v_suf)
                        mod_vector(v_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)) = ADD_STRINGVECTOR(mod_vector(v_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)), Word(a, 5), "CI025" & v_suf)
                        mod_vector(v_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)) = ADD_STRINGVECTOR(mod_vector(v_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)), Word(a, 6), "CI975" & v_suf)
                    ElseIf IsImputation Then
                        mod_vector(v_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)) = ADD_STRINGVECTOR(mod_vector(v_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)), Word(a, 2), "Estimate" & v_suf)
                        mod_vector(v_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)) = ADD_STRINGVECTOR(mod_vector(v_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)), Word(a, 3), "SE" & v_suf)
                        mod_vector(v_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)) = ADD_STRINGVECTOR(mod_vector(v_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)), Word(a, 4), "t" & v_suf)
                        mod_vector(v_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)) = ADD_STRINGVECTOR(mod_vector(v_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)), Word(a, 5), "p" & v_suf)
                        mod_vector(v_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)) = ADD_STRINGVECTOR(mod_vector(v_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)), Word(a, 6), "RateMissing" & v_suf)
                    ElseIf IsMonteCarlo Then
                        mod_vector(v_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)) = ADD_STRINGVECTOR(mod_vector(v_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)), Word(a, 2), "Population" & v_suf)
                        mod_vector(v_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)) = ADD_STRINGVECTOR(mod_vector(v_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)), Word(a, 3), "Estimate" & v_suf)
                        mod_vector(v_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)) = ADD_STRINGVECTOR(mod_vector(v_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)), Word(a, 4), "SD" & v_suf)
                        mod_vector(v_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)) = ADD_STRINGVECTOR(mod_vector(v_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)), Word(a, 5), "SE" & v_suf)
                        mod_vector(v_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)) = ADD_STRINGVECTOR(mod_vector(v_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)), Word(a, 6), "MSE" & v_suf)
                        mod_vector(v_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)) = ADD_STRINGVECTOR(mod_vector(v_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)), Word(a, 7), "95Cov" & v_suf)
                        mod_vector(v_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)) = ADD_STRINGVECTOR(mod_vector(v_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)), Word(a, 8), "%Sig" & v_suf)
                    Else
                        mod_vector(v_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)) = ADD_STRINGVECTOR(mod_vector(v_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)), Word(a, 2), "Estimate" & v_suf)
                        mod_vector(v_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)) = ADD_STRINGVECTOR(mod_vector(v_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)), Word(a, 3), "SE" & v_suf)
                        mod_vector(v_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)) = ADD_STRINGVECTOR(mod_vector(v_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)), Word(a, 4), "t" & v_suf)
                        mod_vector(v_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)) = ADD_STRINGVECTOR(mod_vector(v_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)), Word(a, 5), "p" & v_suf)
                    End If
                    Debug.Print v_num & ": " & mod_vector(v_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num))
                Next
                start_line2 = a
            ElseIf Word(start_line2, 1) = "Variances" Or (Word(start_line2, 1) = "Residual" And Word(start_line2, 2) = "Variances") Then
                error_count = 0
                Debug.Print "Found variances / residual variances"
                For a = start_line2 + 1 To Next_Blank_Line(start_line2) - 1
                    v_num = VarNum(Word(a, 1))
                    mod_matrix_nondirectional(v_num, v_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)) = ADD_STRINGVECTOR(mod_matrix_nondirectional(v_num, v_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)), Word(start_line2, 1), "Type")
                    If IsCInterval = True Then
                        mod_matrix_nondirectional(v_num, v_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)) = ADD_STRINGVECTOR(mod_matrix_nondirectional(v_num, v_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)), Word(a, 2), "CI005")
                        mod_matrix_nondirectional(v_num, v_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)) = ADD_STRINGVECTOR(mod_matrix_nondirectional(v_num, v_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)), Word(a, 3), "CI025")
                        mod_matrix_nondirectional(v_num, v_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)) = ADD_STRINGVECTOR(mod_matrix_nondirectional(v_num, v_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)), Word(a, 4), "CI050")
                        mod_matrix_nondirectional(v_num, v_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)) = ADD_STRINGVECTOR(mod_matrix_nondirectional(v_num, v_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)), Word(a, 6), "CI950")
                        mod_matrix_nondirectional(v_num, v_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)) = ADD_STRINGVECTOR(mod_matrix_nondirectional(v_num, v_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)), Word(a, 7), "CI975")
                        mod_matrix_nondirectional(v_num, v_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)) = ADD_STRINGVECTOR(mod_matrix_nondirectional(v_num, v_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)), Word(a, 8), "CI995")
                    ElseIf Estimator = "BAYES" Then
                        mod_matrix_nondirectional(v_num, v_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)) = ADD_STRINGVECTOR(mod_matrix_nondirectional(v_num, v_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)), Word(a, 2), "Estimate")
                        mod_matrix_nondirectional(v_num, v_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)) = ADD_STRINGVECTOR(mod_matrix_nondirectional(v_num, v_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)), Word(a, 3), "SD")
                        mod_matrix_nondirectional(v_num, v_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)) = ADD_STRINGVECTOR(mod_matrix_nondirectional(v_num, v_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)), Word(a, 4), "p")
                        mod_matrix_nondirectional(v_num, v_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)) = ADD_STRINGVECTOR(mod_matrix_nondirectional(v_num, v_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)), Word(a, 5), "CI025")
                        mod_matrix_nondirectional(v_num, v_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)) = ADD_STRINGVECTOR(mod_matrix_nondirectional(v_num, v_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)), Word(a, 6), "CI975")
                    ElseIf IsImputation Then
                        mod_matrix_nondirectional(v_num, v_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)) = ADD_STRINGVECTOR(mod_matrix_nondirectional(v_num, v_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)), Word(a, 2), "Estimate")
                        mod_matrix_nondirectional(v_num, v_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)) = ADD_STRINGVECTOR(mod_matrix_nondirectional(v_num, v_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)), Word(a, 3), "SE")
                        mod_matrix_nondirectional(v_num, v_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)) = ADD_STRINGVECTOR(mod_matrix_nondirectional(v_num, v_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)), Word(a, 4), "t")
                        mod_matrix_nondirectional(v_num, v_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)) = ADD_STRINGVECTOR(mod_matrix_nondirectional(v_num, v_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)), Word(a, 5), "p")
                        mod_matrix_nondirectional(v_num, v_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)) = ADD_STRINGVECTOR(mod_matrix_nondirectional(v_num, v_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)), Word(a, 6), "RateMissing")
                    ElseIf IsMonteCarlo Then
                        mod_matrix_nondirectional(v_num, v_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)) = ADD_STRINGVECTOR(mod_matrix_nondirectional(v_num, v_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)), Word(a, 2), "Population")
                        mod_matrix_nondirectional(v_num, v_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)) = ADD_STRINGVECTOR(mod_matrix_nondirectional(v_num, v_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)), Word(a, 3), "Estimate")
                        mod_matrix_nondirectional(v_num, v_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)) = ADD_STRINGVECTOR(mod_matrix_nondirectional(v_num, v_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)), Word(a, 4), "SD")
                        mod_matrix_nondirectional(v_num, v_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)) = ADD_STRINGVECTOR(mod_matrix_nondirectional(v_num, v_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)), Word(a, 5), "SE")
                        mod_matrix_nondirectional(v_num, v_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)) = ADD_STRINGVECTOR(mod_matrix_nondirectional(v_num, v_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)), Word(a, 6), "MSE")
                        mod_matrix_nondirectional(v_num, v_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)) = ADD_STRINGVECTOR(mod_matrix_nondirectional(v_num, v_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)), Word(a, 7), "95Cov")
                        mod_matrix_nondirectional(v_num, v_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)) = ADD_STRINGVECTOR(mod_matrix_nondirectional(v_num, v_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)), Word(a, 8), "%Sig")
                    Else
                        mod_matrix_nondirectional(v_num, v_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)) = ADD_STRINGVECTOR(mod_matrix_nondirectional(v_num, v_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)), Word(a, 2), "Estimate")
                        mod_matrix_nondirectional(v_num, v_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)) = ADD_STRINGVECTOR(mod_matrix_nondirectional(v_num, v_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)), Word(a, 3), "SE")
                        mod_matrix_nondirectional(v_num, v_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)) = ADD_STRINGVECTOR(mod_matrix_nondirectional(v_num, v_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)), Word(a, 4), "t")
                        mod_matrix_nondirectional(v_num, v_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)) = ADD_STRINGVECTOR(mod_matrix_nondirectional(v_num, v_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)), Word(a, 5), "p")
                    End If
                    Debug.Print v_num & ": " & mod_matrix_nondirectional(v_num, v_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num))
                Next
                start_line2 = a
            ElseIf IsRSquare = True And ((Word(start_line2, 1) = "Observed" Or Word(start_line2, 1) = "Latent") And Word(start_line2 + 1, 1) = "Variable" And Word(start_line2 + 1, 2) = "Estimate") Or (Word(start_line2, 1) = "Posterior" And Word(start_line2 + 1, 1) = "Variable") Then
                error_count = 0
                Debug.Print "Parsing R-Square"
                For a = start_line2 + 3 To Next_Blank_Line(start_line2 + 3) - 1
                    v_num = VarNum(Word(a, 1))
                    'MsgBox Word(a, 1) & " '" & VarList(v_num, 4) & "'"
                    mod_vector(v_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)) = ADD_STRINGVECTOR(mod_vector(v_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)), Word(start_line2, 1), "Type")
                    If Estimator = "BAYES" Then
                        mod_rsquare(v_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)) = ADD_STRINGVECTOR(mod_rsquare(v_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)), Word(a, 2), "Estimate")
                        mod_rsquare(v_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)) = ADD_STRINGVECTOR(mod_rsquare(v_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)), Word(a, 3), "SD")
                        mod_rsquare(v_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)) = ADD_STRINGVECTOR(mod_rsquare(v_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)), Word(a, 4), "p")
                        mod_rsquare(v_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)) = ADD_STRINGVECTOR(mod_rsquare(v_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)), Word(a, 5), "CI025")
                        mod_rsquare(v_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)) = ADD_STRINGVECTOR(mod_rsquare(v_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)), Word(a, 6), "CI975")
                    ElseIf IsImputation Then
                        mod_rsquare(v_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)) = ADD_STRINGVECTOR(mod_rsquare(v_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)), Word(a, 2), "Estimate")
                        mod_rsquare(v_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)) = ADD_STRINGVECTOR(mod_rsquare(v_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)), Word(a, 3), "SE")
                        mod_rsquare(v_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)) = ADD_STRINGVECTOR(mod_rsquare(v_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)), Word(a, 4), "t")
                        mod_rsquare(v_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)) = ADD_STRINGVECTOR(mod_rsquare(v_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)), Word(a, 5), "p")
                        mod_rsquare(v_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)) = ADD_STRINGVECTOR(mod_rsquare(v_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)), Word(a, 6), "RateMissing")
                    ElseIf IsMonteCarlo Then
                        mod_rsquare(v_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)) = ADD_STRINGVECTOR(mod_rsquare(v_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)), Word(a, 2), "Population")
                        mod_rsquare(v_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)) = ADD_STRINGVECTOR(mod_rsquare(v_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)), Word(a, 3), "Estimate")
                        mod_rsquare(v_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)) = ADD_STRINGVECTOR(mod_rsquare(v_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)), Word(a, 4), "SD")
                        mod_rsquare(v_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)) = ADD_STRINGVECTOR(mod_rsquare(v_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)), Word(a, 5), "SE")
                        mod_rsquare(v_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)) = ADD_STRINGVECTOR(mod_rsquare(v_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)), Word(a, 6), "MSE")
                        mod_rsquare(v_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)) = ADD_STRINGVECTOR(mod_rsquare(v_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)), Word(a, 7), "95Cov")
                        mod_rsquare(v_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)) = ADD_STRINGVECTOR(mod_rsquare(v_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)), Word(a, 8), "%Sig")
                    ElseIf VarList(v_num, 4) = 2 Or VarList(v_num, 4) = 3 Then
                        ' Write residual variance if the variable is ordered categorical or nominal
                        mod_rsquare(v_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)) = ADD_STRINGVECTOR(mod_rsquare(v_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)), Word(a, 2), "Estimate")
                        mod_matrix_nondirectional(v_num, v_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)) = ADD_STRINGVECTOR(mod_matrix_nondirectional(v_num, v_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)), Word(a, 3), "Estimate")
                        mod_matrix_nondirectional(v_num, v_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)) = ADD_STRINGVECTOR(mod_matrix_nondirectional(v_num, v_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)), "Residual", "Type")
                    Else
                        mod_rsquare(v_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)) = ADD_STRINGVECTOR(mod_rsquare(v_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)), Word(a, 2), "Estimate")
                        mod_rsquare(v_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)) = ADD_STRINGVECTOR(mod_rsquare(v_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)), Word(a, 3), "SE")
                        mod_rsquare(v_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)) = ADD_STRINGVECTOR(mod_rsquare(v_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)), Word(a, 4), "t")
                        mod_rsquare(v_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)) = ADD_STRINGVECTOR(mod_rsquare(v_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)), Word(a, 5), "p")
                    End If
                Next
                start_line2 = a
            ' If none of these apply, break the loop
            ElseIf line(start_line2) = "QUALITY OF NUMERICAL RESULTS" Then
                error_count = 0
                start_line2 = Next_Blank_Line(start_line2 + 2)
            Else
                Debug.Print "Looking for additional input"
                Debug.Print line(start_line2)
                'If start_line2 >= SMR_line And start_line2 < CISMR_line Then
                '    start_line2 = CISMR_line
                'ElseIf start_line2 < SMR_line Then
                '    start_line2 = SMR_line
                If error_count <= 3 Then
                    error_count = error_count + 1
                    start_line2 = start_line2 + 1
                Else
                    Debug.Print "No more model results found."
                    results_found = False
                End If
            End If
        Loop
    End If
End Sub

Sub FindComposites()
    ' This part of the program finds composites (observed variables correlated with latent variables)
    For a = 1 To UBound(VarList, 1)
        Debug.Print "Checking whether variable "; a; " is a composite."
        VarList(a, 25) = False ' Set to false by default
        If IsFactor(a) = False Then ' Check whether the variable is observed
            For b = 1 To UBound(mod_matrix_nondirectional, 2)
                If NonDirectionalPath(a, b) <> "NA" And IsFactor(b) = True Then ' If there is a correlation and b is a factor, then nominate it as a composite
                    Debug.Print "Composite found!"
                    VarList(a, 25) = True
                    Exit For
                End If
            Next
        End If
    Next
End Sub


Function NumLines()
    ' Exit if the number of lines has already been worked out
    If pNumLines > 0 Then
        NumLines = pNumLines
        Exit Function
    End If

    ' Exit if the line is blank
    If Len(out_text) < 1 Then
        NumLines = 0
        Exit Function
    End If
    
    NumLines = 1
    a = 1
    Do
        b = InStr(a, out_text, Chr(10))
        If b <> 0 Then NumLines = NumLines + 1
        a = b + 1
    Loop Until b = 0
    
    pNumLines = NumLines
End Function

Function line(num_line)
    start_pos = pLinePos(num_line)

    If num_line < pNumLines Then
        end_pos = pLinePos(num_line + 1) - 2
    Else
        end_pos = Len(out_text)
    End If
    
    line = Replace(Mid(out_text, start_pos, end_pos - start_pos + 1), "+" & Chr(10) & "  !", "")
End Function

Function Word(num_line, num_word)
    If num_line = 0 Then
        Word = 0
        Exit Function
    End If

    Dim temp_word: temp_word = remove_whitespace(line(num_line))

    Dim start_pos: start_pos = 1
    Dim end_pos: end_pos = Len(temp_word)

    If num_word > 1 Then start_pos = instr_inst(temp_word, " ", num_word - 1) + 1
    If num_word < Num_Words(temp_word) Then end_pos = instr_inst(temp_word, " ", num_word) - 1
    If num_word > Num_Words(temp_word) Then
        Word = "NA"
        Exit Function
    End If
    
    Word = Mid(temp_word, start_pos, end_pos - start_pos + 1)
End Function

Function Word_right(num_line, num_word)
    ' Returns word number "num_word" from line number "num_line", counting from the right
    Word_right = Word(num_line, Num_Words_on_Line(num_line) + 1 - num_word)
End Function

Function Words_exceptfor(num_line, Optional num_fromright = 0)
    ' Returns all of the words in a line (separated by a space), except for "num_fromright" words from the right.
    Words_exceptfor = ""
    cntr = Num_Words_on_Line(num_line) - num_fromright
    For a = 1 To cntr
        Words_exceptfor = Words_exceptfor & Word(num_line, a)
        If a < cntr Then Words_exceptfor = Words_exceptfor & " "
    Next
End Function

Function Num_Words(text)
    text = remove_whitespace(text)
    
    If InStr(1, text, " ") = 0 Then
        If Len(text) > 0 Then
            Num_Words = 1
        Else
            Num_Words = 0
        End If
        Exit Function
    Else
        Num_Words = 1
    End If

    start = 1
    Do While InStr(start, text, " ") > 0
        start = InStr(start, text, " ") + 1
        Num_Words = Num_Words + 1
    Loop
End Function

Function Num_Words_on_Line(num_line)

    text = remove_whitespace(line(num_line))

    If InStr(1, text, " ") = 0 Then
        If Len(text) > 0 Then
            Num_Words_on_Line = 1
        Else
            Num_Words_on_Line = 0
        End If
        Exit Function
    Else
        Num_Words_on_Line = 1
    End If
  
    start = 1
    Do While InStr(start, text, " ") > 0
        start = InStr(start, text, " ") + 1
        Num_Words_on_Line = Num_Words_on_Line + 1
    Loop
    
End Function

Function Num_NonBlankLines(num_line)
    ' Returns the number of lines until a blank line is encountered
    
    Dim a: a = 0
    Do While Num_Words_on_Line(num_line + a) > 0 And a <= pNumLines
        a = a + 1
    Loop
    
    Num_NonBlankLines = a
End Function

Function Next_NonBlank_Line(ByVal num_line)
    num_line = num_line + 1
    
    Do While Num_Words_on_Line(num_line) = 0 And num_line <= pNumLines
        num_line = num_line + 1
    Loop
    
    Next_NonBlank_Line = num_line
End Function

Function Next_Blank_Line(ByVal num_line)
    num_line = num_line + 1
    
    Do While Num_Words_on_Line(num_line) > 0 And num_line <= pNumLines
        num_line = num_line + 1
    Loop
    
    Next_Blank_Line = num_line
End Function

Function remove_whitespace(orig_string)
    ' Remove whitespace
    remove_whitespace = Trim(orig_string)
    
    ' Remove double spaces (make them single spaces)
    Do While InStr(1, remove_whitespace, "  ") > 0
        remove_whitespace = Replace(remove_whitespace, "  ", " ")
    Loop
End Function

Function instr_inst(string1, string2, instance)
    Dim temp_text: temp_text = Right(string1, Len(string1) + 1)
    Dim start_pos: start_pos = 0

    For a = 1 To instance
        start_pos = InStr(start_pos + 1, string1, string2)
    Next

    instr_inst = start_pos
End Function

Function FindLine(string1, Optional start = 1, Optional case_sens = True, Optional whole_word = False)
    ' Finds the line at which "string1" is located
            
    Dim temp_line
    
    If start < 1 Then start = 1
    If case_sens = False Then string1 = UCase(string1)
    
    Debug.Print "Searching for: " & string1
    For a = start To pNumLines
        'Debug.Print "Searching line " & a & " of " & pNumLines
        ' Used in the whole word function. By default, they are set to OK
        start_ok = True
        end_ok = True
 
        temp_line = line(a)
        If case_sens = False Then
            temp_line = UCase(temp_line)
            string1 = UCase(string1)
        End If
        word_start = InStr(1, temp_line, string1)
        
        ' If the whole word setting is true, then confirm the presence of spaces or line breaks before / after the word
        If word_start > 0 And whole_word = True Then
            pos_before = word_start - 1
            pos_after = word_start + Len(string1)
            'Debug.Print a & ": " & temp_line
            'Debug.Print "Before: " & pos_before, "After: " & pos_after
            
            ' First check that the position before the word occurs after the beginning of the line, then check to see if it's a space
            If pos_before >= 1 Then
                If Mid(temp_line, pos_before, 1) <> Chr(32) Then
                    start_ok = False
                End If
                'Debug.Print "Start OK: " & start_ok, Asc(Mid(temp_line, pos_before, 1))
            End If
            
            ' First check that the position before the word occurs before the end of the line, then check to see if it's a space
            If pos_after <= Len(temp_line) Then
                'Debug.Print "Checking end"
                If Mid(temp_line, pos_after, 1) <> Chr(32) Then
                    end_ok = False
                End If
            End If
        End If
        
        If (word_start > 0 Or temp_line = string1) And start_ok = True And end_ok = True Then
            FindLine = a
            Exit Function
        End If
    Next

    FindLine = 0
    ' Set FindLine to 0 if the relevant string is not found
End Function

Function DoubleBlank(Optional start = 1)
    ' Find the next instance of a double blank line

    Dim temp_line
    Dim found_first: found_first = False

    For a = start To NumLines()
        temp_line = line(a)
        If temp_line = "" Then
            If found_first = True Then
                DoubleBlank = a - 1
                Exit Function
            End If
            found_first = True
        Else
            found_first = False
        End If
    Next
End Function

Function r_to_t(r_val, n)
' Convert Pearson's r to t value
    If r_val = 1 Then
        r_to_t = 0
        Exit Function
    End If

    r_to_t = r_val * Sqr((n - 2) / (1 - (r_val ^ 2)))
    'MsgBox r_val & " " & n & " " & r_to_t
End Function

Public Property Get obs_Group(model_num)
    obs_Group = obs_matrix_lookup(1, model_num)
End Property

Public Property Get obs_Level(model_num)
    obs_Level = obs_matrix_lookup(2, model_num)
End Property

Public Property Get mod_Group(model_num)
' Returns the number of the group from the nominated model
    mod_Group = mod_matrix_lookup(1, model_num)
End Property

Public Property Get mod_Level(model_num)
    mod_Level = mod_matrix_lookup(2, model_num)
End Property

Public Property Get mod_Mixture(model_num)
    mod_Mixture = mod_matrix_lookup(3, model_num)
End Property

Public Property Get GroupName(group_num)
    GroupName = grouplist(group_num)
End Property

Public Property Get LevelName(level_num)
    LevelName = levellist(level_num)
End Property

Public Property Get MixtureName(mixture_num)
    MixtureName = mixturelist(mixture_num)
End Property

Public Property Get Std_n()
    Std_n = standardlist.Count
End Property

Public Property Get StdName(num)
    If num = 1 Then
        StdName = "Unstandardized"
    Else
        StdName = standardlist(num)
    End If
End Property

Public Property Get VarNum(ByVal var_name)
' Returns the variable number, based on the variable name
' Returns 0 if the variable name is not found

    If IsNumeric(var_name) = True Then
        VarNum = var_name
        Exit Property
    End If
    
    cat = InStr(1, var_name, "#")
    thr = InStr(1, var_name, "$")
    If cat > 0 Then var_name = Left(var_name, cat - 1)
    If thr > 0 Then var_name = Left(var_name, thr - 1)

    For a = 1 To UBound(VarList, 1)
        If VarList(a, 1) = var_name Then
            VarNum = a
            Exit Property
        End If
    Next
    
    VarNum = 0
End Property

Function CategoryNumber(var_name)
    ' Returns the category or threshold number in the variable name
    ' Returns '1' if there is only one category (i.e., not named)
    
    cat = InStr(1, var_name, "#")
    thr = InStr(1, var_name, "$")
           
    CategoryNumber = 1
    If cat > 0 Then CategoryNumber = CInt(Mid(var_name, cat + 1, Len(var_name) - cat))
    If thr > 0 Then CategoryNumber = CInt(Mid(var_name, thr + 1, Len(var_name) - thr))
End Function

Public Property Get ObsVarNum(Optional var_num = 0)
' Returns the variable number of the nominated observed variable
' If "var_num" is left blank, it returns the number of observed variables

    ObsVarNum = 0
    cntr = 0
        
    For a = 1 To UBound(VarList, 1)
        If VarList(a, 5) = False Then
            cntr = cntr + 1
            If var_num = cntr Then
                ObsVarNum = a
                Exit Property
            End If
        End If
    Next
    
    ObsVarNum = cntr
    
    'If var_num = 0 Then
    '    For a = 1 To UBound(VarList, 1)
    '        If VarList(a, 21) <> "" Then ObsVarNum = ObsVarNum + 1
    '    Next
    'Else
    '    obs_cnt = 0
    '    For a = 1 To UBound(VarList, 1)
    '        If VarList(a, 21) <> "" Then
    '            obs_cnt = obs_cnt + 1
    '            If var_num = obs_cnt Then
    '                ObsVarNum = a
    '                Exit Property
    '            End If
    '        End If
    '    Next
    'End If
End Property

Public Property Get Sample_Covariance(x, y, Optional matrix_num = 1)
    'lx = VarList(VarNum(x), 21)
    'ly = VarList(VarNum(y), 21)
    Sample_Covariance = obs_matrix(y, x, 1, obs_Group(matrix_num), obs_Level(matrix_num))
End Property

Public Property Get Sample_Correlation(x, y, Optional matrix_num = 1)
    'lx = VarList(VarNum(x), 21)
    'ly = VarList(VarNum(y), 21)
    Sample_Correlation = obs_matrix(y, x, 2, obs_Group(matrix_num), obs_Level(matrix_num))
End Property

Public Property Get Sample_P(x, y, Optional matrix_num = 1, Optional tails = 2)
    'lx = VarList(VarNum(x), 21)
    'ly = VarList(VarNum(y), 21)
    n = Sample_Cell_N(x, y, matrix_num)
    t = Abs(r_to_t(obs_matrix(y, x, 2, obs_Group(matrix_num), obs_Level(matrix_num)), n))
    Debug.Print "Finding p statistic:", x, y, matrix_num, n, obs_matrix(y, x, 2, obs_Group(matrix_num), obs_Level(matrix_num)), t
    Sample_P = WorksheetFunction.TDist(t, n - 2, tails)
End Property

Public Property Get Sample_Variance(x, Optional matrix_num = 1)
    'lx = VarList(VarNum(x), 21)
    If obs_matrix(x, x, 1, obs_Group(matrix_num), obs_Level(matrix_num)) = "" Then
        Sample_Variance = "NA"
    Else
        Sample_Variance = CDbl(obs_matrix(x, x, 1, obs_Group(matrix_num), obs_Level(matrix_num)))
    End If
    Debug.Print "Sample variance of " & VarName(x, 0) & " is " & Sample_Variance
End Property

Public Property Get Sample_SD(x, Optional matrix_num = 1)
    If Sample_Variance(x, matrix_num) = "NA" Then
        Sample_SD = "NA"
    Else
        Sample_SD = Sqr(Sample_Variance(x, matrix_num))
    End If
End Property

Public Property Get Sample_Mean(x, Optional matrix_num = 1)
    'lx = VarList(VarNum(x), 21)
    Sample_Mean = obs_vector(x, obs_Group(matrix_num), obs_Level(matrix_num))
End Property
Public Property Get Sample_N(Optional matrix_num = 1)
    Sample_N = obs_n(obs_Group(matrix_num), obs_Level(matrix_num), 1)
End Property

Public Property Get Sample_Cell_N(x, y, Optional matrix_num = 1)
' Returns the N of a particular cell, based on covariance coverage and sample size
    Debug.Print "Finding the sample size for an individual cell"

    grp_num = obs_Group(matrix_num)
    lev_num = obs_Level(matrix_num)
    multiplier = 1
    
    If IsCovCoverage = True And lev_num < 2 Then
        multiplier = cov_coverage(y, x, grp_num)
    End If
    
    Sample_Cell_N = WorksheetFunction.Round(multiplier * Sample_N(matrix_num), 0)
End Property

Public Property Get Sample_ICC1(x, Optional matrix_num = 1)
    Sample_ICC1 = GET_STRINGVECTOR(VarList(VarNum(x), 22), obs_Group(matrix_num))
    If Sample_ICC1 <> "NA" Then Sample_ICC1 = CDbl(Sample_ICC1)
End Property

Public Property Get Sample_ICC2(x, Optional matrix_num = 1, Optional ByVal formula = True)
    ICC1 = Sample_ICC1(x, matrix_num)
    clus_size = avg_csize(obs_Group(matrix_num), 2)
    
    If ICC1 = "NA" Or clus_size = "NA" Then
        Sample_ICC2 = "NA"
        Exit Function
    End If
    
    If formula = True Then
        Sample_ICC2 = "(" & clus_size & "*" & ICC1 & ") / (1 + (" & clus_size & " - 1) * " & ICC1 & ")"
    Else
        Sample_ICC2 = (clus_size * ICC1) / (1 + (clus_size - 1) * ICC1)
    End If
End Property

Public Property Get VarName(var_num, Optional format_option = 1)
' Returns the name of a variable
' format_option = 0: Short name only
' format_option = 1: Long name only
' format_option = 0: Long name and short name in parentheses

    If var_num < LBound(VarList, 1) Or var_num > UBound(VarList, 1) Then
        VarName = "NA"
    ElseIf format_option = 0 Or VarList(var_num, 2) = "" Then
        VarName = VarList(var_num, 1)
    ElseIf format_option = 1 Then
        VarName = VarList(var_num, 2)
    ElseIf format_option = 2 Then
        VarName = VarList(var_num, 2) & " (" & VarList(var_num, 1) & ")"
    End If
End Property

Public Property Get ScaleName(scale_num, Optional no_label = False)
    If scale_num < LBound(VarList, 1) Or var_num > UBound(VarList, 1) Then
        ScaleName = "NA"
    ElseIf no_label = True Or VarList(var_num, 4) = "" Then
        ScaleName = VarList(var_num, 3)
    Else
        ScaleName = VarList(var_num, 4)
    End If
End Property

Public Property Get ObsMatrix_n()
    ObsMatrix_n = obs_matrix_list.Count()
End Property

Public Property Get ObsMatrixName(num)
    ObsMatrixName = obs_matrix_list(num)
End Property

Public Property Get Model_n()
    Debug.Print "Counting number of models"
    Model_n = mod_matrix_list.Count
    Debug.Print Model_n
End Property

Public Property Get ModelName(num)
    ModelName = CONCAT_SEP(" ", GET_STRINGVECTOR(mod_matrix_list(num), 1), GET_STRINGVECTOR(mod_matrix_list(num), 2), GET_STRINGVECTOR(mod_matrix_list(num), 3))
End Property

Public Property Get ModelNum(Optional group_num = 1, Optional level_num = 1, Optional mixture_num = 1)
' Returns the model number, based on the group number, the level number, and the mixture number
    Debug.Print "Searching for model number, using group " & group_num & ", level " & level_num & ", mixture " & mixture_num & "."
        
    If group_num > nGroups Or level_num > nLevels Or mixture_number > nMixtures Then
        ModelNum = "NA"
        Exit Property
    End If
    
    temp_count = 0
    For a = 1 To nGroups
        For b = 1 To nLevels
            For c = 1 To nMixtures
                temp_count = temp_count + 1
                If group_num = a And level_num = b And mixture_num = c Then
                    ModelNum = temp_count
                End If
            Next
        Next
    Next
    
    Debug.Print "Model number is: " & ModelNum
End Property


Public Property Get DV(Optional dv_num = 0, Optional model_num = 1)
' Returns the variable number of the nominated dependent variable
' The macro scans the model directional matrix for path leading to each DV
' If "dv_num" is left as 0, it returns the number of DVs.

    Dim dv_count: dv_count = 0
    
    For a = 1 To UBound(VarList, 1)
        ' Increment dv_count if a DV has been found
        temp_val = GET_STRINGVECTOR(VarList(a, 23), model_num)
        If temp_val <> "NA" And temp_val <> "" Then
            dv_count = dv_count + 1
            If dv_num = CInt(temp_val) Then
                DV = a
                Debug.Print "Found DV"; DV; "(" & VarName(a) & ")"
                Exit Property
            End If
        End If
    Next
              
    Debug.Print "Number of DVs:" & dv_count
    If dv_num = 0 Then
        DV = dv_count
    Else
        DV = "NA"
    End If
End Property

Property Get IV(Optional iv_num = 0, Optional model_num = 1)
' Returns the variable number of the nominated independent variable
' The macro scans the model directional matrix for path leading to each DV
' If "iv_num" is left as 0, it returns the number of IVs.

    Dim IV_count: IV_count = 0
    
    For a = 1 To UBound(VarList, 1)
        ' Increment iv_count if a iv has been found
        temp_val = GET_STRINGVECTOR(VarList(a, 24), model_num)
        If temp_val <> "NA" And temp_val <> "" Then
            IV_count = IV_count + 1
            If iv_num = CInt(temp_val) Then
                IV = a
                Debug.Print "Found IV"; IV; "(" & VarName(a) & ")"
                Exit Property
            End If
        End If
    Next
    
    Debug.Print "Number of IVs:" & IV_count
    
    If iv_num = 0 Then
        IV = IV_count
    Else
        IV = "NA"
    End If
End Property

Property Get Antecedent(dv_num, ant_num, Optional model_num = 1)
' A wrapper for the DV_antecedent function
' If ant_num is 0, it returns the number of antecedents for a particular DV
' Returns the variable number of a particular antedent

    var_num = DV(dv_num, model_num)
    Antecedent = DV_antecedent(var_num, ant_num, 1, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num))
End Property


Function DV_antecedent(var_num, Optional ant_num = 0, Optional std_num = 1, Optional group = 1, Optional level = 1, Optional mixture = 1, Optional ON_only = True)
' Returns the variable number of the nominated antecedent
' If ant_num is 0, it returns the number of antecedents of a particular DV
' std_num refers to the standardisation number (1 by default)
    
    If ant_num = 0 Then
        Debug.Print "Finding number of antecedents of " & VarName(var_num, 2)
    Else
        Debug.Print "Locating antecedent " & ant_num & " of " & VarName(var_num, 2)
    End If
    
    ant_counter = 0

    For a = LBound(mod_matrix_directional, 1) To UBound(mod_matrix_directional, 1)
       ' Include antecedent only if type is ON, or ON_only is set to false
        If mod_matrix_directional(a, var_num, 1, group, level, mixture) <> "" And a <> var_num Then
            If GET_STRINGVECTOR(mod_matrix_directional(a, var_num, 1, group, level, mixture), "Type") = "ON" Or ON_only = False Then
                ant_counter = ant_counter + 1
                If ant_counter = ant_num Then
                    DV_antecedent = a
                    Debug.Print "Found antecedent: " & VarName(DV_antecedent)
                    Exit Function
                End If
            End If
        End If
    Next
    
    Debug.Print ant_counter & " antecedents found."
    DV_antecedent = ant_counter
End Function

Function DV_antecedents(dv1, dv2, Optional ant_num = 0, Optional std_num = 1, Optional group = 1, Optional level = 1, Optional mixture = 1, Optional ON_only = True)
' Returns the variable number of the nominated antecedent, which is common to two DVs
' If ant_num is 0, it returns the number of antecedents that are common to two DVs

    If ant_num = 0 Then
        Debug.Print "Finding number of antecedents of " & VarName(dv1, 2) & " and " & VarName(dv2, 2)
    Else
        Debug.Print "Finding mutual antecedent " & ant_num & " of " & VarName(dv1, 2) & " and " & VarName(dv2, 2)
    End If
    
    ant_counter = 0

    For a = LBound(mod_matrix_directional, 1) To UBound(mod_matrix_directional, 1)
        If mod_matrix_directional(a, dv1, 1, group, level, mixture) <> "" And mod_matrix_directional(a, dv2, 1, group, level, mixture) <> "" Then
            ' Include antecedent only if type is ON, or ON_only is set to false
            If (GET_STRINGVECTOR(mod_matrix_directional(a, dv1, 1, group, level, mixture), "Type") = "ON" And GET_STRINGVECTOR(mod_matrix_directional(a, dv2, 1, group, level, mixture), "Type") = "ON") Or ON_only = False Then
                ant_counter = ant_counter + 1
                If ant_counter = ant_num Then
                    DV_antecedents = a
                    Debug.Print "Found mutual antecedent: " & VarName(DV_antecedents)
                    Exit Function
                End If
            End If
        End If
    Next
    
    Debug.Print ant_counter & " antecedents found."
    DV_antecedents = ant_counter
End Function

Public Property Get IsFactor(var_num)
' Checks whether the nominated variable number is a latent variable; returns True if so
    var_num = CInt(var_num)
    If var_num = 0 Then ' If the number is an intercept (0)
        IsFactor = False
    Else
        IsFactor = VarList(var_num, 5)
    End If
End Property

Public Property Get Factor(Optional num = 0)
' Returns the variable number of the nominated latent variable
' If num is left blank or set to 0, it returns the number of latent variables

    Debug.Print "Looking for factors"

    factor_count = 0
    For a = 1 To UBound(VarList, 1)
        'Debug.Print a; " of "; UBound(VarList, 1), VarList(a, 4), VarList(a, 5)
        If VarList(a, 5) = True Then
            factor_count = factor_count + 1
            'Debug.Print "Found factor "; factor_count; "!"
            If num = factor_count Then
                'Debug.Print "Quitting..."
                Factor = a
                Exit Property
            End If
        End If
    Next

    Debug.Print "Made it to the second last line"
    Factor = factor_count
    
End Property

Public Property Get LatentVariable(Optional num = 0)
' A wrapper for "Factor"
    LatentVariable = Factor(num)
End Property

Public Property Get ObservedVariable(Optional num = 0)
    Debug.Print "Looking for observed variables"

    obs_count = 0
    For a = 1 To UBound(VarList, 1)
        If VarList(a, 5) = False Then
            obs_count = obs_count + 1
            If num = obs_count Then
                ObservedVariable = a
                Exit Property
            End If
        End If
    Next

    ObservedVariable = obs_count
End Property


Public Property Get FactorIndicatorSumVariance(fac_num, Optional std_num = 1, Optional model_num = 1, Optional ByVal formula = True)
    ' Returns the total variance of a set of indicators associated with a factor
    
    var_num = Factor(fac_num)
    Debug.Print "Calculating the total variance associated with the indicators of " & VarName(var_num)
    
    n_ind = FactorIndicator(fac_num, 0, std_num, model_num, False)
    
    If formula = True Then
        FactorIndicatorSumVariance = "SUM("
    Else
        FactorIndicatorSumVariance = 0
    End If
    
    'MsgBox VarName(var_num) & ", indicators: " & n_ind

    For y = 1 To n_ind
        ind_y = FactorIndicator(fac_num, y, std_num, model_num, False)
        'MsgBox y & ": " & VarName(ind_y)
        For x = 1 To y
            If x = y Then
                If formula = True Then
                    If y > 1 Then FactorIndicatorSumVariance = FactorIndicatorSumVariance & ","
                    FactorIndicatorSumVariance = FactorIndicatorSumVariance & Variance(ind_y, std_num, model_num, True, formula)
                Else
                    FactorIndicatorSumVariance = FactorIndicatorSumVariance + Variance(ind_y, std_num, model_num, True, formula)
                End If
            Else
                ind_x = FactorIndicator(fac_num, x, std_num, model_num, False)
                If formula = True Then
                    FactorIndicatorSumVariance = FactorIndicatorSumVariance & ",(2*" & Covariance(ind_y, ind_x, std_num, model_num, True, formula) & ")"
                Else
                    FactorIndicatorSumVariance = FactorIndicatorSumVariance + 2 * Covariance(ind_y, ind_x, std_num, model_num, True, formula)
                End If
            End If
        Next
    Next
    
    If formula = True Then FactorIndicatorSumVariance = FactorIndicatorSumVariance & ")"
    
End Property

Public Property Get FactorIndicatorSumVarianceExplained(fac_num, Optional std_num = 1, Optional model_num = 1, Optional all_antecedents = False, Optional ByVal formula = True)
    ' Returns the variance explained by a factor for a set of indicators
    
    var_num = Factor(fac_num)
    n_ind = FactorIndicator(fac_num, 0, std_num, model_num, False)
    
    If formula = True Then
        FactorIndicatorSumVarianceExplained = "SUM("
    Else
        FactorIndicatorSumVarianceExplained = 0
    End If
    
    If all_antecedents = False Then
        Debug.Print "Calculating the variance explained by factor " & VarName(var_num) & " with the indicators."
    
        fac_var = Variance(var_num, std_num, model_num, True, formula)
            
        For y = 1 To n_ind
            ind_y = FactorIndicator(fac_num, y, std_num, model_num, False)
            Path_Y = Path(var_num, ind_y, std_num, model_num)
            For x = 1 To y
                If x = y Then
                    If formula = True Then
                        If y > 1 Then FactorIndicatorSumVarianceExplained = FactorIndicatorSumVarianceExplained & ","
                        FactorIndicatorSumVarianceExplained = FactorIndicatorSumVarianceExplained & "(" & fac_var & "*" & Path_Y & "*" & Path_Y & ")"
                    Else
                        FactorIndicatorSumVarianceExplained = FactorIndicatorSumVarianceExplained + fac_var * Path_Y * Path_Y
                    End If
                Else
                    ind_x = FactorIndicator(fac_num, x, std_num, model_num, False)
                    Path_X = Path(var_num, ind_x, std_num, model_num)
                    If formula = True Then
                        FactorIndicatorSumVarianceExplained = FactorIndicatorSumVarianceExplained & ",(2*" & fac_var & "*" & Path_Y & "*" & Path_X & ")"
                    Else
                        FactorIndicatorSumVarianceExplained = FactorIndicatorSumVarianceExplained + 2 * fac_var * Path_Y * Path_X
                    End If
                End If
            Next
        Next
    Else
        Debug.Print "Calculating the total variance explained by all factors associated with a set of indicators"
        
        For y = 1 To n_ind
            ind_y = FactorIndicator(fac_num, y, std_num, model_num, False)
            If formula = True Then
                If y > 1 Then FactorIndicatorSumVarianceExplained = FactorIndicatorSumVarianceExplained & ","
                FactorIndicatorSumVarianceExplained = FactorIndicatorSumVarianceExplained & GetVarianceExplained(ind_y, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num), formula)
            Else
                FactorIndicatorSumVarianceExplained = FactorIndicatorSumVarianceExplained + GetVarianceExplained(ind_y, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num), formula)
            End If
            For x = 1 To y - 1
                ind_x = FactorIndicator(fac_num, x, std_num, model_num, False)
                If formula = True Then
                    FactorIndicatorSumVarianceExplained = FactorIndicatorSumVarianceExplained + ",(2*" & GetCovarianceExplained(ind_y, ind_x, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num), formula) & ")"
                Else
                    FactorIndicatorSumVarianceExplained = FactorIndicatorSumVarianceExplained + 2 * GetCovarianceExplained(ind_y, ind_x, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num), formula)
                End If
            Next
        Next
    End If
    
    If formula = True Then FactorIndicatorSumVarianceExplained = FactorIndicatorSumVarianceExplained & ")"
End Property

Public Property Get FactorName(fac_num, Optional no_label = False)
    If fac_num < 1 Or fac_num > Factor() Then
        FactorName = "NA"
        Exit Property
    End If
    
    var_num = Factor(fac_num)
    If no_label = True Or VarList(var_num, 2) = "" Then
        FactorName = VarList(var_num, 1)
    Else
        FactorName = VarList(var_num, 2)
    End If
End Property

Public Property Get FactorFind(var_num)
    ' Tests whether the nominated variable number is a factor; returns the factor number if this is the case
    ' Returns a 0 if it is not a factor

    FactorFind = 0
    If var_num < LBound(VarList, 1) Or var_num > UBound(VarList, 1) Or VarList(var_num, 5) = False Then Exit Property

    For a = 1 To var_num
        If VarList(a, 5) = True Then
            FactorFind = FactorFind + 1
        End If
    Next
End Property

Public Property Get FactorIndicator(fac_num, Optional ind_num = 0, Optional std_num = 1, Optional model_num = 1, Optional obs_only = False)
' Returns the variable number of the nominated indicator
' If ind_num is 0 or left blank, it returns the number of indicators associated with a factor
' std_num refers to the standardisation number (1 by default)
' model_num refers to the model number (1 by default)
' If obs_only is TRUE, then only look for observed indicators rather than latent ones
    
    var_num = Factor(fac_num)
    ind_counter = 0
    
    If ind_num = 0 Then
        Debug.Print "Counting the number of indicators for the factor " & VarName(var_num)
    Else
        Debug.Print "Locating indicator " & ind_num & " of the factor " & VarName(var_num)
    End If

    For a = LBound(mod_matrix_directional, 2) To UBound(mod_matrix_directional, 2)
        'Debug.Print a & " " & var_num
        If mod_matrix_directional(var_num, a, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)) <> "" Then
            'Debug.Print a; ":"; VarList(a, 5)
            If Not (obs_only = True And VarList(a, 5) = True) Then ' Do not count result if not an observed variable and obs_only is set to True
                ind_counter = ind_counter + 1
                If ind_counter = ind_num Then
                    FactorIndicator = a
                    Exit Property
                End If
            End If
        End If
    Next
    
    FactorIndicator = ind_counter
End Property

Public Property Get FactorIndicatorIsCategorical(fac_num, Optional std_num = 1, Optional model_num = 1, Optional obs_only = False)
    FactorIndicatorIsCategorical = False
    For a = 1 To FactorIndicator(fac_num, 0, std_num, model_num, obs_only)
        temp_ind = FactorIndicator(fac_num, a, std_num, model_num, obs_only)
        If VarList(temp_ind, 4) = 2 Or VarList(temp_ind, 4) = 3 Then
            FactorIndicatorIsCategorical = True
            Exit Property
        End If
    Next
End Property


Public Property Get FactorIndicatorArray(fac_num, Optional std_num = 1, Optional model_num = 1, Optional obs_only = False, Optional SortBySize = True)
' Returns an array of indicators sorted in descending order of factor loading size

    'Debug.Print "Step 1"
    Dim IndArray()
    var_num = Factor(fac_num)
    n_indicator = FactorIndicator(fac_num, 0, std_num, model_num, obs_only)
    
    'MsgBox fac_num & " " & var_num & " " & model_num & " " & n_indicator
    
    If n_indicator > 0 Then
        ReDim IndArray(1 To FactorIndicator(fac_num, 0, std_num, model_num, obs_only), 1 To 2)
        'Debug.Print "Step 2"
        For a = 1 To UBound(IndArray, 1)
            IndArray(a, 1) = FactorIndicator(fac_num, a, std_num, model_num, obs_only)
            IndArray(a, 2) = Path(var_num, IndArray(a, 1), std_num, model_num)
        Next
        
        'Debug.Print "Step 3"
        If SortBySize = True Then
            FactorIndicatorArray = ARRAY_SORT_COLUMN(IndArray, 2, False, True)
        Else
            FactorIndicatorArray = IndArray
        End If
    Else
        FactorIndicatorArray = "NA"
    End If
End Property

Public Property Get FactorMatch(var_num, Optional model_num1 = 1, Optional model_num2 = 1, Optional obs_only = False, Optional metric = False)
' A wrapper for the matchFactor function
' It searches for a match between two factors on different levels

    FactorMatch = matchFactor(var_num, mod_Group(model_num1), mod_Level(model_num1), mod_Mixture(model_num1), mod_Group(model_num2), mod_Level(model_num2), mod_Mixture(model_num2), obs_only, metric)
End Property


Public Property Get IndicatorName(fac_num, ind_num, Optional std_num = 1, Optional model_num = 1, Optional obs_only = False, Optional no_label = False)
    If fac_num < 1 Or fac_num > Factor() Then
        IndicatorName = "NA"
        Exit Property
    End If

    var_num = FactorIndicator(fac_num, ind_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num))
    
    If no_label = True Or VarList(var_num, 2) = "" Then
        IndicatorName = VarList(var_num, 1)
    Else
        IndicatorName = VarList(var_num, 2)
    End If
End Property


Public Property Get ModelVariable_n(Optional directional = True)
' Returns the number of variables in the model
    If directional = True Then ModelVariable_n = UBound(mod_matrix_directional, 1)
    If directional = False Then ModelVariable_n = UBound(mod_matrix_nondirectional, 1)
End Property
Public Property Get CI_Intercept(y, Optional std_num = 1, Optional model_num = 1, Optional CIval = "025", Optional category = 1)
    Debug.Print "Locating confidence interval for intercept"
    suffix = ""
    If category > 1 Then suffix = "_" & category
    CI_Intercept = GET_STRINGVECTOR(mod_vector(y, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)), "CI" & CIval & suffix)
End Property

Public Property Get CI(y, x, Optional std_num = 1, Optional model_num = 1, Optional CIval = "025", Optional category = 1)
    Debug.Print "Locating confidence interval", y, x
    suffix = ""
    If category > 1 Then suffix = "_" & category
    CI = GET_STRINGVECTOR(mod_matrix_directional(y, x, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)), "CI" & CIval & suffix)
End Property

Public Property Get NonDirectionalPath(y, x, Optional std_num = 1, Optional model_num = 1)
    Debug.Print "Locating covariance", y, x
    NonDirectionalPath = GET_STRINGVECTOR(mod_matrix_nondirectional(y, x, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)), "Estimate")
End Property

Public Property Get NonDirectionalPathP(y, x, Optional std_num = 1, Optional model_num = 1)
    Debug.Print "Locating covariance", y, x
    NonDirectionalPathP = GET_STRINGVECTOR(mod_matrix_nondirectional(y, x, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)), "p")
End Property

Public Property Get Path(from_var, to_var, Optional std_num = 1, Optional model_num = 1, Optional category = 1)
    ' A wrapper for GetPath
    Debug.Print "Locating path coefficient", from_var, to_var
    Path = GetPath(from_var, to_var, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num), category)
End Property

Function GetPath(from_var, to_var, Optional std_num = 1, Optional group = 1, Optional level = 1, Optional mixture = 1, Optional category = 1)
    search_for = "Estimate"
    If category > 1 Then search_for = search_for & "_" & category
    GetPath = GET_STRINGVECTOR(mod_matrix_directional(from_var, to_var, std_num, group, level, mixture), search_for)
End Function

Public Property Get PathSE(y, x, Optional std_num = 1, Optional model_num = 1, Optional category = 1)
    Debug.Print "Locating path coefficient", y, x
    search_for = "SE"
    If category > 1 Then search_for = search_for & "_" & category
    PathSE = GET_STRINGVECTOR(mod_matrix_directional(y, x, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)), search_for)
End Property

Public Property Get PathP(y, x, Optional std_num = 1, Optional model_num = 1, Optional category = 1)
    Debug.Print "Locating path coefficient", y, x
    search_for = "p"
    If category > 1 Then search_for = search_for & "_" & category
    PathP = GET_STRINGVECTOR(mod_matrix_directional(y, x, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)), search_for)
End Property

Public Property Get PathNCategories(y, x, Optional std_num = 1, Optional model_num = 1)
    PathNCategories = GET_STRINGVECTOR(mod_matrix_directional(y, x, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)), "NCategories")
    If PathNCategories = "NA" Then
        PathNCategories = 1
    Else
        PathNCategories = CInt(PathNCategories)
    End If
End Property

Public Property Get Covariance(y, x, Optional std_num = 1, Optional model_num = 1, Optional no_residual = False, Optional ByVal formula = True)
' Gets the covariance of the nominated variable
' A wrapper for GetCovariance
    
    Covariance = GetCovariance(y, x, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num), no_residual, formula)
End Property

Function FindCorrelate(var_num, Optional cor_num = 0, Optional std_num = 1, Optional group = 1, Optional level = 1, Optional mixture = 1, Optional no_residual = False)
' Returns the variable number of a numbered correlate (cor_num), based on WITH in the model
' If cor_num is set to 0, it returns the number of correlates of the nominated variable

    FindCorrelate = 0
    For a = 1 To UBound(mod_matrix_nondirectional, 1)
        var_type = GET_STRINGVECTOR(mod_matrix_nondirectional(var_num, a, std_num, group, level, mixture), "Type")
        If mod_matrix_nondirectional(var_num, a, std_num, group, level, mixture) <> "" And var_num <> a And (no_residual = False Or var_type <> "Residual") Then
            FindCorrelate = FindCorrelate + 1
            If cor_num = FindCorrelate Then ' Return the variable number the correlate
                FindCorrelate = a
                Exit Function
            End If
        End If
    Next

End Function

Function IsVarianceResidual(var_num, Optional std_num = 1, Optional model_num = 1)
' Returns TRUE if a nominated variable is endogenous / has a residual variance
    Debug.Print "Checking whether variance of " & VarName(var_num, 0) & " is residual."

    IsVarianceResidual = False
    If GET_STRINGVECTOR(mod_matrix_nondirectional(var_num, var_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)), "Type") = "Residual" Then IsVarianceResidual = True

    Debug.Print "Residual? " & IsVarianceResidual
End Function

Function IsCovarianceResidual(y, x, Optional std_num = 1, Optional model_num = 1)
' Returns TRUE if a nominated variable is endogenous / has a residual variance
    Debug.Print "Checking whether covariance of " & y & " " & x & " is residual."

    IsCovarianceResidual = False
    If GET_STRINGVECTOR(mod_matrix_nondirectional(y, x, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)), "Type") = "Residual" Then IsCovarianceResidual = True

    Debug.Print "Residual? " & IsCovarianceResidual
End Function

Public Property Get Variance(var_num, Optional std_num = 1, Optional model_num = 1, Optional no_residual = False, Optional ByVal formula = True)
' Gets the variance of the nominated variable
' A wrapper for GetVar

    Variance = GetVariance(var_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num), no_residual, formula)
End Property

Function GetVariance(var_num, Optional std_num = 1, Optional group = 1, Optional level = 1, Optional mixture = 1, Optional no_residual = False, Optional ByVal formula = True)
' Gets the variance of the nominated variable
' It retrieves the variance of the variable from the diagonal of the model non-directional table, or calculates the variance from the model if necessary

    Debug.Print "Locating variance of " & VarName(var_num, 2)
    temp_v = GET_STRINGVECTOR(mod_matrix_nondirectional(var_num, var_num, std_num, group, level, mixture), "Estimate")
    
    'On Error GoTo error_handling
    If (no_residual = True And IsVarianceResidual(var_num, std_num, ModelNum(group, level, mixture))) Or temp_v = "NA" Then
        ' Fallback 1: If the variable is categorical, assume a variance of 1
        If VarList(var_num, 4) = 2 Or VarList(var_num, 4) = 3 Then
            Debug.Print "Fallback 1: Variable is categorical or nominal. Setting total variance to 1."
            GetVariance = 1
        ' Fallback 2: Get estimate from TECH4 (variables are latent)
        ElseIf VarList(var_num, 5) = True And IsTECH4 = True Then
            Debug.Print "Fallback 2; getting estimate from TECH4"
            GetVariance = GET_STRINGVECTOR(mod_matrix_nondirectional(var_num, var_num, std_num, group, level, mixture), "TECH4_Estimate")
        ' Fallback 3: Get estimate from SAMPSTAT (observed variable array)
        ElseIf VarList(var_num, 5) = False And IsSAMPSTAT = True Then
            Debug.Print "Fallback 3; getting estimate from SAMPSTAT"
            If StdName(std_num) = "STDYX" Or StdName(std_num) = "STDY" Then
                GetVariance = 1
            Else
                GetVariance = Sample_Variance(var_num, ModelNum(group, level, mixture))
            End If
        ' Fallback 4: If the std_num > 1, set variance to 1
        ElseIf std_num > 1 And temp_v <> "NA" Then
            Debug.Print "Fallback 4; setting variance to 1"
            GetVariance = 1
        ' Fallback 5: Estimate the total variance from the set of antecedents (explained variance + residual variance + residual covariance)
        ElseIf temp_v <> "NA" Then
            Debug.Print "Fallback 5; estimating from the model"
            If formula = True Then
                GetVariance = "(" & GetVarianceExplained(var_num, std_num, group, level, mixture, formula) & " + " & temp_v & ")"
            Else
                GetVariance = GetVarianceExplained(var_num, std_num, group, level, mixture, formula) + temp_v
            End If
        Else ' Run out of ideas!
            GoTo error_handling
        End If
    Else
        Debug.Print "Get variance from the model directly"
        GetVariance = temp_v
    End If
    
    Exit Function
    
error_handling:
    MsgBox "The variance of " & VarName(var_num, 0) & " cannot be located in the output. Either: (1) request SAMPSTAT output for observed variables, (2) request TECH4 for latent variables, or (3) estimate the variance in the model."
    On Error GoTo 0
    End
End Function

Function GetCovariance(y, x, Optional std_num = 1, Optional group = 1, Optional level = 1, Optional mixture = 1, Optional no_residual = False, Optional ByVal formula = True)
' Gets the covariance of the nominated variable
' It retrieves the covariance of the variable from the model non-directional table, or calculates the covariance from the model if necessary

    Debug.Print "Locating covariance of " & VarName(y, 2) & " and " & VarName(x, 2)
    temp_cv = GET_STRINGVECTOR(mod_matrix_nondirectional(y, x, std_num, group, level, mixture), "Estimate")
    var_x = GET_STRINGVECTOR(mod_matrix_nondirectional(x, x, std_num, group, level, mixture), "Estimate")
    var_y = GET_STRINGVECTOR(mod_matrix_nondirectional(y, y, std_num, group, level, mixture), "Estimate")

    rv1 = IsVarianceResidual(y, std_num, ModelNum(group, level, mixture))
    rv2 = IsVarianceResidual(x, std_num, ModelNum(group, level, mixture))

    ' Resort to fallbacks if: total variance is requested, but Y or X are residual
    ' OR: The variance of X or Y are not named in the model (meaning they are exogenous)
    If (no_residual = True And (rv1 = True Or rv2 = True)) Or (var_x = "NA" Or var_y = "NA") Then
        ' Fallback 1: Get estimate from TECH4 (both variables are latent)
        If VarList(y, 5) = True And VarList(x, 5) = True And IsTECH4 = True Then
            GetCovariance = GET_STRINGVECTOR(mod_matrix_nondirectional(y, x, std_num, group, level, mixture), "TECH4_Estimate")
        ' Fallback 2: Get the estimate from the observed covariance matrix
        ElseIf VarList(y, 5) = False And VarList(x, 5) = False And IsSAMPSTAT = True Then
            If StdName(std_num) = "STDYX" Or StdName(std_num) = "STDY" Then
                GetCovariance = Sample_Correlation(x, y, ModelNum(group, level, mixture))
            Else
                GetCovariance = Sample_Covariance(x, y, ModelNum(group, level, mixture))
            End If
        ' Fallback 3: Calculate the estimate from the model
        ElseIf (rv1 = True Or rv2 = True) And temp_cv <> "NA" Then
            If formula = True Then
                GetCovariance = "(" & temp_cv & " + " & GetCovarianceExplained(y, x, std_num, group, level, mixture, formula) & ")"
            Else
                GetCovariance = temp_cv + GetCovarianceExplained(y, x, std_num, group, level, mixture, formula)
            End If
        Else
            GoTo error_handling
        End If
    Else
        Debug.Print "Obtaining covariance directly from matrix: " & y & ", " & x
        GetCovariance = temp_cv
    End If
    Exit Function
    
error_handling:
    MsgBox "The covariance of " & VarName(y) & " and " & VarName(x) & " cannot be located in the output. Either: (1) request SAMPSTAT output for observed variables, (2) request TECH4 for latent variables, or (3) estimate the covariance between the two variables in the model."
    End
            
End Function

Function GetCovarianceExplained(y, x, Optional std_num = 1, Optional group = 1, Optional level = 1, Optional mixture = 1, Optional ByVal formula = True)
    
    If formula = True Then
        GetCovarianceExplained = "SUM("
    Else
        GetCovarianceExplained = 0
    End If
    
    For a = 1 To DV_antecedents(y, x, 0, std_num, group, level, mixture, False)
        ant1 = DV_antecedents(y, x, a, std_num, group, level, mixture, False)
        var1 = GetVariance(ant1, std_num, group, level, mixture, True, formula)
        If var1 = "NA" Then GoTo error_handling1
        Path1 = Path(ant1, y, std_num, ModelNum(group, level, mixture))
        Path2 = Path(ant1, x, std_num, ModelNum(group, level, mixture))
        If Path1 = "NA" Then Path1 = 0
        If Path2 = "NA" Then Path2 = 0
        For b = 1 To a
            If a = b Then ' Same antecedent; a fork
                If formula = True Then
                    If a > 1 Then GetCovarianceExplained = GetCovarianceExplained & ","
                    GetCovarianceExplained = GetCovarianceExplained & "(" & var1 & "*" & Path1 & "*" & Path2 & ")"
                Else
                    GetCovarianceExplained = GetCovarianceExplained + var1 * Path1 * Path2
                End If
            Else ' Different antecedents; indirect path via covariance
                ant2 = DV_antecedents(y, x, b, std_num, group, level, mixture, False)
                var2 = GetVariance(ant2, std_num, group, level, mixture, True, formula)
                If var2 = "NA" Then GoTo error_handling1
                cv_ant12 = GetCovariance(ant1, ant2, std_num, group, level, mixture, True, formula)
                If cv_ant12 = "NA" Then GoTo error_handling3
                Path3 = Path(ant2, y, std_num, ModelNum(group, level, mixture))
                Path4 = Path(ant2, x, std_num, ModelNum(group, level, mixture))
                If Path3 = "NA" Then Path3 = 0
                If Path4 = "NA" Then Path4 = 0
                If formula = True Then
                    GetCovarianceExplained = GetCovarianceExplained & ",(" & cv_ant12 & "*" & Path1 & "*" & Path4 & ")"
                    GetCovarianceExplained = GetCovarianceExplained & ",(" & cv_ant12 & "*" & Path2 & "*" & Path3 & ")"
                Else
                    GetCovarianceExplained = GetCovarianceExplained + cv_ant12 * Path1 * Path4
                    GetCovarianceExplained = GetCovarianceExplained + cv_ant12 * Path2 * Path3
                End If
            End If
        Next
    Next
    
    If formula = True Then GetCovarianceExplained = GetCovarianceExplained & ")"
    Exit Function
    
error_handling1:
    MsgBox "The variance of " & VarName(ant1) & " cannot be located in the output. Either: (1) request SAMPSTAT output for observed variables, (2) request TECH4 for latent variables, or (3) estimate the covariance between the two variables in the model."
    End
error_handling2:
    MsgBox "The variance of " & VarName(ant2) & " cannot be located in the output. Either: (1) request SAMPSTAT output for observed variables, (2) request TECH4 for latent variables, or (3) estimate the covariance between the two variables in the model."
    End
error_handling3:
    MsgBox "The covariance of " & VarName(ant1) & " and " & VarName(ant2) & " cannot be located in the output. Either: (1) request SAMPSTAT output for observed variables, (2) request TECH4 for latent variables, or (3) estimate the covariance between the two variables in the model."
    End
End Function


Public Property Get VarianceExplained(dv_num, Optional std_num = 1, Optional model_num = 1, Optional ByVal formula = True)
' A wrapper for the VarExp function
' Returns the variance explained by a set of covariants
    VarianceExplained = GetVarianceExplained(dv_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num), formula)
End Property

Function GetVarianceExplained(var_num, Optional std_num = 1, Optional group = 1, Optional level = 1, Optional mixture = 1, Optional ByVal formula = True)
' Returns the variance explained by a set of covariants

    Debug.Print "Calculating variance explained for " & VarName(var_num, 2)
    
    If formula = True Then
        GetVarianceExplained = "SUM("
    Else
        GetVarianceExplained = 0
    End If
    
    For y = 1 To DV_antecedent(var_num, 0, 1, group, level, mixture, False)
        ant_y = DV_antecedent(var_num, y, std_num, group, level, mixture, False)
        var_y = GetVariance(ant_y, std_num, group, level, mixture, True, formula)
        Debug.Print "1"
        If var_y = "NA" Then GoTo error_handling1
        Debug.Print "2"
        Path1 = GetPath(ant_y, var_num, std_num, group, level, mixture)
        Debug.Print "3"
        If Path1 = "NA" Then Path1 = 0
        Debug.Print "4"
        For x = 1 To y
            Debug.Print "4 " & x & " of " & y
            If x = y Then ' Add variance
                If formula = True Then
                    Debug.Print "5 Using Formula"
                    If y > 1 Then GetVarianceExplained = GetVarianceExplained & ","
                    GetVarianceExplained = GetVarianceExplained & "(" & var_y & "*" & Path1 & "*" & Path1 & ")"
                Else
                    Debug.Print "5 No Formula"
                    Debug.Print var_y
                    Debug.Print Path1
                    GetVarianceExplained = GetVarianceExplained + (var_y * Path1 * Path1)
                End If
            Else ' Add covariance
                Debug.Print "6"
                ant_x = DV_antecedent(var_num, x, std_num, group, level, mixture, False)
                Debug.Print "7"
                var_x = GetVariance(ant_x, std_num, group, level, mixture, True, formula)
                If var_x = "NA" Then GoTo error_handling2
                Path2 = GetPath(ant_x, var_num, std_num, group, level, mixture)
                If Path2 = "NA" Then Path2 = 0
                cv_xy = GetCovariance(ant_y, ant_x, std_num, group, level, mixture, True, formula)
                If cv_xy = "NA" Then GoTo error_handling3
                If formula = True Then
                    GetVarianceExplained = GetVarianceExplained & ",(2*" & cv_xy & "*" & Path1 & "*" & Path2 & ")"
                Else
                    Debug.Print "Calculating 2*" & cv_xy & " * " & Path1 & "*" & Path2
                    GetVarianceExplained = GetVarianceExplained + (2 * cv_xy * Path1 * Path2)
                End If
            End If
        Next
    Next
    
    If formula = True Then GetVarianceExplained = GetVarianceExplained & ")"
    Exit Function

error_handling1:
    MsgBox "The variance of " & VarName(var_y) & " cannot be located in the output. Either: (1) request SAMPSTAT output for observed variables, (2) request TECH4 for latent variables, or (3) include the variance in the model."
    End
error_handling2:
    MsgBox "The variance of " & VarName(var_x) & " cannot be located in the output. Either: (1) request SAMPSTAT output for observed variables, (2) request TECH4 for latent variables, or (3) include the variance in the model."
    End
error_handling3:
    MsgBox "The covariance of " & VarName(var_x) & " and " & VarName(var_y) & " cannot be located in the output. Either: (1) request SAMPSTAT output for observed variables, (2) request TECH4 for latent variables, or (3) estimate the covariance between the two variables in the model."
    End

End Function

Public Property Get SD(var_num, Optional std_num = 1, Optional model_num = 1, Optional no_residual = False, Optional ByVal formula = True)
    If formula = True Then
        SD = "SQRT(" & Variance(var_num, std_num, model_num, no_residual, True) & ")"
    Else
        SD = Sqr(Variance(var_num, std_num, model_num, no_residual, False))
    End If
End Property

Public Property Get Mean(var_num, Optional std_num = 1, Optional model_num = 1, Optional category = 1)
    search_for = "Estimate"
    If category > 1 Then search_for = search_for & "_" & category
    Mean = GET_STRINGVECTOR(mod_vector(var_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)), "Estimate")
End Property

Public Property Get Intercept(var_num, Optional std_num = 1, Optional model_num = 1, Optional category = 1)
    search_for = "Estimate"
    If category > 1 Then search_for = search_for & "_" & category
    Intercept = GET_STRINGVECTOR(mod_vector(var_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)), search_for)
End Property

Public Property Get InterceptSE(var_num, Optional std_num = 1, Optional model_num = 1, Optional category = 1)
    search_for = "SE"
    If category > 1 Then search_for = search_for & "_" & category
    InterceptSE = GET_STRINGVECTOR(mod_vector(var_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)), search_for)
End Property

Public Property Get InterceptP(var_num, Optional std_num = 1, Optional model_num = 1, Optional category = 1)
    search_for = "p"
    If category > 1 Then search_for = search_for & "_" & category
    InterceptP = GET_STRINGVECTOR(mod_vector(var_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)), search_for)
End Property

Public Property Get InterceptNCategories(var_num, Optional std_num = 1, Optional model_num = 1)
    InterceptNCategories = GET_STRINGVECTOR(mod_vector(var_num, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)), "NCategories")
    If InterceptNCategories = "NA" Then
        InterceptNCategories = 1
    Else
        InterceptNCategories = CInt(InterceptNCategories)
    End If
End Property

Public Property Get Omega(fac_num, Optional std_num = 1, Optional model_num = 1, Optional obs_only = False, Optional ByVal formula = True)
' Returns Coefficient Omega / Composite Reliability for the nominated factor.
' In bifactor models, it assumes that the factor is orthogonal to other factors where there are cross-loadings.
    
    ' Check whether any of the indicators are categorical; quit if that's the case
    If FactorIndicatorIsCategorical(fac_num, std_num, model_num, obs_only) = True Then
        Omega = "NA"
        Exit Property
    End If

    If formula = True Then
        Omega = FactorIndicatorSumVarianceExplained(fac_num, std_num, model_num, False, True) & " / " & FactorIndicatorSumVariance(fac_num, std_num, model_num, True)
    Else
        Omega = FactorIndicatorSumVarianceExplained(fac_num, std_num, model_num, False, False) / FactorIndicatorSumVariance(fac_num, std_num, model_num, False)
    End If
End Property

Public Property Get OmegaTotal(fac_num, Optional std_num = 1, Optional model_num = 1, Optional obs_only = False, Optional ByVal formula = True)
' Returns the total variance explained across a set of indicators (which may be 1 or more factors)

    ' Check whether any of the indicators are categorical; quit if that's the case
    If FactorIndicatorIsCategorical(fac_num, std_num, model_num, obs_only) = True Then
        OmegaTotal = "NA"
        Exit Property
    End If

    If formula = True Then
        OmegaTotal = FactorIndicatorSumVarianceExplained(fac_num, std_num, model_num, True, True) & " / " & FactorIndicatorSumVariance(fac_num, std_num, model_num, True)
    Else
        OmegaTotal = FactorIndicatorSumVarianceExplained(fac_num, std_num, model_num, True, False) / FactorIndicatorSumVariance(fac_num, std_num, model_num, False)
    End If
End Property


Public Property Get AVE(fac_num, Optional std_num = 1, Optional model_num = 1, Optional obs_only = False, Optional ByVal formula = True)
' Returns AVE for the nominated factor

    If formula = True Then
        true_var = "SUM("
        error_var = "SUM("
        error_covar = "("
    Else
        true_var = 0
        error_var = 0
        error_covar = 0
    End If

    var_num = Factor(fac_num)

    ' Change to STDYX if the factor is endogenous
    fac_var = Variance(var_num, std_num, model_num, True)
    If fac_var = "Residual" Then
        MsgBox "Factor " & VarName(var_num, 0) & " is endogenous. Request STDYX output."
        End
    End If

    Debug.Print "Getting factor and indicator number"

    num_ind = FactorIndicator(fac_num, 0, std_num, model_num, obs_only)
    
    If num_ind < 2 Then
        AVE = "NA"
        Exit Property
    End If
          
    For a = 1 To num_ind
        Debug.Print a
        ind1 = FactorIndicator(fac_num, a, std_num, model_num, obs_only)
        Debug.Print var_num, ind1, std_num, model_num, obs_only
        If formula = True Then
            If a > 1 Then
                true_var = true_var & ","
                error_var = error_var & ","
            End If
            true_var = true_var & Path(var_num, ind1, std_num, model_num) & "^2"
            error_var = error_var & Variance(ind1, std_num, model_num)
        Else
            true_var = true_var + Path(var_num, ind1, std_num, model_num) ^ 2
            error_var = error_var + Variance(ind1, std_num, model_num)
        End If
    Next
    
    If formula = True Then
        true_var = "(" & fac_var & "*" & true_var & "))"
        error_var = error_var & ")"
        AVE = "(" & true_var & ") / (" & true_var & "+" & error_var & ")"
    Else
        true_var = true_var * fac_var
        AVE = true_var / (true_var + error_var)
    End If
    
    Debug.Print "Made it to end"
End Property

Public Property Get Alpha(fac_num, Optional std_num = 1, Optional model_num = 1, Optional obs_only = False, Optional ByVal formula = True)
    ' Check that the number of cells is equal to the number of columns
    
    Debug.Print "Calculating Cronbach's alpha"
    
    Dim num_v: num_v = FactorIndicator(fac_num, 0, std_num, model_num, True)
    'Debug.Print "Number of observed indicators is: "; num_v
    
    ' Do not try to compute Cronbach's alpha if the number of indicators is too small
    If num_v < 2 Then
        Alpha = "NA"
        Exit Property
    End If
    
    matrix_num = model_num
    
    ' Get variances
    Dim vars()
    ReDim vars(1 To num_v)
    For a = 1 To UBound(vars)
        vars(a) = Sample_Variance(FactorIndicator(fac_num, a, std_num, model_num, True), model_num)
    Next
    
    ' Get covariances / correlations
    Dim covars()
    c = 1
    
    ReDim covars(1 To (num_v ^ 2 - num_v) / 2)
    num_cv = UBound(covars)
    For a = 1 To num_v
        For b = 1 To a
            ind1 = FactorIndicator(fac_num, a, std_num, model_num, True)
            ind2 = FactorIndicator(fac_num, b, std_num, model_num, True)
            If a <> b Then
                covars(c) = Sample_Covariance(ind1, ind2, matrix_num)
                c = c + 1
            End If
        Next
    Next
       
    ' Get variance mean
    Dim vars_mean
    If formula = True Then
        vars_mean = "AVERAGE("
        For a = 1 To num_v
            If a > 1 Then vars_mean = vars_mean & ","
            vars_mean = vars_mean & vars(a)
        Next
        vars_mean = vars_mean & ")"
    Else
        vars_mean = 0
        For a = 1 To num_v
            vars_mean = vars_mean + (vars(a) / num_v)
        Next
    End If
    
    ' Get covariance mean
    Dim covars_mean
    If formula = True Then
        covars_mean = "AVERAGE("
        For a = 1 To num_cv
            If a > 1 Then covars_mean = covars_mean & ","
            covars_mean = covars_mean & covars(a)
        Next
        covars_mean = covars_mean & ")"
    Else
        covars_mean = 0
        For a = 1 To num_cv
            covars_mean = covars_mean + (covars(a) / num_cv)
        Next
    End If
    
    ' Return alpha
    If formula = True Then
        Alpha = "(" & num_v & " * " & covars_mean & ") / (" & vars_mean & " + (" & num_v & " - 1) * " & covars_mean & ")"
    Else
        Alpha = (num_v * covars_mean) / (vars_mean + (num_v - 1) * covars_mean)
    End If
End Property
Public Property Get Composite(Optional num = 0)
' Returns the variable number of the nominated composite (which may be a latent variable OR an observed variable correlated with a latent variable)

' Returns the variable number of the nominated factor
' If num is left blank or set to 0, it returns the number of factors

    Debug.Print "Looking for composites (factors and observed variables correlated with factors)"

    factor_count = 0
    For a = 1 To UBound(VarList, 1)
        If VarList(a, 5) = True Or VarList(a, 25) = True Then
            factor_count = factor_count + 1
            If num = factor_count Then
                Composite = a
                Exit Property
            End If
        End If
    Next

    Composite = factor_count
End Property
Public Property Get Composite_Mean(ByVal y, ByVal model_num, Optional ByVal formula = True)
' Calculates the composite (weighted mean) of the indicators associated with a factor
' Returns the sample mean of the variable if it is not a latent factor

    ' Get list of ys
    Dim y_list: y_list = ""
    Dim y_count: y_count = 1
    Dim weighted_sum
    Dim weight
    
    If formula = True Then
        weighted_sum = "SUM("
        weight = "SUM("
    Else
        weighted_sum = 0
        weight = 0
    End If

    ' Check to see whether it is a latent variable
    Debug.Print "Part 1"
    If VarList(y, 5) = False Then
        Composite_Mean = Sample_Mean(y, model_num)
        Exit Property
    Else
        For a = 1 To UBound(mod_matrix_directional, 2)
            If GET_STRINGVECTOR(mod_matrix_directional(y, a, 1, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)), "Type") = "BY" Then y_list = ADD_STRINGVECTOR(y_list, a)
            'If mod_matrix_directional(y, a, 1, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)) <> "" Then y_list = ADD_STRINGVECTOR(y_list, a)
        Next
    End If
    y_count = COUNT_STRINGVECTOR(y_list)
    
    ReDim YY(1 To y_count)
    ReDim y_weight(1 To y_count)
    
    For a = 1 To y_count
        var_num = GET_STRINGVECTOR(y_list, a)
        s_mean = Sample_Mean(var_num, model_num)
        s_c_n = Sample_Cell_N(var_num, var_num, model_num)
        If formula = True Then
            If a > 1 Then
                weighted_sum = weighted_sum & ","
                weight = weight & ","
            End If
            weighted_sum = weighted_sum & s_mean & "*" & s_c_n
            weight = weight & s_c_n
        Else
            weighted_sum = weighted_sum + s_mean * s_c_n
            weight = weight + s_c_n
        End If
    Next
    
    If formula = True Then
        weighted_sum = weighted_sum & ")"
        weight = weight & ")"
        Composite_Mean = weighted_sum & " / " & weight
    Else
        Composite_Mean = weighted_sum / weight
    End If

End Property
Public Property Get Composite_SD(ByVal y, ByVal model_num, Optional ByVal formula = True)
    Debug.Print "Calculating composite SD"
    Debug.Print Composite_Variance(y, model_num, formula)
    
    If formula = True Then
        Composite_SD = "SQRT(" & Composite_Variance(y, model_num, formula) & ")"
    Else
        Composite_SD = Sqr(Composite_Variance(y, model_num, formula))
    End If
End Property

Public Property Get Composite_Variance(ByVal y, ByVal model_num, Optional ByVal formula = True)
' Calculate the composite variance of the indicators associated with a factor
' Returns the total variance of the variable if it is not a latent factor
    Debug.Print ""
    Debug.Print "Calculating composite variance"

    ' Get list of ys and xs
    Dim y_list: y_list = ""
    Dim y_count: y_count = 1
    Dim YY()
    Dim y_weight()

    ' Check to see whether it is a latent variable

    Debug.Print "Part 1"
    If VarList(y, 5) = False Then
        Composite_Variance = obs_matrix(y, y, 1, obs_Group(model_num), obs_Level(model_num))
        Exit Property
    Else
        For a = 1 To UBound(mod_matrix_directional, 2)
            If mod_matrix_directional(y, a, 1, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)) <> "" Then y_list = ADD_STRINGVECTOR(y_list, a)
        Next
    End If
    y_count = COUNT_STRINGVECTOR(y_list)

    ReDim YY(1 To y_count, 1 To y_count)
    ReDim y_weight(1 To y_count)
    
    Debug.Print "Part 2"
    For a = 1 To y_count
        y_num1 = CInt(GET_STRINGVECTOR(y_list, a))
        For b = 1 To a
            Debug.Print a, b
            y_num2 = CInt(GET_STRINGVECTOR(y_list, b))
            Debug.Print y_num1, y_num2, 1, obs_Group(model_num), obs_Level(model_num)
            YY(a, b) = obs_matrix(y_num1, y_num2, 1, obs_Group(model_num), obs_Level(model_num))
            YY(b, a) = YY(a, b)
        Next
        y_weight(a) = Sample_Cell_N(y_num1, y_num1, model_num)
    Next
    
    Debug.Print "Part 3"
    
    Debug.Print "3.1 Transposing y weight"
    y_weight_transpose = ARRAY_TRANSPOSE(y_weight)

    Debug.Print "3.2 Multiplying y weight and YY"
    VarY = ARRAY_MULTIPLY(y_weight_transpose, YY, formula)
    
    Debug.Print "3.3 Multiplying Var Y and y weight"
    VarY = ARRAY_MULTIPLY(VarY, y_weight, formula)
    
    Debug.Print "3.4 Multiplying y weight and a matrix of all ones"
    weight = ARRAY_MULTIPLY(y_weight_transpose, ARRAY_ALLONES(y_count, y_count), formula)
    
    Debug.Print "3.5 Multiplying weight and y weight"
    weight = ARRAY_MULTIPLY(weight, y_weight, formula)
    
    Debug.Print "3.6 Putting together formula..."
    'MsgBox DISPLAY_ARRAY(VarY)
    'MsgBox DISPLAY_ARRAY(weight)
    
    If formula = True Then
        Composite_Variance = "(" & VarY & "/" & weight & ")"
    Else
        Composite_Variance = VarY / weight
    End If
End Property


Public Property Get Composite_Correlation(ByVal y, ByVal x, Optional ByVal std_num = 2, Optional ByVal model_num = 1, Optional ByVal formula = True)
    Debug.Print "Finding composite correlation"
    
    ' Get list of ys and xs
    Dim y_list: y_list = ""
    Dim x_list: x_list = ""
    Dim y_count: y_count = 1
    Dim x_count: x_count = 1
    Dim XX()
    Dim YY()
    Dim XY()
    Dim x_weight()
    Dim y_weight()
    
    ' If both variables are observed, simply return the correlation
    If VarList(x, 5) = False And VarList(y, 5) = False Then
        Composite_Correlation = obs_matrix(y, x, 1, obs_Group(model_num), obs_Level(model_num))
        Exit Property
    End If
    
    ' Check whether y is a factor
    Debug.Print "Is Y a factor"
    If VarList(y, 5) = True Then
        For a = 1 To UBound(mod_matrix_directional, 2)
            If GET_STRINGVECTOR(mod_matrix_directional(y, a, 1, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)), "Type") = "BY" Then y_list = ADD_STRINGVECTOR(y_list, a)
            'If mod_matrix_directional(y, a, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)) <> "" Then y_list = ADD_STRINGVECTOR(y_list, a)
        Next
    Else
        y_list = ADD_STRINGVECTOR(y_list, y)
    End If
    y_count = COUNT_STRINGVECTOR(y_list)
        
    ' Check whether x is a factor
    Debug.Print "Is X a factor"
    If VarList(x, 5) = True Then
        For a = 1 To UBound(mod_matrix_directional, 2)
            If GET_STRINGVECTOR(mod_matrix_directional(x, a, 1, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)), "Type") = "BY" Then x_list = ADD_STRINGVECTOR(x_list, a)
            'If mod_matrix_directional(x, a, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)) <> "" Then x_list = ADD_STRINGVECTOR(x_list, a)
        Next
    Else
        x_list = ADD_STRINGVECTOR(x_list, x)
    End If
    x_count = COUNT_STRINGVECTOR(x_list)
    
    Debug.Print "Populating arrays"
    ReDim YY(1 To y_count, 1 To y_count)
    ReDim XX(1 To x_count, 1 To x_count)
    ReDim XY(1 To x_count, 1 To y_count)
    ReDim y_weight(1 To y_count)
    ReDim x_weight(1 To x_count)
   
    Debug.Print "Populating YY array"
    ' Populate YY array (and y_weight vector)
    For a = 1 To y_count
        Debug.Print a
        y_num1 = CInt(GET_STRINGVECTOR(y_list, a))
        Debug.Print y_num1
        For b = 1 To a
            Debug.Print a, b
            y_num2 = CInt(GET_STRINGVECTOR(y_list, b))
            Debug.Print y_num1, y_num2, std_num, obs_Group(model_num), obs_Level(model_num)
            YY(a, b) = obs_matrix(y_num1, y_num2, 2, obs_Group(model_num), obs_Level(model_num))
            YY(b, a) = YY(a, b)
        Next
        y_weight(a) = Sample_Cell_N(y_num1, y_num1, model_num)
    Next
   
    Debug.Print "Populating XX array"
    ' Populate XX array
    For a = 1 To x_count
        x_num1 = CInt(GET_STRINGVECTOR(x_list, a))
        For b = 1 To a
            x_num2 = CInt(GET_STRINGVECTOR(x_list, b))
            XX(a, b) = obs_matrix(x_num1, x_num2, 2, obs_Group(model_num), obs_Level(model_num))
            XX(b, a) = XX(a, b)
        Next
        x_weight(a) = Sample_Cell_N(x_num1, x_num1, model_num)
    Next
   
    Debug.Print "Populating XY array"
    ' Populate XY array
    For a = 1 To x_count
        x_num = CInt(GET_STRINGVECTOR(x_list, a))
        For b = 1 To y_count
            y_num = CInt(GET_STRINGVECTOR(y_list, b))
            XY(a, b) = obs_matrix(x_num, y_num, 2, obs_Group(model_num), obs_Level(model_num))
        Next
    Next
    
    ' Calculate composite correlation
    Debug.Print "Calculating composites"
    x_weight_transpose = ARRAY_TRANSPOSE(x_weight)
    y_weight_transpose = ARRAY_TRANSPOSE(y_weight)
    
    'MsgBox DISPLAY_ARRAY(x_weight_transpose)
    'MsgBox DISPLAY_ARRAY(XY)
    
    Debug.Print "Numerator"
    temp_numer = ARRAY_MULTIPLY(x_weight_transpose, XY, formula)
    
    Debug.Print "Part 2"
    
    temp_numer = ARRAY_MULTIPLY(temp_numer, y_weight, formula)
    
    Debug.Print "SDx"
    SDx = ARRAY_MULTIPLY(x_weight_transpose, XX, formula)
    SDx = ARRAY_MULTIPLY(SDx, x_weight, formula)
    
    Debug.Print "SDy"
    SDY = ARRAY_MULTIPLY(y_weight_transpose, YY, formula)
    SDY = ARRAY_MULTIPLY(SDY, y_weight, formula)
    
    If formula = True Then
        SDx = "SQRT(" & SDx & ")"
        SDY = "SQRT(" & SDY & ")"
        Composite_Correlation = temp_numer & " / (" & SDx & "*" & SDY & ")"
    Else
        SDx = Sqr(SDx)
        SDY = Sqr(SDY)
        Composite_Correlation = temp_numer / (SDx * SDY)
    End If
End Property


Public Property Get Composite_N(ByVal y, ByVal x, Optional ByVal std_num = 1, Optional ByVal model_num = 1, Optional ByVal formula = False)
' Returns the estimated N of the composite (a weighted average of the rXY sample N's)
    Debug.Print "Finding Composite N"

    ' Get list of ys and xs
    Dim y_list: y_list = ""
    Dim x_list: x_list = ""
    Dim y_count: y_count = 1
    Dim x_count: x_count = 1
    Dim XX()
    Dim YY()
    Dim XY()
    Dim x_weight()
    Dim y_weight()
    
    ' If both variables are observed, simply return the Sample N
    Debug.Print "Are both variables observed?"
    If VarList(x, 5) = False And VarList(y, 5) = False Then
        Composite_N = Sample_Cell_N(x, y, model_num)
        Exit Property
    End If

    ' Check whether y is a factor
    Debug.Print "Is Y a factor"
    If VarList(y, 5) = True Then
        For a = 1 To UBound(mod_matrix_directional, 2)
            If mod_matrix_directional(y, a, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)) <> "" Then y_list = ADD_STRINGVECTOR(y_list, a)
        Next
    Else
        y_list = ADD_STRINGVECTOR(y_list, y)
    End If
    y_count = COUNT_STRINGVECTOR(y_list)
        
    ' Check whether x is a factor
    Debug.Print "Is X a factor"
    If VarList(x, 5) = True Then
        For a = 1 To UBound(mod_matrix_directional, 2)
            If mod_matrix_directional(x, a, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)) <> "" Then x_list = ADD_STRINGVECTOR(x_list, a)
        Next
    Else
        x_list = ADD_STRINGVECTOR(y_list, x)
    End If
    x_count = COUNT_STRINGVECTOR(x_list)

    If formula = True Then
        Composite_N = "AVERAGE("
    Else
        Composite_N = 0
    End If
    
    For a = 1 To y_count
        y_num = GET_STRINGVECTOR(y_list, a)
        For b = 1 To x_count
            x_num = GET_STRINGVECTOR(x_list, b)
            If formula = True Then
                If a > 1 Or b > 1 Then Composite_N = Composite_N & ","
                Composite_N = Composite_N & Sample_Cell_N(x_num, y_num, model_num)
                If a = y_count And b = x_count Then Composite_N = Composite_N & ")"
            Else
                Composite_N = Composite_N + (Sample_Cell_N(x_num, y_num, model_num) / (y_count * x_count))
            End If
        Next
    Next
End Property

Public Property Get Composite_P(y, x, Optional model_num = 1, Optional tails = 2)
    Debug.Print "Finding Composite p"
    n = Composite_N(y, x, 2, model_num)
    r = Composite_Correlation(y, x, 2, model_num, False)
    t = Abs(r_to_t(r, n))
    Composite_P = WorksheetFunction.TDist(t, n - 2, tails)
End Property


Function matchFactor(ByVal var_num1, Optional ByVal group1 = 1, Optional ByVal level1 = 1, Optional ByVal mixture1 = 1, Optional ByVal group2 = 1, Optional ByVal level2 = 1, Optional ByVal mixture2 = 1, Optional ByVal obs_only = False, Optional ByVal metric = False)
' Attempts to match a factor on a different matrix by looking for the same indicators. It will return the variable number of the matched factor.
' obs_only: If set to TRUE, then only observed indicators will be used for the matching.
' metric: If set to TRUE, then the macro will search for metric equivalent factor loadings (i.e., corresponding loadings that have the same value). If FALSE, it will look only for the same pattern of loadings across the same set of indicators.
' Returns 0 if a matched factor is not found.

    Dim temp_val1, temp_val2
    matchFactor = 0

    For a = 1 To Factor()
        var_num2 = Factor(a)
        temp_val1 = ""
        temp_val2 = ""
        counter = 0
        For b = 1 To UBound(mod_matrix_directional, 2)
            Debug.Print "Factor"; a; "of"; Factor(); "Variable"; var_num1; "of"; UBound(mod_matrix_directional, 1); ", Indicator "; b; " of "; UBound(mod_matrix_directional, 2); ", "; group1; ":"; group2, level1; ":"; level2, mixture1; ":"; mixture2
            
            temp_val1 = mod_matrix_directional(var_num1, b, 1, group1, level1, mixture1)
            temp_val2 = mod_matrix_directional(var_num2, b, 1, group2, level2, mixture2)
                                                                          
            ' If the indicator is latent (and a named indicator), then find its variable number on the second matrix
            If VarList(b, 5) = True And GET_STRINGVECTOR(temp_val1, "Type") = "BY" Then
                If obs_only = True Then
                    Debug.Print "Skipping to the next iteration: not an observed indicator"
                    GoTo nextiteration
                End If
                b2 = matchFactor(b, group1, level1, mixture1, group2, level2, mixture2, obs_only, metric)
                temp_val2 = mod_matrix_directional(var_num2, b2, 1, group2, level2, mixture2)
                Debug.Print "Reassigned"; b; "to"; b2
            End If
                                              
            ' Skip this iteration if neither is a named factor loading
            type1 = GET_STRINGVECTOR(temp_val1, "Type")
            type2 = GET_STRINGVECTOR(temp_val2, "Type")
            
            If (type1 = "BY" And type2 <> "BY") Or (type1 <> "BY" And type2 = "BY") Then
                Debug.Print "Breaking loop; one of them is a factor loading and the other isn't"
                Exit For
            ElseIf type1 <> "BY" Or type2 <> "BY" Then
                Debug.Print "Skipping to the next iteration; one or more of these is not a factor loading"
                GoTo nextiteration
            End If
                        
            temp_val1 = GET_STRINGVECTOR(temp_val1, "Estimate")
            temp_val2 = GET_STRINGVECTOR(temp_val2, "Estimate")
                        
            ' Drop out of the inner loop if the two values aren't equal (metric equivalence)
            If temp_val1 <> temp_val2 And metric = True Then
                Debug.Print "Breaking loop; factor loadings are not metric equivalent"
                Exit For
            End If
            
            counter = counter + 1
            
nextiteration:
        Next
        If b > UBound(mod_matrix_directional, 2) And counter > 0 Then
            Debug.Print "Matched the factors:"; var_num1; "("; VarName(var_num1); ") and"; var_num2; "("; VarName(var_num2); ")"
            matchFactor = var_num2
            Exit Function
        End If
    Next
End Function

Public Property Get RSquare(ByVal dv_num, Optional ByVal model_num = 1)
    'MsgBox "Requested DV number: " & dv_num
    RSquare = GET_STRINGVECTOR(mod_rsquare(dv_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num)), "Estimate")
End Property

Public Property Get ImportanceWeight(dv_num, iv_num1, Optional std_num = 1, Optional model_num = 1, Optional ByVal formula = True)
    ' Returns the unique variance in a dependent variable (dv_num) accounted for by a predictor (iv_num)
    ' Uses model-based relative importance to determine the weighting
    
    iv_var1 = Variance(iv_num1, std_num, model_num, True, formula)
    Path1 = Path(iv_num1, dv_num, std_num, model_num)
    
    If iv_var1 = "NA" Or Path1 = "NA" Then
        ImportanceWeight = "NA"
        Exit Property
    End If
    
    If formula = True Then
        ImportanceWeight = "SUM((" & iv_var1 & "*" & Path1 & "*" & Path1 & ")"
    Else
        Debug.Print "Calculating importance weight, step 1"
        Debug.Print iv_var1
        Debug.Print Path1
        ImportanceWeight = iv_var1 * Path1 * Path1
    End If
    
    For a = 1 To DV_antecedent(dv_num, 0, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num), True)
        iv_num2 = DV_antecedent(dv_num, a, std_num, mod_Group(model_num), mod_Level(model_num), mod_Mixture(model_num), True)
        If iv_num1 <> iv_num2 Then
            iv_var2 = Variance(iv_num2, std_num, model_num, True, formula)
            iv_cov = Covariance(iv_num1, iv_num2, std_num, model_num, True, formula)
            If iv_var2 = "NA" Or iv_cov = "NA" Then
                ImportanceWeight = "NA"
                Exit Property
            End If
            Path2 = Path(iv_num2, dv_num, std_num, model_num)
            If formula = True Then
                ratio = "((" & iv_var1 & "*" & Path1 & "*" & Path1 & ") / ((" & iv_var1 & "*" & Path1 & "*" & Path1 & ") + (" & iv_var2 & "*" & Path2 & "*" & Path2 & ")))"
                ImportanceWeight = ImportanceWeight & ", (2*" & iv_cov & "*" & Path1 & "*" & Path2 & "*" & ratio & ")"
            Else
                Debug.Print "Calculating importance weight, step 2"
                ratio = (iv_var1 * Path1 * Path1) / ((iv_var1 * Path1 * Path1) + (iv_var2 * Path2 * Path2))
                ImportanceWeight = ImportanceWeight + (2 * iv_cov * Path1 * Path2 * ratio)
            End If
        End If
    Next
    
    If formula = True Then ImportanceWeight = ImportanceWeight & ")"
End Property

Public Property Get RelativeImportance(dv_num, iv_num, Optional std_num = 1, Optional model_num = 1, Optional ByVal formula = True)
    Debug.Print "Getting importance weight"
    IW = ImportanceWeight(dv_num, iv_num, std_num, model_num, formula)
    
    Debug.Print "Getting variance explained"
    VE = VarianceExplained(dv_num, std_num, model_num, formula)
    
    If IW = "NA" Or VE = "NA" Then
        RelativeImportance = "NA"
        Exit Property
    End If
    
    If formula = True Then
        RelativeImportance = IW & " / " & VE
    Else
        RelativeImportance = IW / VE
    End If
End Property


