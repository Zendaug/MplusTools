VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cDataStructure"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private pVarNames()
' 1: Variable short name
' 2: Variable long name / label
' 3: Scale short name

Private pScaleNames()
' 1: Scale short name
' 2: Scale long name / label
' 3: Include scale in CFA

Private pDataset(), pDatasetFiltered()
Private filter
Private miss_value
Private temp_dataset
Public n_variables, n_cases, start_row, end_row
Public var_names, var_labels, scale_names

Public Sub UpdateDataStructure(Optional wsheet = "")
    ' This procedure loads the data structure from the active worksheet
    
' Count the number of cases
    Debug.Print "Determining the number of cases"
    For a = start_row To 1048576
        If WorksheetFunction.Count(ActiveSheet.Rows(a)) = 0 Then
            n_cases = a - start_row
            Exit For
        End If
    Next

    end_row = a - 1

' Load the raw dataset
    Debug.Print "Loading the raw dataset"
    temp_dataset = Range(Cells(start_row, 1), Cells(end_row, n_variables)).Value
    var_ctr = 0
    
' Change all of the errors to blanks
    Debug.Print "Converting all of the errors to blanks"
    For y = 1 To UBound(temp_dataset, 1)
        For x = 1 To UBound(temp_dataset, 2)
            If VarType(temp_dataset(y, x)) = vbError Then temp_dataset(y, x) = ""
        Next
    Next
End Sub

Public Sub CreateDataStructure(Optional wsheet = "")
    'Thisprocedure

    Debug.Print "Creating Data Structure"
    
    ' Defaults
    start_row = 1
    var_names = False
    var_labels = False
    scale_names = False
    miss_value = ""
    filter = False
    
' Count the number of n_variables
    Debug.Print "Counting"
    For a = 1 To 10000
        temp_var = Cells(1, a).Value
        ' Quit if any of the variable names are numeric
        If IsNumeric(temp_var) And temp_var <> "" Then
            MsgBox "Error: Data (numbers) found in Row #1. This row should only contain a list of variable names."
            End
        End If
        If temp_var = "" Then
            n_variables = a - 1
            Exit For
        End If
    Next

' Quit if there are 0 variables
    Debug.Print "Quitting if there are no variables"
    If n_variables = 0 Then
        MsgBox "Error: No variable names were found in Row #1."
        End
    End If

' Check to see whether row 1 contains variable names
    Debug.Print "Checking row 1 for variable names"
    If WorksheetFunction.Count(Range(Cells(1, 1), Cells(1, n_variables))) = 0 Then
        'InputForm.CheckBox2 = True
        var_names = True
        start_row = 2
    End If

' Check to see whether row 2 contains item texts
    Debug.Print "Checking row 2 for item texts"
    If WorksheetFunction.Count(Range(Cells(2, 1), Cells(2, n_variables))) = 0 Then
        var_labels = True
        'InputForm.CheckBox3 = True
        start_row = 3
    End If

' Check to see whether row 3 contains item texts
    Debug.Print "Checking row 3 for scale names"
    If WorksheetFunction.Count(Range(Cells(3, 1), Cells(3, n_variables))) = 0 Then
        'MsgBox WorksheetFunction.Count(Range(Cells(3, 1), Cells(3, n_variables)))
        'InputForm.ScaleNames = True
        'CFASyntax.Enabled = True
        scale_names = True
        start_row = 5
    End If

    Call UpdateDataStructure(wsheet)
    Debug.Print "Finished updating the dataset"

' Create filtered dataset and variable names
    For x = 1 To UBound(temp_dataset, 2)
        temp_cells = ARRAY_COLUMN(temp_dataset, x)
        If WorksheetFunction.Count(temp_cells) > 1 Then ' Only enter the loop if the column contains numbers
            If WorksheetFunction.StDev_S(temp_cells) > 0 Then ' Only enter the loop if the standard deviation is more than 0
                
                ' Add the variable information
                var_ctr = var_ctr + 1
                var_name = ActiveSheet.Cells(1, x)
                var_label = ""
                scale_name = ""
                scale_label = ""
                If var_labels = True Then var_label = ActiveSheet.Cells(2, x)
                If scale_names = True Then
                    scale_name = ActiveSheet.Cells(3, x)
                    scale_label = ActiveSheet.Cells(4, x)
                End If
                AddVariable var_name, var_label, scale_name, scale_label
      
                ' Only keep the columns with associated variables (i.e., get rid of all of the non-numeric and 0 SD columns)
                ReDim Preserve pDataset(1 To UBound(temp_dataset, 1), 1 To var_ctr)
                For y = 1 To UBound(temp_dataset, 1)
                    pDataset(y, var_ctr) = temp_dataset(y, x)
                Next
                n_variables = var_ctr
            End If
        End If
    Next
End Sub

Public Sub RefreshDatastructure(Optional wsheet = "")
    Call UpdateDataStructure(wsheet)
    Debug.Print "Finished updating the dataset"

' Create filtered dataset and variable names
    For x = 1 To UBound(temp_dataset, 2)
        temp_cells = ARRAY_COLUMN(temp_dataset, x)
        If WorksheetFunction.Count(temp_cells) > 1 Then ' Only enter the loop if the column contains numbers
            If WorksheetFunction.StDev_S(temp_cells) > 0 Then ' Only enter the loop if the standard deviation is more than 0
                
                ' Add the variable information
                var_ctr = var_ctr + 1
      
                ' Only keep the columns with associated variables (i.e., get rid of all of the non-numeric and 0 SD columns)
                ReDim Preserve pDataset(1 To UBound(temp_dataset, 1), 1 To var_ctr)
                For y = 1 To UBound(temp_dataset, 1)
                    pDataset(y, var_ctr) = temp_dataset(y, x)
                Next
                n_variables = var_ctr
            End If
        End If
    Next
    filter = False
End Sub
 

Private Sub FilterData()
    ' This procedure filters the data, removing all of the non-blank and missing values. It does it the only the first time it is requested.
    If filter = False Then
        pDatasetFiltered = pDataset
        If miss_value <> "" Then ' Start the process if the missing value is defined
            For y = 1 To UBound(pDatasetFiltered, 1)
                For x = 1 To UBound(pDatasetFiltered, 2)
                    If WorksheetFunction.IsText(pDatasetFiltered(y, x)) = True Or pDatasetFiltered(y, x) = "" Then
                        pDatasetFiltered(y, x) = miss_value
                    End If
                Next
            Next
        End If
        filter = True
    End If
End Sub

Public Property Let SetMissingValue(m_value)
    miss_value = m_value
    filter = False
End Property

Public Property Get MissingValue()
    MissingValue = miss_value
End Property

Public Property Get Dataset(y, x)
    FilterData ' Filter the data if it hasn't been done yet
    Dataset = pDatasetFiltered(y, x)
End Property

Public Property Get Cases_N()
    ' Returns the number of cases
    Cases_N = UBound(pDataset, 1)
End Property

Public Property Get Variables_N()
    ' Returns the number of variables
    Variables_N = UBound(pDataset, 2)
End Property

Public Sub AddVariable(var_name, var_label, scale_name, scale_label)
    Debug.Print "Adding entries to variable array"; var_name; var_label; scale_name; scale_label
    
    ' Resize array; if it can't be done, create the size for the first time
    On Error GoTo err1
    ReDim Preserve pVarNames(1 To 3, 1 To UBound(pVarNames, 2) + 1)
    GoTo err1_end
err1:
    Debug.Print "Resizing variable array"
    ReDim pVarNames(1 To 3, 1 To 1)
    
err1_end:
    On Error GoTo 0
    
    ' Checking to see whether duplicate variable name exists
    For a = 1 To UBound(pVarNames, 2) - 1
        If var_name = pVarNames(1, a) Or format_mplus(var_name) = format_mplus(pVarNames(1, a)) Then
            MsgBox "Error - duplicate variable name found: " & var_name & " (formatted as: '" & format_mplus(var_name) & "'), positions " & a & " and " & UBound(pVarNames, 2) & ". Halting execution."
            End
        End If
    Next
    
    Debug.Print "Adding data to variable names array: " & UBound(pVarNames, 2) & " - " & var_name & " " & var_label & " " & scale_name
    pVarNames(1, UBound(pVarNames, 2)) = var_name
    pVarNames(2, UBound(pVarNames, 2)) = var_label
    pVarNames(3, UBound(pVarNames, 2)) = scale_name

    If scale_name <> "" Then AddScale scale_name, scale_label
End Sub

Public Sub AddData(array_name, miss_value)
    pDataset = array_name
    
    For y = LBound(pDataset, 1) To UBound(pDataset, 1)
        For x = LBound(pDataset, 2) To UBound(pDataset, 2)
            If pDataset(y, x) = miss_value Then
                pDataset(y, x) = "NA"
            End If
        Next
    Next
End Sub
    
Private Sub AddScale(scale_name, scale_label)
    Dim scale_found: scale_found = False
    On Error GoTo err2

    Debug.Print "Searching to see if the scale exists"
    num_scales = UBound(pScaleNames, 2)
    Debug.Print "Number of scales: "; num_scales
    For a = 1 To num_scales
        If pScaleNames(1, a) = scale_name Then
            scale_found = True
            Exit For
        End If
    Next
    If scale_found = False Then ReDim Preserve pScaleNames(1 To 3, 1 To num_scales + 1)
    GoTo err2_end

err2:
    Debug.Print "Resizing variable array"
    ReDim pScaleNames(1 To 3, 1 To 1)

err2_end:
    On Error GoTo 0
    
    Debug.Print "Adding data to scales names array"
    If scale_found = False Then
        pScaleNames(1, UBound(pScaleNames, 2)) = scale_name
        pScaleNames(2, UBound(pScaleNames, 2)) = scale_label
        pScaleNames(3, UBound(pScaleNames, 2)) = 0
    End If
End Sub
Public Property Get VarName(Optional var_num = 0)
    If var_num = 0 Then
        VarName = UBound(pVarNames, 2)
        Exit Property
    End If

    VarName = format_mplus(pVarNames(1, var_num))
End Property

Public Property Get VarNumber(var_name)
' Returns the variable number for a nominated name
    For a = 1 To UBound(pVarNames, 2)
        If pVarNames(1, a) = var_name Then
            VarNumber = a
        End If
    Next
End Property

Public Property Get VarLabel(var_num)
    If IsNumeric(var_num) Then
        VarLabel = pVarNames(2, var_num)
    Else
        For a = 1 To UBound(pVarNames, 2)
            If pVarNames(1, a) = var_num Then
                VarLabel = pVarNames(2, a)
                Exit Property
            End If
        Next
    End If
End Property

Public Property Get ScaleName(Optional ByVal scale_num = 0, Optional only_included = False)
    If only_included = True Then scale_num = ScaleIncluded(scale_num)
    
    If scale_num = 0 Then
        ScaleName = UBound(pScaleNames, 2)
        Exit Property
    Else
        If scale_num > 0 And scale_num <= UBound(pScaleNames, 2) Then
            ScaleName = pScaleNames(1, scale_num)
        End If
    End If
End Property

Public Property Get ScaleNumber(scale_name)
' Returns the scale number for a nominated name (looks for either the name or label)
    For a = 1 To UBound(pScaleNames, 2)
        If pScaleNames(1, a) = scale_name Or pScaleNames(2, a) = scale_name Then
            ScaleNumber = a
        End If
    Next
End Property

Public Property Get ScaleLabel(Optional scale_num = 0, Optional only_included = False)
    If only_included = True Then scale_num = ScaleIncluded(scale_num)
    
    If scale_num = 0 Then
        ScaleLabel = UBound(pScaleNames, 2)
        Exit Property
    Else
        If scale_num > 0 And scale_num <= UBound(pScaleNames, 2) Then
            ScaleLabel = pScaleNames(2, scale_num)
        End If
    End If
End Property

Public Property Get ScaleInclude(Optional scale_num = 0)
    If scale_num = 0 Then ' Return the number of included scales
        ScaleInclude = 0
        For a = 1 To UBound(pScaleNames, 2)
            If pScaleNames(3, a) = 1 Then ScaleInclude = ScaleInclude + 1
        Next
    Else
        ScaleInclude = pScaleNames(3, scale_num)
    End If
End Property

Private Function ScaleIncluded(ByVal scale_num)
    ' Returns the number of the included scale
    counter = 0
    For a = 1 To UBound(pScaleNames, 2)
        If pScaleNames(3, a) = 1 Then
            counter = counter + 1
            If counter = scale_num Then
                ScaleIncluded = a
                Exit Function
            End If
        End If
    Next
    
    ScaleIncluded = 0
End Function

Public Property Let IncludeScale(scale_num)
    pScaleNames(3, scale_num) = 1
End Property

Public Property Let ExcludeScale(scale_num)
    pScaleNames(3, scale_num) = 0
End Property

Public Property Get ScaleIndicator(ByVal scale_num, Optional ByVal ind_num = 0, Optional ByVal only_included = False)
    If only_included = True Then scale_num = ScaleIncluded(scale_num)
    
    scal_name = ScaleName(scale_num)
    ind_count = 0
    
    For a = 1 To UBound(pVarNames, 2)
        If scal_name = pVarNames(3, a) Then
            ind_count = ind_count + 1
            If ind_count = ind_num Then
                ScaleIndicator = pVarNames(1, a)
                Exit Property
            End If
        End If
    Next
    
    ' If indicator number is 0, return the number of indicators
    ScaleIndicator = ind_count
End Property

Public Property Get Variance(ByVal var_num)
    If IsNumeric(var_num) = False Then var_num = VarNumber(var_num)
    temp_array = ARRAY_COLUMN(pDataset, var_num)
    Variance = WorksheetFunction.Var_S(temp_array)
End Property
