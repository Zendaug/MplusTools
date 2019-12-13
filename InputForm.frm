VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} InputForm 
   Caption         =   "Export to Mplus"
   ClientHeight    =   8145
   ClientLeft      =   30
   ClientTop       =   270
   ClientWidth     =   4335
   OleObjectBlob   =   "InputForm.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "InputForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public execute

Private Sub create_datafile_Click()
    If create_datafile = True Then
        data_filename.Enabled = True
        recode_mv.Enabled = True
        miss_value.Enabled = True
        mv_label.Enabled = True
        data_file_label.Enabled = True
        CommandButton1.Enabled = True
        
        MCSimulation.Enabled = True
        MCSyntax.Enabled = True
        MCDatasets.Enabled = True
        MCDatasets_label.Enabled = True
        
    Else
        data_filename.Enabled = False
        recode_mv.Enabled = False
        miss_value.Enabled = False
        mv_label.Enabled = False
        data_file_label.Enabled = False
        If create_syntax = False Then CommandButton1.Enabled = False
        
        MCSimulation.Enabled = False
        MCSyntax.Enabled = False
        MCDatasets.Enabled = False
        MCDatasets_label.Enabled = False
        MCSyntax.Value = False
    End If
End Sub

Private Sub create_syntax_Click()
    If create_syntax = True Then
        syntax_label.Enabled = True
        input_filename.Enabled = True
        If DataStructure.var_labels = True Then var_labels.Enabled = True
        If DataStructure.scale_names = True Then
            scale_names.Enabled = True
            CFASyntax.Enabled = True
        End If
        MCSimulation.Enabled = True
        MCSyntax.Enabled = True
        MCDatasets.Enabled = True
        MCDatasets_label.Enabled = True
    Else
        syntax_label.Enabled = False
        input_filename.Enabled = False
        var_labels.Enabled = False
        scale_names.Enabled = False
        CFASyntax.Enabled = False
        
        MCSimulation.Enabled = False
        MCSyntax.Enabled = False
        MCDatasets.Enabled = False
        MCDatasets_label.Enabled = False
        MCSyntax.Value = False
    End If
End Sub

Private Sub export_to_Change()
    If export_to.ListIndex = 0 Then
        input_filename.text = ActiveSheet.Name & ".inp"
    ElseIf export_to.ListIndex = 1 Then
        input_filename.text = ActiveSheet.Name & ".R"
    End If
End Sub

Private Sub CommandButton1_Click()
    execute = True
    If recode_mv.Value = False Then
        DataStructure.SetMissingValue = ""
    Else
        DataStructure.SetMissingValue = miss_value
    End If
    InputForm.Hide
End Sub

Private Sub MCSyntax_Click()
    If MCSyntax = False Then
        MCDatasets.Enabled = False
        MCDatasets_label.Enabled = False
    Else
        MCDatasets.Enabled = True
        MCDatasets_label.Enabled = True
    End If
End Sub

Private Sub UserForm_Initialize()
    execute = False
    data_filename.text = ActiveSheet.Name & ".csv"
    input_filename.text = ActiveSheet.Name & ".inp"
    
    If DataStructure.var_names = False Then
        MsgBox "Error: Variable names not detected on first row."
        End
    End If
    
    If DataStructure.var_labels = True Then
        var_labels.Enabled = True
        var_labels = True
    End If
    If DataStructure.scale_names = True Then
        scale_names.Enabled = True
        scale_names.Value = True
        CFASyntax.Enabled = True
    End If
    
    'variables = DataStructure.n_variables
    'start_row = DataStructure.start_row
    'end_row = DataStructure.end_row
    
    Status = "Number of variables detected: " & DataStructure.n_variables & Chr(10) & "Number of cases detected: " & DataStructure.n_cases & " (data starts on row " & DataStructure.start_row & " and ends on row " & DataStructure.end_row & ")."
    
    export_to.AddItem ("Mplus")
    export_to.AddItem ("R")
    export_to.ListIndex = 0
End Sub


