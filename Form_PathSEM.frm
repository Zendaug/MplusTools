VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Form_PathSEM 
   Caption         =   "Path Analysis / SEM"
   ClientHeight    =   6360
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8535
   OleObjectBlob   =   "Form_PathSEM.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Form_PathSEM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public group_index
Public level_index
Public mixture_index
Public stand_index
Public ci_index
Public ri_index
Public FitStats
Public ci90
Public ci95
Public ci99
Public CI
Public execute
Public nGroups
Public nLevels
Public nMixtures
Public group_output
Public level_output
Public mixture_output
Public stand_output
Public ci_output
Public ri_output

Private Sub ModelNum_Change()
    If ModelNum.ListIndex > 0 Then
        Label8.Enabled = False
        Label9.Enabled = False
        Label11.Enabled = False
        Label12.Enabled = False
        
        PopCol.Enabled = False
        PopRow.Enabled = False
        PopAdjRow.Enabled = False
        PopAdjCol.Enabled = False
        
        PopCol.ListIndex = 0
        PopRow.ListIndex = 0
        PopAdjRow.ListIndex = 0
        PopAdjCol.ListIndex = 0
    Else
        If PopCol.ListCount > 1 Then
            Label8.Enabled = True
            Label9.Enabled = True
            Label11.Enabled = True
            Label12.Enabled = True
            
            PopCol.Enabled = True
            PopRow.Enabled = True
            PopAdjRow.Enabled = True
            PopAdjCol.Enabled = True
        End If
    End If
End Sub

Private Sub change_note()
    Note = "Note. "
    If PVal.Value = True Then Note = Note & "* p < .05, ** p < .01, *** p < .001. "
    If ModelFit = True And FitStats <> "" Then Note = Note & "Model Fit: " & FitStats
End Sub

Private Sub PopRow_Change()

End Sub

Private Sub PVal_Click()
    Call change_note
End Sub

Private Sub CommandButton1_Click()
    Dim temp_array(), temp_lookup(), temp_count
    
    CI = 0
    If CLevel.ListIndex > 0 Then
        If ci90 = CLevel.ListIndex Then CI = 90
        If ci95 = CLevel.ListIndex Then CI = 95
        If ci99 = CLevel.ListIndex Then CI = 99
    End If
    
    group_output = 0
    level_output = 0
    mixture_output = 0
    ci_output = 0
    ri_output = 0
    
    If PopRow.ListIndex = group_index Then group_output = 1
    If PopCol.ListIndex = group_index Then group_output = 2
    If PopAdjRow.ListIndex = group_index Then group_output = 3
    If PopAdjCol.ListIndex = group_index Then group_output = 4
    
    If PopRow.ListIndex = level_index Then level_output = 1
    If PopCol.ListIndex = level_index Then level_output = 2
    If PopAdjRow.ListIndex = level_index Then level_output = 3
    If PopAdjCol.ListIndex = level_index Then level_output = 4
    
    If PopRow.ListIndex = mixture_index Then mixture_output = 1
    If PopCol.ListIndex = mixture_index Then mixture_output = 2
    If PopAdjRow.ListIndex = mixture_index Then mixture_output = 3
    If PopAdjCol.ListIndex = mixture_index Then mixture_output = 4
    
    If PopRow.ListIndex = stand_index Then stand_output = 1
    If PopCol.ListIndex = stand_index Then stand_output = 2
    If PopAdjRow.ListIndex = stand_index Then stand_output = 3
    If PopAdjCol.ListIndex = stand_index Then stand_output = 4

    If PopRow.ListIndex = ci_index Then ci_output = 1
    If PopCol.ListIndex = ci_index Then ci_output = 2
    If PopAdjRow.ListIndex = ci_index Then ci_output = 3
    If PopAdjCol.ListIndex = ci_index Then ci_output = 4

    If PopRow.ListIndex = ri_index Then ri_output = 1
    If PopCol.ListIndex = ri_index Then ri_output = 2
    If PopAdjRow.ListIndex = ri_index Then ri_output = 3
    If PopAdjCol.ListIndex = ri_index Then ri_output = 4

    execute = True
    Me.Hide
End Sub

Private Sub UserForm_Initialize()
    execute = False
End Sub
