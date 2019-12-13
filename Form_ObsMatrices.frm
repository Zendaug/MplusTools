VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Form_ObsMatrices 
   Caption         =   "Observed Correlation / Covariance Matrix"
   ClientHeight    =   8415
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4380
   OleObjectBlob   =   "Form_ObsMatrices.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Form_ObsMatrices"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public execute As Boolean

Private Sub CheckBox2_Click()
    Call change_note
End Sub

Private Sub CommandButton1_Click()
    execute = True
    Me.Hide
End Sub

Private Sub Correlation_Click()
    Option_Above.Enabled = True
    Call change_note
End Sub

Private Sub Covariance_Click()
    Option_Above.Enabled = False
    Option_None.Value = True
    Call change_note
End Sub

Private Sub Label3_Click()
    Call change_note
End Sub

Private Sub Label4_Click()
    Call change_note
End Sub

Private Sub matrix1_Change()
    Call change_note
End Sub

Sub change_note()
    n1 = MplusOutput.Sample_N(matrix1.ListIndex + 1)
    n2 = 0

    t05 = WorksheetFunction.TInv(0.05, n1 - 1)
    t01 = WorksheetFunction.TInv(0.01, n1 - 1)
    t001 = WorksheetFunction.TInv(0.001, n1 - 1)
    r05 = WorksheetFunction.Round(t_to_r(t05, n1), 2)
    r01 = WorksheetFunction.Round(t_to_r(t01, n1), 2)
    r001 = WorksheetFunction.Round(t_to_r(t001, n1), 2)
    
    If matrix2.ListIndex > 0 Then
        n2 = MplusOutput.Sample_N(matrix2.ListIndex)
        t05_a = WorksheetFunction.TInv(0.05, n2 - 1)
        t01_a = WorksheetFunction.TInv(0.01, n2 - 1)
        t001_a = WorksheetFunction.TInv(0.001, n2 - 1)
        r05_a = WorksheetFunction.Round(t_to_r(t05_a, n2), 2)
        r01_a = WorksheetFunction.Round(t_to_r(t01_a, n2), 2)
        r001_a = WorksheetFunction.Round(t_to_r(t001_a, n2), 2)
    End If
    
    Note = "Note. "
    If Option_APA = True Then
        Note = Note & "* p < .05, ** p < .01, *** p < .001; N = " & n1 & " (below diagonal)"
        If n2 > 0 Then Note = Note & ", " & "N = " & n2 & " (above diagonal)"
        Note = Note & "."
    End If
    If Option_Above = True Then
        If n2 > 0 Then
            Note = Note & "Correlations below the diagonal are"
        Else
            Note = Note & "Correlations are"
        End If
        Note = Note & " significant above " & format(r05, ".00") & " (p < .05), " & format(r01, ".00") & " (p < .01) and " & format(r001, ".00") & " (p < .001); N = " & n1 & "."
        If n2 > 0 Then
            Note = Note & " Correlations above the diagonal are significant above " & format(r05_a, ".00") & " (p < .05), " & format(r01_a, ".00") & " (p < .01) and " & format(r001_a, ".00") & " (p < .001); N = " & n2 & "."
        End If
    End If
    If Option_None = True Then
        Note = Note & "N = " & n1 & " (below diagonal)"
        If n2 > 0 Then Note = Note & ", " & "N = " & n2 & " (above diagonal)"
        Note = Note & "."
    End If
End Sub

Private Sub matrix2_Change()
    Call change_note
End Sub

Private Sub Option_Above_Click()
    Call change_note
End Sub

Private Sub Option_APA_Click()
    Call change_note
End Sub
Private Sub Option_None_Click()
    Call change_note
End Sub

Private Sub UserForm_Initialize()
    execute = False
End Sub
