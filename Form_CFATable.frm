VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Form_CFATable 
   Caption         =   "Confirmatory Factor Analysis"
   ClientHeight    =   6825
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8625
   OleObjectBlob   =   "Form_CFATable.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Form_CFATable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public execute As Boolean
Public FitStats As String

Private Sub ModelFit_Click()
    Call change_note
End Sub

Private Sub PVal_Click()
    Call change_note
End Sub

Private Sub StandNum_Change()
    If StandNum.ListIndex > 0 Then
        Heading2 = "Standardized Factor Loadings for the Measurement Model"
    Else
        Heading2 = "Unstandardized Factor Loadings for the Measurment Model"
    End If
End Sub

Private Sub UserForm_Initialize()
    execute = False
    
    CoefAction.AddItem ("")
    CoefAction.AddItem ("Hide coefficients below this size")
    CoefAction.AddItem ("Bold coefficients above this size")
    
    CoefAction.ListIndex = 0

End Sub

Private Sub UserForm_Activate()
    Call change_note
End Sub

Private Sub CommandButton1_Click()
    If IsNumeric(HideBelow) = False Then HideBelow = ""
    execute = True
    Me.Hide
End Sub

Private Sub change_note()
    Note = "Note. "
    If PVal.Value = True Then Note = Note & "* p < .05, ** p < .01, *** p < .001. "
    If ModelFit = True And FitStats <> "" Then Note = Note & "Model Fit: " & FitStats
End Sub
