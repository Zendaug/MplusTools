VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Form_FactorCorrelations 
   Caption         =   "Factor / Composite Covariances / Correlations"
   ClientHeight    =   5520
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8625
   OleObjectBlob   =   "Form_FactorCorrelations.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Form_FactorCorrelations"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public execute

Private Sub UserForm_Initialize()
    execute = False
    
    VarDisplay.AddItem ("Short variable names only")
    VarDisplay.AddItem ("Variable labels")
    VarDisplay.AddItem ("Variable labels (and short names)")
    VarDisplay.ListIndex = 1

End Sub

Private Sub CommandButton1_Click()
    execute = True
    Me.Hide
End Sub


