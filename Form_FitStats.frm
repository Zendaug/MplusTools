VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Form_FitStats 
   Caption         =   "Model Fit Statistics"
   ClientHeight    =   7185
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3630
   OleObjectBlob   =   "Form_FitStats.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Form_FitStats"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public execute

Private Sub Automatic_Click()
    If Automatic.Value = True Then
        Frame1.Enabled = False
        ChiSq.Enabled = False
        CFI.Enabled = False
        TLI.Enabled = False
        RMSEA.Enabled = False
        SRMR.Enabled = False
        Bayes.Enabled = False
        PPP.Enabled = False
        Frame2.Enabled = False
        AIC.Enabled = False
        BIC.Enabled = False
        BICssa.Enabled = False
        DIC.Enabled = False
        PD.Enabled = False
    Else
        Frame1.Enabled = True
        ChiSq.Enabled = True
        CFI.Enabled = True
        TLI.Enabled = True
        RMSEA.Enabled = True
        SRMR.Enabled = True
        Bayes.Enabled = True
        PPP.Enabled = True
        Frame2.Enabled = True
        AIC.Enabled = True
        BIC.Enabled = True
        BICssa.Enabled = True
        DIC.Enabled = True
        PD.Enabled = True
    End If
End Sub

Private Sub CommandButton1_Click()
    execute = True
    Form_FitStats.Hide
End Sub

Private Sub tabular_Click()
    If tabular = True Then
        HeadRow.Enabled = True
    Else
        HeadRow.Enabled = False
        HeadRow.Value = False
    End If
End Sub

Private Sub UserForm_Initialize()
    execute = False
End Sub
