VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Settings 
   Caption         =   "Settings"
   ClientHeight    =   3000
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "Settings.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Settings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Continue_Click()
    If useformula.ListIndex = 0 Then use_formula = False
    If useformula.ListIndex = 1 Then use_formula = True
    
    n_decimals = DecimalPlaces.ListIndex + 1
    var_disp_mode = VarDisplay.ListIndex
    Unload Me
End Sub

Private Sub UserForm_Initialize()
    useformula.AddItem ("No")
    useformula.AddItem ("Yes")
    
    If use_formula = True Then
        useformula.ListIndex = 1
    Else
        useformula.ListIndex = 0
    End If
    
    DecimalPlaces.AddItem ("1")
    DecimalPlaces.AddItem ("2")
    DecimalPlaces.AddItem ("3")
    DecimalPlaces.ListIndex = n_decimals - 1
    
    VarDisplay.AddItem ("Short variable names only")
    VarDisplay.AddItem ("Variable labels")
    VarDisplay.AddItem ("Variable labels (and short names)")
    VarDisplay.ListIndex = var_disp_mode
End Sub
