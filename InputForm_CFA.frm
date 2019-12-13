VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} InputForm_CFA 
   Caption         =   "Input Syntax for Confirmatory Factor Analysis"
   ClientHeight    =   7305
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10335
   OleObjectBlob   =   "InputForm_CFA.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "InputForm_CFA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public execute

Private Sub BCFA_CL_Click()
    If BCFA_CL = True Or BCFA_RC = True Then
        LV1.Enabled = False
        LV2.Enabled = False
        LV3.Enabled = False
        LV3b.Enabled = False
        LV4.Enabled = False
        LV1.Value = True
        LV2.Value = False
        LV3.Value = False
        LV3b.Value = False
        LV4.Value = False
    Else
        LV1.Enabled = True
        LV2.Enabled = True
        LV3.Enabled = True
        LV4.Enabled = True
    End If
End Sub

Private Sub BCFA_RC_Click()
    Call BCFA_CL_Click
End Sub

Private Sub Command_Add_Click()
    For a = 0 To ExcludeList.ListCount - 1
        If ExcludeList.Selected(a) = True Then
            scale_num = DataStructure.ScaleNumber(ExcludeList.List(a))
            DataStructure.IncludeScale = scale_num
        End If
    Next
    populate_list
End Sub

Private Sub Command_Remove_Click()
    For a = 0 To IncludeList.ListCount - 1
        If IncludeList.Selected(a) = True Then
            scale_num = DataStructure.ScaleNumber(IncludeList.List(a))
            DataStructure.ExcludeScale = scale_num
        End If
    Next
    populate_list
End Sub

Private Sub Continue_Click()
    execute = True
    Me.Hide
End Sub


Private Sub LV1_Click()
    LV3b.Enabled = False
    LV3b = False
End Sub

Private Sub LV2_Click()
    LV3b.Enabled = False
    LV3b = False
End Sub

Private Sub LV3_Click()
    LV3b.Enabled = True
    LV3b = True
End Sub

Private Sub UserForm_Initialize()
    execute = False
    populate_list
End Sub

Private Sub populate_list()
    ExcludeList.Clear
    IncludeList.Clear

    Debug.Print "Populating list..."
    For a = 1 To DataStructure.ScaleName
        Debug.Print a; " of "; DataStructure.ScaleName
        If DataStructure.ScaleInclude(a) = 0 Then
            ExcludeList.AddItem DataStructure.ScaleLabel(a)
        Else
            IncludeList.AddItem DataStructure.ScaleLabel(a)
        End If
    Next
End Sub
