VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} LoadMplusOutput 
   Caption         =   "Load Mplus Output"
   ClientHeight    =   4680
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   13500
   OleObjectBlob   =   "LoadMplusOutput.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "LoadMplusOutput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public execute As Boolean


Private Sub CommandButton6_Click()

End Sub

Private Sub Continue_Click()
    syntax_text = LoadMplusOutput.MPlusInput
    execute = True
    Me.Hide
End Sub

Private Sub LoadOutput_Click()
    Dim fd As Office.FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
    
       .AllowMultiSelect = False
    
       ' Set the title of the dialog box.
       .TITLE = "Please select the file."
    
       ' Clear out the current filters, and add our own.
       .Filters.Clear
       .Filters.Add "Mplus Output File", "*.out"
       .Filters.Add "All Files", "*.*"
    
       ' Show the dialog box. If the .Show method returns True, the
       ' user picked at least one file. If the .Show method returns
       ' False, the user clicked Cancel.
       If .Show = True Then
         txtfilename = .SelectedItems(1) 'replace txtFileName with your textbox
    
       End If
    End With
        
    Open txtfilename For Input As #1
    
    Do Until EOF(1)
        Line Input #1, textline
        text = text & textline & Chr(10)
    Loop
    Close #1
    
    MPlusInput = text
End Sub

Private Sub MPlusInput_Change()

End Sub

Private Sub UserForm_Initialize()
    execute = False
    If IsEmpty(syntax_text) = False Then LoadMplusOutput.MPlusInput = syntax_text
End Sub
