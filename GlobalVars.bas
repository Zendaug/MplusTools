Attribute VB_Name = "GlobalVars"
' Version 1.0.3:
' Fixed a bug causing the Monte Carlo procedure to produce exactly the same dataset

' Version 1.0.2:
' Fixed the duplicates algorithm so that compares the two variable names that have been formatted to 8 characters or fewer using the "format_mplus" function
' Modified the "format_mplus" algorithm for formatting the Mplus variable names.
' Fixed a bug in the observed correlation tables that would not allow you to display variable names / labels.

' Version 1.0.1:
' Fixed bug in Fit Statistics tabular output where the DF and p were not displayed for the chi-square statistic
' Fixed a bug in which MEANS/INTERCEPTS/THRESHOLDS were not properly loaded from the sample statistics for categorical variables

' Global variables

Public MplusOutput, DataStructure
Public syntax_text

Public use_formula
Public n_decimals     ' Number of decimal places (1, 2, 3)
Public var_disp_mode  ' 0: display variable name only, 1: display variable label, 2: display variable label and [name]


Sub GotoSettings()
    Call ResetDefaults
    Settings.Show
End Sub

Sub ResetDefaults()
    If IsEmpty(use_formula) Then use_formula = False
    If IsEmpty(n_decimals) Then n_decimals = 2
    If IsEmpty(var_disp_mode) Then var_disp_mode = 1
End Sub

Public Sub ExportModules()
' A sub that will export all of the modules for upload to Github
    Dim bExport As Boolean
    Dim wkbSource As Excel.Workbook
    Dim szSourceWorkbook As String
    Dim szExportPath As String
    Dim szFileName As String
    Dim cmpComponent As VBIDE.VBComponent

    ''' The code modules will be exported in a folder named.
    ''' VBAProjectFiles in the Documents folder.
    ''' The code below create this folder if it not exist
    ''' or delete all files in the folder if it exist.
    If FolderWithVBAProjectFiles = "Error" Then
        MsgBox "Export Folder not exist"
        Exit Sub
    End If
    
    On Error Resume Next
        Kill FolderWithVBAProjectFiles & "\*.*"
    On Error GoTo 0

    ''' NOTE: This workbook must be open in Excel.
    'szSourceWorkbook = ActiveWorkbook.Name
    'Set wkbSource = Application.Workbooks(szSourceWorkbook)

    Set wkbSource = Application.Workbooks("Mplus Tools.xlam")
    
    If wkbSource.VBProject.Protection = 1 Then
    MsgBox "The VBA in this workbook is protected," & _
        "not possible to export the code"
    Exit Sub
    End If
    
    Dim sFolder As String
    ' Open the select folder prompt
    With Application.FileDialog(msoFileDialogFolderPicker)
        If .Show = -1 Then ' if OK is pressed
            sFolder = .SelectedItems(1)
        End If
    End With
    
    If sFolder <> "" Then ' if a file was chosen
        szExportPath = sFolder & "\"
    Else
        Exit Sub
    End If
    
    For Each cmpComponent In wkbSource.VBProject.VBComponents
        
        bExport = True
        szFileName = cmpComponent.Name

        ''' Concatenate the correct filename for export.
        Select Case cmpComponent.Type
            Case vbext_ct_ClassModule
                szFileName = szFileName & ".cls"
            Case vbext_ct_MSForm
                szFileName = szFileName & ".frm"
            Case vbext_ct_StdModule
                szFileName = szFileName & ".bas"
            Case vbext_ct_Document
                ''' This is a worksheet or workbook object.
                ''' Don't try to export.
                bExport = False
        End Select
        
        If bExport Then
            ''' Export the component to a text file.
            cmpComponent.Export szExportPath & szFileName
            
        ''' remove it from the project if you want
        '''wkbSource.VBProject.VBComponents.Remove cmpComponent
        
        End If
   
    Next cmpComponent

    MsgBox "Export is ready"
End Sub


Function FolderWithVBAProjectFiles() As String
    Dim WshShell As Object
    Dim FSO As Object
    Dim SpecialPath As String

    Set WshShell = CreateObject("WScript.Shell")
    Set FSO = CreateObject("scripting.filesystemobject")

    SpecialPath = WshShell.SpecialFolders("MyDocuments")

    If Right(SpecialPath, 1) <> "\" Then
        SpecialPath = SpecialPath & "\"
    End If
    
    If FSO.FolderExists(SpecialPath & "VBAProjectFiles") = False Then
        On Error Resume Next
        MkDir SpecialPath & "VBAProjectFiles"
        On Error GoTo 0
    End If
    
    If FSO.FolderExists(SpecialPath & "VBAProjectFiles") = True Then
        FolderWithVBAProjectFiles = SpecialPath & "VBAProjectFiles"
    Else
        FolderWithVBAProjectFiles = "Error"
    End If
    
End Function

