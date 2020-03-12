Attribute VB_Name = "Output_FitStats"
Sub Start_ModelFit()
    LoadMplusOutput.Show
    
    ' Halt operations if the user exits the form without clicking on the "Proceed" button
    If LoadMplusOutput.execute = False Then Exit Sub
    
    Set MplusOutput = New cMplusOutput
    MplusOutput.ParseOutput = LoadMplusOutput.MPlusInput.text
    Unload LoadMplusOutput
    
    If MplusOutput.IsModel = False Then
        MsgBox "Model not found."
        Exit Sub
    End If
    
' Reset defaults
    Call ResetDefaults
    
    If MplusOutput.Estimator = "BAYES" Then
        Form_FitStats.Bayes = True
        Form_FitStats.PPP = True
    Else
        Form_FitStats.ChiSq = True
        Form_FitStats.CFI = True
        Form_FitStats.RMSEA = True
        Form_FitStats.SRMR = True
    End If
    
    Form_FitStats.Show
    
    If Form_FitStats.execute = False Then Exit Sub
    
    If Form_FitStats.Automatic = True Then
        ActiveCell.Value = ModelFitAuto(MplusOutput)
        Call WriteModelFitAuto(MplusOutput, _
                           Form_FitStats.tabular, _
                           Form_FitStats.HeadRow)
    Else
        Call WriteModelFit(MplusOutput, _
                           Form_FitStats.ChiSq, _
                           Form_FitStats.CFI, _
                           Form_FitStats.TLI, _
                           Form_FitStats.RMSEA, _
                           Form_FitStats.SRMR, _
                           Form_FitStats.Bayes, _
                           Form_FitStats.PPP, _
                           Form_FitStats.AIC, _
                           Form_FitStats.BIC, _
                           Form_FitStats.BICssa, _
                           Form_FitStats.DIC, _
                           Form_FitStats.PD, _
                           Form_FitStats.tabular, _
                           Form_FitStats.HeadRow)
    End If
    
    Unload Form_FitStats
End Sub

Public Sub WriteModelFitAuto(MplusOutput, Optional tabular = False, Optional heading = False)
    y_start = ActiveCell.Row
    x_start = ActiveCell.Column
    If MplusOutput.Estimator = "BAYES" Then
        If tabular = True Then
            If heading = True Then
                If MplusOutput.BayesLower <> Empty Then
                    Cells(y_start, x_start) = ChrW(916) & ChrW(967) & "² 95% CI"
                    Cells(y_start, x_start + 1) = "PPp"
                End If
                y_start = y_start + 1
            End If
            If MplusOutput.BayesLower <> Empty Then
                Cells(y_start, x_start) = "[" & format(MplusOutput.BayesLower, ".000") & ", " & format(MplusOutput.BayesUpper, ".000") & "]"
                Cells(y_start, x_start + 1) = format(MplusOutput.PPP, ".000")
            End If
            Cells(y_start + 1, x_start).Activate
        Else
            Cells(y_start, x_start) = ModelFitAuto(MplusOutput)
        End If
    Else
        If tabular = True Then
            x_offset = 0
            If heading = True Then
                If MplusOutput.ChiSq <> Empty Then
                    Cells(y_start, x_start) = ChrW(967) & "²"
                    Cells(y_start, x_start + 1) = "DF"
                    Cells(y_start, x_start + 2) = "p"
                    x_offset = x_offset + 3
                End If
                If MplusOutput.CFI <> Empty Then
                    Cells(y_start, x_start + x_offset) = "CFI"
                    x_offset = x_offset + 1
                End If
                If MplusOutput.RMSEA <> Empty Then
                    Cells(y_start, x_start + x_offset) = "RMSEA"
                    x_offset = x_offset + 1
                End If
                If MplusOutput.SRMR <> Empty Then
                    Cells(y_start, x_start + x_offset) = "SRMR"
                    x_offset = x_offset + 1
                End If
                If MplusOutput.SRMR_W <> Empty Then
                    Cells(y_start, x_start + x_offset) = "SRMR-W"
                    Cells(y_start, x_start + x_offset + 1) = "SRMR-B"
                    x_offset = x_offset + 2
                End If
                y_start = y_start + 1
                x_offset = 0
            End If
            If MplusOutput.ChiSq <> Empty Then
                Cells(y_start, x_start) = MplusOutput.ChiSq
                Cells(y_start, x_start + 1) = MplusOutput.DF
                Cells(y_start, x_start + 2) = format(MplusOutput.ChiSqP, ".000")
                x_offset = x_offset + 3
            End If
            If MplusOutput.CFI <> Empty Then
                Cells(y_start, x_start + x_offset) = format(MplusOutput.CFI, ".000")
                x_offset = x_offset + 1
            End If
            If MplusOutput.RMSEA <> Empty Then
                Cells(y_start, x_start + x_offset) = format(MplusOutput.RMSEA, ".000")
                x_offset = x_offset + 1
            End If
            If MplusOutput.SRMR <> Empty Then
                Cells(y_start, x_start + x_offset) = format(MplusOutput.SRMR, ".000")
                x_offset = x_offset + 1
            End If
            If MplusOutput.SRMR_W <> Empty Then
                Cells(y_start, x_start + x_offset) = format(MplusOutput.SRMR_W, ".000")
                Cells(y_start, x_start + x_offset + 1) = format(MplusOutput.SRMR_B, ".000")
                x_offset = x_offset + 2
            End If
            x_offset = 0
            Cells(y_start + 1, x_start).Activate
        Else
            Cells(y_start, x_start) = ModelFitAuto(MplusOutput)
        End If
    End If
End Sub

Public Function ModelFitAuto(MplusOutput)
    If MplusOutput.Estimator = "BAYES" Then
        If MplusOutput.BayesLower <> Empty Then
            ModelFitAuto = ChrW(916) & ChrW(967) & "² 95% CI: [" & format(MplusOutput.BayesLower, ".000") & ", " & format(MplusOutput.BayesUpper, ".000") & "]"
            ModelFitAuto = ModelFitAuto & ", PPp = " & format(MplusOutput.PPP, ".000")
        End If
    Else
        If MplusOutput.ChiSq <> Empty Then ModelFitAuto = ChrW(967) & "²(df = " & MplusOutput.DF & ") = " & MplusOutput.ChiSq & ", p = " & format(MplusOutput.ChiSqP, ".000")
        If MplusOutput.CFI <> Empty Then ModelFitAuto = ModelFitAuto & ", CFI = " & format(MplusOutput.CFI, ".000")
        If MplusOutput.RMSEA <> Empty Then ModelFitAuto = ModelFitAuto & ", RMSEA = " & format(MplusOutput.RMSEA, ".000")
        If MplusOutput.SRMR <> Empty Then ModelFitAuto = ModelFitAuto & ", SRMR = " & format(MplusOutput.SRMR, ".000")
        If MplusOutput.SRMR_W <> Empty Then ModelFitAuto = ModelFitAuto & ", SRMR-W = " & format(MplusOutput.SRMR_W, ".000")
        If MplusOutput.SRMR_B <> Empty Then ModelFitAuto = ModelFitAuto & ", SRMR-B = " & format(MplusOutput.SRMR_B, ".000")
    End If
End Function

Public Function ModelComparisonAuto(MplusOutput, Optional tabular = False, Optional heading = False)
    ModelComparisonAuto = ""
    If MplusOutput.Estimator = "BAYES" Then
        If MplusOutput.DIC <> Empty Then
            ModelComparisonAuto = "DIC = " & MplusOutput.DIC
            If MplusOutput.PD <> Empty Then ModelComparisonAuto = ModelComparisonAuto & ", pD = " & MplusOutput.PD
            If MplusOutput.BIC <> Empty Then ModelComparisonAuto = ModelComparisonAuto & ", BIC = " & MplusOutput.BIC
        End If
    Else
        If MplusOutput.AIC <> Empty Then ModelComparisonAuto = ModelComparisonAuto & "AIC = " & MplusOutput.AIC
        If MplusOutput.BIC <> Empty Then ModelComparisonAuto = ModelComparisonAuto & ", BIC = " & MplusOutput.BIC
        If MplusOutput.SSABIC <> Empty Then ModelComparisonAuto = ModelComparisonAuto & ", BIC (sample size adjusted) = " & MplusOutput.SSABIC
    End If
End Function

Sub WriteModelFit(MplusOutput, _
                  Optional ChiSq, Optional CFI, Optional TLI, Optional RMSEA, Optional SRMR, Optional Bayes, Optional PPP, _
                  Optional AIC, Optional BIC, Optional BICssa, Optional DIC, Optional PD, Optional tabular = False, Optional heading = False)
    
    temp_text = ""
    
    If tabular = True Then
        start_y = ActiveCell.Row
        start_x = ActiveCell.Column
        offset_x = 0
        offset_y = 0
        If heading = True Then offset_y = 1
        
        If ChiSq = True And MplusOutput.ChiSq <> Empty Then
            If heading = True Then
                Cells(start_y, start_x + offset_x) = ChrW(967) & "²"
                Cells(start_y, start_x + offset_x + 1) = "DF"
                Cells(start_y, start_x + offset_x + 2) = "p"
            End If
            Cells(start_y + offset_y, start_x + offset_x) = MplusOutput.ChiSq
            Cells(start_y + offset_y, start_x + offset_x + 1) = MplusOutput.DF
            Cells(start_y + offset_y, start_x + offset_x + 2) = format(MplusOutput.ChiSqP, ".000")
            offset_x = offset_x + 3
        End If
        If CFI = True And MplusOutput.CFI <> Empty Then
            If heading = True Then Cells(start_y, start_x + offset_x) = "CFI"
            Cells(start_y + offset_y, start_x + offset_x) = MplusOutput.CFI
            offset_x = offset_x + 1
        End If
        If TLI = True And MplusOutput.TLI <> Empty Then
            If heading = True Then Cells(start_y, start_x + offset_x) = "TLI"
            Cells(start_y + offset_y, start_x + offset_x) = MplusOutput.TLI
            offset_x = offset_x + 1
        End If
        If RMSEA = True And MplusOutput.RMSEA <> Empty Then
            If heading = True Then Cells(start_y, start_x + offset_x) = "RMSEA"
            Cells(start_y + offset_y, start_x + offset_x) = MplusOutput.RMSEA
            offset_x = offset_x + 1
        End If
        If SRMR = True And (MplusOutput.SRMR <> Empty Or MplusOutput.SRMR_W <> Empty) Then
            If heading = True Then
                If MplusOutput.SRMR_W <> Empty Then
                    Cells(start_y, start_x + offset_x) = "SRMR-W"
                    Cells(start_y, start_x + offset_x + 1) = "SRMR-B"
                Else
                    Cells(start_y, start_x + offset_x) = "SRMR"
                End If
            End If
            If MplusOutput.SRMR_W <> Empty Then
                Cells(start_y + offset_y, start_x + offset_x) = MplusOutput.SRMR_W
                Cells(start_y + offset_y, start_x + offset_x + 1) = MplusOutput.SRMR_B
                offset_x = offset_x + 2
            Else
                Cells(start_y + offset_y, start_x + offset_x) = MplusOutput.SRMR
                offset_x = offset_x + 1
            End If
        End If
        If Bayes = True And MplusOutput.BayesLower <> Empty Then
            If heading = True Then Cells(start_y, start_x + offset_x) = ChrW(916) & ChrW(967) & "² 95% CI"
            Cells(start_y + offset_y, start_x + offset_x) = "[" & format(MplusOutput.BayesLower, ".000") & ", " & format(MplusOutput.BayesUpper, ".000") & "]"
            offset_x = offset_x + 1
        End If
        If PPP = True And MplusOutput.PPP <> Empty Then
            If heading = True Then Cells(start_y, start_x + offset_x) = "PPp"
            Cells(start_y + offset_y, start_x + offset_x) = MplusOutput.PPP
            offset_x = offset_x + 1
        End If
        If AIC = True And MplusOutput.AIC <> Empty Then
            If heading = True Then Cells(start_y, start_x + offset_x) = "AIC"
            Cells(start_y + offset_y, start_x + offset_x) = MplusOutput.AIC
            offset_x = offset_x + 1
        End If
        If BIC = True And MplusOutput.BIC <> Empty Then
            If heading = True Then Cells(start_y, start_x + offset_x) = "BIC"
            Cells(start_y + offset_y, start_x + offset_x) = MplusOutput.BIC
            offset_x = offset_x + 1
        End If
        If BICssa = True And MplusOutput.BICssa <> Empty Then
            If heading = True Then Cells(start_y, start_x + offset_x) = "BICssa"
            Cells(start_y + offset_y, start_x + offset_x) = MplusOutput.BICssa
            offset_x = offset_x + 1
        End If
        If DIC = True And MplusOutput.DIC <> Empty Then
            If heading = True Then Cells(start_y, start_x + offset_x) = "DIC"
            Cells(start_y + offset_y, start_x + offset_x) = MplusOutput.DIC
            offset_x = offset_x + 1
        End If
        If PD = True And MplusOutput.PD <> Empty Then
            If heading = True Then Cells(start_y, start_x + offset_x) = "Est N Parameters"
            Cells(start_y + offset_y, start_x + offset_x) = MplusOutput.PD
            offset_x = offset_x + 1
        End If
    Else
        If ChiSq = True And MplusOutput.ChiSq <> Empty Then temp_text = temp_text & ChrW(967) & "²(df = " & MplusOutput.DF & ") = " & MplusOutput.ChiSq & ", p = " & format(MplusOutput.ChiSqP, ".000")
        If CFI = True And MplusOutput.CFI <> Empty Then temp_text = temp_text & ", CFI = " & MplusOutput.CFI
        If TLI = True And MplusOutput.TLI <> Empty Then temp_text = temp_text & ", TLI = " & MplusOutput.TLI
        If RMSEA = True And MplusOutput.RMSEA <> Empty Then temp_text = temp_text & ", RMSEA = " & MplusOutput.RMSEA
        If SRMR = True And MplusOutput.SRMR <> Empty Then temp_text = temp_text & ", SRMR = " & MplusOutput.SRMR
        If SRMR = True And MplusOutput.SRMR_W <> Empty Then temp_text = temp_text & ", SRMR-W = " & MplusOutput.SRMR_W
        If SRMR = True And MplusOutput.SRMR_B <> Empty Then temp_text = temp_text & ", SRMR-B = " & MplusOutput.SRMR_B
        If Bayes = True And MplusOutput.BayesLower <> Empty Then temp_text = temp_text & ChrW(916) & ChrW(967) & "² 95% CI: [" & format(MplusOutput.BayesLower, ".000") & ", " & format(MplusOutput.BayesUpper, ".000") & "]"
        If PPP = True And MplusOutput.PPP <> Empty Then temp_text = temp_text & ", PPp = " & MplusOutput.PPP
        If AIC = True And MplusOutput.AIC <> Empty Then temp_text = temp_text & ", AIC = " & MplusOutput.AIC
        If BIC = True And MplusOutput.BIC <> Empty Then temp_text = temp_text & ", BIC = " & MplusOutput.BIC
        If BICssa = True And MplusOutput.BICssa <> Empty Then temp_text = temp_text & ", BIC (sample size adjusted) = " & MplusOutput.BICssa
        If DIC = True And MplusOutput.DIC <> Empty Then temp_text = temp_text & ", DIC = " & MplusOutput.DIC
        If PD = True And MplusOutput.PD <> Empty Then temp_text = temp_text & ", Est n parameters = " & MplusOutput.PD
        If Left(temp_text, 2) = ", " Then temp_text = Right(temp_text, Len(temp_text) - 2)
        ActiveCell.Value = temp_text
    End If
End Sub


Sub testmfauto()
    LoadMplusOutput.Show
    
    ' Halt operations if the user exits the form without clicking on the "Proceed" button
    If LoadMplusOutput.execute = False Then Exit Sub
    
    Set MplusOutput = New cMplusOutput
    MplusOutput.ParseOutput = LoadMplusOutput.MPlusInput.text
    
    Unload LoadMplusOutput
    
    Call WriteModelFitAuto(MplusOutput, True, True)
End Sub

