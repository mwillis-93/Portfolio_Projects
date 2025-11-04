Attribute VB_Name = "ParaCalc_v5_IP"
Option Explicit

Function MonthRound(EffDate As Date) As Date
    
    If Day(EffDate) = 1 Then
        MonthRound = EffDate
    ElseIf Day(EffDate) <= 15 Then
        MonthRound = DateSerial(Year(EffDate), Month(EffDate), 1)
    Else
        MonthRound = DateSerial(Year(EffDate), Month(EffDate) + 1, 1)
    End If
        
End Function
Sub ParaCalc()
    
    Dim AYRng As Range, EDRng As Range
    Dim PAR As Double
    Dim AYCell As Range, EDCell As Range
    Dim AccYr As Integer
    Dim EffDate As Date, EDm As Date, EDp As Date
    Dim EDNew As Long, EDOld As Long
    Dim RI As Double, RunTot As Double, CumlRI As Double
    
    Set EDRng = Range("A13:A35")
    Set AYRng = Range("A49:A59")
        
    'Loop through Accident Years
    For Each AYCell In AYRng
        If IsNumeric(AYCell.Value) Then
            AccYr = AYCell.Value
            RunTot = 0
            
            'Loop through Effective Dates
            For Each EDCell In EDRng
                If Not IsDate(EDCell.Value) Then
                    GoTo ContinueLoop
                End If
                
                EffDate = EDCell.Value
                If Not IsDate(EDCell.Offset(-1, 0).Value) Then
                    EDm = DateSerial(Year(EDCell.Value) - 1, 1, 1)
                Else
                    EDm = EDCell.Offset(-1, 0).Value
                End If
                
                If Not IsDate(EDCell.Offset(1, 0).Value) Then
                    EDp = DateSerial(AccYr + 1, 1, 1)
                Else
                    EDp = EDCell.Offset(1, 0).Value
                End If
                
                PAR = 0
                
                'EVALUATE FOR CORRECTNESS, DOES NOT LINE UP WITH EXPECTED VALUE (LOOKING AT DEALING MORE THAN ONE CHANGE FOR GIVEN EFFDATE)
                If EDCell = EDp Then
                    GoTo ContinueLoop
                    
                'Year(EffDate) =< AccYr - 2 --> Skip
                ElseIf Year(EffDate) < AccYr - 2 Then
                    GoTo ContinueLoop
                    
                'Year(EffDate) >= AccYr + 1 --> Skip
                ElseIf Year(EffDate) > AccYr + 1 Then
                    GoTo ContinueLoop
                    
                Else
                    If Year(EffDate) = AccYr And IsNumeric(EDCell.Offset(-1, 0).Value) = False And RunTot = 0 Then
                        EDNew = Month(MonthRound(EffDate))
                        EDOld = 1
                        If IsNumeric(EDCell.Offset(-1, 2).Value) = True Then
                            RI = EDCell.Offset(-1, 2).Value
                        Else
                            RI = 1
                        End If
                        PAR = 72 * RI + (0.5 * ((13 - EDOld) ^ 2 - (13 - EDNew) ^ 2)) * RI + (0.5 * (13 - EDNew) ^ 2) * EDCell.Offset(0, 2)
                    'No AY - 1
                    ElseIf EffDate <= DateSerial(AccYr - 1, 1, 1) And Year(EDp) >= AccYr Then
                        If IsNumeric(EDCell.Offset(0, 2).Value) = True Then
                            RI = EDCell.Offset(0, 2).Value
                        Else
                            RI = 1
                        End If
                        PAR = 72 * RI
                    'EffDate = AY - 1
                    ElseIf Year(EffDate) = AccYr - 1 Then
                        EDNew = Month(MonthRound(EffDate))
                        EDOld = Month(MonthRound(EDm))
                        'AY - 1
                        If Year(EDm) < AccYr - 1 And Year(EDp) = AccYr - 1 Then
                            EDOld = 1
                            If IsNumeric(EDCell.Offset(-1, 2).Value) = True Then
                                RI = EDCell.Offset(-1, 2).Value
                            Else
                                RI = 1
                            End If
                            PAR = (0.5 * ((EDNew - 1) ^ 2 - (EDOld - 1) ^ 2)) * RI
                        ElseIf Year(EDm) < AccYr - 1 And Year(EDp) > AccYr - 1 Then
                            EDOld = 1
                            If IsNumeric(EDCell.Offset(-1, 2).Value) = True Then
                                RI = EDCell.Offset(-1, 2).Value
                            Else
                                RI = 1
                            End If
                            PAR = (0.5 * ((EDNew - 1) ^ 2 - (EDOld - 1) ^ 2)) * RI + (72 - 0.5 * ((EDNew - 1) ^ 2)) * EDCell.Offset(0, 2)
                        ElseIf Year(EDm) = AccYr - 1 And Year(EDp) = AccYr - 1 Then
                            If IsNumeric(EDCell.Offset(-1, 2).Value) = True Then
                                RI = EDCell.Offset(-1, 2).Value
                            Else
                                RI = 1
                            End If
                            PAR = (0.5 * ((EDNew - 1) ^ 2 - (EDOld - 1) ^ 2)) * RI
                        ElseIf Year(EDm) = AccYr - 1 And Year(EDp) > AccYr - 1 Then
                            If IsNumeric(EDCell.Offset(-1, 2).Value) = True Then
                                RI = EDCell.Offset(-1, 2).Value
                            Else
                                RI = 1
                            End If
                            PAR = (0.5 * ((EDNew - 1) ^ 2 - (EDOld - 1) ^ 2)) * RI + (72 - 0.5 * ((EDNew - 1) ^ 2)) * EDCell.Offset(0, 2)
                        End If
                    'No AY
                    ElseIf Year(EDm) < AccYr And Year(EffDate) > AccYr Then
                        If IsNumeric(EDCell.Offset(-1, 2).Value) = True Then
                            RI = EDCell.Offset(-1, 2).Value
                        Else
                            RI = 1
                        End If
                        PAR = 72 * RI
                            
                    'EffDate = AY
                    ElseIf Year(EffDate) = AccYr Then
                        EDNew = Month(MonthRound(EffDate))
                        EDOld = Month(MonthRound(EDm))

                        'AY
                        If Year(EDm) < AccYr And Year(EDp) = AccYr Then
                            EDOld = 1
                            If IsNumeric(EDCell.Offset(-1, 2).Value) = True Then
                                RI = EDCell.Offset(-1, 2).Value
                            Else
                                RI = 1
                            End If
                            PAR = (0.5 * ((13 - EDOld) ^ 2 - (13 - EDNew) ^ 2)) * RI
                        ElseIf Year(EDm) < AccYr And Year(EDp) > AccYr Then
                            EDOld = 1
                            If IsNumeric(EDCell.Offset(-1, 2).Value) = True Then
                                RI = EDCell.Offset(-1, 2).Value
                            Else
                                RI = 1
                            End If
                            PAR = (0.5 * ((13 - EDOld) ^ 2 - (13 - EDNew) ^ 2)) * RI + (0.5 * (13 - EDNew) ^ 2) * EDCell.Offset(0, 2)
                        ElseIf Year(EDm) = AccYr And Year(EDp) = AccYr Then
                            If IsNumeric(EDCell.Offset(-1, 2).Value) = True Then
                                RI = EDCell.Offset(-1, 2).Value
                            Else
                                RI = 1
                            End If
                            PAR = (0.5 * ((13 - EDOld) ^ 2 - (13 - EDNew) ^ 2)) * RI
                        ElseIf Year(EDm) = AccYr And Year(EDp) > AccYr Then
                            If IsNumeric(EDCell.Offset(-1, 2).Value) = True Then
                                RI = EDCell.Offset(-1, 2).Value
                            Else
                                RI = 1
                            End If
                            PAR = (0.5 * ((13 - EDOld) ^ 2 - (13 - EDNew) ^ 2)) * RI + (0.5 * (13 - EDNew) ^ 2) * EDCell.Offset(0, 2)
                        End If
                    
                End If
              End If
                    'Keep running total of partial areas x RI
                    RunTot = RunTot + PAR
ContinueLoop:
            Next EDCell
            
                CumlRI = RunTot / 144
            
            AYCell.Offset(0, 1).Value = CumlRI
            
        End If
        
    Next AYCell
    
End Sub
                   

