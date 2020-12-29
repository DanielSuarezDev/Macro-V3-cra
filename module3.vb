Sub MarcaBajasRetanqueos()
Dim LARGOCancel As Long
    MismoMes = Sheets("CONTABILIZADOS").Cells(3, 2)
    MesSiguiente = DateAdd("M", 1, MismoMes)
    MesSubsiguiente = DateAdd("m", 2, MismoMes)
    LARGOCancel = Sheets("CANCELADOS").Range("A" & Rows.Count).End(xlUp).Row
    With Sheets("CANCELADOS")
    For i = 2 To LARGOCancel
            If Not IsError(Application.VLookup(.Cells(i, 1), Sheets("ACTIVOS").Range("A:AY"), 37, 0)) Then
                FECHA = CDate(Application.VLookup(.Cells(i, 1), Sheets("ACTIVOS").Range("A:AY"), 37, 0))
                Select Case Sheets("CONTABILIZADOS").Cells(1, 2).Value
                 Case Is = "MISMO MES"
                    If Month(MismoMes) = Month(FECHA) And Year(MismoMes) = Year(FECHA) Then
                      'If (CDate(Application.VLookup(.Cells(i, 1), Sheets("ACTIVOS").Range("A:AY"), 15, 0)) - CDate(.Cells(i, 33))) < 15 Or (CDate(Application.VLookup(.Cells(i, 1), Sheets("ACTIVOS").Range("A:AY"), 15, 0)) - .Cells(i, 33)) > -15 Then
                      'se cambio los retanqueos ahora retanqueo es lo que sea igual a 0
                      If (CDate(Application.VLookup(.Cells(i, 1), Sheets("ACTIVOS").Range("A:AY"), 15, 0)) - CDate(.Cells(i, 33))) = 0 Then
                            .Cells(i, 49) = "RETANQUEO"
                      End If
                    End If
                 Case Is = "MES SIGUIENTE"
                     If Month(MesSiguiente) = Month(FECHA) And Year(MesSiguiente) = Year(FECHA) Then
                     'If (CDate(Application.VLookup(.Cells(i, 1), Sheets("ACTIVOS").Range("A:AY"), 15, 0)) - CDate(.Cells(i, 33))) < 15 Or (CDate(Application.VLookup(.Cells(i, 1), Sheets("ACTIVOS").Range("A:AY"), 15, 0)) - .Cells(i, 33)) > -15 Then
                     If (CDate(Application.VLookup(.Cells(i, 1), Sheets("ACTIVOS").Range("A:AY"), 15, 0)) - CDate(.Cells(i, 33))) = 0 Then
                        .Cells(i, 49) = "RETANQUEO"
                     End If
                    End If
                 Case Is = "MES SUBSIGUIENTE"
                    If Month(MesSubsiguiente) = Month(FECHA) And Year(MesSubsiguiente) = Year(FECHA) Then
                    'If (CDate(Application.VLookup(.Cells(i, 1), Sheets("ACTIVOS").Range("A:AY"), 15, 0)) - CDate(.Cells(i, 33))) < 15 Or (CDate(Application.VLookup(.Cells(i, 1), Sheets("ACTIVOS").Range("A:AY"), 15, 0)) - .Cells(i, 33)) > -15 Then
                    If (CDate(Application.VLookup(.Cells(i, 1), Sheets("ACTIVOS").Range("A:AY"), 15, 0)) - CDate(.Cells(i, 33))) = 0 Then
                           .Cells(i, 49) = "RETANQUEO"
                    End If
                    End If
                End Select
            End If
    Next i
    End With
End Sub