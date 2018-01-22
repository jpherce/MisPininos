Attribute VB_Name = "Modulo2"
' Ultima modificación: 18-nov-17
' Simplificación del ciclo de FCH y FCR
' Formato condicional en vacas DNB 18.11.17
' Mod calculo de EM305d, visualización de columnas 16 y 17
' Mod forma presentación edad becerras
Option Explicit
Dim rCelda As Object
Dim rCelda1 As Object
Dim bActivarRutina As Boolean
Dim scampo, sHoja, sTabla As String
Dim sOrder As String
' Subrutinas utilizadas frecuentemente

Private Sub ADEL()
' Calcula los Días en Leche, Fecha x Secar y Fecha x Parir
' Trabaja en hoja HATO
    Dim iEM305 As Double
    Dim lCounter, lTotal As Long
    If Range("Desarrollador!B20") = UCase("T") Then _
      Exit Sub
    Application.ScreenUpdating = False
    Desproteger
    bActivarRutina = True
    Application.DisplayStatusBar = True
    lTotal = Range("Tabla1").Rows.Count
    For Each rCelda In Range("Tabla1[Arete]")
        Application.StatusBar = _
          "Actualizando Fechas... " & _
          Format(lCounter / lTotal, "0%")
        ' Calculando días en leche
        If Not IsDate(rCelda.Offset(0, 5)) Then _
          GoTo 4321
        If Not IsEmpty(rCelda.Offset(0, 5)) And _
          Not rCelda.Offset(0, 11) = "**SECA**" Then
                If IsDate(rCelda.Offset(0, 5)) Then _
                  rCelda.Offset(0, 3) = _
                  Date - CDate(rCelda.Offset(0, 5))
            Else
4321:
                rCelda.Offset(0, 3).Clear
        End If
        ' Calculando días de servicio
        If Not IsDate(CDate(rCelda.Offset(0, 7))) And _
          Not IsEmpty(rCelda.Offset(0, 7)) Then
            rCelda.Offset(0, 11).Clear 'FxSecar
            rCelda.Offset(0, 12).Clear 'FxParir
            'rCelda.Offset(0, 16).Clear
        End If
        If (IsDate(rCelda.Offset(0, 7))) And _
          rCelda.Offset(0, 8) _
          <> UCase("Calor") Then
            If rCelda.Offset(0, 10) = "P" Then
                    If Not rCelda.Offset(0, 11) = "**SECA**" Then
                        With rCelda.Offset(0, 11)   'FxSecar
                            .Value = rCelda.Offset(0, 7) + 213
                            .NumberFormat = "d-mmm-yy"
                        End With
                        With rCelda.Offset(0, 12)   'FxParir
                            .Value = rCelda.Offset(0, 7) + 273
                            .NumberFormat = "d-mmm-yy"
                        End With
                    End If
                Else
                    If Not rCelda.Offset(0, 11) = "**SECA**" Then _
                        rCelda.Offset(0, 11).Clear 'FxSecar
                    rCelda.Offset(0, 12).Clear 'FxParir
            End If
        End If
        ' Proyección a 305 días
        If CBool(Range("Configuracion!B65")) Then _
          Sheets("Hato").Columns("P").Hidden = False Else _
          Sheets("Hato").Columns("P").Hidden = True
        ' Calculando ValorRelativo
        If CBool(Range("Configuracion!B66")) Then _
          Sheets("Hato").Columns("Q").Hidden = False Else _
          Sheets("Hato").Columns("Q").Hidden = True
        If CBool(Range("Configuracion!B66")) And Not IsEmpty(rCelda.Offset(0, 15)) And _
          IsNumeric(rCelda.Offset(0, 15)) Then
            iEM305 = 1
            ' Calcular Valor Relativo en relación al Eq. Madurez
            If Not CBool(Range("Configuracion!B67")) Then GoTo 9753
            Select Case rCelda.Offset(0, 4)
                Case Is = 1
                    iEM305 = Range("Configuracion!L3")
                Case Is = 2
                    iEM305 = Range("Configuracion!L4")
                Case Is >= 3
                    iEM305 = Range("Configuracion!L5")
            End Select
9753:
                rCelda.Offset(0, 16) = Int(((Val(rCelda.Offset(0, 15)) * iEM305) / _
                  pProy305d()) * 100)
            Else
                rCelda.Offset(0, 16).Clear
        End If
        lCounter = lCounter + 1
    Next rCelda
    bActivarRutina = False
    If Not CBool(Range("Configuracion!C39")) Then _
      Proteger
    Application.StatusBar = False
    Application.ScreenUpdating = True
End Sub

Private Sub ADEL1()
' Calcula los Días en Leche, Fecha x Secar y Fecha x Parir
' rCelda.Offset(0,3),rCelda.Offset(0,11),rCelda.Offset(0,12)
    Dim iEM305 As Double
    bActivarRutina = True
        If Not IsDate(rCelda.Offset(0, 5)) Then _
          GoTo 4321
        If Not IsEmpty(rCelda.Offset(0, 5)) And _
          Not rCelda.Offset(0, 11) = "**SECA**" Then
                If IsDate(rCelda.Offset(0, 5)) Then _
                  rCelda.Offset(0, 3) = _
                  Date - CDate(rCelda.Offset(0, 5))
            Else
4321:
                rCelda.Offset(0, 3).Clear
        End If
        ' Calculando días de servicio
        If Not IsDate(CDate(rCelda.Offset(0, 7))) And _
          Not IsEmpty(rCelda.Offset(0, 7)) Then
            rCelda.Offset(0, 11).Clear 'FxSecar
            rCelda.Offset(0, 12).Clear 'FxParir
            'rCelda.Offset(0, 16).Clear
        End If
        If (IsDate(rCelda.Offset(0, 7))) And _
          rCelda.Offset(0, 8) <> UCase("Calor") _
          Then
            If rCelda.Offset(0, 10) = "P" Then
                    If Not rCelda.Offset(0, 11) = "**SECA**" Then
                        With rCelda.Offset(0, 11)   'FxSecar
                            .Value = rCelda.Offset(0, 7) + 213
                            .NumberFormat = "d-mmm-yy"
                        End With
                        With rCelda.Offset(0, 12)   'FxParir
                            .Value = rCelda.Offset(0, 7) + 273
                            .NumberFormat = "d-mmm-yy"
                        End With
                    End If
                Else
                    If Not rCelda.Offset(0, 11) = "**SECA**" Then _
                        rCelda.Offset(0, 11).Clear 'FxSecar
                    rCelda.Offset(0, 12).Clear 'FxParir
            End If
        End If
        ' Proyección a 305 días
        If CBool(Range("Configuracion!B65")) Then _
          Sheets("Hato").Columns("P").Hidden = False Else _
          Sheets("Hato").Columns("P").Hidden = True
        ' Calculando ValorRelativo
        If CBool(Range("Configuracion!B66")) Then _
          Sheets("Hato").Columns("Q").Hidden = False Else _
          Sheets("Hato").Columns("Q").Hidden = True
        If CBool(Range("Configuracion!B66")) And Not IsEmpty(rCelda.Offset(0, 15)) And _
          IsNumeric(rCelda.Offset(0, 15)) Then
            iEM305 = 1
            ' Calcular Valor Relativo en relación al Eq. Madurez
            If Not CBool(Range("Configuracion!B67")) Then GoTo 9753
            Select Case rCelda.Offset(0, 4)
                Case Is = 1
                    iEM305 = Range("Configuracion!L3")
                Case Is = 2
                    iEM305 = Range("Configuracion!L4")
                Case Is >= 3
                    iEM305 = Range("Configuracion!L5")
            End Select
9753:
                rCelda.Offset(0, 16) = Int(((Val(rCelda.Offset(0, 15)) * iEM305) / _
                  pProy305d()) * 100)
            Else
                rCelda.Offset(0, 16).Clear
        End If
    bActivarRutina = False
End Sub

Private Sub AE()
' Calcula la Edad de los animales
' Trabaja en Hoja Reemplazos
    Dim rCelda As Range
    Dim lCounter, lTotal As Long
    Dim sN, sD As Long
    If Range("Desarrollador!B20") = UCase("T") Then _
      Exit Sub
    Application.ScreenUpdating = False
    Desproteger
    bActivarRutina = True
    Application.DisplayStatusBar = True
    lTotal = Range("Tabla2").Rows.Count
    For Each rCelda In Range("Tabla2[Arete]")
        Application.StatusBar = _
          "Actualizando Fechas... " & _
          Format(lCounter / lTotal, "0%")
        If Not IsEmpty(rCelda.Offset(0, 4)) Then
                On Error GoTo 2201
                
                sN = Date - CDate(rCelda.Offset(0, 4))
                sD = 365
                rCelda.Offset(0, 3) = Format(Int(sN / sD), "0") & "-" & _
                  Format(Int((sN - sD * Int(sN / sD)) / 30.4), "00")
                
                On Error GoTo 0
            Else
2201:
                rCelda.Offset(0, 3).Clear 'Si F.Nacim en blanco
        End If
        If Not IsDate(CDate(rCelda.Offset(0, 6))) Or _
          IsEmpty(rCelda.Offset(0, 6)) Then 'F.Servicio
            rCelda.Offset(0, 10).Clear 'FxParir
            'rCelda.Offset(0, 23).Clear
        End If
        If (IsDate(rCelda.Offset(0, 6))) Then
            If rCelda.Offset(0, 9) = "P" Then
                With rCelda.Offset(0, 10)   'FxParir
                    .Value = rCelda.Offset(0, 6) + 273
                    .NumberFormat = "d-mmm-yy"
                End With
            End If
        End If
        lCounter = lCounter + 1
    Next rCelda
    bActivarRutina = False
    If Not CBool(Range("Configuracion!C39")) Then _
      Proteger
    Application.StatusBar = False
    Application.ScreenUpdating = True
End Sub

Private Sub AE1()
' Calcula la Edad de los animales
' Trabaja en Hoja Reemplazos
    Dim sN, sD As Long
    bActivarRutina = True
    If Not IsEmpty(rCelda.Offset(0, 4)) Then
            On Error GoTo 2201
            
            sN = Date - CDate(rCelda.Offset(0, 4))
            sD = 365
            rCelda.Offset(0, 3) = Format(Int(sN / sD), "0") & "-" & _
              Format(Int((sN - sD * Int(sN / sD)) / 30.4), "00")
            
            On Error GoTo 0
        Else
2201:
            rCelda.Offset(0, 3).Clear 'Si F.Nacim en blanco
    End If
    If Not IsDate(CDate(rCelda.Offset(0, 6))) Or _
      IsEmpty(rCelda.Offset(0, 6)) Then 'F.Servicio
        rCelda.Offset(0, 10).Clear 'FxParir
        'rCelda.Offset(0, 23).Clear
    End If
    If (IsDate(rCelda.Offset(0, 6))) Then
        If rCelda.Offset(0, 9) = "P" Then
            With rCelda.Offset(0, 10)   'FxParir
                .Value = rCelda.Offset(0, 6) + 273
                .NumberFormat = "d-mmm-yy"
            End With
        End If
    End If
End Sub

Private Sub BorrarRenglon()
    Dim mRenglon As Long
    Selection.ListObject.ListRows(ActiveCell.Row).Delete
End Sub

Private Sub CerrarTodo()
' Oculta todas las hojas del archivo
    Dim i As Long
    On Error Resume Next
    Sheets("Inicio").Visible = True
    For i = 1 To ThisWorkbook.Sheets.Count
        
        Sheets(i).Visible = xlVeryHidden
        If Sheets(i).Name = "Inicio" Then _
          Sheets(i).Visible = True
        'If Sheets(i).Name = "Eventos" Then _
          Sheets(i).Visible = True
    Next i
    On Error GoTo 0
End Sub

Private Sub Desproteger()
' Desprotege la hoja activa
    If ActiveSheet.ProtectContents Then _
      ActiveSheet.Unprotect Password:="0246813579"
End Sub

Private Sub FCH()
'Formato Condicional Hato
' Aplica formatos condicionales a ciertas celdas de una Tabla
    'Dim rCelda As Range
    Dim lCounter, lTotal As Long
    If Range("Desarrollador!B20") = UCase("T") Then _
      Exit Sub
    If bActivarRutina = True Then _
      Exit Sub
    Application.ScreenUpdating = False
    Application.DisplayStatusBar = True
    lTotal = Range("Tabla1").Rows.Count
    Desproteger 'Módulo2
   
   For Each rCelda In Range("Tabla1[Arete]")
        Application.StatusBar = _
          "Comprobando contenido en celdas... " & _
          Format(lCounter / lTotal, "0%")
        'Reduce un ciclo completo de revisión
        ADEL1
    'Arete
        Set rCelda1 = rCelda.Offset(0, 0)
        FormatoEstandard
        If IsEmpty(rCelda.Offset(0, 0)) Then _
          GoTo ColA
        ' Arete Repetido
        If WorksheetFunction. _
          CountIf(Range("Tabla1[Arete]"), _
          rCelda.Offset(0, 0)) > 1 Then _
          FormatoErrorEntrada
        If WorksheetFunction. _
          CountIf(Range("Tabla2[Arete]"), _
          rCelda.Offset(0, 0)) > 0 Then _
          FormatoErrorEntrada
ColA:
        ' Arete en Blanco
        If IsEmpty(rCelda.Offset(0, 0)) Then _
          FormatoErrorFaltante
    
    'Corral
        Set rCelda1 = rCelda.Offset(0, 1)
        FormatoEstandard
        If IsEmpty(rCelda.Offset(0, 1)) Then _
          GoTo ColB
        ' Vaca en Producción en el Corral de Secas
        If rCelda.Offset(0, 1) >= Range("Configuracion!C9") And _
          Not rCelda.Offset(0, 11) = "**SECA**" Then _
            FormatoErrorManejo
        ' Vaca Seca en Corral de Produccion
        If rCelda.Offset(0, 1) < Range("Configuracion!C9") And _
          rCelda.Offset(0, 11) = "**SECA**" Then _
          FormatoErrorManejo
ColB:
        ' Corral en Blanco
        If IsEmpty(rCelda.Offset(0, 1)) And _
          IsDate(rCelda.Offset(0, 5)) Then _
          FormatoErrorFaltante
    
    ' Produccion
        Set rCelda1 = rCelda.Offset(0, 2)
        FormatoEstandard
        If IsEmpty(rCelda.Offset(0, 2)) Then _
          GoTo ColC
        ' Vaca Seca
        If IsNumeric(rCelda.Offset(0, 2)) And _
          rCelda.Offset(0, 11) = "**SECA**" Then _
          FormatoErrorManejo
        ' Produccion mínima
        If rCelda.Offset(0, 2) < _
          Range("Configuracion!C24") And _
          rCelda.Offset(0, 3) > 90 Then _
          FormatoErrorManejo
        ' No numéro
        If Not IsNumeric(rCelda.Offset(0, 2)) Then _
          FormatoErrorEntrada
ColC:
        
        ' En Blanco
        If rCelda.Offset(0, 2) = vbNullString And _
          rCelda.Offset(0, 1) < _
          Range("Configuracion!C9") And _
          IsDate(rCelda.Offset(0, 5)) Then _
          FormatoErrorFaltante
ColD:
    ' DEL
        Set rCelda1 = rCelda.Offset(0, 3)
        FormatoEstandard
        ' Vacas atrasadas
        If rCelda.Offset(0, 3) > _
          Range("Configuracion!C6") And _
          IsEmpty(rCelda.Offset(0, 7)) And _
          Not rCelda.Offset(0, 13) = "DNB" Then _
          FormatoErrorManejo
    
    'Parto
        Set rCelda1 = rCelda.Offset(0, 4)
        FormatoEstandard
        If IsEmpty(rCelda.Offset(0, 4)) Then _
          GoTo ColE
         ' No es numero
        If Not IsNumeric(rCelda.Offset(0, 4)) Then _
          FormatoErrorEntrada
ColE:
        ' En Blanco
        If IsEmpty(rCelda.Offset(0, 4)) And _
          Not IsEmpty(rCelda.Offset(0, 5)) Then _
          FormatoErrorFaltante
    
    'Fecha Parto
        Set rCelda1 = rCelda.Offset(0, 5)
        FormatoEstandard
        If IsEmpty(rCelda.Offset(0, 5)) Then _
          GoTo ColF
        ' Fecha Futura
        On Error Resume Next
        If CDate(rCelda.Offset(0, 5)) > Date Then _
          FormatoErrorEntrada
        On Error GoTo 0
        ' No Fecha
        If Not IsDate(rCelda.Offset(0, 5)) And _
          IsEmpty(rCelda.Offset(0, 5)) Then _
          FormatoErrorEntrada
ColF:
        ' En Blanco
        If IsEmpty(rCelda.Offset(0, 5)) And _
          (Not IsEmpty(rCelda.Offset(0, 4)) _
          Or Not IsEmpty(rCelda.Offset(0, 1))) Then _
          FormatoErrorFaltante
    
    'Servicio
        Set rCelda1 = rCelda.Offset(0, 6)
        FormatoEstandard
        If IsEmpty(rCelda.Offset(0, 6)) Or _
          rCelda.Offset(0, 13) = "DNB" Then _
          GoTo ColG
        ' Repetidora
        If rCelda.Offset(0, 6) > 3 And _
          Not rCelda.Offset(0, 10) = "P" Then _
          FormatoErrorManejo
        ' No es numero
        If Not IsNumeric(rCelda.Offset(0, 6)) Then _
        FormatoErrorEntrada
ColG:
        ' En Blanco
        If IsEmpty(rCelda.Offset(0, 6)) And _
          Not IsEmpty(rCelda.Offset(0, 7)) Then _
          FormatoErrorFaltante
        
        ' Calor
        If rCelda.Offset(0, 8) = "Calor" Then _
          FormatoEstandard
    
    ' Fecha Servicio
        Set rCelda1 = rCelda.Offset(0, 7)
        FormatoEstandard
        If IsEmpty(rCelda.Offset(0, 7)) Or _
         rCelda.Offset(0, 13) = "DNB" Then _
          GoTo ColH
        ' Fecha Futura
        On Error Resume Next
        If CDate(rCelda.Offset(0, 7)) > Date Then _
          FormatoErrorEntrada
        On Error GoTo 0
        ' Falta Dx
        On Error Resume Next
        If IsDate(rCelda.Offset(0, 7)) And _
          Date - CDate(rCelda.Offset(0, 7)) > _
          Range("Configuracion!C5") And _
          Not rCelda.Offset(0, 10) = "P" Then _
          FormatoErrorManejo
        ' Fecha Anterior al Parto
        If IsDate(rCelda.Offset(0, 7)) And _
          CDate(rCelda.Offset(0, 7)) <= _
          CDate(rCelda.Offset(0, 5)) Then _
          FormatoErrorEntrada
        On Error GoTo 0
        ' No Fecha
        If Not IsDate(rCelda.Offset(0, 7)) And _
          Not IsEmpty(rCelda.Offset(0, 7)) Then _
          FormatoErrorEntrada
ColH:
        ' En Blanco
        If IsEmpty(rCelda.Offset(0, 7)) And _
          Not IsEmpty(rCelda.Offset(0, 6)) Then _
          FormatoErrorFaltante
        
    ' Semental
        Set rCelda1 = rCelda.Offset(0, 8)
        FormatoEstandard
ColI:
        ' En Blanco
        If IsEmpty(rCelda.Offset(0, 8)) And _
          CBool(Range("Configuracion!C15")) And _
          Not IsEmpty(rCelda.Offset(0, 7)) Then _
          FormatoErrorFaltante

    ' Técnico
        Set rCelda1 = rCelda.Offset(0, 9)
        FormatoEstandard
ColJ:
        ' En Blanco
        If IsEmpty(rCelda.Offset(0, 9)) And _
          CBool(Range("Configuracion!C16")) And _
          Not IsEmpty(rCelda.Offset(0, 7)) Then _
          FormatoErrorFaltante
    
    ' Status
        Set rCelda1 = rCelda.Offset(0, 10)
        FormatoEstandard
        ' contenido diferente a G o V
        If (Not IsEmpty(rCelda.Offset(0, 10)) And _
          Not rCelda.Offset(0, 10) = "P") And _
          (Not IsEmpty(rCelda.Offset(0, 10)) And _
          Not rCelda.Offset(0, 10) = "O") Then _
          FormatoErrorEntrada
          
     ' FxSecar
        Set rCelda1 = rCelda.Offset(0, 11)
        FormatoEstandard
        If IsEmpty(rCelda.Offset(0, 11)) Then _
          GoTo ColL
        ' Fecha pasada
        On Error Resume Next
        If Date >= CDate(rCelda.Offset(0, 11)) Then _
          FormatoErrorManejo
        On Error GoTo 0
        ' Vaca Seca
        If rCelda.Offset(0, 11) = "**SECA**" Then _
          FormatoEstandard
ColL:
        ' En blanco
        If IsEmpty(rCelda.Offset(0, 11)) And _
          rCelda.Offset(0, 10) = "P" Then _
          FormatoErrorFaltante
        
    ' FxParir
        Set rCelda1 = rCelda.Offset(0, 12)
        FormatoEstandard
        If IsEmpty(rCelda.Offset(0, 12)) Then _
          GoTo ColM
        ' Fecha pasada
        On Error Resume Next
        If Date >= CDate(rCelda.Offset(0, 12)) Then _
          FormatoErrorManejo
        On Error GoTo 0
ColM:
        ' En blanco
        If IsEmpty(rCelda.Offset(0, 12)) And _
          rCelda.Offset(0, 10) = "P" Then _
          FormatoErrorFaltante
          
    ' Proy305d
        Set rCelda1 = rCelda.Offset(0, 15)
        FormatoEstandard
ColP:
        lCounter = lCounter + 1
    Next rCelda
    Proteger
    'Range("Desarrollador!B20") = "T"
    Application.StatusBar = False
    Application.ScreenUpdating = True
End Sub


Private Sub FCH1()
'Formato Condicional Hato
' Aplica formatos condicionales a ciertas celdas de una Tabla
    'Dim rCelda As Range
    Dim lCounter, lTotal As Long
    If Range("Desarrollador!B20") = UCase("T") Then _
      Exit Sub
    If bActivarRutina = True Then _
      Exit Sub
    Application.ScreenUpdating = False
    Application.DisplayStatusBar = True
    lTotal = Range("Tabla1").Rows.Count
    Desproteger 'Módulo2
   
   For Each rCelda In Range("Tabla1[Arete]")
        Application.StatusBar = _
          "Comprobando contenido en celdas... " & _
          Format(lCounter / lTotal, "0%")
        
    'Arete
        Set rCelda1 = rCelda.Offset(0, 0)
        FormatoEstandard
        If IsEmpty(rCelda.Offset(0, 0)) Then _
          GoTo ColA
        ' Arete Repetido
        If WorksheetFunction. _
          CountIf(Range("Tabla1[Arete]"), _
          rCelda.Offset(0, 0)) > 1 Then _
          FormatoErrorEntrada
        If WorksheetFunction. _
          CountIf(Range("Tabla2[Arete]"), _
          rCelda.Offset(0, 0)) > 0 Then _
          FormatoErrorEntrada
ColA:
        ' Arete en Blanco
        If IsEmpty(rCelda.Offset(0, 0)) Then _
          FormatoErrorFaltante
    
    'Corral
        Set rCelda1 = rCelda.Offset(0, 1)
        FormatoEstandard
        If IsEmpty(rCelda.Offset(0, 1)) Then _
          GoTo ColB
        ' Vaca en Producción en el Corral de Secas
        If rCelda.Offset(0, 1) >= Range("Configuracion!C9") And _
          Not rCelda.Offset(0, 11) = "**SECA**" Then _
            FormatoErrorManejo
        ' Vaca Seca en Corral de Produccion
        If rCelda.Offset(0, 1) < Range("Configuracion!C9") And _
          rCelda.Offset(0, 11) = "**SECA**" Then _
          FormatoErrorManejo
ColB:
        ' Corral en Blanco
        If IsEmpty(rCelda.Offset(0, 1)) And _
          IsDate(rCelda.Offset(0, 5)) Then _
          FormatoErrorFaltante
    
    ' Produccion
        Set rCelda1 = rCelda.Offset(0, 2)
        FormatoEstandard
        If IsEmpty(rCelda.Offset(0, 2)) Then _
          GoTo ColC
        ' Vaca Seca
        If IsNumeric(rCelda.Offset(0, 2)) And _
          rCelda.Offset(0, 11) = "**SECA**" Then _
          FormatoErrorManejo
        ' Produccion mínima
        If rCelda.Offset(0, 2) < _
          Range("Configuracion!C24") And _
          rCelda.Offset(0, 3) > 90 Then _
          FormatoErrorManejo
        ' No numéro
        If Not IsNumeric(rCelda.Offset(0, 2)) Then _
          FormatoErrorEntrada
ColC:
        ' En Blanco
        If rCelda.Offset(0, 2) = vbNullString And _
          rCelda.Offset(0, 1) < _
          Range("Configuracion!C9") And _
          IsDate(rCelda.Offset(0, 5)) Then _
          FormatoErrorFaltante
    
    ' DEL
        Set rCelda1 = rCelda.Offset(0, 3)
        FormatoEstandard
        ' Vacas atrasadas
        If rCelda.Offset(0, 3) > _
          Range("Configuracion!C6") And _
          IsEmpty(rCelda.Offset(0, 7)) And _
          Not rCelda.Offset(0, 13) = "DNB" Then _
          FormatoErrorManejo
    
    'Parto
        Set rCelda1 = rCelda.Offset(0, 4)
        FormatoEstandard
        If IsEmpty(rCelda.Offset(0, 4)) Then _
          GoTo ColE
         ' No es numero
        If Not IsNumeric(rCelda.Offset(0, 4)) Then _
          FormatoErrorEntrada
ColE:
        ' En Blanco
        If IsEmpty(rCelda.Offset(0, 4)) And _
          Not IsEmpty(rCelda.Offset(0, 5)) Then _
          FormatoErrorFaltante
    
    'Fecha Parto
        Set rCelda1 = rCelda.Offset(0, 5)
        FormatoEstandard
        If IsEmpty(rCelda.Offset(0, 5)) Then _
          GoTo ColF
        ' Fecha Futura
        On Error Resume Next
        If CDate(rCelda.Offset(0, 5)) > Date Then _
          FormatoErrorEntrada
        On Error GoTo 0
        ' No Fecha
        If Not IsDate(rCelda.Offset(0, 5)) And _
          IsEmpty(rCelda.Offset(0, 5)) Then _
          FormatoErrorEntrada
ColF:
        ' En Blanco
        If IsEmpty(rCelda.Offset(0, 5)) And _
          (Not IsEmpty(rCelda.Offset(0, 4)) _
          Or Not IsEmpty(rCelda.Offset(0, 1))) Then _
          FormatoErrorFaltante
    
    'Servicio
        Set rCelda1 = rCelda.Offset(0, 6)
        FormatoEstandard
        If IsEmpty(rCelda.Offset(0, 6)) Or _
          rCelda.Offset(0, 13) = "DNB" Then _
          GoTo ColG
        ' Repetidora
        If rCelda.Offset(0, 6) > 3 And _
          Not rCelda.Offset(0, 10) = "P" Then _
          FormatoErrorManejo
        ' No es numero
        If Not IsNumeric(rCelda.Offset(0, 6)) Then _
        FormatoErrorEntrada
ColG:
        ' En Blanco
        If IsEmpty(rCelda.Offset(0, 6)) And _
          Not IsEmpty(rCelda.Offset(0, 7)) Then _
          FormatoErrorFaltante
        
        ' Calor
        If rCelda.Offset(0, 8) = "Calor" Then _
          FormatoEstandard
    
    ' Fecha Servicio
        Set rCelda1 = rCelda.Offset(0, 7)
        FormatoEstandard
        If IsEmpty(rCelda.Offset(0, 7)) Or _
         rCelda.Offset(0, 13) = "DNB" Then _
          GoTo ColH
        ' Fecha Futura
        On Error Resume Next
        If CDate(rCelda.Offset(0, 7)) > Date Then _
          FormatoErrorEntrada
        On Error GoTo 0
        ' Falta Dx
        On Error Resume Next
        If IsDate(rCelda.Offset(0, 7)) And _
          Date - CDate(rCelda.Offset(0, 7)) > _
          Range("Configuracion!C5") And _
          Not rCelda.Offset(0, 10) = "P" Then _
          FormatoErrorManejo
        'On Error GoTo 0
        ' Fecha Anterior al Parto
        If IsDate(rCelda.Offset(0, 7)) And _
          CDate(rCelda.Offset(0, 7)) <= _
          CDate(rCelda.Offset(0, 5)) Then _
          FormatoErrorEntrada
        On Error GoTo 0
        ' No Fecha
        If Not IsDate(rCelda.Offset(0, 7)) And _
          Not IsEmpty(rCelda.Offset(0, 7)) Then _
          FormatoErrorEntrada
ColH:
        ' En Blanco
        If IsEmpty(rCelda.Offset(0, 7)) And _
          Not IsEmpty(rCelda.Offset(0, 6)) Then _
          FormatoErrorFaltante
        
    ' Semental
        Set rCelda1 = rCelda.Offset(0, 8)
        FormatoEstandard
ColI:
        ' En Blanco
        If IsEmpty(rCelda.Offset(0, 8)) And _
          CBool(Range("Configuracion!C15")) And _
          Not IsEmpty(rCelda.Offset(0, 7)) Then _
          FormatoErrorFaltante

    ' Técnico
        Set rCelda1 = rCelda.Offset(0, 9)
        FormatoEstandard
ColJ:
        ' En Blanco
        If IsEmpty(rCelda.Offset(0, 9)) And _
          CBool(Range("Configuracion!C16")) And _
          Not IsEmpty(rCelda.Offset(0, 7)) Then _
          FormatoErrorFaltante
    
    ' Status
        Set rCelda1 = rCelda.Offset(0, 10)
        FormatoEstandard
        ' contenido diferente a G o V
        If (Not IsEmpty(rCelda.Offset(0, 10)) And _
          Not rCelda.Offset(0, 10) = "P") And _
          (Not IsEmpty(rCelda.Offset(0, 10)) And _
          Not rCelda.Offset(0, 10) = "O") Then _
          FormatoErrorEntrada
          
     ' FxSecar
        Set rCelda1 = rCelda.Offset(0, 11)
        FormatoEstandard
        If IsEmpty(rCelda.Offset(0, 11)) Then _
          GoTo ColL
        ' Fecha pasada
        On Error Resume Next
        If Date >= CDate(rCelda.Offset(0, 11)) Then _
          FormatoErrorManejo
        On Error GoTo 0
        ' Vaca Seca
        If rCelda.Offset(0, 11) = "**SECA**" Then _
          FormatoEstandard
ColL:
        ' En blanco
        If IsEmpty(rCelda.Offset(0, 11)) And _
          rCelda.Offset(0, 10) = "P" Then _
          FormatoErrorFaltante
        
    ' FxParir
        Set rCelda1 = rCelda.Offset(0, 12)
        FormatoEstandard
        If IsEmpty(rCelda.Offset(0, 12)) Then _
          GoTo ColM
        ' Fecha pasada
        On Error Resume Next
        If Date >= CDate(rCelda.Offset(0, 12)) Then _
          FormatoErrorManejo
        On Error GoTo 0
ColM:
        ' En blanco
        If IsEmpty(rCelda.Offset(0, 12)) And _
          rCelda.Offset(0, 10) = "P" Then _
          FormatoErrorFaltante
          
    ' Proy305d
        Set rCelda1 = rCelda.Offset(0, 15)
        FormatoEstandard
ColP:
        lCounter = lCounter + 1
    Next rCelda
    Proteger
    'Range("Desarrollador!B20") = "T"
    Application.StatusBar = False
    Application.ScreenUpdating = True
End Sub

Private Sub FCR()
' Formato Condicional Reemplazos
' Aplica formatos condicionales a ciertas celdas de una Tabla
    Dim lCounter, lTotal As Long
    If Range("Desarrollador!B20") = UCase("T") Then _
      Exit Sub
    If bActivarRutina = True Then _
      Exit Sub
    Application.ScreenUpdating = False
    Application.DisplayStatusBar = True
    lTotal = Range("Tabla2").Rows.Count
    Desproteger 'Módulo2
    
    For Each rCelda In Range("Tabla2[Arete]")
        Application.StatusBar = _
        "Comprobando contenido en celdas... " & _
        Format(lCounter / lTotal, "0%")
        AE1
    'Arete
        Set rCelda1 = rCelda.Offset(0, 0)
        FormatoEstandard
        If IsEmpty(rCelda.Offset(0, 0)) Then _
          GoTo ColA
        ' Arete Repetido
        If WorksheetFunction. _
          CountIf(Range("Tabla2[Arete]"), _
          rCelda.Offset(0, 0)) > 1 Then _
          FormatoErrorEntrada
        If WorksheetFunction. _
          CountIf(Range("Tabla1[Arete]"), _
            rCelda.Offset(0, 0)) > 0 Then _
            FormatoErrorEntrada
ColA:
        ' Arete en Blanco
        If IsEmpty(rCelda.Offset(0, 0)) Then _
          FormatoErrorFaltante
    
    'Corral
        Set rCelda1 = rCelda.Offset(0, 1)
        FormatoEstandard
        If IsEmpty(rCelda.Offset(0, 1)) Then _
          GoTo ColB
        ' Vaquilla en corral de lactancia
        On Error Resume Next
        If rCelda.Offset(0, 1) = _
          Range("Configuracion!C13") And _
          (Date - CDate(rCelda.Offset(0, 4))) > _
          Range("Configuracion!C34") Then _
          FormatoErrorManejo
        On Error GoTo 0
        ' Vaca No en corral de Preparación
        If IsDate(rCelda.Offset(0, 10)) And _
          (Date - CDate(rCelda.Offset(0, 10)) < 30) And _
          rCelda.Offset(0, 1) <> _
          Range("Configuracion!C10") Then _
          FormatoErrorManejo
ColB:
        ' Corral en Blanco
        If IsEmpty(rCelda.Offset(0, 1)) And _
          IsDate(rCelda.Offset(0, 4)) Then _
          FormatoErrorFaltante
    
    ' PesoCorporal
        Set rCelda1 = rCelda.Offset(0, 2)
        FormatoEstandard
        ' Prod. en Blanco
        Sheets("Reemplazos").Columns("C").Hidden = True
        If CBool(Range("Configuracion!C35")) Then
            Sheets("Reemplazos").Columns("C").Hidden = False
            If IsEmpty(rCelda.Offset(0, 2)) And _
              IsDate(rCelda.Offset(0, 4)) Then _
              FormatoErrorFaltante
        End If
        ' No número
        If Not IsNumeric(rCelda.Offset(0, 2)) Then _
          FormatoErrorEntrada
    
    ' Edad
        Set rCelda1 = rCelda.Offset(0, 3)
        FormatoEstandard
        ' Vaquillas atrasadas
        On Error Resume Next
        If Int((Date - CDate(rCelda.Offset(0, 4))) / 30.4) > _
          Range("Configuracion!C47") And _
          IsEmpty(rCelda.Offset(0, 6)) Then _
          FormatoErrorManejo
        ' Vaquillas por destetar
        If rCelda.Offset(0, 1) = _
          Range("Configuracion!C13") And _
          (Date - CDate(rCelda.Offset(0, 4))) > _
          Range("Configuracion!C34") Then _
          FormatoErrorManejo
        On Error GoTo 0
    'Fecha Nacimiento
        Set rCelda1 = rCelda.Offset(0, 4)
        FormatoEstandard
        If IsEmpty(rCelda.Offset(0, 4)) Then _
          GoTo ColC
        ' Fecha Futura
        On Error Resume Next
        If CDate(rCelda.Offset(0, 4)) > Date Then _
          FormatoErrorEntrada
        On Error GoTo 0
        ' No Fecha
        If Not IsDate(rCelda.Offset(0, 4)) And _
          Not IsEmpty(rCelda.Offset(0, 4)) Then _
          FormatoErrorEntrada
ColC:
    ' Control de Pesos
        If CBool(Range("Configuracion!C35")) Then
            Set rCelda1 = rCelda.Offset(0, 2)
            ' En Blanco
            If IsEmpty(rCelda.Offset(0, 2)) Then _
              GoTo ColE
            If CurvaPesos(Val(rCelda.Offset(0, 0)), _
              Val(rCelda.Offset(0, 2)), rCelda.Offset(0, 3)) _
              = "Bajo" Then FormatoErrorManejo
            If CurvaPesos(Val(rCelda.Offset(0, 0)), _
              Val(rCelda.Offset(0, 2)), rCelda.Offset(0, 3)) _
              = "Alto" Then FormatoErrorManejo1
        End If
ColE:
        ' En Blanco
        If IsEmpty(rCelda.Offset(0, 4)) And _
          Not IsEmpty(rCelda.Offset(0, 0)) Then _
          FormatoErrorFaltante
    
    'Servicio
        Set rCelda1 = rCelda.Offset(0, 5)
        FormatoEstandard
        If IsEmpty(rCelda.Offset(0, 5)) Then _
          GoTo ColF
        ' Repetidora
        If rCelda.Offset(0, 5) > 2 And _
          Not rCelda.Offset(0, 9) = "P" Then _
          FormatoErrorManejo
        ' No es numero
        If Not IsNumeric(rCelda.Offset(0, 5)) Then _
          FormatoErrorEntrada
ColF:
        ' En Blanco
        If IsEmpty(rCelda.Offset(0, 5)) And _
          Not IsEmpty(rCelda.Offset(0, 6)) Then _
          FormatoErrorFaltante
        
        ' Calor
        If rCelda.Offset(0, 7) = "Calor" Then _
          FormatoEstandard
        
    ' Fecha Servicio
        Set rCelda1 = rCelda.Offset(0, 6)
        FormatoEstandard
        If IsEmpty(rCelda.Offset(0, 6)) Then _
          GoTo ColG
        ' Fecha Futura
        On Error Resume Next
        If CDate(rCelda.Offset(0, 6)) > Date Then _
          FormatoErrorEntrada
        On Error GoTo 0
        ' Falta Dx
        On Error Resume Next
        If IsDate(rCelda.Offset(0, 6)) And _
          Date - CDate(rCelda.Offset(0, 6)) > _
          Range("Configuracion!C5") And _
          Not rCelda.Offset(0, 9) = "P" Then _
          FormatoErrorManejo
        On Error GoTo 0
        ' No Fecha
        If Not IsDate(rCelda.Offset(0, 6)) And _
          Not IsEmpty(rCelda.Offset(0, 6)) Then _
          FormatoErrorEntrada
ColG:
        ' En Blanco
        If IsEmpty(rCelda.Offset(0, 6)) And _
          Not IsEmpty(rCelda.Offset(0, 5)) Then _
          FormatoErrorFaltante
        
    ' Semental
        Set rCelda1 = rCelda.Offset(0, 7)
        FormatoEstandard
        ' En Blanco
        If IsEmpty(rCelda.Offset(0, 7)) And _
          CBool(Range("Configuracion!C15")) And _
          Not IsEmpty(rCelda.Offset(0, 6)) Then _
          FormatoErrorFaltante

    ' Técnico
        Set rCelda1 = rCelda.Offset(0, 8)
        FormatoEstandard
        ' En Blanco
        If IsEmpty(rCelda.Offset(0, 8)) And _
          CBool(Range("Configuracion!C16")) And _
          Not IsEmpty(rCelda.Offset(0, 7)) Then _
          FormatoErrorFaltante
    
    ' Status
        Set rCelda1 = rCelda.Offset(0, 9)
        FormatoEstandard
        ' Contenido diferente a G o V
        If (Not IsEmpty(rCelda.Offset(0, 9)) And _
          Not rCelda.Offset(0, 9) = "P") And _
          (Not IsEmpty(rCelda.Offset(0, 9)) And _
          Not rCelda.Offset(0, 9) = "O") Then _
          FormatoErrorEntrada

    ' FxSecar
        Set rCelda1 = rCelda.Offset(0, 10)
        FormatoEstandard
        If IsEmpty(rCelda.Offset(0, 10)) Then _
          GoTo ColK
        ' Fecha pasada
        On Error Resume Next
        If Date >= CDate(rCelda.Offset(0, 10)) Then _
          FormatoErrorManejo
        On Error GoTo 0
ColK:
        ' En blanco
        If IsEmpty(rCelda.Offset(0, 10)) And _
          rCelda.Offset(0, 9) = "P" Then _
          FormatoErrorFaltante
    
    ' Sexo
       Set rCelda1 = rCelda.Offset(0, 13)
        FormatoEstandard
        If IsEmpty(rCelda.Offset(0, 13)) Then _
          GoTo ColP
        ' Contenido
        If Not rCelda.Offset(0, 13) = "H" And _
          Not rCelda.Offset(0, 13) = "M" And _
          Not rCelda.Offset(0, 13) = "FM" Then _
          FormatoErrorEntrada
ColP:
        'En Blanco
        If IsEmpty(rCelda.Offset(0, 13)) Then _
          FormatoErrorFaltante
    
        lCounter = lCounter + 1
    Next rCelda
    Proteger
    Application.StatusBar = False
    Application.ScreenUpdating = True
End Sub

Private Sub FiltrosOff()
' Quita los Autofiltros de la hoja activa
    Selection.EntireColumn.Hidden = False
    Range("A2").Select
    Selection.AutoFilter
    Selection.AutoFilter
End Sub

Private Sub FinalLista()
' Busca último renglón, independientemente de la versión de excel
    Cells(Rows.Count, 1).End(xlUp).Select
    'iUltimoRegistro = ActiveCell.Row - 1
End Sub

Private Sub FormatoErrorEntrada()
' Cambia Formato
    With rCelda1
        .Interior.Color = 255 'Fondo Rojo
        .Font.ThemeColor = xlThemeColorDark1 'Letra Blanca
        .Font.Bold = True 'Negritas
    End With
End Sub

Private Sub FormatoErrorFaltante()
' Cambia Formato
    With rCelda1.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 49407
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
End Sub

Private Sub FormatoErrorManejo()
' Cambia Formato
    With rCelda1
        .Interior.Pattern = xlNone 'Fondo Estándard
        .Font.Color = -16776961
        .Font.Bold = True 'Normal
    End With
End Sub

Private Sub FormatoErrorManejo1()
' Cambia Formato
    With rCelda1
        .Interior.Pattern = xlNone 'Fondo Estándard
        .Font.ColorIndex = 29
        '.Font.Color = -16776961
        .Font.Bold = True 'Normal
    End With
End Sub


Private Sub FormatoEstandard()
' Cambia Formato
    With rCelda1
    'With Selection
        .Interior.Pattern = xlNone 'Fondo Estándard
        .Font.ColorIndex = xlAutomatic 'Letra estándard
        .Font.Bold = False 'Normal
        .HorizontalAlignment = xlCenter
    End With
End Sub

Private Sub LoginRecord()
' Diario de Actividades
    Dim sPath As String
    Dim sArch As String
    Dim bArch As Boolean
    Dim ilogin
    'MsgBox = Dir(Application.ActiveWorkbook.Path)
    sPath = Application.ActiveWorkbook.Path
    'MsgBox = sPath
'    sArch = Dir("C:\Users\yo\Documents\My Box Files\INIFAP\Log101.txt")
    sArch = Dir("Log101.txt")
    If sArch <> vbNullString Then bArch = True
    On Error GoTo 100
'    Open "C:\Users\yo\Documents\My Box Files\INIFAP\Log101.txt" For _
      Append As #1
    Open "Log101.txt" For Append As #1
    On Error GoTo 0
200
    If bArch = False Then Write #1, _
      "Fecha", "Arete", "Evento", "Clave", "Detalles"
    If ilogin = 1 Then    'Entrada al sistema
            'Write #1, TextDate, cmboIdArete, cmboEvento,
            'Write #1, Me.txtdate.Text, formLogin.txtUsuario.Text, Date, _
              Format(Time, "h:mm")
        Else              'Salida del sistema
            'Write #1, "Cerrado", formLogin.txtUsuario.Text, Date, _
              Format(Time, "h:mm")
    End If
    Close #1
    Exit Sub
100
    Open "Log101.txt" For Append As #1
    GoTo 200
End Sub

Private Sub MensajeCierre()
    MsgBox _
      "Utiliza 'Cerrar' para salir de este formulario.", _
      vbExclamation, _
      "Generación de Códigos"
End Sub

Private Sub OrdenarEventos()
' Ordenar Hoja "Eventos" (Fecha y Arete)
    With ActiveWorkbook.Worksheets("Eventos")
        With .ListObjects("Tabla6").Sort
            .SortFields.Clear
            .SortFields.Add _
              Key:=Range("Tabla6[Fecha]"), _
              SortOn:=xlSortOnValues, _
              Order:=xlAscending, _
              DataOption:=xlSortNormal
            .SortFields.Add _
              Key:=Range("Tabla6[Arete]"), _
              SortOn:=xlSortOnValues, _
              Order:=xlAscending, _
              DataOption:=xlSortNormal
            .Header = xlYes
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
            .SortFields.Clear
        End With
    End With
End Sub

Private Sub OrdenarHato()
    ' Ordenar Hoja "Hato"
    With ActiveWorkbook.Worksheets("Hato")
        With .ListObjects("Tabla1").Sort
            .SortFields.Clear
            .SortFields.Add _
              Key:=Range("Tabla1[Arete]"), _
              SortOn:=xlSortOnValues, _
              Order:=xlAscending, _
              DataOption:=xlSortNormal
            .Header = xlYes
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
            .SortFields.Clear
        End With
    End With
End Sub

Private Sub OrdenarHato2()
    ' Ordenar Hoja "Hato2"
    With ActiveWorkbook.Worksheets("Hato2")
        With .ListObjects("Tabla15").Sort
            .SortFields.Clear
            .SortFields.Add _
              Key:=Range("Tabla15[Arete]"), _
              SortOn:=xlSortOnValues, _
              Order:=xlAscending, _
              DataOption:=xlSortNormal
            .Header = xlYes
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
            .SortFields.Clear
        End With
    End With
End Sub

Private Sub OrdenarHojas()
    sHoja = "Hato2": sTabla = "Tabla15": scampo = "Arete"
    OrdenarTablas
    'OrdenarHato2
    sHoja = "Reemplazos": sTabla = "Tabla2": scampo = "Arete"
    OrdenarTablas
    'OrdenarReemplazos
    sHoja = "Hato": sTabla = "Tabla1": scampo = "Arete"
    OrdenarTablas
    'OrdenarHato
    sHoja = "Eventos": sTabla = "Tabla6": scampo = "Fecha"
    OrdenarTablas
    'OrdenarEventos
End Sub

Private Sub OrdenarTablas()
    ' Ordenar Hoja "Reemplazos"
    With ActiveWorkbook.Worksheets(sHoja)
        With .ListObjects(sTabla).Sort
            .SortFields.Clear
            .SortFields.Add _
              Key:=Range(sTabla & "[" & scampo & "]"), _
              SortOn:=xlSortOnValues, _
              Order:=xlAscending, _
              DataOption:=xlSortNormal
            .Header = xlYes
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
            .SortFields.Clear
        End With
    End With
End Sub





Private Sub OrdenarReemplazos()
    ' Ordenar Hoja "Reemplazos"
    With ActiveWorkbook.Worksheets("Reemplazos")
        With .ListObjects("Tabla2").Sort
            .SortFields.Clear
            .SortFields.Add _
              Key:=Range("Tabla2[Arete]"), _
              SortOn:=xlSortOnValues, _
              Order:=xlAscending, _
              DataOption:=xlSortNormal
            .Header = xlYes
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
            .SortFields.Clear
        End With
        '.ActiveCell.Activate
    End With
End Sub

Private Sub Proteger()
    If Not CBool(Range("Configuracion!C39")) Then _
      ProtegerHoja
End Sub

Private Sub ProtegerHoja()
' Proteger hoja activa, dejando algunas atribuciones
    ActiveSheet.Protect Password:="0246813579", _
      DrawingObjects:=True, _
      Contents:=True, _
      Scenarios:=True, _
      AllowFormattingCells:=True, _
      AllowFormattingColumns:=False, _
      AllowFormattingRows:=True, _
      AllowSorting:=True, _
      AllowFiltering:=True, _
      AllowUsingPivotTables:=True
End Sub

Private Sub QuitarAutofiltros()
' Quitar los autofiltros de la hojas
    Selection.AutoFilter
    Selection.AutoFilter
End Sub

Private Sub QuitarTabla()
' Convierte una tabla en un rango de celdas
    With Application
        .DisplayAlerts = False
        .CommandBars.ExecuteMso "TableConvertToRange"
        .DisplayAlerts = True
    End With
End Sub

Private Sub UnderConstruction()
    Dim sTextoMsj, sMsjTitulo, sMensaje As String
    sMsjTitulo = Range("Configuracion!C3").Text
    sTextoMsj = "Código en Construcción"
    sMensaje = MsgBox(sTextoMsj, vbInformation, sMsjTitulo)
End Sub

Private Function CurvaPesos(Id As Double, Peso As Double, Edad As String)
    Dim a1, col As Double
    Dim sRaza As String
    Dim rCelda As Range
    sRaza = "CrossBreed"
    a1 = Val(Mid(Edad, 1, Len(Edad) - 1))
    For Each rCelda In Range("Tabla8[Arete]")
        If rCelda.Offset(0, 0) = Id Then sRaza = rCelda.Offset(0, 5): Exit For
    Next
    Select Case UCase(sRaza)
        Case Is = "HOLSTEIN"
            col = 7
        Case Is = "JERSEY"
            col = 1
        Case Is = "BROWN SWISS"
            col = 7
        Case Is = "GUERNSEY"
            col = 3
        Case Is = "AYRSHIRE"
            col = 6
        Case Is = "CROSSBREED"
            col = 9
        Case Else
            col = 7
    End Select
    a1 = Val(Mid(Edad, 1, Len(Edad) - 1))
    For Each rCelda In Range("Tabla114[Age]")
        If rCelda.Offset(0, 0) = a1 Then
            Select Case Peso
                Case Is < Val(Mid(rCelda.Offset(0, col), 1, 4)) * 0.45359237
                    CurvaPesos = "Bajo"
                Case Is > Val(Mid(rCelda.Offset(0, col), 6, 4)) * 0.45359237
                    CurvaPesos = "Alto"
                Case Else
                    CurvaPesos = "Ok"
            End Select
            Exit For
        End If
    Next
1234:
End Function

