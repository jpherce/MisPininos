Attribute VB_Name = "ModImportarInfo"
'Última Modificación: 21-Oct-17
' Calcular persistencia
Option Explicit
Dim bFlagError As Boolean
Dim rCelda As Range
Dim wb As Workbook
Dim iARowInfoVital, iARH, iARH2, iAR As Long
Dim sLocAnimal, sLocPrevia As String
Dim sEvento, sObserv, sResp As String
Dim sMetadato As String
Dim ws, wsH, wsR, wsH2, wsIV, wsP As Worksheet

Private Sub AgregarHoja()
' Agrega Hoja y escribe encabezados
    On Error GoTo 1234
    Worksheets.Add.Name = "Importación de Datos"
1234:
    On Error GoTo 0
    Worksheets("Importación de Datos").Select
    With Range("A1")
        .Value = "Fecha"
        .Font.Bold = True
    End With
    With Range("B1")
        .Value = "Arete"
        .Font.Bold = True
    End With
    With Range("C1")
        .Value = "Clave"
        .Font.Bold = True
    End With
    With Range("D1")
        .Value = "Observación"
        .Font.Bold = True
    End With
    With Range("E1")
        .Value = "Técnico"
        .Font.Bold = True
    End With
    With Range("F1")
        .Clear
        .Font.Bold = True
    End With
End Sub

Private Sub PrepImportarDatos()
    Dim sMsjTitulo, sRespuesta As String
    Dim nR As Long

    sMsjTitulo = "Importación de Datos"
    ' Verificar existencia Hoja
    If BuscarHoja("Importación de Datos") Then
            With Worksheets("Importación de Datos")
                .Visible = True
                .Select
            End With
            Range("A1").Select
        Else
            AgregarHoja
    End If
    ' Verificar Hoja limpia
    nR = WorksheetFunction.CountA(Range("A:A"))
    If nR >= 2 Then
            sRespuesta = MsgBox("¿Borrar los datos existentes?", _
              vbYesNo + vbDefaultButton2 + vbQuestion, sMsjTitulo)
            If sRespuesta = vbYes Then
                Range(Cells(2, 1), Cells(nR, 6)).Clear
            End If
        Else
            Range("F1").Clear
    End If
    'Mostrar Instrucciones
    sRespuesta = MsgBox("        INSTRUCCIONES" & Chr(13) & Chr(13) & _
      "1° Copiar y Pegar los datos que se desean importar" & Chr(13) & _
      "    en el orden estipulado en encabezados." & Chr(13) & _
      "2° Presionar Botón [Importar Datos] para proceder." & Chr(13) & _
      "3° Revisar mensaje devuelto por el sistema.", _
      vbInformation + vbOKOnly, sMsjTitulo)
End Sub
    
Sub ImportarDatos()
    Dim a As Range
    Dim sMsjTitulo As String
    Dim lCounter, lTotal As Long
    Application.ScreenUpdating = CBool(Range("Desarrollador!B6"))
    sMsjTitulo = "Importación de Datos"
    Set wsH = Worksheets("Hato")
    Set wsR = Worksheets("Reemplazos")
    Set wsH2 = Worksheets("Hato2")
    Set wsIV = Worksheets("InfoVitalicia")
    Worksheets("Importación de Datos").Select
    Range("Desarrollador!B20") = "T"
    ' Verificar que existan datos
    If WorksheetFunction.CountA(Range("A:A")) = 1 Then
        MsgBox "No existen datos que importar", vbCritical + vbOKOnly, _
          sMsjTitulo
        Exit Sub
    End If
    Set a = Range("A1:" & Range("A1").End(xlDown).Address)
    With Range("F1")
        .Value = "DatosImportados"
        .Font.Bold = True
    End With
    Application.DisplayStatusBar = True
    lTotal = a.Rows.Count
    For Each rCelda In a
        Application.StatusBar = _
          "Importando... " & _
          Format(lCounter / lTotal, "0%")
        bFlagError = False
        ' Checar título
        If UCase(rCelda.Offset(0, 0)) = "FECHA" Then GoTo 8576
        ' Validar Col A como Fecha
        If Not IsDate(rCelda.Offset(0, 0)) Then
            rCelda.Offset(0, 5) = "No es Fecha"
            bFlagError = True
            GoTo 5678
        End If
        ' Fecha posterior al sistema
        If CDate(rCelda.Offset(0, 0)) > Date Then
            rCelda.Offset(0, 5) = "¡La Fecha es para el Futuro!"
            bFlagError = True
            GoTo 5678
        End If
        ' Validar Col C Como Arete
        If Not IsNumeric(rCelda.Offset(0, 1)) Or _
          IsEmpty(rCelda.Offset(0, 1)) Then
            rCelda.Offset(0, 5) = "No es Arete"
            bFlagError = True
            GoTo 5678
        End If
        ChecarMismoEvento
        If bFlagError = True Then GoTo 5678
        ' Validar existencia del animal
        LocalizaArete
        If bFlagError = True Then
            rCelda.Offset(0, 5) = "Arete no Encontrado"
            GoTo 5678
        End If
        ' Validar Col B como Evento
        Select Case UCase(rCelda.Offset(0, 2))
            Case Is = "SERVICIO"
                ImportarServicios
            Case Is = "CALOR"
                ImportarCalores
            Case Is = "PRODUCCIÓN"
                ImportarProd
            'Case Is = "MOVIMIENTO"
            'Case Is = "ENFERMEDAD"
            Case Is = "REVISIÓN"
                ImportarRevision
            Case Is = "DX GEST."
                ImportarDxGest
                Case Is = "SECAR"
                ImportarSecado
            'Case Is = "NOTA"
            'Case Is = "PARTO"
            'Case Is = "IMANTACIÓN"
            'Case Is = "OTRO"
            'Case Is = "BAJA"
            Case Else
                rCelda.Offset(0, 5) = "Clave no programada"
                bFlagError = True
        End Select
        'LocalizaArete
5678:
        If Not bFlagError Then rCelda.Offset(0, 5) = "Ok"
8576:
        lCounter = lCounter + 1
    Next rCelda
    Range("Desarrollador!B20").Clear
    Application.StatusBar = False
    Application.ScreenUpdating = True
End Sub

Private Sub ImportarRevision()
    CalcularProxRevision
    If InStr(UCase(rCelda.Offset(0, 3)), "OE") Or (InStr(UCase(rCelda.Offset(0, 3)), "ODE") _
      And InStr(UCase(rCelda.Offset(0, 3)), "OIE")) Then sMetadato = "Anestro"
    ConsecutivoDeEventos
End Sub

Private Sub CalcularProxRevision()
' Calcula la prox. fecha de revisión en base a las
'   observaciones.
    Dim iDias As Long
    Dim sCadena As String
    sCadena = UCase(Trim(rCelda.Offset(0, 3)))
    If Right(sCadena, 2) = "R8" Then
        iDias = 8
    End If
    If Right(sCadena, 3) = "R15" Then
        iDias = 14
    End If
    If Right(sCadena, 3) = "R21" Then
        iDias = 21
    End If
    If Right(sCadena, 3) = "R30" Then
        iDias = 28
    End If
    If iDias > 0 Then
            With wsIV.Cells(iARowInfoVital, 16)
                .Value = CDate(rCelda.Offset(0, 0)) + iDias
                .NumberFormat = "d-mmm-yy"
            End With
        Else
            wsIV.Cells(iARowInfoVital, 16).Clear
     End If
End Sub

Private Sub Factorizacion()
    Dim n As Long
    'N = nMM
    If n >= 2 ^ 4 Then '16
        'Tabla8
        'If bFlagError Then Exit Sub
        If n - 2 ^ 4 >= 0 Then n = n - 2 ^ 4
    End If
    If n >= 2 ^ 3 Then '8
        'Tabla2
        'If bFlagError Then Exit Sub
        If n - 2 ^ 3 >= 0 Then n = n - 2 ^ 3
    End If
    If n >= 2 ^ 2 Then '4
        'Tabla15
        'If bFlagError Then Exit Sub
        If n - 2 ^ 2 >= 0 Then n = n - 2 ^ 2
    End If
    If n >= 2 Then
        'Tabla1
        'If bFlagError Then Exit Sub
        If n - 2 ^ 1 >= 0 Then n = n - 2 ^ 1
    End If
End Sub

Private Sub LocalizaArete()
 ' Localiza al animal
   'Dim N As Long
    'N = nMM
    Dim bAreteNoEncontrado As Boolean
    Dim sArete As Variant
    bAreteNoEncontrado = False
    sLocPrevia = sLocAnimal
    sLocAnimal = vbNullString
    If rCelda.Offset(0, 1) = vbNullString Then GoTo 4231
    If IsNumeric(rCelda.Offset(0, 1)) Then sArete = _
      CDbl(rCelda.Offset(0, 1)) Else sArete = rCelda.Offset(0, 1)
    ' Buscar y posicionarse
    iARH = IndiceTabla(rCelda.Offset(0, 1), "Tabla1")
    If iARH > 0 Then
            sLocAnimal = "H"
            '********
            'Application.Run "Desproteger"
            Cells(iARH, 1).Activate
            Set ws = Worksheets("Hato2")
            'Application.Run "Proteger"

1234:
            iARH2 = _
              IndiceTabla(rCelda.Offset(0, 1), "Tabla15")
            If iARH2 = 0 Then GoTo 3412 'Sí no existe
        Else
            iAR = IndiceTabla(rCelda.Offset(0, 1), "Tabla2")
            If iAR = 0 Then GoTo ControlDeErrores 'Sí no existe
            sLocAnimal = "R"
    End If
    Set ws = Worksheets("InfoVitalicia")
    Application.Run "Desproteger"
2341:
    iARowInfoVital = IndiceTabla(rCelda.Offset(0, 1), "Tabla8")
    If iARowInfoVital = 0 Then GoTo 4123 'Sí no existe
    Exit Sub

3412:
    Alta_Hoja2
    GoTo 1234
4123:
    'AltaInfovital
    GoTo 2341

ControlDeErrores:
    bAreteNoEncontrado = True
4231:
    'bProceder = False
End Sub

Private Sub LoginRecord1()
' Bitácora de Eventos en formato CSV
'    Dim sPath As String
    Dim sArch As String
    Dim bArch As Boolean
'    sPath = Application.ActiveWorkbook.Path
'    sArch = Dir("C:\Users\yo\Documents\My Box Files\INIFAP\Log101.txt")
    sArch = Dir("Log101.txt")
    If sArch <> vbNullString Then bArch = True
    On Error GoTo 100
'    Open "C:\Users\yo\Documents\My Box Files\INIFAP\Log101.txt" For Append As #1
    Open "Log101.txt" For Append As #1
    On Error GoTo 0
200
    If bArch = False Then Write #1, "IdHato", "Arete", "Fecha", "Evento", _
      "Observaciones", "Responsable", "Usuario", "F.Captura", "H.Captura"
    Write #1, Range("Configuracion!D3"), rCelda.Offset(0, 1), _
      rCelda.Offset(0, 0), rCelda.Offset(0, 2), rCelda.Offset(0, 3), _
      rCelda.Offset(0, 4), Range("Configuracion!C49"), Date, Format(Time, "h:mm")
    Close #1
    Exit Sub
100
    Open "Log101.txt" For Append As #1
    GoTo 200
End Sub

Private Sub ImportarProd()
    'Dim rCelda As Range
    Dim mProdAcum, mDiasProd, mProdMax, _
      mPicoProd, mProy305, mPersist As Double
    Dim ws, wsIV As Worksheet
    Dim iCol, iParto, i As Long
    Dim dFParto As Date
    Dim bTest As Boolean
    ' Producciones en blanco
    If Not IsNumeric(rCelda.Offset(0, 3)) Or _
      IsEmpty(rCelda.Offset(0, 3)) Then
        rCelda.Offset(0, 5) = "No hay producción"
        bFlagError = True
        Exit Sub
    End If
    ' Producciones fuera de rango
    If rCelda.Offset(0, 3) < 0 Or _
      rCelda.Offset(0, 3) > 69 Then
        rCelda.Offset(0, 5) = "Prod. fuera de rango"
        bFlagError = True
        Exit Sub
    End If
    ' Producciones no numéricas
    If Not IsNumeric(rCelda.Offset(0, 3)) Then
        rCelda.Offset(0, 5) = "Prod. NO numérica"
        bFlagError = True
        Exit Sub
    End If
    ' Corral en blanco
    If Not rCelda.Offset(0, 4) = vbNullString And _
      Not IsNumeric(rCelda.Offset(0, 4)) Then
        rCelda.Offset(0, 5) = "Corral NO especificado"
        bFlagError = True
        Exit Sub
    End If
    ' Producciones negativas
    If Val(rCelda.Offset(0, 3)) < 0 Then
        rCelda.Offset(0, 5) = "Producción Negativa"
        bFlagError = True
        Exit Sub
    End If
    If Not CheckFecha(rCelda.Offset(0, 1), 6) Then Exit Sub
    iParto = WorksheetFunction.VLookup(rCelda.Offset(0, 1), _
      Range("Tabla1"), 5, False)
    dFParto = CDate(WorksheetFunction.VLookup(rCelda.Offset(0, 1), _
      Range("Tabla1"), 6, False))
    Set ws = Worksheets("Hato2")
    Application.Run "Desproteger"
    ws.Visible = xlSheetVisible
    ' Vaca Seca
    If WorksheetFunction.VLookup(rCelda.Offset(0, 1), _
      Range("Tabla1"), 2, False) = Range("Configuracion!C9") Or _
      WorksheetFunction.VLookup(rCelda.Offset(0, 1), _
      Range("Tabla1"), 2, False) = Range("Configuracion!C10 ") Then
        rCelda.Offset(0, 5) = "Vaca Seca"
        bFlagError = True
    End If
    ' Escribir Datos
    mDiasProd = CDate(rCelda.Offset(0, 0)) - dFParto
    If wsH.Cells(iARH, 3) > 0 Then _
      mPersist = Int(CDbl(rCelda.Offset(0, 3)) / wsH.Cells(iARH, 3) * 100) _
      Else mPersist = 0
    ' DEL-Persistencia-Parto
    sMetadato = Format(mDiasProd, "000") & "-" & Format(mPersist, "000") _
      & "-" & Format(wsH.Cells(iARH, 5), "00")
    Select Case mDiasProd
        Case Is <= 30
            iCol = 4
        Case Is <= 60
            iCol = 5
        Case Is <= 90
            iCol = 6
        Case Is <= 120
            iCol = 7
        Case Is <= 150
            iCol = 8
        Case Is <= 180
            iCol = 9
        Case Is <= 210
            iCol = 10
        Case Is <= 240
            iCol = 11
        Case Is <= 270
            iCol = 12
        Case Is <= 305
            iCol = 13
        Case Is > 305
            GoTo 365
    End Select
    With ws.Cells(iARH2, iCol)
        .Value = CDbl(rCelda.Offset(0, 3))
        .NumberFormat = "0.0"
    End With
365:
    ' Acumular Producción
    mProdAcum = 0
    mProdMax = 0
    mPicoProd = WorksheetFunction. _
      Max(Range(ws.Cells(iARH2, 4), ws.Cells(iARH2, 13)))
    For iCol = 4 To 13
        ' ***** FALTA HACER UN AJUSTE PARA CUANDO NO SE HAGAN PESAJES
        ' (prod actual + ultima prod)/(fecha actual - ultima fecha)
        'If Sheets("Hato2").Cells(iARH2, iCol) > mProdMax Then _
          mProdMax = Sheets("Hato2").Cells(iARH2, iCol)
        mProdAcum = mProdAcum + _
          ws.Cells(iARH2, iCol) * 30
    Next iCol
    mProdAcum = Int(mProdAcum / 10) * 10
    ws.Cells(iARH2, 14) = mProdAcum 'Prod.Acum
    ' Proyectar producción a 305d
    'mProy305 = Int(mPicoProd * Range("Configuracion!C36") / 10) * 10
    
    'If Not IsEmpty(wsIV.Cells(iARowInfoVital, 6)) Then _
      Range("CurvaLact!M2") = wsIV.Cells(iARowInfoVital, 6) _
      Else Range("CurvaLact!M2") = "CrossBreed"
    
    If UCase(Range("CurvaLact!M2")) = "CRUZA" Then _
      Range("CurvaLact!M2") = "CrossBreed"
    If iParto < 3 Then Range("CurvaLact!M3") = iParto _
      Else iParto = 3
    If mDiasProd <= 305 Then Range("CurvaLact!N9") = mDiasProd _
      Else Range("CurvaLact!N9") = 305
    ' Utilizar la función [Buscar Objetivo]
    'On Error Resume Next
    Range("CurvaLact!N10").GoalSeek Goal:=Val(rCelda.Offset(0, 3)), _
      ChangingCell:=Range("CurvaLact!N4")
    'On Error GoTo 0
    mProy305 = Range("CurvaLact!N11")
    ws.Cells(iARH2, 15) = mProy305 'Proy.305d
    If mPersist > 0 Then ws.Cells(iARH2, 19) = mPersist
    Set wsP = wsH
    Application.Run "Desproteger1"
    With wsH.Cells(iARH, 3) 'Prod
        .Value = CDbl(rCelda.Offset(0, 3))
        .NumberFormat = "0.0"
        If CBool(Range("Configuracion!B65")) Then _
          wsH.Cells(iARH, 16) = mProy305
          'wsH.Cells(iARH, 16) = Format(mProy305, "#,#")
    End With
    If Not rCelda.Offset(0, 4) = vbNullString Then _
      wsH.Cells(iARH, 2) = CDbl(rCelda.Offset(0, 4))
    If Not CBool(Range("Desarrollador!B6")) Then _
      ws.Visible = xlSheetVeryHidden
    ConsecutivoDeEventos
End Sub

Private Sub ImportarSecado()
    ' Vaca ya Seca
    If wsH.Cells(iARH, 2) = _
      Range("Configuracion!C9") Then
        rCelda.Offset(0, 5) = "Animal previamente reportado como Seca"
        bFlagError = True
        Exit Sub
    End If
    ' Más de 152 dias en leche
    If CDate(rCelda.Offset(0, 0)) - _
      CDate(wsH.Cells(iARH, 8)) <= 152 Then
        rCelda.Offset(0, 6) = "La lactancia sólo tuvo " & _
              wsH.Cells(iARH, 4) & " días de duración."
    End If
    ' Animal no gestante
    If Not wsH.Cells(iARH, 11) = "P" Then
        rCelda.Offset(0, 5) = "Este animal no está gestante."
    End If
    Set ws = wsH
    Application.Run "Desproteger1"
    ' Escribir Datos
    'Parto-DEL
    sMetadato = Format(wsH.Cells(iARH, 5), "00") & "-" & _
      Format(CDate(rCelda.Offset(0, 0)) - _
      CDate(wsH.Cells(iARH, 6)), "000")
    wsH.Cells(iARH, 2) = _
      Range("Configuracion!C9") 'Corral
    wsH.Cells(iARH, 3).Clear 'Produccion
    wsH.Cells(iARH, 4).Clear 'del
    wsH.Cells(iARH, 12) = "**SECA**" 'FxSecar
    With wsH2.Cells(iARH2, 16) 'F.Secado
        .Value = CDate(rCelda.Offset(0, 0))
        .NumberFormat = "d-mmm-yy"
    End With
    If Not CBool(Range("Desarrollador!B6")) Then _
      ws.Visible = xlSheetVeryHidden
    ConsecutivoDeEventos
End Sub

Private Sub ImportarCalores()
    If IsEmpty(rCelda.Offset(0, 4)) And CBool(Range("Configuracion!C16")) Then
        rCelda.Offset(0, 5) = "Falta técnico responsable"
        Exit Sub
    End If
    If sLocAnimal = "H" Then
        ' Último Servicio o calor
        If CDate(rCelda.Offset(0, 0)) <= CDate(wsH.Cells(iARH, 8)) Then
            rCelda.Offset(0, 5) = "Fecha del Evento es igual o anterior a la Fecha del " _
                & rCelda.Offset(0, 2) & " registrado"
            bFlagError = True
            Exit Sub
        End If
        If CDate(rCelda.Offset(0, 0)) - CDate(wsH.Cells(iARH, 8)) <= 18 And _
          Not IsEmpty(wsH.Cells(iARH, 8)) Then
            rCelda.Offset(0, 6) = _
              "El intervalo entre Calores o Servicios es menor a 18 días"
        End If
        If CDate(rCelda.Offset(0, 0)) - CDate(wsH.Cells(iARH, 8)) >= 36 And _
          Not IsEmpty(wsH.Cells(iARH, 8)) Then
            rCelda.Offset(0, 6) = _
              "El intervalo entre Calores o Servicios es mayor a 36 días"
        End If
        ' Escribir Datos
        ' 00-DíasÚltimoServicio-DEL
        If IsEmpty(wsH.Cells(iARH, 7)) Then sMetadato = "00-000" Else _
          sMetadato = "00-" & Format(CDate(rCelda.Offset(0, 0)) - _
          CDate(wsH.Cells(iARH, 8)), "000")
        sMetadato = sMetadato & "-" & _
          Format(CDate(rCelda.Offset(0, 0)) - CDate(wsH.Cells(iARH, 6)), "000")
        If wsH.Cells(iARH, 8) = vbNullString Then
            With wsH2.Cells(iARH2, 17) 'd1Calor
                .Value = CDate(rCelda.Offset(0, 0)) - CDate(wsH.Cells(iARH, 6))
            End With
        End If
        If wsH.Cells(iARH, 11) = "P" Then
            wsH.Cells(iARH, 12).Clear 'FxSecar
            wsH.Cells(iARH, 13).Clear 'FxParir
            wsH.Cells(iARH, 14) = "pAb" 'Clave1
            wsH2.Cells(iARH2, 3).Clear 'DAbiertos
            rCelda.Offset(0, 6) = _
              "Animal previamente reportado como Gestante."
        End If
        With wsH.Cells(iARH, 8)    'F.Servicio
            .Value = CDate(rCelda.Offset(0, 0))
            .NumberFormat = "d-mmm-yy"
        End With
        wsH.Cells(iARH, 9) = "Calor" 'Semental
        wsH.Cells(iARH, 10) = UCase(rCelda.Offset(0, 4)) 'Tecnico
        wsH.Cells(iARH, 11).Clear 'Status
    End If
    If sLocAnimal = "R" Then
        ' Checar errores
        ChecarSexo
        If bFlagError Then Exit Sub
        ' Último Servicio o calor
        If CDate(rCelda.Offset(0, 0)) <= CDate(wsR.Cells(iAR, 7)) Then
            rCelda.Offset(0, 5) = _
              "Fecha del Evento es igual o anterior a la Fecha del " _
                & rCelda.Offset(0, 2) & " registrado"
            bFlagError = True
            Exit Sub
        End If
        If CDate(rCelda.Offset(0, 0)) - CDate(wsR.Cells(iAR, 7)) <= 18 And _
          Not IsEmpty(wsR.Cells(iAR, 7)) Then
            rCelda.Offset(0, 6) = _
              "El intervalo entre Calores o Servicios es menor a 18 días"
        End If
        If CDate(rCelda.Offset(0, 0)) - CDate(wsR.Cells(iAR, 7)) >= 36 And _
          Not IsEmpty(wsR.Cells(iAR, 7)) Then
            rCelda.Offset(0, 6) = _
              "El intervalo entre Calores o Servicios es mayor a 36 días"
        End If
        ' Escribir Datos
        ' 00-DíasÚltimoServicio-DEL
        If IsEmpty(wsR.Cells(iAR, 6)) Then sMetadato = "00-000" Else _
          sMetadato = "00-" & _
          Format(CDate(rCelda.Offset(0, 0)) - CDate(wsR.Cells(iAR, 7)), "000")
        sMetadato = sMetadato & "-" & _
          Format(CDate(rCelda.Offset(0, 0)) - CDate(wsR.Cells(iAR, 5)), "000")
        If wsR.Cells(iAR, 10) = "P" Then
            wsR.Cells(iAR, 11).Clear 'FxParir
            wsR.Cells(iAR, 12) = "pAb" 'Clave1
            wsIV.Cells(iARowInfoVital, 11).Clear 'Edad1Parto
            rCelda.Offset(0, 6) = _
              "Animal previamente reportado como Gestante."
        End If
        With wsR.Cells(iAR, 7)  'F.Servicio
            .Value = CDate(rCelda.Offset(0, 0))
            .NumberFormat = "d-mmm-yy"
        End With
        wsR.Cells(iAR, 8) = "Calor"   'Semental
        wsR.Cells(iAR, 9) = UCase(rCelda.Offset(0, 4))  'Tecnico
        wsR.Cells(iAR, 10).Clear  'Status
    End If
    ConsecutivoDeEventos
End Sub

Private Sub ChecarSexo()
    ' Checar Sexo
    If Not wsR.Cells(iAR, 14) = "H" Then
        rCelda.Offset(0, 5) = "El animal no es una hembra"
        bFlagError = True
    End If
End Sub

Private Sub ImportarDxGest()
    Dim sDx As String
    Set ws = Worksheets("Eventos")
    If rCelda.Offset(0, 3) = "Gestante" Then sDx = "P" Else sDx = "O"
    ' Mínimo un servicio
    If sLocAnimal = "H" Then
        If rCelda.Offset(0, 3) = "Gestante" Then
                ' Sin servicios
                If wsH.Cells(iARH, 7) < 1 Then
                    rCelda.Offset(0, 5) = "Animal sin Servicios"
                    bFlagError = True
                    Exit Sub
                End If
                ' Mínimo con 45 días post servicio
                If CDate(rCelda.Offset(0, 0)) - CDate(wsH.Cells(iARH, 8)) _
                  < Range("Configuracion!C5") Then
                    rCelda.Offset(0, 6) = "Animal con menos de " & _
                      Range("Configuracion!C5") & " días de servicio."
                End If
                ' Último servicio
                If UCase(wsH.Cells(iARH, 9)) = "CALOR" Then
                    rCelda.Offset(0, 5) = "Animal sin Servicios."
                    bFlagError = True
                    Exit Sub
                End If
                ' Escribir Datos
                'Servicio-DíasCarga
                sMetadato = Format(wsH.Cells(iARH, 7), "00") & "-" _
                  & Format(CDate(rCelda.Offset(0, 0)) - _
                  CDate(wsH.Cells(iARH, 8)), "000")
                wsH.Cells(iARH, 11) = "P" 'Status
                On Error Resume Next
                With wsH.Cells(iARH, 12)
                    .Value = CDate(wsH.Cells(iARH, 8)) + 213 'FxSecar
                    .NumberFormat = "d-mmm-yy"
                End With
                With wsH.Cells(iARH, 13)
                    .Value = CDate(wsH.Cells(iARH, 8)) + 273 'FxParir
                    .NumberFormat = "d-mmm-yy"
                End With
                If Range(iARH, 14) = "pAb" Then _
                      Range(iARH, 14).Clear 'Clave1
                wsH2.Cells(iARH2, 3) = _
                  CDate(wsH.Cells(iARH, 8)) - _
                  CDate(wsH.Cells(iARH, 6)) 'DAbiertos
                On Error GoTo 0
            Else
                ' Previamente Gestante
                If wsH.Cells(iARH, 11) = "P" Then
                    ' Avisos informativo
                    rCelda.Offset(0, 6) = "Animal previamente reportado como Gestante."
                    ' Registrar Dato
                    wsH.Cells(iARH, 11) = "O" 'Status
                    wsH.Cells(iARH, 12).Clear 'FxSecar
                    wsH.Cells(iARH, 13).Clear 'FxParir
                    wsH2.Cells(iARH2, 3).Clear 'DAbiertos
                End If
        End If
'+++++++
        ws.Cells(BuscarEvento(rCelda.Offset(0, 1), "Serv", wsH.Cells(iARH, 8)), 9) = _
          ws.Cells(BuscarEvento(rCelda.Offset(0, 1), "Serv", wsH.Cells(iARH, 8)), 9) _
          & "-" & sDx
'+++++++
    End If
    If sLocAnimal = "R" Then
        ChecarSexo
        If bFlagError Then Exit Sub
         ' Mínimo un servicio
        If rCelda.Offset(0, 3) = "Gestante" Then
                ' Sin servicios
                If wsR.Cells(iAR, 6) < 1 Then
                    rCelda.Offset(0, 5) = "Animal sin Servicios"
                    bFlagError = True
                    Exit Sub
                End If
                ' Mínimo con 45 días post servicio
                If CDate(rCelda.Offset(0, 0)) - _
                  CDate(wsR.Cells(iAR, 7)) < _
                  Range("Configuracion!C5") Then
                    rCelda.Offset(0, 6) = "Animal con menos de " & _
                      Range("Configuracion!C5") & " días de servicio."
                End If
                ' Último servicio
                If UCase(wsR.Cells(iAR, 8)) = UCase("Calor") Then
                    rCelda.Offset(0, 5) = "Animal sin Servicios."
                    bFlagError = True
                    Exit Sub
                End If
                ' Escribir Datos
                'Servicio-DíasGestación
                sMetadato = Format(wsR.Cells(iAR, 6), "00") & "-" _
                  & Format(CDate(rCelda.Offset(0, 0)) - _
                  CDate(wsR.Cells(iAR, 7)), "000")
                wsR.Cells(iAR, 10) = "P"  'Status
                On Error Resume Next
                With wsR.Cells(iAR, 11)  'FxParir
                    .Value = CDate(wsR.Cells(iAR, 7)) + 273
                    .NumberFormat = "d-mmm-yy"
                End With
                With wsIV.Cells(iARowInfoVital, 11) 'EdadAlParto
                     .Value = Int(((CDate(wsR.Cells(iAR, 11)) _
                      - CDate(wsR.Cells(iAR, 5)))) / 30.4) '& " m"
                End With
                If wsR.Cells(iAR, 12) = "pAb" Then _
                  wsR.Cells(iAR, 13).Clear 'Clave1
                On Error GoTo 0
            Else
                ' Previamente Gestante
                If wsR.Cells(iAR, 10) = "P" Then
                    ' Avisos informativo
                    rCelda.Offset(0, 6) = "Animal previamente reportado como Gestante."
                    ' Registrar Dato
                    wsR.Cells(iAR, 10).Clear 'Status
                    wsR.Cells(iAR, 11).Clear 'FxParir
                    If wsR.Cells(iAR, 12) = "pAb" Then _
                      wsR.Cells(iAR, 13).Clear 'Clave1
                    wsIV.Cells(iARowInfoVital, 11).Clear 'EdadAlParto
                End If
        End If
'+++
        ws.Cells(BuscarEvento(rCelda.Offset(0, 1), "Serv", wsR.Cells(iAR, 7)), 9) = _
          ws.Cells(BuscarEvento(rCelda.Offset(0, 1), "Serv", wsR.Cells(iAR, 7)), 9) _
          & "-" & sDx
'+++
    End If
    ConsecutivoDeEventos
End Sub

Private Sub ImportarServicios()
    Dim iD1S As Long
    If IsEmpty(rCelda.Offset(0, 4)) And CBool(Range("Configuracion!C16")) Then
        rCelda.Offset(0, 5) = "Falta técnico responsable"
        bFlagError = True
        Exit Sub
    End If
    If sLocAnimal = "H" Then
        ' Último Servicio
        If CDate(rCelda.Offset(0, 0)) <= CDate(wsH.Cells(iARH, 8)) Then
            rCelda.Offset(0, 5) = "Fecha del Evento es igual o anterior a la Fecha del " _
                & rCelda.Offset(0, 2) & " registrado"
            bFlagError = True
            Exit Sub
        End If
        If CDate(rCelda.Offset(0, 0)) - CDate(wsH.Cells(iARH, 8)) <= 18 And _
          Not IsEmpty(wsH.Cells(iARH, 8)) Then
            rCelda.Offset(0, 6) = "El intervalo entre Calores o Servicios es menor a 18 días"
        End If
        If CDate(rCelda.Offset(0, 0)) - CDate(wsH.Cells(iARH, 8)) >= 36 And _
          Not IsEmpty(wsH.Cells(iARH, 8)) Then
            rCelda.Offset(0, 6) = "El intervalo entre Calores o Servicios es mayor a 36 días"
        End If
        ' Escribir Datos
        'Serv-DíasÚltimoServicio-DEL
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
        If IsEmpty(wsH.Cells(iARH, 7)) Then sMetadato = "01-000" Else _
          sMetadato = Format(Val(wsH.Cells(iARH, 7)) + 1, "00") & "-" & _
          Format(CDate(rCelda.Offset(0, 0)) - CDate(wsH.Cells(iARH, 8)), "000")
        sMetadato = sMetadato & "-" & _
          Format(CDate(rCelda.Offset(0, 0)) - CDate(wsH.Cells(iARH, 6)), "000")
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
        
'+++++++++++++++++++++++++++++++++++++
'+                                   +
'+       Arreglar este desmadre      +
'+                                   +
'+++++++++++++++++++++++++++++++++++++
        
        If wsH.Cells(iARH, 11) = "P" Then
            wsH.Cells(iARH, 11).Clear 'Status
            wsH.Cells(iARH, 12).Clear 'FxSecarwsH.
            wsH.Cells(iARH, 13).Clear 'FxParir
            wsH.Cells(iARH, 14) = "pAb" 'Clave1
            Set ws = Worksheets("Hato2")
            Desproteger1
            ws.Cells(iARH2, 3).Clear 'dAbiertos
        End If
        '1er Servicio
        If wsH.Cells(iARH, 7) + 1 = 1 Then
            wsH2.Cells(iARH2, 2) = _
              CDate(rCelda.Offset(0, 0)) - _
              CDate((wsH.Cells(iARH, 6))) 'Dias1Serv
            If wsH2.Cells(iARH2, 17) = vbNullString Then _
              wsH2.Cells(iARH2, 17) = _
                CDate(rCelda.Offset(0, 0)) - CDate(wsH.Cells(iARH, 6)) 'd1Calor
        End If
        With wsH.Cells(iARH, 8) 'F.Servicio
            .Value = CDate(rCelda.Offset(0, 0))
            .NumberFormat = "d-mmm-yy"
        End With
        wsH.Cells(iARH, 7) = wsH.Cells(iARH, 7) + 1 'Servicio
        If rCelda.Offset(0, 3) = vbNullString Then
                wsH.Cells(iARH, 9) = "N.D."
            Else
                'Semental
                wsH.Cells(iARH, 9) = UCase(rCelda.Offset(0, 3))
        End If
        If rCelda.Offset(0, 4) = vbNullString Then
                wsH.Cells(iARH, 10) = "N.D."
            Else
                'Tecnico
                wsH.Cells(iARH, 10) = UCase(rCelda.Offset(0, 4))
        End If
        wsH.Cells(iARH, 11).Clear 'Status
    End If
    If sLocAnimal = "R" Then
        ChecarSexo
        If bFlagError Then Exit Sub
        ' Último Servicio
        If CDate(rCelda.Offset(0, 0)) <= CDate(wsR.Cells(iAR, 7)) Then
            rCelda.Offset(0, 5) = "Fecha del Evento es igual o anterior a la Fecha del " _
                & rCelda.Offset(0, 2) & " registrado"
            bFlagError = True
            Exit Sub
        End If
        If CDate(rCelda.Offset(0, 0)) - CDate(wsR.Cells(iAR, 7)) <= 18 And _
          Not IsEmpty(wsR.Cells(iAR, 7)) Then
            rCelda.Offset(0, 6) = "El intervalo entre Calores o Servicios es menor a 18 días"
        End If
        If CDate(rCelda.Offset(0, 0)) - CDate(wsR.Cells(iAR, 7)) >= 36 And _
          Not IsEmpty(wsR.Cells(iAR, 7)) Then
            rCelda.Offset(0, 6) = "El intervalo entre Calores o Servicios es mayor a 36 días"
        End If
        ' Escribir Datos
        'Serv-DíasÚltimoServicio-DEL
        If IsEmpty(wsR.Cells(iAR, 6)) Then sMetadato = "01-000" Else _
          sMetadato = Format(Val(wsR.Cells(iAR, 6)) + 1, "00") _
          & "-" & Format(CDate(rCelda.Offset(0, 0)) - CDate(wsR.Cells(iAR, 7)), "000")
        sMetadato = sMetadato & "-" & _
          Format(CDate(rCelda.Offset(0, 0)) - CDate(wsR.Cells(iAR, 5)), "000")
        If wsR.Cells(iAR, 10) = "P" Then
            wsR.Cells(iAR, 10).Clear 'Status
            wsR.Cells(iAR, 11).Clear 'FxParir
            wsR.Cells(iAR, 12) = "pAb" 'Clave1
            wsIV.Cells(iARowInfoVital, 11).Clear 'EdadAlParto
        End If
        wsR.Cells(iAR, 6) = _
          wsR.Cells(iAR, 6) + 1 'Servicio
        With wsR.Cells(iAR, 7) 'F.Servicio
            .Value = CDate(rCelda.Offset(0, 0))
            .NumberFormat = "d-mmm-yy"
        End With
        If rCelda.Offset(0, 3) = vbNullString Then
                wsR.Cells(iAR, 8) = "N.D."
            Else
                wsR.Cells(iAR, 8) = _
                  UCase(rCelda.Offset(0, 3)) 'Semental
        End If
        If rCelda.Offset(0, 4) = vbNullString Then
                wsR.Cells(iAR, 9) = "N.D."
            Else
                wsR.Cells(iAR, 9) = _
                  UCase(rCelda.Offset(0, 4)) 'Tecnico
        End If
        If wsR.Cells(iAR, 6) = 1 Then _
          wsIV.Cells(iARowInfoVital, 10) = _
          Int((CDate(wsR.Cells(iAR, 7)) _
          - CDate(wsR.Cells(iAR, 5))) / 30.4) 'Edad1Serv
    End If
    If Not CBool(Range("Desarrollador!B6")) Then _
      wsH2.Visible = xlSheetVeryHidden
    ConsecutivoDeEventos
End Sub

Private Sub Alta_Hoja2()
' Alta de animales en Hato2
    Dim iRenglon As Long
    Set ws = Worksheets("Hato2")
    Set wsP = ws
    Application.Run "Desproteger1"
    iRenglon = TamañoTabla("Tabla15") + 2
    ws.Cells(iRenglon, 1) = rCelda.Offset(0, 1)
    If Not CBool(Range("Desarrollador!B6")) Then _
      ws.Visible = xlSheetVeryHidden
End Sub

Private Sub ChecarFechaParto()
    ' Checar errores
    ChecarMismoEvento
    If sLocAnimal = "R" Then Exit Sub
    If Not IsDate(WorksheetFunction.VLookup(rCelda.Offset(0, 1), _
      Range("Tabla1"), 6, False)) _
      And Not IsEmpty(WorksheetFunction.VLookup(rCelda.Offset(0, 1), _
      Range("Tabla1"), 6, False)) Then
        bFlagError = True
        Exit Sub
    End If
End Sub

Private Sub ChecarMismoEvento()
    Dim sEvento As String
        Select Case UCase(rCelda.Offset(0, 2))
            Case Is = "SERVICIO"
                sEvento = "Serv"
            Case Is = "CALOR"
                sEvento = "Calor"
            Case Is = "PRODUCCIÓN"
                sEvento = "Prod"
            'Case Is = "MOVIMIENTO"
                'sEvento = "Mov"
            'Case Is = "ENFERMEDAD"
            Case Is = "REVISIÓN"
                sEvento = "Rev"
            Case Is = "DX GEST."
                sEvento = "DxGst"
            Case Is = "SECAR"
                sEvento = "Seca"
            'Case Is = "NOTA"
            'Case Is = "PARTO"
            'Case Is = "IMANTACIÓN"
            'Case Is = "OTRO"
            'Case Is = "BAJA"
            Case Else
                rCelda.Offset(0, 5) = "Clave no programada"
                'GoTo 5678
        End Select
    If BuscarEvento(rCelda.Offset(0, 1), sEvento, rCelda.Offset(0, 0)) > 0 Then
        rCelda.Offset(0, 5) = "Evento anteriormente capturado"
        bFlagError = True
    End If
End Sub

Private Function CheckFecha(Arete, col)
    ' Revisa que exista una Fecha válida
    CheckFecha = True
    On Error GoTo 1234
    If Not IsDate(CDate(WorksheetFunction.VLookup(Arete, _
      Range("Tabla1"), col, False))) Or _
      IsEmpty(WorksheetFunction.VLookup(Arete, _
      Range("Tabla1"), col, False)) Then
1234:
        CheckFecha = False
        Exit Function
    End If
End Function

Private Sub ConsecutivoDeEventos()
    ' Agregar Información al Consecutivo
    Dim iRenglon As Long
    Set ws = Worksheets("Eventos")
    Application.Run "Desproteger"
    iRenglon = TamañoTabla("Tabla6") + 2
    ws.Cells(iRenglon, 1) = CDbl(rCelda.Offset(0, 1))
    With ws.Cells(iRenglon, 2)
        .Value = CDate(rCelda.Offset(0, 0))
        .NumberFormat = "d-mmm-yy"
    End With
    Select Case UCase(rCelda.Offset(0, 2))
        Case "SERVICIO"
            sEvento = "Serv"
            If rCelda.Offset(0, 3) = vbNullString Then
                    sObserv = "DESCONOCIDO"
                Else
                    sObserv = UCase(rCelda.Offset(0, 3))
            End If
            If rCelda.Offset(0, 4) = vbNullString Then
                    sResp = " "
                Else
                    sResp = UCase(rCelda.Offset(0, 4))
            End If
        Case "CALOR"
            sEvento = "Calor"
            sObserv = "Calor"
            sResp = UCase(rCelda.Offset(0, 4))
        Case "PRODUCCIÓN"
            sEvento = "Prod"
            sObserv = Format(CDbl(rCelda.Offset(0, 3)), "0.0")
        'Case "PESAJE"
            'sEvento = "Pesaje"
            'If IsEmpty(rCelda.Offset(0, 3)) Then
                    'sObserv = Format(CDbl(rCelda.Offset(0, 4)), "0.0")
                'Else
                    'sObserv = Format(CDbl(rCelda.Offset(0, 3)), "0.0") _
                     & "-> " & rCelda.Offset(0, 4)
            'End If
        'Case "MOVIMIENTO"
            'sEvento = "Mov"
            'sObserv = CDbl(rCelda.Offset(0, 3))
        'Case "ENFERMEDAD"
            'sEvento = "Enf-" & UCase(sEnf)
            'sObserv = UCase(Trim(rCelda.Offset(0, 3)))
            'sResp = UCase(Trim(rCelda.Offset(0, 4)))
        Case "REVISIÓN"
            sEvento = "Rev"
            sObserv = Trim(UCase(rCelda.Offset(0, 3)))
            sResp = Trim(UCase(rCelda.Offset(0, 4)))
        Case "DX GEST."
            sEvento = "DxGst"
            If rCelda.Offset(0, 3) = "Gestante" Then
                    sObserv = "Gest"
                Else
                    sObserv = "Vacía"
            End If
            sResp = Trim(UCase(rCelda.Offset(0, 4)))
        Case "SECAR"
            sEvento = "Seca"
        'Case "PARTO"
            'sEvento = "Parto"
            'If rCelda.Offset(0, 3) = "Aborto" Then
                    'sEvento = "Aborto"
                'Else
                    'sEvento = "Parto"
                    'sObserv = UCase(sPartoDet)
            'End If
        'Case "IMANTACIÓN"
            'sEvento = "Iman"
            'sResp = UCase(Me.TextBox4)
        'Case "NOTA"
            'sEvento = "Nota"
            'sObserv = UCase(Trim(Me.TextBox4))
            'sResp = UCase(Trim(Me.TextBox5))
        'Case "OTRO"
            'Select Case Me.ComboBox4
                'Case "Vacunación"
                    'sObserv = Me.ComboBox5
                    'sResp = Me.TextBox6
                'Case "DNB"
                    'sEvento = "DNB"
                    'sObserv = Me.TextBox5 'Me.ComboBox5
                    'sResp = Me.TextBox6
                'Case "TB+"
                    'sEvento = "TB+"
                    'sObserv = Me.TextBox5
                    'sResp = Me.TextBox6
                'Case "K Caseína"
                    'sEvento = "KCaseína"
                    'sObserv = Me.TextBox5
                    'sResp = Me.TextBox6
                'Case "Id Adicional"
                    'sEvento = "Id+"
                    'sObserv = Me.TextBox5
                    'sResp = Me.TextBox6
            'End Select
            'sEvento = Me.ComboBox4
            'sObserv = UCase(Trim(Me.TextBox5))
            'sResp = UCase(Trim(Me.TextBox6))
        'Case "BAJA"
            'sEvento = "Baja"
            'sObserv = Me.ComboBox4
        'Case "Alta"
            'sEvento = "Alta"
    End Select
    With ws
        .Cells(iRenglon, 3) = sEvento
        .Cells(iRenglon, 4) = sObserv
        .Cells(iRenglon, 5) = sResp
        .Cells(iRenglon, 6) = Range("Configuracion!C49")
        .Cells(iRenglon, 7) = Format(Date, "d-mmm-yy")
        .Cells(iRenglon, 8) = Format(Time, "hh:mm")
        .Cells(iRenglon, 9) = sMetadato
    End With
    If CBool(Range("Configuracion!C25")) Then _
      LoginRecord1
    If Not CBool(Range("Desarrollador!B6")) Then _
      ws.Visible = xlSheetVeryHidden
End Sub

Private Sub Desproteger1()
' Desprotege la hoja activa
    If wsH.ProtectContents Then _
      wsH.Unprotect Password:="0246813579"
End Sub

Private Sub ProtegerHoja1()
' Proteger hoja activa, dejando algunas atribuciones
    wsH.Protect Password:="0246813579", _
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

