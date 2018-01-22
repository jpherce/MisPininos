Attribute VB_Name = "ModParametros"
' Ultima Modificacion: 28-Nov-17
' Adición IntervaloCalores
' Corrección en tVacasServidas
Option Explicit
Dim rCelda As Range
Dim dDxPct As Double
Dim iC, iR, iRenglon As Long
Dim ws As Worksheet

Function DUE(Arete_Buscado As Variant, _
  Evento_Buscado As String, Col_Buscada As Long)
' (D)ato (U)ltimo (E)vento
' Devuelve el dato del último evento buscado
' Arete_Buscado
' Evento_Buscado:
' Ejemplo: =DUE(1084,"Serv",1)
    Dim cont, i As Long
    Dim Ocurrencia As Double
    Dim rCelda As Range
    ' Contar las ocurrencias de estos eventos
    Ocurrencia = WorksheetFunction. _
      CountIfs(Range("Tabla6[Arete]"), Arete_Buscado, _
        Range("Tabla6[Clave]"), Evento_Buscado)
    DUE = "27-jun-1959" 'comienzo del mundo :)
    For Each rCelda In Range("Tabla6[Arete]")
       If Val(rCelda.Offset(i, 0)) = _
         Arete_Buscado And _
         rCelda.Offset(i, 2) = Evento_Buscado Then
           cont = cont + 1
           ' Si es la última ocurrencia del evento
           If cont = Ocurrencia Then
               ' Devuelve fecha del ultimo evento
               DUE = rCelda.Offset(0, Col_Buscada)
               GoTo 100
           End If
       End If
    Next
100:
End Function

Function PicoProd(Arete_Buscado As Variant, _
  Prod_DEL As Long)
' Devuelve Pico de Producción (1) o _
  días al pico de producción _
  =PicoProd(1025,2) = días al pico _
  =PicoProd(1025,1) = pico producción
    Dim cont, i, Pico, DEL As Long
    Dim Ocurrencia As Double
    Dim rCelda As Range
    Dim fParto As Date
    fParto = CDate(DUE(Arete_Buscado, "Parto", 1))
    Pico = 0
    DEL = 0
    For Each rCelda In Range("Tabla6[Arete]")
       If Val(rCelda.Offset(i, 0)) = _
         Arete_Buscado And _
         rCelda.Offset(i, 1) >= _
         fParto And _
         rCelda.Offset(i, 2) = _
         "Prod" Then
            If Val(rCelda.Offset(i, 3)) > Pico Then
              Pico = Val(rCelda.Offset(i, 3))
              ' Extrae Días al pico de Metadatos
              DEL = Val(Left(rCelda.Offset(i, 8), 3))
            End If
       End If
    Next
    If Prod_DEL = 2 Then PicoProd = DEL Else _
      PicoProd = Pico
100:
End Function

Function AnimPorParir(Optional Dias, Optional Tabla)
    Dim rCelda, Rango As Range
    Dim iDesde, iHasta As Long
    Dim iContador As Long
    iContador = 0
    Select Case Dias
        Case 30
            iDesde = 243: iHasta = 730
        Case 60
            iDesde = 213: iHasta = 243
        Case 90
            iDesde = 183: iHasta = 213
        Case 120
            iDesde = 153: iHasta = 183
        Case 150
            iDesde = 123: iHasta = 153
        Case 180
            iDesde = 93: iHasta = 123
        Case 210
            iDesde = 63: iHasta = 93
        Case 240
            iDesde = 33: iHasta = 63
        Case 270
            iDesde = 0: iHasta = 33
        Case Else
            iDesde = 0: iHasta = 730
    End Select
    If Tabla = "H" Then
            For Each rCelda In Range("Tabla1[Arete]")
                If rCelda.Offset(0, 7) <= Date - iDesde And _
                  rCelda.Offset(0, 7) >= Date - iHasta And _
                  rCelda.Offset(0, 10) = "P" Then _
                  iContador = iContador + 1
            Next rCelda
        Else
            For Each rCelda In Range("Tabla2[Arete]")
                If rCelda.Offset(0, 6) <= Date - iDesde And _
                  rCelda.Offset(0, 6) >= Date - iHasta And _
                  rCelda.Offset(0, 9) = "P" Then _
                  iContador = iContador + 1
            Next rCelda
    End If
    AnimPorParir = iContador
End Function

Function AnimPorParir2(Optional Dias, Optional Tabla)
    Dim rCelda, Rango As Range
    Dim iDesde, iHasta As Long
    Dim i As Long
    Dim iContador As Long
    Dim dd, MM, m1, m2, m3, m4, m5, m6, m7, m8 As Long
    Dim a(8)
    Dim AA As Long
    m1 = Month(Date)
    If XLMod(m1 + 1, 12) = 0 Then m2 = 12 Else _
      m2 = XLMod(m1 + 1, 12)
    If XLMod(m1 + 2, 12) = 0 Then m3 = 12 Else _
      m3 = XLMod(m1 + 2, 12)
    If XLMod(m1 + 3, 12) = 0 Then m4 = 12 Else _
      m4 = XLMod(m1 + 3, 12)
    If XLMod(m1 + 4, 12) = 0 Then m5 = 12 Else _
      m5 = XLMod(m1 + 4, 12)
    If XLMod(m1 + 5, 12) = 0 Then m6 = 12 Else _
      m6 = XLMod(m1 + 5, 12)
    If XLMod(m1 + 6, 12) = 0 Then m7 = 12 Else _
      m7 = XLMod(m1 + 6, 12)
    If XLMod(m1 + 7, 12) = 0 Then m8 = 12 Else _
              m8 = XLMod(m1 + 7, 12)
    For i = 1 To 8
        a(i) = Year(Date)
    Next i
    Select Case Month(Date)
        Case 6
            For i = 8 To 8
                a(i) = a(i) + 1
            Next i
        Case 7
            For i = 7 To 8
                a(i) = a(i) + 1
            Next i
        Case 8
            For i = 6 To 8
                a(i) = a(i) + 1
            Next i
        Case 9
            For i = 5 To 8
                a(i) = a(i) + 1
            Next i
        Case 10
            For i = 4 To 8
                a(i) = a(i) + 1
            Next i
        Case 11
            For i = 3 To 8
                a(i) = a(i) + 1
            Next i
        Case 12
            For i = 2 To 8
                a(i) = a(i) + 1
            Next i
    End Select
    Select Case Dias
        Case 1
            MM = m1
            AA = a(1)
        Case 2
            MM = m2
            AA = a(2)
        Case 3
            MM = m3
            AA = a(3)
        Case 4
            MM = m4
            AA = a(4)
        Case 5
            MM = m5
            AA = a(5)
        Case 6
            MM = m6
            AA = a(6)
        Case 7
            MM = m7
            AA = a(7)
        Case 8
            MM = m8
            AA = a(8)
    End Select
    Select Case MM
        Case 1, 3, 5, 7, 8, 10, 12
            dd = 31
        Case 2
            dd = 28
        Case Else
            dd = 30
    End Select
    AnimPorParir2 = 0
    iContador = 0
    For Each rCelda In Range("Tabla1[Arete]")
        If rCelda.Offset(0, 12) >= CDate(1 & "," & MM & "," & AA) And _
          rCelda.Offset(0, 12) <= CDate(dd & "," & MM & "," & AA) Then _
          iContador = iContador + 1
    Next rCelda
    For Each rCelda In Range("Tabla2[Arete]")
        If rCelda.Offset(0, 11) >= CDate(1 & "," & MM & "," & AA) And _
          rCelda.Offset(0, 11) <= CDate(dd & "," & MM & "," & AA) Then _
          iContador = iContador + 1
    Next rCelda
    AnimPorParir2 = iContador
End Function

Function AnimPorSecar(Optional Dias)
    Dim rCelda, Rango As Range
    Dim iDesde, iHasta As Long
    Dim i As Long
    Dim iContador As Long
    Dim dd, MM, m1, m2, m3, m4, m5, m6, m7, m8 As Long
    Dim a(8)
    Dim AA As Long
    m1 = Month(Date)
    If XLMod(m1 + 1, 12) = 0 Then m2 = 12 Else _
      m2 = XLMod(m1 + 1, 12)
    If XLMod(m1 + 2, 12) = 0 Then m3 = 12 Else _
      m3 = XLMod(m1 + 2, 12)
    If XLMod(m1 + 3, 12) = 0 Then m4 = 12 Else _
      m4 = XLMod(m1 + 3, 12)
    If XLMod(m1 + 4, 12) = 0 Then m5 = 12 Else _
      m5 = XLMod(m1 + 4, 12)
    If XLMod(m1 + 5, 12) = 0 Then m6 = 12 Else _
      m6 = XLMod(m1 + 5, 12)
    If XLMod(m1 + 6, 12) = 0 Then m7 = 12 Else _
      m7 = XLMod(m1 + 6, 12)
    If XLMod(m1 + 7, 12) = 0 Then m8 = 12 Else _
              m8 = XLMod(m1 + 7, 12)
    For i = 1 To 8
        a(i) = Year(Date)
    Next i
    Select Case Month(Date)
        Case 6
            For i = 8 To 8
                a(i) = a(i) + 1
            Next i
        Case 7
            For i = 7 To 8
                a(i) = a(i) + 1
            Next i
        Case 8
            For i = 6 To 8
                a(i) = a(i) + 1
            Next i
        Case 9
            For i = 5 To 8
                a(i) = a(i) + 1
            Next i
        Case 10
            For i = 4 To 8
                a(i) = a(i) + 1
            Next i
        Case 11
            For i = 3 To 8
                a(i) = a(i) + 1
            Next i
        Case 12
            For i = 2 To 8
                a(i) = a(i) + 1
            Next i
    End Select
    Select Case Dias
        Case 1
            MM = m1
            AA = a(1)
        Case 2
            MM = m2
            AA = a(2)
        Case 3
            MM = m3
            AA = a(3)
        Case 4
            MM = m4
            AA = a(4)
        Case 5
            MM = m5
            AA = a(5)
        Case 6
            MM = m6
            AA = a(6)
        Case 7
            MM = m7
            AA = a(7)
        Case 8
            MM = m8
            AA = a(8)
    End Select
    Select Case MM
        Case 1, 3, 5, 7, 8, 10, 12
            dd = 31
        Case 2
            dd = 28
        Case Else
            dd = 30
    End Select
    AnimPorSecar = 0
    iContador = 0
    For Each rCelda In Range("Tabla1[Arete]")
        If rCelda.Offset(0, 11) >= CDate(1 & "," & MM & "," & AA) And _
          rCelda.Offset(0, 11) <= CDate(dd & "," & MM & "," & AA) Then _
          iContador = iContador + 1
    Next rCelda
    AnimPorSecar = iContador
End Function

Private Sub BorrarHoja()
' Limpia los datos de la Hoja
    Dim i As Long
    Application.ScreenUpdating = True
    Range(Range("Estadísticas!C5"), _
      Range("Estadísticas!L10")).ClearContents
    Range(Range("Estadísticas!C16"), _
      Range("Estadísticas!L21")).ClearContents
    Range(Range("Estadísticas!C27"), _
      Range("Estadísticas!K29")).ClearContents
    Range(Range("Estadísticas!C39"), _
      Range("Estadísticas!H48")).ClearContents
    For i = 1 To 10000: Next i
    Application.ScreenUpdating = False
End Sub

Function BreedingInterval()
' Fuente DRIM5 "Interpreting Reproductive Efficiency Indexes"
    BreedingInterval = (pDAb() - pD1S()) / (pServicios(1, "P") - 1)
End Function

Private Sub CalcularEstadisticas()
    BorrarHoja
    On Error Resume Next
    'Columna B
    'Range("Estadísticas!B1") = Range("Configuracion!C3")
    Range("Estadísticas!B4") = "Parto"
    Range("Estadísticas!B5") = 1
    Range("Estadísticas!B6") = 2
    Range("Estadísticas!B7") = 3
    Range("Estadísticas!B8") = 4
    Range("Estadísticas!B9") = "5+"
    Range("Estadísticas!B14") = "REEMPLAZOS"
    Range("Estadísticas!B15") = "Etapa"
    Range("Estadísticas!B16") = "Lactancia"
    Range("Estadísticas!B17") = "Desarrollo"
    Range("Estadísticas!B18") = "Novillas"
    Range("Estadísticas!B19") = "Vaquillas"
    Range("Estadísticas!B20") = "Vaq. Gestantes"
    Range("Estadísticas!B25") = "ANIMALES POR PARIR"
    Range("Estadísticas!B27") = "Vacas"
    Range("Estadísticas!B28") = "Reemplazos"
    Range("Estadísticas!B37") = "ANIMALES POR CORRAL"
    Range("Estadísticas!B38") = "Corral"
    Range("Estadísticas!B39") = 1
    Range("Estadísticas!B40") = 2
    Range("Estadísticas!B41") = 3
    Range("Estadísticas!B42") = 4
    Range("Estadísticas!B43") = 5
    Range("Estadísticas!B44") = 6
    Range("Estadísticas!B45") = 7
    Range("Estadísticas!B46") = 8
    Range("Estadísticas!B47") = 9
    
    'Columna C
    Range("Estadísticas!C4") = "Número de Animales"
    Range("Estadísticas!C5") = _
      WorksheetFunction.CountIfs(Range("Tabla1[Parto]"), 1)
    Range("Estadísticas!C6") = _
      WorksheetFunction.CountIfs(Range("Tabla1[Parto]"), 2)
    Range("Estadísticas!C7") = _
      WorksheetFunction.CountIfs(Range("Tabla1[Parto]"), 3)
    Range("Estadísticas!C8") = _
      WorksheetFunction.CountIfs(Range("Tabla1[Parto]"), 4)
    Range("Estadísticas!C9") = _
      WorksheetFunction.CountIfs(Range("Tabla1[Parto]"), ">=5")
    Range("Estadísticas!C10") = _
      WorksheetFunction.Sum(Range("Estadísticas!C5:C9"))
    Range("Estadísticas!C15") = "Número de Animales"
    Range("Estadísticas!C16") = tLactantes()
    Range("Estadísticas!C17") = tDesarrollo()
    Range("Estadísticas!C18") = tNovillas()
    Range("Estadísticas!C19") = tVaquillas("O")
    Range("Estadísticas!C20") = tVaquillas("P")
    Range("Estadísticas!C21") = _
      WorksheetFunction.Sum(Range("Estadísticas!C16:C20"))
    Range("Estadísticas!C26") = "30 Días"
    Range("Estadísticas!C27") = AnimPorParir(30, "H")
    Range("Estadísticas!C28") = AnimPorParir(30, "R")
    Range("Estadísticas!C29") = _
      Range("Estadísticas!C27") + Range("Estadísticas!C28")
    Range("Estadísticas!C38") = "Número de Animales"
    Range("Estadísticas!C39") = _
      WorksheetFunction.CountIfs(Range("Tabla1[Corral]"), 1) + _
      WorksheetFunction.CountIfs(Range("Tabla2[Corral]"), 1)
    Range("Estadísticas!C40") = WorksheetFunction.CountIfs _
      (Range("Tabla1[Corral]"), 2) + _
      WorksheetFunction.CountIfs(Range("Tabla2[Corral]"), 2)
    Range("Estadísticas!C41") = _
      WorksheetFunction.CountIfs(Range("Tabla1[Corral]"), 3) + _
      WorksheetFunction.CountIfs(Range("Tabla2[Corral]"), 3)
    Range("Estadísticas!C42") = WorksheetFunction.CountIfs( _
      Range("Tabla1[Corral]"), 4) + _
      WorksheetFunction.CountIfs(Range("Tabla2[Corral]"), 4)
    Range("Estadísticas!C43") = WorksheetFunction.CountIfs( _
      Range("Tabla1[Corral]"), 5) + _
      WorksheetFunction.CountIfs(Range("Tabla2[Corral]"), 5)
    Range("Estadísticas!C44") = WorksheetFunction.CountIfs( _
      Range("Tabla1[Corral]"), 6) + _
      WorksheetFunction.CountIfs(Range("Tabla2[Corral]"), 6)
    Range("Estadísticas!C45") = WorksheetFunction.CountIfs( _
      Range("Tabla1[Corral]"), 7) + _
      WorksheetFunction.CountIfs(Range("Tabla2[Corral]"), 7)
    Range("Estadísticas!C46") = WorksheetFunction.CountIfs( _
      Range("Tabla1[Corral]"), 8) + _
      WorksheetFunction.CountIfs(Range("Tabla2[Corral]"), 8)
    Range("Estadísticas!C47") = WorksheetFunction.CountIfs( _
      Range("Tabla1[Corral]"), 9) + _
      WorksheetFunction.CountIfs(Range("Tabla2[Corral]"), 9)
    Range("Estadísticas!C48") = WorksheetFunction. _
      Sum(Range("Estadísticas!C39:C47"))
    
    'Columna D
    Range("Estadísticas!D4") = "% del Hato"
    Range("Estadísticas!D5") = Format(Range("Estadísticas!C5") / _
      Range("Estadísticas!C10"), "#%")
    Range("Estadísticas!D6") = Format(Range("Estadísticas!C6") / _
      Range("Estadísticas!C10"), "#%")
    Range("Estadísticas!D7") = Format(Range("Estadísticas!C7") / _
      Range("Estadísticas!C10"), "#%")
    Range("Estadísticas!D8") = Format(Range("Estadísticas!C8") / _
      Range("Estadísticas!C10"), "#%")
    Range("Estadísticas!D9") = Format(Range("Estadísticas!C9") / _
      Range("Estadísticas!C10"), "#%")
    Range("Estadísticas!D10") = Format(WorksheetFunction. _
      Sum(Range("Estadísticas!D5:D9")), "#%")
    Range("Estadísticas!D15") = "% de Reemplazos"
    Range("Estadísticas!D16") = Format(Range("Estadísticas!C16") / _
      Range("Estadísticas!C21"), "#.0%")
    Range("Estadísticas!D17") = Format(Range("Estadísticas!C17") / _
      Range("Estadísticas!C21"), "#.0%")
    Range("Estadísticas!D18") = Format(Range("Estadísticas!C18") / _
      Range("Estadísticas!C21"), "#.0%")
    Range("Estadísticas!D19") = Format(Range("Estadísticas!C19") / _
      Range("Estadísticas!C21"), "#.0%")
    Range("Estadísticas!D20") = Format(Range("Estadísticas!C20") / _
      Range("Estadísticas!C21"), "#.0%")
    Range("Estadísticas!D21") = Format(WorksheetFunction. _
      Sum(Range("Estadísticas!D16:D20")), "#%")
    Range("Estadísticas!D26") = "60 Días"
    Range("Estadísticas!D27") = AnimPorParir(60, "H")
    Range("Estadísticas!D28") = AnimPorParir(60, "R")
    Range("Estadísticas!D29") = Format(Range("Estadísticas!D27") + _
      Range("Estadísticas!D28"), "#")
    Range("Estadísticas!D38") = "%"
    Range("Estadísticas!D39") = Format(Range("Estadísticas!C39") / _
      Range("Estadísticas!C48"), "#%")
    Range("Estadísticas!D40") = Format(Range("Estadísticas!C40") / _
      Range("Estadísticas!C48"), "#%")
    Range("Estadísticas!D41") = Format(Range("Estadísticas!C41") / _
      Range("Estadísticas!C48"), "#%")
    Range("Estadísticas!D42") = Format(Range("Estadísticas!C42") / _
      Range("Estadísticas!C48"), "#%")
    Range("Estadísticas!D43") = Format(Range("Estadísticas!C43") / _
      Range("Estadísticas!C48"), "#%")
    Range("Estadísticas!D44") = Format(Range("Estadísticas!C44") / _
      Range("Estadísticas!C48"), "#%")
    Range("Estadísticas!D45") = Format(Range("Estadísticas!C45") / _
      Range("Estadísticas!C48"), "#%")
    Range("Estadísticas!D46") = Format(Range("Estadísticas!C46") / _
      Range("Estadísticas!C48"), "#%")
    Range("Estadísticas!D47") = Format(Range("Estadísticas!C47") / _
      Range("Estadísticas!C48"), "#%")
    Range("Estadísticas!D48") = Format(WorksheetFunction. _
      Sum(Range("Estadísticas!D39:D47")), "#%")

    'Columna E
    Range("Estadísticas!E4") = "Días en Leche"
    'E5
    Range("Estadísticas!E5") = pDEL(1)
    'E6
    Range("Estadísticas!E6") = pDEL(2)
    'E7
    Range("Estadísticas!E7") = pDEL(3)
    'E8
    Range("Estadísticas!E8") = pDEL(4)
    'E9
    Range("Estadísticas!E9") = pDEL(5)
    'E10
    Range("Estadísticas!E10") = Format(pDEL(), "#")
    Range("Estadísticas!E15") = "Edad"
    'E16
    If Not IsError(WorksheetFunction. _
      AverageIfs(Range("Tabla2[Edad2]"), Range("Tabla2[Edad2]"), _
      "<=45")) Then Range("Estadísticas!E16") = _
      Format(WorksheetFunction.AverageIfs(Range("Tabla2[Edad2]"), _
      Range("Tabla2[Edad2]"), "<=45"), "#") & "d"
    'E17
    If Not IsError(WorksheetFunction. _
      AverageIfs(Range("Tabla2[Edad2]"), Range("Tabla2[Edad2]"), _
      ">45", Range("Tabla2[Edad2]"), "<365", _
      Range("Tabla2[Sexo]"), "H")) Then _
      Range("Estadísticas!E17") = Format(Int(WorksheetFunction. _
      AverageIfs(Range("Tabla2[Edad2]"), Range("Tabla2[Edad2]"), _
      ">45", Range("Tabla2[Edad2]"), "<365", _
      Range("Tabla2[Sexo]"), "H") / 30.4), "#") & "m"
    'E18
    If Not IsError(WorksheetFunction. _
      AverageIfs(Range("Tabla2[Edad2]"), Range("Tabla2[Edad2]"), _
      ">365", Range("Tabla2[Edad2]"), "<" & _
      Int(Range("Configuracion!C47") * 30.4), _
      Range("Tabla2[Sexo]"), "H")) Then _
      Range("Estadísticas!E18") = Format(Int(WorksheetFunction. _
      AverageIfs(Range("Tabla2[Edad2]"), Range("Tabla2[Edad2]"), _
      ">365", Range("Tabla2[Edad2]"), "<" & _
      Int(Range("Configuracion!C47") * 30.4), _
      Range("Tabla2[Sexo]"), "H") / 30.4), "#") & "m"
    'E19
    If Not IsError(WorksheetFunction. _
      AverageIfs(Range("Tabla2[Edad2]"), Range("Tabla2[Edad2]"), _
      ">" & Int(Range("Configuracion!C47") * 30.4), _
      Range("Tabla2[Status]"), "<>G", _
      Range("Tabla2[Sexo]"), "H")) Then _
      Range("Estadísticas!E19") = Format(Int(WorksheetFunction. _
      AverageIfs(Range("Tabla2[Edad2]"), _
      Range("Tabla2[Edad2]"), ">" & _
      Int(Range("Configuracion!C47") * 30.4), _
      Range("Tabla2[Status]"), "<>G", _
      Range("Tabla2[Sexo]"), "H") / 30.4), "#") & "m"
    'E20
    If Not IsError(WorksheetFunction. _
      AverageIfs(Range("Tabla2[Edad2]"), Range("Tabla2[Edad2]"), _
      ">" & Int(Range("Configuracion!C47") * 30.4), _
      Range("Tabla2[Status]"), "P", _
      Range("Tabla2[Sexo]"), "H")) Then _
      Range("Estadísticas!E20") = Format(Int(WorksheetFunction. _
      AverageIfs(Range("Tabla2[Edad2]"), _
      Range("Tabla2[Edad2]"), ">" & _
      Int(Range("Configuracion!C47") * 30.4), _
      Range("Tabla2[Status]"), "P", _
      Range("Tabla2[Sexo]"), "H") / 30.4), "#") & "m"
    Range("Estadísticas!E26") = "90 Días"
    Range("Estadísticas!E27") = AnimPorParir(90, "H")
    Range("Estadísticas!E28") = AnimPorParir(90, "R")
    Range("Estadísticas!E29") = Range("Estadísticas!E27") + _
      Range("Estadísticas!E28")
    Range("Estadísticas!E38") = "Días en Leche"
    'E39
    Range("Estadísticas!E39") = pDEL(1)
    'E40
    Range("Estadísticas!E40") = pDEL(2)
    'E41
    Range("Estadísticas!E41") = pDEL(3)
    'E42
    Range("Estadísticas!E42") = pDEL(4)
    'E43
    Range("Estadísticas!E43") = pDEL(5)

    'Columna F
    Range("Estadísticas!F4") = "Producción Promedio"
    'F5
    Range("Estadísticas!F5") = pProd(1)
    'F6
    Range("Estadísticas!F6") = pProd(2)
    'F7
    Range("Estadísticas!F7") = pProd(3)
    'F8
    Range("Estadísticas!F8") = pProd(4)
    'F9
    Range("Estadísticas!F9") = pProd(5)
    'F10
    Range("Estadísticas!F10") = Format(pProd(), "#.#")
    Range("Estadísticas!F15") = "Peso Promedio"
    'On Error Resume Next
    'F16
    If Not IsError(WorksheetFunction. _
      AverageIfs(Range("Tabla2[PesoCorporal]"), _
      Range("Tabla2[Edad2]"), "<=45")) Then _
      Range("Estadísticas!F16") = _
      Format(WorksheetFunction. _
      AverageIfs(Range("Tabla2[PesoCorporal]"), _
      Range("Tabla2[Edad2]"), "<=45"), "#")
    'F17
    If Not IsError(WorksheetFunction. _
      AverageIfs(Range("Tabla2[PesoCorporal]"), _
      Range("Tabla2[Edad2]"), ">45", _
      Range("Tabla2[Edad2]"), "<365", _
      Range("Tabla2[Sexo]"), "H")) Then _
      Range("Estadísticas!F17") = _
      Format(Int(WorksheetFunction. _
      AverageIfs(Range("Tabla2[PesoCorporal]"), _
      Range("Tabla2[Edad2]"), ">45", _
      Range("Tabla2[Edad2]"), "<365", _
      Range("Tabla2[Sexo]"), "H") / 30.4), "#")
    'F18
    If Not IsError(WorksheetFunction. _
      AverageIfs(Range("Tabla2[PesoCorporal]"), _
      Range("Tabla2[Edad2]"), ">365", _
      Range("Tabla2[Edad2]"), "<" & _
      Int(Range("Configuracion!C47") * 30.4), _
      Range("Tabla2[Sexo]"), "H")) Then _
      Range("Estadísticas!F18") = _
      Format(Int(WorksheetFunction. _
      AverageIfs(Range("Tabla2[PesoCorporal]"), _
      Range("Tabla2[Edad2]"), ">365", _
      Range("Tabla2[Edad2]"), "<" & _
      Int(Range("Configuracion!C47") * 30.4), _
      Range("Tabla2[Sexo]"), "H") / 30.4), "#")
    'F19
    If Not IsError(WorksheetFunction. _
      AverageIfs(Range("Tabla2[PesoCorporal]"), _
      Range("Tabla2[Edad2]"), ">" & _
      Int(Range("Configuracion!C47") * 30.4), _
      Range("Tabla2[Status]"), "<>G", _
      Range("Tabla2[Sexo]"), "H")) Then _
      Range("Estadísticas!F19") = _
      Format(Int(WorksheetFunction. _
      AverageIfs(Range("Tabla2[PesoCorporal]"), _
      Range("Tabla2[Edad2]"), ">" & _
      Int(Range("Configuracion!C47") * 30.4), _
      Range("Tabla2[Status]"), "<>G", _
      Range("Tabla2[Sexo]"), "H") / 30.4), "#")
    'F20
    If Not IsError(WorksheetFunction. _
      AverageIfs(Range("Tabla2[PesoCorporal]"), _
      Range("Tabla2[Edad2]"), ">" & _
      Int(Range("Configuracion!C47") * 30.4), _
      Range("Tabla2[Status]"), "P", _
      Range("Tabla2[Sexo]"), "H")) Then _
      Range("Estadísticas!F20") = _
      Format(Int(WorksheetFunction. _
      AverageIfs(Range("Tabla2[PesoCorporal]"), _
      Range("Tabla2[Edad2]"), ">" & _
      Int(Range("Configuracion!C47") * 30.4), _
      Range("Tabla2[Status]"), "P", _
      Range("Tabla2[Sexo]"), "H") / 30.4), "#")
    On Error GoTo 0
    Range("Estadísticas!F26") = "120 Días"
    Range("Estadísticas!F27") = AnimPorParir(120, "H")
    Range("Estadísticas!F28") = AnimPorParir(120, "R")
    Range("Estadísticas!F29") = _
      Range("Estadísticas!F27") + Range("Estadísticas!F28")
    Range("Estadísticas!F38") = "Producción Promedio"
    'F39
    Range("Estadísticas!F39") = pProd(1)
    'F40
    Range("Estadísticas!F40") = pProd(2)
    'F41
    Range("Estadísticas!F41") = pProd(3)
    'F42
    Range("Estadísticas!F42") = pProd(4)
    'F43
    Range("Estadísticas!F43") = pProd(5)
   
    'Columna G
    Range("Estadísticas!G4") = "Prom. Pico de Producción"
    Range("Estadísticas!G26") = "150 Días"
    Range("Estadísticas!G27") = AnimPorParir(150, "H")
    Range("Estadísticas!G28") = AnimPorParir(150, "R")
    Range("Estadísticas!G29") = _
      Range("Estadísticas!G27") + Range("Estadísticas!G28")
    Range("Estadísticas!G38") = "Lactancia mínima"
      
    'Columna H
    Range("Estadísticas!H4") = "Proyección a 305d"
    Range("Estadísticas!H5") = Format(pProy305d(1), "#,#")
    Range("Estadísticas!H6") = Format(pProy305d(2), "#,#")
    Range("Estadísticas!H7") = Format(pProy305d(3), "#,#")
    Range("Estadísticas!H8") = Format(pProy305d(4), "#,#")
    Range("Estadísticas!H9") = Format(pProy305d(5), "#,#")
    'H10
    Range("Estadísticas!H10") = Format(pProy305d(), "#,#")
    Range("Estadísticas!H26") = "180 Días"
    Range("Estadísticas!H27") = AnimPorParir(180, "H")
    Range("Estadísticas!H28") = AnimPorParir(180, "R")
    Range("Estadísticas!H29") = _
      Range("Estadísticas!H27") + Range("Estadísticas!H28")
    Range("Estadísticas!H38") = "Lactancia Máxima"

    'Columna I
    Range("Estadísticas!I4") = "Número de Servicios por Vaca"
    'I5
    Range("Estadísticas!I5") = pServicios(1, , 1)
    'I6
    Range("Estadísticas!I6") = pServicios(1, , 2)
    'I7
    Range("Estadísticas!I7") = pServicios(1, , 3)
    'I8
    Range("Estadísticas!I8") = pServicios(1, , 4)
    'I9
    Range("Estadísticas!I9") = pServicios(1, , 5)
    'I10
    Range("Estadísticas!I10") = Format(pServicios(1), "#.0")
    Range("Estadísticas!I15") = _
      "Número de Servicios por Animal"
    'I19
    Range("Estadísticas!I19") = pServicios(2)
    'I27
    Range("Estadísticas!I27") = AnimPorParir(210, "H")
    'I28
    Range("Estadísticas!I28") = AnimPorParir(210, "R")
    'I29
    Range("Estadísticas!I29") = _
      (Range("Estadísticas!I27") + Range("Estadísticas!I28"))
      
    'Columna J
    Range("Estadísticas!J4") = _
      "Número de Servicios por Concepción"
    'J5
      Range("Estadísticas!J5") = pServicios(1, "P", 1)
    'J6
    Range("Estadísticas!J6") = pServicios(1, "P", 2)
    'J7
    Range("Estadísticas!J7") = pServicios(1, "P", 3)
    'J8
    Range("Estadísticas!J8") = pServicios(1, "P", 4)
    'J9
    Range("Estadísticas!J9") = pServicios(1, "P", 5)
    'J10
    Range("Estadísticas!J10") = Format(pServicios(1, "P"), "#.0")
    Range("Estadísticas!J15") = _
      "Número de Servicios por Concepción"
    'J20
    Range("Estadísticas!J20") = pServicios(2, "P")
    Range("Estadísticas!J26") = "240 Días*"
    'J27
    Range("Estadísticas!J27") = AnimPorParir(240, "H")
    'J28
    Range("Estadísticas!J28") = AnimPorParir(240, "R")
    'J29
    Range("Estadísticas!J29") = _
      (Range("Estadísticas!J27") + Range("Estadísticas!J28"))
    
    'Columna K
    'Range("Estadísticas!K1") = "Situación del Establo al día:"
    Range("Estadísticas!K4") = "Promedio días Abiertos"
    Range("Estadísticas!K5") = Format(pDAb(1), "#")
    Range("Estadísticas!K6") = Format(pDAb(2), "#")
    Range("Estadísticas!K7") = Format(pDAb(3), "#")
    Range("Estadísticas!K8") = Format(pDAb(4), "#")
    Range("Estadísticas!K9") = Format(pDAb(5), "#")
    Range("Estadísticas!K10") = Format(pDAb(), "#")
    Range("Estadísticas!K15") = "Promedio Edad al Parto"
    Range("Estadísticas!K20") = Format(pEdadAlParto, "#") & "m"
    Range("Estadísticas!K26") = "270 Días*"
    Range("Estadísticas!K27") = AnimPorParir(270, "H")
    Range("Estadísticas!K28") = AnimPorParir(270, "R")
    Range("Estadísticas!K29") = _
      (Range("Estadísticas!K27") + Range("Estadísticas!K28"))
    Range("Estadísticas!J30") = _
      "*Incluye animales sin Dx. Gestación"
    'Columna L
    'Range("Estadísticas!L1") = Format(Date, "dd-mmm-yy")
    Range("Estadísticas!L4") = "Promedio días a 1er. Servicio"
    Range("Estadísticas!L5") = Format(pD1S(1), "#")
    Range("Estadísticas!L6") = Format(pD1S(2), "#")
    Range("Estadísticas!L7") = Format(pD1S(3), "#")
    Range("Estadísticas!L8") = Format(pD1S(4), "#")
    Range("Estadísticas!L9") = Format(pD1S(5), "#")
    Range("Estadísticas!L10") = Format(pD1S(), "#")
    Range("Estadísticas!L15") = "Promedio Edad a 1er. Servicio"
    If Not pEdad1Serv() = "ND" Then _
      Range("Estadísticas!L19") = pEdad1Serv() & "m" Else _
      Range("Estadísticas!L19") = pEdad1Serv()
    If Not pEdad1Serv("P") = "ND" Then _
      Range("Estadísticas!L20") = pEdad1Serv("P") & "m" Else _
      Range("Estadísticas!L20") = pEdad1Serv("P")
    LactMinMaxCorral
    'Otros
    On Error GoTo 0
End Sub

Function cvProy305d()
' Calcula el coeficiente de variación de la _
  Proyección a 305 días
    Dim deProy305d, pProy305d As Long
    cvProy305d = "ND"
    If WorksheetFunction.Count(Range("Tabla15[Proy305d]")) > 0 Then
        deProy305d = WorksheetFunction. _
          StDevP(Range("Tabla15[Proy305d]"))
        pProy305d = WorksheetFunction. _
          Average(Range("Tabla15[Proy305d]"))
        cvProy305d = Int((deProy305d / pProy305d) * 100)
    End If
End Function

Function DxGstPositivos(Optional Dias)
' Calcula el % de Dx Gestantes Positivos en un período _
  por los dias. _
  Si los dias = 0 entonces se calculará para el año anterior.
    Dim dTotDx, dDxPositivos As Double
    Dim DesdeFecha As Date
    If IsMissing(Dias) Then
            DesdeFecha = Date - 365
        Else
            Dias = 0
            DesdeFecha = Date - Dias
    End If
    For Each rCelda In Range("Tabla6[Arete]")
        If CDate(rCelda.Offset(0, 1)) >= CDate(DesdeFecha) Then
            If rCelda.Offset(0, 2) = "DxGst" Then
                dTotDx = dTotDx + 1
                If rCelda.Offset(0, 3) = "Gest" Then _
                  dDxPositivos = dDxPositivos + 1
            End If
        End If
    Next rCelda
    If dTotDx = 0 Then DxGstPositivos = 0 Else _
      DxGstPositivos = dDxPositivos / dTotDx * 100
End Function

Private Function Elapsed_Time(start, finish As Date)
      Dim hours, minutes, seconds As Long
      hours = Hour(finish - start)
      minutes = Minute(finish - start)
      seconds = Second(finish - start)
      'Elapsed_Time = Application.Transpose(Array(hours, minutes, seconds))
      Elapsed_Time = Array(hours, minutes, seconds)
End Function

Function FertilidadTecnico(NombreTecnico As String)
' Calcula el % de Gestaciones de un toro determinado
    Dim nAnimServ As Double
    Dim nAnimServG As Double
    Dim rCelda As Range
    FertilidadTecnico = "ND"
    For Each rCelda In Range("Tabla6[Arete]")
        'Si Serv, fecha anterior a la espera voluntaria
        If rCelda.Offset(0, 2) = "Serv" And _
          CDate(rCelda.Offset(0, 1)) <= Date - CDate(Range("Configuracion!C6")) _
          And UCase(rCelda.Offset(0, 4)) = UCase(NombreTecnico) Then
            nAnimServ = nAnimServ + 1
            ' Si animal Gestante
            If InStr(rCelda.Offset(0, 8), "P") Then _
              nAnimServG = nAnimServG + 1
        End If
    Next rCelda
    If nAnimServ > 0 Then FertilidadTecnico = _
      Format(nAnimServG / nAnimServ, "0.0%")
End Function

Function FertilidadToro(NombreToro As String)
' Calcula el % de Gestaciones de un toro determinado
    Dim nAnimServ As Double
    Dim nAnimServG As Double
    Dim rCelda As Range
    FertilidadToro = "ND"
    For Each rCelda In Range("Tabla6[Arete]")
        'Si Serv, fecha anterior a la espera voluntaria
        If rCelda.Offset(0, 2) = "Serv" And _
          CDate(rCelda.Offset(0, 1)) <= Date - CDate(Range("Configuracion!C6")) _
          And UCase(rCelda.Offset(0, 3)) = UCase(NombreToro) Then
            nAnimServ = nAnimServ + 1
            ' Si animal Gestante
            If InStr(rCelda.Offset(0, 8), "P") Then _
              nAnimServG = nAnimServG + 1
        End If
    Next rCelda
    If nAnimServ > 0 Then FertilidadToro = _
      Format(nAnimServG / nAnimServ, "0.0%")
End Function

Function HeatsDetected()
' Fuente DRIM5 "Interpreting Reproductive Efficiency Indexes"
    HeatsDetected = Format((21 / BreedingInterval), "0%")
End Function

Private Sub LactMinMaxCorral()
' Contar Parto,Lactancia Mínima y Maxima por Corral
    Dim LmMC(10, 3)
    Dim i As Long
    LmMC(1, 1) = 1
    LmMC(2, 1) = 2
    LmMC(3, 1) = 3
    LmMC(4, 1) = 4
    LmMC(5, 1) = 5
    LmMC(6, 1) = 6
    LmMC(7, 1) = 7
    LmMC(8, 1) = 8
    LmMC(9, 1) = 9
    LmMC(10, 1) = 10
    ' Inicializar variables
    For i = i To 10
        LmMC(i, 2) = 10
        LmMC(i, 3) = 0
    Next i
    ' Tomar datos y agregarlos a matriz
    For Each rCelda In Range("Tabla1[Corral]")
        If rCelda.Offset(0, 3) < LmMC(Val(rCelda), 2) Then _
          LmMC(Val(rCelda), 2) = rCelda.Offset(0, 3)
        If rCelda.Offset(0, 3) > LmMC(Val(rCelda), 3) Then _
          LmMC(Val(rCelda), 3) = rCelda.Offset(0, 3)
    Next rCelda
    ' Escribir resultados
    Range("Estadísticas!G39") = LmMC(1, 2)
    Range("Estadísticas!H39") = LmMC(1, 3)
    Range("Estadísticas!G40") = LmMC(2, 2)
    Range("Estadísticas!H40") = LmMC(2, 3)
    Range("Estadísticas!G41") = LmMC(3, 2)
    Range("Estadísticas!H41") = LmMC(3, 3)
    Range("Estadísticas!G42") = LmMC(4, 2)
    Range("Estadísticas!H42") = LmMC(4, 3)
    Range("Estadísticas!G43") = LmMC(5, 2)
    Range("Estadísticas!H43") = LmMC(5, 3)
End Sub

Function nCaloresPerdidos()
' Calcula el número de Calores Perdidos
' Fórmula desarrollada por JP
' Se agregan 11 días por desviación del ciclo estral
    nCaloresPerdidos = Format((pDAb() - ( _
      Range("Configuracion!C6") + ((pServicios(1, "P") _
      - 1) * 21) + 11)) / 21, "#.#")
End Function

Function numAbortos()
' Contabiliza los abortos de los ultimos 12 meses
    Dim rCelda As Range
    numAbortos = 0
    For Each rCelda In Range("Tabla6[Arete]")
        If rCelda.Offset(0, 1) >= Date - 365 _
          And rCelda.Offset(0, 2) = "Aborto" Then
            numAbortos = numAbortos + 1
        End If
    Next
End Function

Function nAbortos(Mes, Año)
    ' Contabiliza los abortos del mes
    Dim nA As Long
    Dim fDesde, fHasta As Date
    Dim rCelda As Range
    nA = 0
    fDesde = CDate("1," & Mes & "," & Año)
    fHasta = WorksheetFunction. _
    EoMonth _
    (fDesde, 0)
    For Each rCelda In Range("Tabla6[Arete]")
        If rCelda.Offset(0, 1) >= fDesde And _
          rCelda.Offset(0, 1) <= fHasta And _
          rCelda.Offset(0, 2) = "Aborto" Then
            nA = nA + 1
        End If
    Next
    nAbortos = nA
End Function

Function nPartos(Mes, Año)
    ' Contabiliza los partos del mes
    Dim nA As Long
    Dim fDesde, fHasta As Date
    Dim rCelda As Range
    nA = 0
    fDesde = CDate("1," & Mes & "," & Año)
    fHasta = WorksheetFunction. _
    EoMonth _
    (fDesde, 0)
    For Each rCelda In Range("Tabla6[Arete]")
        If rCelda.Offset(0, 1) >= fDesde And _
          rCelda.Offset(0, 1) <= fHasta And _
          rCelda.Offset(0, 2) = "Parto" Then
            nA = nA + 1
        End If
    Next
    nPartos = nA
End Function

Function nBajas(Mes, Año)
    ' Contabiliza los abortos del mes
    Dim nA As Long
    Dim fDesde, fHasta As Date
    Dim rCelda As Range
    nA = 0
    fDesde = CDate("1," & Mes & "," & Año)
    fHasta = WorksheetFunction. _
    EoMonth _
    (fDesde, 0)
    For Each rCelda In Range("Tabla6[Arete]")
        If rCelda.Offset(0, 1) >= fDesde And _
          rCelda.Offset(0, 1) <= fHasta And _
          rCelda.Offset(0, 2) = "Baja" Then
            nA = nA + 1
        End If
    Next
    nBajas = nA
End Function

Function pctAbortos(Optional Dias)
' Calcula el porcentaje de abortos desde una fecha determinada
    Dim dFecha As Date
    Dim nP, nA As Long
    Dim rCelda As Range
    pctAbortos = 0
    If IsMissing(Dias) Then Dias = 365
    nA = 0
    nP = 0
    For Each rCelda In Range("Tabla6[Arete]")
        If rCelda.Offset(0, 1) >= Date - Dias Then
            Select Case rCelda.Offset(0, 2)
                Case Is = "Parto"
                    nP = nP + 1
                Case Is = "Aborto"
                    nA = nA + 1
            End Select
        End If
    Next
    pctAbortos = nA / (nA + nP)
End Function

Function pctGest1Serv()
' Calcula el % de Gestaciones al Primer Servicio
    Dim nAnim1Serv As Double
    Dim nAnim1ServG As Double
    Dim rCelda As Range
    pctGest1Serv = "ND"
    For Each rCelda In Range("Tabla6[Arete]")
        'Si 1er Serv, fecha anterior a la espera voluntaria
        If rCelda.Offset(0, 2) = "Serv" And _
          CDate(rCelda.Offset(0, 1)) <= Date - CDate(Range("Configuracion!C6")) _
          And Val(Left(rCelda.Offset(0, 8), 2)) = 1 Then
            nAnim1Serv = nAnim1Serv + 1
            ' Si animal Gestante
            If InStr(rCelda.Offset(0, 8), "P") Then _
              nAnim1ServG = nAnim1ServG + 1
        End If
    Next rCelda
    If nAnim1Serv > 0 Then pctGest1Serv = _
      Format(nAnim1ServG / nAnim1Serv, "0.0%")
End Function

Function pD1Calor(Optional Parto, Optional StatusRepro)
' Calcula el promedio de los días al primer calor de anim. en Hato
    Dim rCelda As Range
    Dim iCuenta, iSuma, iPuntero As Long
    iCuenta = 0: iSuma = 0: 'vRes = "ND"
    On Error Resume Next
    ' Todos los Animales Servidos
    If IsMissing(StatusRepro) Then
            ' Todos los Partos
            If IsMissing(Parto) Then
                    iPuntero = 100
                    For Each rCelda In Range("Tabla1[Arete]")
                        ' Si existe el animal en Hato2
                        If Not IsEmpty(WorksheetFunction. _
                          VLookup _
                          (rCelda.Offset(0, 0), _
                          Range("Tabla15"), 1, False)) Then
                            ' Si animal tiene valor buscado
                            If WorksheetFunction. _
                              VLookup _
                              (rCelda.Offset(0, 0), _
                              Range("Tabla15"), 17, False) > 0 Then
                                GoTo RealizarCalculos
100:
                            End If
                        End If 'If Not IsEmpty(
                    Next rCelda
                Else
                ' Animales con determinado parto
                    iPuntero = 110
                    For Each rCelda In Range("Tabla1[Arete]")
                        If rCelda.Offset(0, 4) = Parto Then
                            ' Si existe el animal en Hato2
                            If Not IsEmpty(WorksheetFunction. _
                              VLookup _
                              (rCelda.Offset(0, 0), _
                              Range("Tabla15"), 1, False)) Then
                                ' Si animal tiene valor buscado
                                If WorksheetFunction. _
                                  VLookup _
                                  (rCelda.Offset(0, 0), _
                                  Range("Tabla15"), 17, False) > 0 Then
                                    GoTo RealizarCalculos
110:
                                End If
                            End If 'IsEmpty
                        End If 'rCelda
                    Next rCelda
            End If 'If IsMissing(Parto)
        Else
        ' Todos los Animales Servidos y Gestantes
            ' Cualquier Parto y Gestante
            If IsMissing(Parto) Then
                    iPuntero = 101
                    For Each rCelda In Range("Tabla1[Arete]")
                        ' Si están Gestantes
                        If rCelda.Offset(0, 10) = "P" Then
                            ' Si existe el animal en Hato2
                            If Not IsEmpty(WorksheetFunction. _
                              VLookup _
                              (rCelda.Offset(0, 0), _
                              Range("Tabla15"), 1, False)) Then
                                ' Si animal tiene valor buscado
                                If WorksheetFunction. _
                                  VLookup _
                                  (rCelda.Offset(0, 0), _
                                  Range("Tabla15"), 17, False) > 0 Then
                                    GoTo RealizarCalculos
101:
                                End If
                            End If
                        End If
                    Next rCelda
                Else
                    ' Animales con Determinado Parto y Gestantes
                    iPuntero = 111
                    For Each rCelda In Range("Tabla1[Arete]")
                        ' Si es del mismo parto
                        If rCelda.Offset(0, 4) = Parto And _
                          rCelda.Offset(0, 10) = "P" Then
                            ' Si existe el animal en Hato2
                            If Not IsEmpty(WorksheetFunction. _
                              VLookup _
                              (rCelda.Offset(0, 0), _
                              Range("Tabla15"), 1, False)) Then
                                ' Si animal tiene valor buscado
                                If WorksheetFunction. _
                                  VLookup _
                                  (rCelda.Offset(0, 0), _
                                  Range("Tabla15"), 17, False) > 0 Then
                                    GoTo RealizarCalculos
111:
                                End If
                            End If
                        End If
                    Next rCelda
            End If 'If IsMissing(Parto)
    End If 'If IsMissing(Status)
    pD1Calor = iSuma / iCuenta
    On Error GoTo 0
    Exit Function

RealizarCalculos:
    ' Acumular Datos
    iSuma = iSuma + WorksheetFunction. _
      VLookup _
      (rCelda.Offset(0, 0), Range("Tabla15"), 17, False)
    ' Contar Eventos
    iCuenta = iCuenta + 1
    Select Case iPuntero
        Case Is = 100
            GoTo 100
        Case Is = 101
            GoTo 101
        Case Is = 110
            GoTo 110
        Case Is = 111
            GoTo 111
    End Select
End Function

Function pD1S(Optional Parto, Optional StatusRepro)
' Calcula promedio de días a 1er servicio de anim. en hato.
    Dim rCelda As Range
    Dim iCuenta, iSuma, iPuntero As Long
    iCuenta = 0: iSuma = 0: 'vRes = "ND"
    On Error Resume Next
    ' Todos los Animales Servidos
    If IsMissing(StatusRepro) Then
            ' Todos los Partos
            If IsMissing(Parto) Then
                    iPuntero = 100
                    For Each rCelda In Range("Tabla1[Arete]")
                        ' Si existe el animal en Hato2
                        If Not IsEmpty(WorksheetFunction. _
                          VLookup _
                          (rCelda.Offset(0, 0), _
                          Range("Tabla15"), 1, False)) Then
                            ' Si animal tiene valor buscado
                            If WorksheetFunction. _
                              VLookup _
                              (rCelda.Offset(0, 0), _
                              Range("Tabla15"), 2, False) > 0 Then
                                GoTo RealizarCalculos
100:
                            End If
                        End If 'If Not IsEmpty(
                    Next rCelda
                Else
                ' Animales con determinado parto
                    iPuntero = 110
                    For Each rCelda In Range("Tabla1[Arete]")
                        If rCelda.Offset(0, 4) = Parto Then
                            ' Si existe el animal en Hato2
                            If Not IsEmpty(WorksheetFunction. _
                              VLookup _
                              (rCelda.Offset(0, 0), _
                              Range("Tabla15"), 1, False)) Then
                                ' Si animal tiene valor buscado
                                If WorksheetFunction. _
                                  VLookup _
                                  (rCelda.Offset(0, 0), _
                                  Range("Tabla15"), 2, False) > 0 Then
                                    GoTo RealizarCalculos
110:
                                End If
                            End If 'IsEmpty
                        End If 'rCelda
                    Next rCelda
            End If 'If IsMissing(Parto)
        Else
        ' Todos los Animales Servidos y Gestantes
            ' Cualquier Parto y Gestante
            If IsMissing(Parto) Then
                    iPuntero = 101
                    For Each rCelda In Range("Tabla1[Arete]")
                        ' Si están Gestantes
                        If rCelda.Offset(0, 10) = "P" Then
                            ' Si existe el animal en Hato2
                            If Not IsEmpty(WorksheetFunction. _
                              VLookup _
                              (rCelda.Offset(0, 0), _
                              Range("Tabla15"), 1, False)) Then
                                ' Si animal tiene valor buscado
                                If WorksheetFunction. _
                                  VLookup _
                                  (rCelda.Offset(0, 0), _
                                  Range("Tabla15"), 2, False) > 0 Then
                                    GoTo RealizarCalculos
101:
                                End If
                            End If
                        End If
                    Next rCelda
                Else
                    ' Animales con Determinado Parto y Gestantes
                    iPuntero = 111
                    For Each rCelda In Range("Tabla1[Arete]")
                        ' Si es del mismo parto
                        If rCelda.Offset(0, 4) = Parto And _
                          rCelda.Offset(0, 10) = "P" Then
                            ' Si existe el animal en Hato2
                            If Not IsEmpty(WorksheetFunction. _
                              VLookup _
                              (rCelda.Offset(0, 0), _
                              Range("Tabla15"), 1, False)) Then
                                ' Si animal tiene valor buscado
                                If WorksheetFunction. _
                                  VLookup _
                                  (rCelda.Offset(0, 0), _
                                  Range("Tabla15"), 2, False) > 0 Then
                                    GoTo RealizarCalculos
111:
                                End If
                            End If
                        End If
                    Next rCelda
            End If 'If IsMissing(Parto)
    End If 'If IsMissing(Status)
    pD1S = iSuma / iCuenta
    On Error GoTo 0
    Exit Function

RealizarCalculos:
    ' Acumular Datos
    iSuma = iSuma + WorksheetFunction. _
      VLookup _
      (rCelda.Offset(0, 0), Range("Tabla15"), 2, False)
    ' Contar Eventos
    iCuenta = iCuenta + 1
    Select Case iPuntero
        Case Is = 100
            GoTo 100
        Case Is = 101
            GoTo 101
        Case Is = 110
            GoTo 110
        Case Is = 111
            GoTo 111
    End Select
End Function

Function pDAb(Optional Parto)
' Calcula promedio de días abiertos, optativo por parto
    Dim rCelda As Range
    Dim iCuenta, iSuma, iPuntero As Long
    iCuenta = 0: iSuma = 0: 'vRes = "ND"
    On Error Resume Next
' Todos los Animales Servidos y Gestantes
    ' Cualquier Parto y Gestante
    If IsMissing(Parto) Then
            iPuntero = 101
            For Each rCelda In Range("Tabla1[Arete]")
                ' Si están Gestantes
                If rCelda.Offset(0, 10) = "P" Then
                    ' Si existe el animal en Hato2
                    If Not IsEmpty(WorksheetFunction. _
                      VLookup _
                      (rCelda.Offset(0, 0), _
                      Range("Tabla15"), 1, False)) Then
                        ' Si animal tiene valor buscado
                        If WorksheetFunction. _
                          VLookup _
                          (rCelda.Offset(0, 0), _
                          Range("Tabla15"), 3, False) > 0 Then
                            GoTo RealizarCalculos
101:
                        End If
                    End If
                End If
            Next rCelda
        Else
            ' Animales con Determinado Parto y Gestantes
            iPuntero = 111
            For Each rCelda In Range("Tabla1[Arete]")
                ' Si es del mismo parto
                If rCelda.Offset(0, 4) = Parto And _
                  rCelda.Offset(0, 10) = "P" Then
                    ' Si existe el animal en Hato2
                    If Not IsEmpty(WorksheetFunction. _
                      VLookup _
                      (rCelda.Offset(0, 0), _
                      Range("Tabla15"), 1, False)) Then
                        ' Si animal tiene valor buscado
                        If WorksheetFunction. _
                          VLookup _
                          (rCelda.Offset(0, 0), _
                          Range("Tabla15"), 3, False) > 0 Then
                            GoTo RealizarCalculos
111:
                        End If
                    End If
                End If
            Next rCelda
    End If 'If IsMissing(Parto)
    pDAb = iSuma / iCuenta
    On Error GoTo 0
    Exit Function

RealizarCalculos:
    ' Acumular Datos
    iSuma = iSuma + WorksheetFunction. _
      VLookup _
      (rCelda.Offset(0, 0), Range("Tabla15"), 3, False)
    ' Contar Eventos
    iCuenta = iCuenta + 1
    Select Case iPuntero
        Case Is = 101
            GoTo 101
        Case Is = 111
            GoTo 111
    End Select
End Function

Function pDEL(Optional Parto)
' Promedio de Días en Leche por Parto
' parto = 0 para promedio de todo el hato
    If IsMissing(Parto) Then Parto = ">=1"
    If Val(Parto) > 3 Then Parto = ">3"
    pDEL = WorksheetFunction. _
      AverageIfs _
      ( _
      Range("Tabla1[DEL]"), _
      Range("Tabla1[Parto]"), Parto _
      )
End Function

Function pProd(Optional Parto)
' Promedio de Producción por Parto
' pProd() para todo los partos
    If IsMissing(Parto) Then
            pProd = WorksheetFunction. _
            Average _
              ( _
              Range("Tabla1[Prod.]") _
              )
        Else
            If Val(Parto) > 3 Then Parto = ">3"
            pProd = WorksheetFunction. _
            AverageIfs _
              ( _
              Range("Tabla1[Prod.]"), _
              Range("Tabla1[Parto]"), Parto _
              )
    End If
100:
End Function

Function pEdad1Serv(Optional StatusRepro As String)
' Promedia la Edad al Primer Servicio de reemplazos
' Entrar P para calcular las gestantes
    Dim nAnim, nSumDias As Long
    Dim rCelda As Range
    pEdad1Serv = "ND"
    nAnim = 0: nSumDias = 0
    If UCase(StatusRepro) = "P" Then
    'If Not IsMissing(Status) Then
            For Each rCelda In Range("Tabla2[Arete]")
                If rCelda.Offset(0, 5) > 0 And _
                  rCelda.Offset(0, 9) = "P" Then
                    nAnim = nAnim + 1
                    nSumDias = nSumDias + WorksheetFunction. _
                      VLookup _
                      (Val(rCelda), Range("Tabla8"), 10, False)
                End If
            Next rCelda
        Else
            For Each rCelda In Range("Tabla2[Arete]")
                If rCelda.Offset(0, 5) > 0 And _
                  Not rCelda.Offset(0, 9) = "P" Then
                    nAnim = nAnim + 1
                    nSumDias = nSumDias + WorksheetFunction. _
                      VLookup _
                      (Val(rCelda), Range("Tabla8"), 10, False)
                End If
            Next rCelda
    End If
    If nAnim > 0 Then pEdad1Serv = Int(nSumDias / nAnim)
End Function

Function pEdadAlParto()
' Promedia la Edad al Parto de reemplazos
    pEdadAlParto = "ND"
    If WorksheetFunction. _
      Count _
      ( _
      Range("Tabla8[Edad1Parto]" _
      ) _
      ) > 0 Then _
      pEdadAlParto = WorksheetFunction. _
      Average _
      ( _
      Range("Tabla8[Edad1Parto]" _
      ) _
      )
End Function

Function pProdLinea()
' Promedio de Producción en Hato
    If Not IsError(WorksheetFunction. _
      Average _
      ( _
      Range("Tabla1[Prod.]") _
      ) _
      ) Then _
      pProdLinea = WorksheetFunction. _
      Average _
      ( _
      Range("Tabla1[Prod.]" _
      ) _
      )
End Function

Function pProy305d(Optional Parto)
' Promedio de Proyección a 305 días en Hato
    Dim rCelda As Range
    Dim iii, nNumAnim As Double
    iii = 0: nNumAnim = 0
    pProy305d = "ND"
    If IsMissing(Parto) Then
            If WorksheetFunction. _
              Count _
              ( _
              Range("Tabla15[Proy305d]") _
              ) _
              > 0 Then _
              pProy305d = Int(WorksheetFunction. _
              Average _
              ( _
              Range("Tabla15[Proy305d]") _
              ) _
              / 100) * 100
        Else
            On Error Resume Next
            For Each rCelda In Range("Tabla1[Arete]")
                If rCelda.Offset(0, 4) = Parto Then
                    If Not IsEmpty(WorksheetFunction. _
                      VLookup _
                      (rCelda.Offset(0, 0), _
                      Range("Tabla15"), 1, False)) Then
                        ' Si animal tiene valor buscado
                        If WorksheetFunction. _
                          VLookup _
                          (rCelda.Offset(0, 0), _
                          Range("Tabla15"), 15, False) > 0 Then
                            iii = iii + WorksheetFunction. _
                              VLookup _
                              (rCelda.Offset(0, 0), _
                              Range("Tabla15"), 15, False)
                            nNumAnim = nNumAnim + 1
                        End If
                    End If
                End If
            Next
            On Error GoTo 0
            If nNumAnim > 0 Then _
              pProy305d = Int((iii / nNumAnim) / 100) * 100
    End If
End Function

Function ProdAcumulada(Arete)
' Extrae la producción mensual de la vaca especificada
    Dim rCelda As Range
    ProdAcumulada = "ND"
    For Each rCelda In Range("Tabla15[Arete]")
         If rCelda.Offset(0, 0) = Arete Then
            ProdAcumulada = rCelda.Offset(0, 13)
        End If
    Next rCelda
End Function

Function IntervaloCalores(Optional Animal)
' Promedio de intervalo entre calores de las vacas existentes
' IntervaloCalores() todos los animales
' IntervaloCalores(1) sólo vacas
' IntervaloCalores(2) sólo reemplazos
' Debido al loop que ejecuta en Eventos, el rendimiento del _
  sistema se ve dramáticamente disminuido
    Dim rCelda As Range
    Dim n, i As Long
    Dim s As String
    IntervaloCalores = 0
    n = 0
    i = 0
    If IsMissing(Animal) Then Animal = 0
    If Animal = 0 Or Animal = 1 Then
        For Each rCelda In Range("Tabla1[Arete]")
            If rCelda.Offset(0, 8) <> "" Then
                If rCelda.Offset(0, 8) = "Calor" Then s = "Calor" Else s = "Serv"
                    If Val(Mid(DUE(rCelda.Offset(0, 0), s, 8), 4, 3)) > 0 Then
                        i = i + Val(Mid(DUE(rCelda.Offset(0, 0), s, 8), 4, 3))
                        n = n + 1
                    End If
            End If
        Next rCelda
    End If
    If Animal = 0 Or Animal = 2 Then
        For Each rCelda In Range("Tabla2[Arete]")
                If rCelda.Offset(0, 7) <> "" Then
                If rCelda.Offset(0, 7) = "Calor" Then s = "Calor" Else s = "Serv"
                If Val(Mid(DUE(rCelda.Offset(0, 0), s, 8), 4, 3)) > 0 Then
                    i = i + Val(Mid(DUE(rCelda.Offset(0, 0), s, 8), 4, 3))
                    n = n + 1
                End If
            End If
        Next rCelda
    End If
    If n = 0 Then IntervaloCalores Else IntervaloCalores = i / n
End Function

Function PromParam(Tipo As String, Optional Parto)
' Promedio de Proyección a 305 días      en Hato
    Dim rCelda As Range
    Dim iii, col, nNumAnim As Double
    iii = 0: nNumAnim = 0
    PromParam = "ND"
    On Error GoTo 100
    Select Case UCase(Tipo)
        Case UCase("d1S")
            PromParam = WorksheetFunction. _
              Average _
              ( _
              Range("Tabla15[d1S]") _
              )
            col = 2 'columna B
        Case UCase("dAbiertos")
            PromParam = WorksheetFunction. _
              Average _
              ( _
              Range("Tabla15[dAbiertos]") _
              )
            col = 3 'columna C
        Case UCase("ProdAcum")
            PromParam = WorksheetFunction. _
              Average _
              ( _
              Range("Tabla15[ProdAcum]") _
              )
            col = 14 'columna N
        Case Is = UCase("Proy305d")
            PromParam = WorksheetFunction. _
              Average _
              ( _
              Range("Tabla15[Proy305d]") _
              )
            col = 15 'columna O
        Case UCase("d1Calor")
            PromParam = WorksheetFunction. _
              Average _
              ( _
              Range("Tabla15[d1Calor]") _
              )
            col = 16 'columna Q
        Case UCase("Persistencia")
            PromParam = WorksheetFunction. _
              Average _
              ( _
              Range("Tabla15[Persistencia]") _
              )
            col = 19 'columna S
        Case UCase("PicoProd")
            PromParam = WorksheetFunction. _
              Average _
              ( _
              Range("Tabla15[PicoProd]") _
              )
            col = 20 'columna T
        Case UCase("DiasPico")
            PromParam = WorksheetFunction. _
              Average _
              ( _
              Range("Tabla15[DiasPico]") _
              )
            col = 21 'columna U
    End Select
    On Error GoTo 0
    If IsMissing(Parto) Then
            GoTo 100
        Else
            ' Sólo se individualiza hasta parto 3
            If Parto >= 4 Then Parto = 4
            On Error Resume Next
            For Each rCelda In Range("Tabla1[Arete]")
                If rCelda.Offset(0, 4) = Parto Then
                    If Not IsEmpty(WorksheetFunction. _
                      VLookup _
                      ( _
                      rCelda.Offset(0, 0), _
                      Range("Tabla15"), col, False _
                      ) _
                      ) Then
                        ' Si animal tiene valor buscado
                        If WorksheetFunction. _
                          VLookup _
                          ( _
                          rCelda.Offset(0, 0), _
                          Range("Tabla15"), col, False _
                          ) > 0 Then
                            iii = iii + WorksheetFunction. _
                              VLookup _
                              ( _
                              rCelda.Offset(0, 0), _
                              Range("Tabla15"), col, False _
                              )
                            nNumAnim = nNumAnim + 1
                        End If
                    End If
                End If
            Next
            On Error GoTo 0
100:
            If nNumAnim > 0 Then _
              PromParam = (iii / nNumAnim)
    End If
End Function

Function ProduccionVaca(Arete As Double, Dias As Double)
' Extrae la producción mensual de la vaca especificada
    Dim rCelda As Range
    Dim iCol As Long
    ProduccionVaca = "ND"
    For Each rCelda In Range("Tabla15[Arete]")
        If rCelda.Offset(0, 0) = Arete Then
            Select Case Dias
                Case Is = 30
                    iCol = 3 'Col C
                Case Is = 60
                    iCol = 4 'Col D
                Case Is = 90
                    iCol = 5 'Col E
                Case Is = 120
                    iCol = 6 'Col F
                Case Is = 150
                    iCol = 7 'Col G
                Case Is = 180
                    iCol = 8 'Col H
                Case Is = 210
                    iCol = 9 'Col I
                Case Is = 240
                    iCol = 10 'Col J
                Case Is = 270
                    iCol = 11 'Col K
                Case Is = 300
                    iCol = 12 'Col L
                Case Else
                    iCol = 20
            End Select
            ProduccionVaca = rCelda.Offset(0, iCol)
        End If
    Next rCelda
End Function

Function pServicios(Animales As Long, _
  Optional StatusRepro As String, Optional Parto As Long)
    Dim rCelda As Range
    Dim iii, nNumAnim As Double
    iii = 0: nNumAnim = 0
    pServicios = "ND"
    Select Case Animales
        Case Is = 1 'Vacas
            Select Case StatusRepro
                Case Is = "P"
                    If Parto = 0 Then
                            For Each rCelda In Range("Tabla1[Arete]")
                                If rCelda.Offset(0, 6) > 0 And _
                                  rCelda.Offset(0, 10) = "P" And _
                                  Not rCelda.Offset(0, 13) = "DNB" Then
                                    iii = iii + rCelda.Offset(0, 6)
                                    nNumAnim = nNumAnim + 1
                                End If
                            Next
                        Else
                            For Each rCelda In Range("Tabla1[Arete]")
                                If rCelda.Offset(0, 4) = Parto And _
                                  rCelda.Offset(0, 6) > 0 And _
                                  rCelda.Offset(0, 10) = "P" And _
                                  Not rCelda.Offset(0, 13) = "DNB" Then
                                    iii = iii + rCelda.Offset(0, 6)
                                    nNumAnim = nNumAnim + 1
                                End If
                            Next
                    End If
                Case Is = "O", vbNullString
                    If Parto = 0 Then
                            For Each rCelda In Range("Tabla1[Arete]")
                                If rCelda.Offset(0, 6) > 0 And _
                                  Not rCelda.Offset(0, 10) = "P" And _
                                  Not rCelda.Offset(0, 13) = "DNB" Then
                                    iii = iii + rCelda.Offset(0, 6)
                                    nNumAnim = nNumAnim + 1
                                End If
                            Next
                        Else
                            For Each rCelda In Range("Tabla1[Arete]")
                                If rCelda.Offset(0, 4) = Parto And _
                                  rCelda.Offset(0, 6) > 0 And _
                                  Not rCelda.Offset(0, 10) = "P" And _
                                  Not rCelda.Offset(0, 13) = "DNB" Then
                                    iii = iii + rCelda.Offset(0, 6)
                                    nNumAnim = nNumAnim + 1
                                End If
                            Next
                    End If
            End Select
        Case Is = 2 'Reemplazos
           Select Case StatusRepro
                Case Is = "P"
                    For Each rCelda In Range("Tabla2[Arete]")
                        If rCelda.Offset(0, 5) > 0 And _
                          rCelda.Offset(0, 9) = "P" And _
                          Not rCelda.Offset(0, 11) = "DNB" Then
                          iii = iii + rCelda.Offset(0, 5)
                          nNumAnim = nNumAnim + 1
                        End If
                    Next
                Case Is = "O", vbNullString
                    For Each rCelda In Range("Tabla2[Arete]")
                        If rCelda.Offset(0, 5) > 0 And _
                          Not rCelda.Offset(0, 9) = "P" And _
                          Not rCelda.Offset(0, 11) = "DNB" Then
                          iii = iii + rCelda.Offset(0, 5)
                          nNumAnim = nNumAnim + 1
                        End If
                    Next
            End Select
    End Select
    If nNumAnim > 0 Then pServicios = iii / nNumAnim
End Function

Function tAnimales(Optional Animales)
    Select Case Animales
        Case IsMissing(Animales)
            tAnimales = WorksheetFunction. _
              Count _
              ( _
              Range("Tabla1[Arete]") _
              ) + _
              WorksheetFunction. _
              CountIfs _
              ( _
              Range("Tabla2[Sexo]"), "H" _
              )
        Case Is = 1
            tAnimales = WorksheetFunction. _
              Count _
              ( _
              Range("Tabla1[Arete]") _
              )
        Case Is = 2
            tAnimales = WorksheetFunction. _
              CountIfs _
              ( _
              Range("Tabla2[Sexo]"), "H" _
              )
    End Select
End Function

Function tAnimCorral(Corral As Variant, _
  Optional Animales)
' Totaliza el número de animales por corral _
  y por tipo de animal, _
  0 o en blanco para todos los animales del hato, _
  1 para vacas y 2 para reemplazos
    Dim nNumAnim As Long
    Select Case Animales
        Case Is = 0 Or IsMissing(Animales)
1234:
            GoTo 2341
            nNumAnim = tAnimCorral
            GoTo 3412
            nNumAnim = nNumAnim + tAnimCorral
            tAnimCorral = nNumAnim
            Exit Function
        Case Is = 1
            GoTo 2341
            Exit Function
        Case Is = 2
            GoTo 3412
    End Select
2341:
    tAnimCorral = _
              WorksheetFunction. _
              CountIfs _
              ( _
              Range("Tabla1[Corral]"), Corral _
              )
3412:
    tAnimCorral = _
              WorksheetFunction. _
              CountIfs _
              ( _
              Range("Tabla2[Corral]"), Corral _
              )
End Function

Function TasaEmbarazo()
' Fuente: "La Evaluación de tasa de embarazo sobre la _
  fertilidad de la vaca" _
  www.aipl.usda.gov/reference/fertility/DPR-rpt.es.htm
  TasaEmbarazo = 21 / (pDAb() - Range("Configuracion!C6") + 11)
End Function

Function tDesarrollo()
' Totaliza el número de hembras en Desarrollo
    Dim rCelda As Range
    tDesarrollo = 0
    For Each rCelda In Range("Tabla2[Arete]")
        If Date - CDate(rCelda.Offset(0, 4)) <= 365 And _
          Date - CDate(rCelda.Offset(0, 4)) > 45 And _
          rCelda.Offset(0, 13) = "H" Then _
          tDesarrollo = tDesarrollo + 1
    Next
End Function

Function tLactantes()
' Totaliza el número de hembras Lactantes
    Dim rCelda As Range
    tLactantes = 0
    For Each rCelda In Range("Tabla2[Arete]")
        If Date - CDate(rCelda.Offset(0, 4)) <= 45 And _
        rCelda.Offset(0, 13) = "H" Then _
          tLactantes = tLactantes + 1
    Next
End Function

Function tNovillas()
' Totaliza el número de Novillas
    Dim rCelda As Range
    tNovillas = 0
    For Each rCelda In Range("Tabla2[Arete]")
        If Date - CDate(rCelda.Offset(0, 4)) > 365 And _
          Date - CDate(rCelda.Offset(0, 4)) < _
          Int(Range("Configuracion!C47") * 30.4) And _
          rCelda.Offset(0, 13) = "H" Then _
          tNovillas = tNovillas + 1
    Next
End Function

Function tProblema(Optional Animales)
' 0 o en blanco para todos los animales del hato, _
  1 para vacas y 2 para reemplazos
    tProblema = 0
    Select Case Animales
        Case Is = 0 Or IsMissing(Animales)
            GoTo 2341
            GoTo 3412
            Exit Function
        Case Is = 1
            GoTo 2341
1234:
            Exit Function
        Case Is = 2
            GoTo 3412
    End Select

2341:
    For Each rCelda In Range("Tabla1[Arete]")
        If rCelda.Offset(0, 3) >= Range("Configuracion!C6") + 21 _
          And IsEmpty(rCelda.Offset(0, 7)) And _
          Not rCelda.Offset(0, 13) = "DNB" Then _
          tProblema = tProblema + 1
    Next
    If Animales = 1 Then GoTo 1234
3412:
    For Each rCelda In Range("Tabla2[Arete]")
        'Edad, F.Servicio, not DNB, H
        If Date - CDate(rCelda.Offset(0, 4)) > _
          Int(Range("Configuracion!C47") * 30.4) _
          And IsEmpty(rCelda.Offset(0, 6)) And _
          Not rCelda.Offset(0, 11) = "DNB" And _
          rCelda.Offset(0, 13) = "H" Then _
          tProblema = tProblema + 1
    Next
End Function

Function tRepetidoras(Optional Animales)
' 0 o en blanco para todos los animales del hato, _
  1 para vacas y 2 para reemplazos
    tRepetidoras = 0
    Select Case Animales
        Case Is = 0 Or IsMissing(Animales)
            GoTo 2341
            GoTo 3412
            Exit Function
        Case Is = 1
            GoTo 2341
1234:
            Exit Function
        Case Is = 2
            GoTo 3412
    End Select

2341:
    For Each rCelda In Range("Tabla1[Arete]")
        If rCelda.Offset(0, 6) > 3 _
          And Not rCelda.Offset(0, 10) = "P" And _
          Not rCelda.Offset(0, 13) = "DNB" Then _
          tRepetidoras = tRepetidoras + 1
    Next
    If Animales = 1 Then GoTo 1234
3412:
    For Each rCelda In Range("Tabla2[Arete]")
        If rCelda.Offset(0, 5) > 3 _
          And Not rCelda.Offset(0, 9) = "P" And _
          Not rCelda.Offset(0, 11) = "DNB" Then _
          tRepetidoras = tRepetidoras + 1
    Next
End Function


Function tVacasProd(Optional Parto)
' Totaliza el número de Vacas en Producción
    If IsMissing(Parto) Then
            tVacasProd = WorksheetFunction. _
              Count _
              ( _
              Range("Tabla1[Prod.]") _
              )
        Else
            If Val(Parto) > 3 Then Parto = ">3"
            tVacasProd = WorksheetFunction. _
              CountIfs _
              ( _
              Range("Tabla1[Prod.]"), "<>" & "", _
              Range("Tabla1[Parto]"), Parto _
              )
    End If
End Function

Function tVacasFrescas(Optional Parto)
' Totaliza el número de Vacas en Producción
    If IsMissing(Parto) Then
        tVacasFrescas = WorksheetFunction. _
          CountIf _
          ( _
          Range("Tabla1[DEL]"), "<=" & 30 _
          )
      Else
        If Val(Parto) > 3 Then Parto = ">3"
        tVacasFrescas = WorksheetFunction. _
          CountIfs _
          ( _
          Range("Tabla1[DEL]"), "<=" & 30, _
          Range("Tabla1[Parto]"), Parto _
          )
    End If
End Function

Function tVacasServidas(Optional Servicio)
    If IsMissing(Servicio) Then
            tVacasServidas = WorksheetFunction. _
              CountIfs _
              ( _
              Range("Tabla1[Servicio]"), ">" & 0, _
              Range("Tabla1[Status]"), "<>" & "P", _
              Range("Tabla1[Clave1]"), "<>" & "DNB" _
              )
        Else
            If Val(Servicio) > 3 Then Servicio = ">3"
            tVacasServidas = WorksheetFunction. _
              CountIfs _
              ( _
              Range("Tabla1[Servicio]"), ">" & 0, _
              Range("Tabla1[Status]"), "<>" & "P", _
              Range("Tabla1[Servicio]"), Servicio, _
              Range("Tabla1[Clave1]"), "<>" & "DNB" _
              )
    End If
End Function

Function tVacasGest()
    tVacasGest = WorksheetFunction. _
      CountIfs _
      ( _
      Range("Tabla1[Servicio]"), ">" & "0", _
      Range("Tabla1[Status]"), "=" & "P" _
      )
End Function

Function tVacasSinServ()
    tVacasSinServ = WorksheetFunction. _
      CountIfs _
      ( _
      Range("Tabla1[DEL]"), ">" & 45, _
      Range("Tabla1[Servicio]"), "=" & "", _
      Range("Tabla1[Clave1]"), "<>" & "DNB" _
      )
End Function

Function tVacasEV()
    tVacasEV = WorksheetFunction. _
      CountIfs _
      ( _
      Range("Tabla1[DEL]"), "<=" & 45, _
      Range("Tabla1[Servicio]"), "=" & "", _
      Range("Tabla1[Clave1]"), "<>" & "DNB" _
      )
End Function

Function tVacasSecas()
' Totaliza el número de Vacas Secas
    tVacasSecas = WorksheetFunction. _
      CountBlank _
      ( _
      Range("Tabla1[Prod.]") _
      )
End Function

Function tVaquillas(Optional StatusRepro As String)
' Totaliza el número de Vaquillas según estatus reproductivo _
  "P" para gestantes, "O" para no gestantes, vbnullstring para todas
    Dim rCelda As Range
    tVaquillas = 0
    Select Case StatusRepro
        Case Is = vbNullString
            For Each rCelda In Range("Tabla2[Arete]")
                'If Not rCelda.Offset(0, 9) = "P" And
                If _
                  Date - CDate(rCelda.Offset(0, 4)) > _
                  Int(Range("Configuracion!C47") * 30.4) And _
                  rCelda.Offset(0, 13) = "H" Then _
                  tVaquillas = tVaquillas + 1
            Next
            Exit Function
        Case Is = "P"
            For Each rCelda In Range("Tabla2[Arete]")
                If rCelda.Offset(0, 9) = "P" And _
                  Date - CDate(rCelda.Offset(0, 4)) > _
                  Int(Range("Configuracion!C47") * 30.4) And _
                  rCelda.Offset(0, 13) = "H" Then _
                  tVaquillas = tVaquillas + 1
            Next
            Exit Function
        Case Is = "O"
            For Each rCelda In Range("Tabla2[Arete]")
                If Not rCelda.Offset(0, 9) = "P" And _
                  Date - CDate(rCelda.Offset(0, 4)) > _
                  Int(Range("Configuracion!C47") * 30.4) And _
                  rCelda.Offset(0, 13) = "H" Then _
                  tVaquillas = tVaquillas + 1
            Next
    End Select
End Function

Function IndiceFecha(nTabla)
    'Devuelve el renglón donde se encuentra la fecha
    Dim rCelda As Object
    Dim nfecha As Date
    ' La fecha del sistema
    nfecha = CDate(Date)
    IndiceFecha = 0
    nTabla = nTabla & "[Fecha]"
    For Each rCelda In Range(nTabla)
        If Month(CDate(rCelda)) = Month(Date) And Year(CDate(rCelda)) = Year(Date) Then
            IndiceFecha = rCelda.Offset.Row
            Exit Function
        End If
    Next rCelda
End Function

Private Sub Instamatic()
' insertar los parámetros del momento
    'Dim ws As Worksheet
    'Dim iR, iRenglon As Integer
    Application.ScreenUpdating = CBool(Range("Desarrollador!B6"))
    Application.DisplayStatusBar = True
    Set ws = Worksheets("AcumStat")
    'iRenglon = TamañoTabla("Tabla14") + 5
    If IndiceFecha("Tabla14") = 0 Then
            iR = TamañoTabla("Tabla14") + 2
        Else
            iR = IndiceFecha("Tabla14")
    End If
    'iR = 2
    iC = 1
    With ws
        ' Mes/año
        .Cells(iR, iC) = WorksheetFunction.EoMonth(Date, 0)
        cntdr
        'Fecha
        .Cells(iR, iC) = Date
        '.Cells(iRenglon, iC) = .Cells(iR, iC)
        cntdr
        ' Total de vacas
        .Cells(iR, iC) = tVacasProd() + tVacasSecas()
        '.Cells(iRenglon, iC) = .Cells(iR, iC)
        cntdr
        ' Vcs Producción
        .Cells(iR, iC) = tVacasProd()
        '.Cells(iRenglon, iC) = .Cells(iR, iC)
        cntdr
        ' Vcas Secas
        .Cells(iR, iC) = tVacasSecas()
        '.Cells(iRenglon, iC) = .Cells(iR, iC)
        cntdr
        ' Producción
        .Cells(iR, iC) = Int(WorksheetFunction.Sum(Range("Tabla1[Prod.]")))
        '.Cells(iRenglon, iC) = .Cells(iR, iC)
        cntdr
        ' Persistencia
        .Cells(iR, iC) = Format(WorksheetFunction.Average(Range("Tabla15[Persistencia]")), "0.0")
        '.Cells(iRenglon, iC) = .Cells(iR, iC)
        cntdr
        ' Producción máxima
        .Cells(iR, iC) = WorksheetFunction.Max(Range("Tabla1[Prod.]"))
        '.Cells(iRenglon, iC) = .Cells(iR, iC)
        cntdr
        ' Pico de lactancia
        .Cells(iR, iC) = PromParam("PicoProd")
        '.Cells(iRenglon, iC) = .Cells(iR, iC)
        cntdr
        .Cells(iR, iC) = pProy305d()
        '.Cells(iRenglon, iC) = .Cells(iR, iC)
        cntdr
        ' días en leche
        .Cells(iR, iC) = pDEL()
        '.Cells(iRenglon, iC) = .Cells(iR, iC)
        cntdr
        ' días seca
        .Cells(iR, iC) = WorksheetFunction.Average(Range("Tabla4[DíasSeca]"))
        '.Cells(iRenglon, iC) = .Cells(iR, iC)
        cntdr
        ' Intervalo entre partos
        .Cells(iR, iC) = WorksheetFunction.Average(Range("Tabla4[IntervaloParto]"))
        '.Cells(iRenglon, iC) = .Cells(iR, iC)
        cntdr
        ' Intervalo entre partos proyectado
        .Cells(iR, iC) = Int((pDAb() + 279) / 30.4)
        '.Cells(iRenglon, iC) = .Cells(iR, iC)
        cntdr
        ' días abiertos
        .Cells(iR, iC) = Format(pDAb(), "0")
        '.Cells(iRenglon, iC) = .Cells(iR, iC)
        cntdr
        ' SxV
        .Cells(iR, iC) = Format(pServicios(1), "0.0")
        '.Cells(iRenglon, iC) = .Cells(iR, iC)
        cntdr
        ' SxC
        .Cells(iR, iC) = Format(pServicios(1, "P"), "0.0")
        '.Cells(iRenglon, iC) = .Cells(iR, iC)
        cntdr
        ' d1S
        .Cells(iR, iC) = Format(pD1S(), "0")
        '.Cells(iRenglon, iC) = .Cells(iR, iC)
        cntdr
        ' d1Sg
        .Cells(iR, iC) = Format(pD1S(, "P"), "0")
        '.Cells(iRenglon, iC) = .Cells(iR, iC)
        cntdr
        ' d1C
        .Cells(iR, iC) = Format(pD1Calor, "0.0")
        '.Cells(iRenglon, iC) = .Cells(iR, iC)
        cntdr
        ' % gestantes 1° servicio
        .Cells(iR, iC) = Format(pctGest1Serv(), "0.0")
        '.Cells(iRenglon, iC) = .Cells(iR, iC)
        cntdr
        ' % Dx gest. positivo
        .Cells(iR, iC) = Format(DxGstPositivos(), "0.0")
        '.Cells(iRenglon, iC) = .Cells(iR, iC)
        cntdr
        ' % Dx gest. positivo + 30 d
        .Cells(iR, iC) = Format(DxGstPositivos(30), "0.0")
        '.Cells(iRenglon, iC) = .Cells(iR, iC)
        cntdr
        ' Calores perdidos
        .Cells(iR, iC) = Format(nCaloresPerdidos(), "0.0")
        '.Cells(iRenglon, iC) = .Cells(iR, iC)
        cntdr
        ' tasa de embarazo
        .Cells(iR, iC) = Format(TasaEmbarazo(), "0.0")
        '.Cells(iRenglon, iC) = .Cells(iR, iC)
        cntdr
        ' Breeding interval
        .Cells(iR, iC) = Format(BreedingInterval(), "0.0")
        '.Cells(iRenglon, iC) = .Cells(iR, iC)
        cntdr
        ' Heats detected
        .Cells(iR, iC) = Format(HeatsDetected(), "0.0")
        '.Cells(iRenglon, iC) = .Cells(iR, iC)
        cntdr
        ' Abortos
        .Cells(iR, iC) = Format(nAbortos(Month(Date), Year(Date)), "0")
        '.Cells(iRenglon, iC) = .Cells(iR, iC)
        cntdr
        ' SxVaq
        .Cells(iR, iC) = Format(pServicios(2), "0.0")
        '.Cells(iRenglon, iC) = .Cells(iR, iC)
        cntdr
        ' SxVaqG
        .Cells(iR, iC) = Format(pServicios(2, "P"), "0.0")
        '.Cells(iRenglon, iC) = .Cells(iR, iC)
        cntdr
        ' Edad al parto
        .Cells(iR, iC) = Format(pEdadAlParto(), "0")
        '.Cells(iRenglon, iC) = .Cells(iR, iC)
        cntdr
        ' Edad 1° Serv
        .Cells(iR, iC) = Format(pEdad1Serv(), "0")
        '.Cells(iRenglon, iC) = .Cells(iR, iC)
        cntdr
        ' Edad 1° Serv Gestante
        .Cells(iR, iC) = Format(pEdad1Serv("P"), "0")
        '.Cells(iRenglon, iC) = .Cells(iR, iC)
        cntdr
        ' AxP30d
        .Cells(iR, iC) = AnimPorParir(30, "H")
        '.Cells(iRenglon, iC) = .Cells(iR, iC)
        cntdr
        ' AxP60d
        .Cells(iR, iC) = AnimPorParir(60, "H")
        '.Cells(iRenglon, iC) = .Cells(iR, iC)
        cntdr
        ' AxP90d
        .Cells(iR, iC) = AnimPorParir(90, "H")
        '.Cells(iRenglon, iC) = .Cells(iR, iC)
        cntdr
        ' AxP120d
        .Cells(iR, iC) = AnimPorParir(120, "H")
        '.Cells(iRenglon, iC) = .Cells(iR, iC)
        cntdr
        ' AxP150d
        .Cells(iR, iC) = AnimPorParir(150, "H")
        '.Cells(iRenglon, iC) = .Cells(iR, iC)
        cntdr
        ' AxP180d
        .Cells(iR, iC) = AnimPorParir(180, "H")
        '.Cells(iRenglon, iC) = .Cells(iR, iC)
        cntdr
        ' AxP210d
        .Cells(iR, iC) = AnimPorParir(210, "H")
        '.Cells(iRenglon, iC) = .Cells(iR, iC)
        cntdr
        ' AxS30d
        .Cells(iR, iC) = AnimPorSecar(1)
        '.Cells(iRenglon, iC) = .Cells(iR, iC)
        cntdr
        ' AxS30d
        .Cells(iR, iC) = AnimPorSecar(2)
        '.Cells(iRenglon, iC) = .Cells(iR, iC)
        cntdr
        ' AxS30d
        .Cells(iR, iC) = AnimPorSecar(3)
        '.Cells(iRenglon, iC) = .Cells(iR, iC)
        cntdr
        ' AxS30d
        .Cells(iR, iC) = AnimPorSecar(4)
        '.Cells(iRenglon, iC) = .Cells(iR, iC)
        cntdr
        ' AxS30d
        .Cells(iR, iC) = AnimPorSecar(5)
        '.Cells(iRenglon, iC) = .Cells(iR, iC)
        cntdr
        ' AxS30d
        .Cells(iR, iC) = AnimPorSecar(6)
        '.Cells(iRenglon, iC) = .Cells(iR, iC)
        cntdr
        ' V gest
        .Cells(iR, iC) = tVacasGest()
        '.Cells(iRenglon, iC) = .Cells(iR, iC)
        cntdr
        ' V Serv
        .Cells(iR, iC) = tVacasServidas()
        '.Cells(iRenglon, iC) = .Cells(iR, iC)
        cntdr
        ' V EV
        .Cells(iR, iC) = tVacasEV()
        '.Cells(iRenglon, iC) = .Cells(iR, iC)
        cntdr
        ' V sS
        .Cells(iR, iC) = tVacasSinServ()
        '.Cells(iRenglon, iC) = .Cells(iR, iC)
        cntdr
        ' vProbl
        .Cells(iR, iC) = tProblema(1)
        '.Cells(iRenglon, iC) = .Cells(iR, iC)
        cntdr
        ' VRep
        .Cells(iR, iC) = tRepetidoras(1)
        '.Cells(iRenglon, iC) = .Cells(iR, iC)
        cntdr
        ' V 1S
        .Cells(iR, iC) = tVacasServidas(1)
        '.Cells(iRenglon, iC) = .Cells(iR, iC)
        cntdr
        ' V 2S
        .Cells(iR, iC) = tVacasServidas(2)
        '.Cells(iRenglon, iC) = .Cells(iR, iC)
        cntdr
        ' V 3S
        .Cells(iR, iC) = tVacasServidas(3)
        '.Cells(iRenglon, iC) = .Cells(iR, iC)
        cntdr
        ' V 4+S
        .Cells(iR, iC) = tVacasServidas() - _
          tVacasServidas(1) - tVacasServidas(2) - _
          tVacasServidas(3)
        '.Cells(iRenglon, iC) = .Cells(iR, iC)
        cntdr
        ' Tot Reempl
        .Cells(iR, iC) = tAnimales(2)
        '.Cells(iRenglon, iC) = .Cells(iR, iC)
        cntdr
        ' Tot Lactantes
        .Cells(iR, iC) = tLactantes()
        '.Cells(iRenglon, iC) = .Cells(iR, iC)
        cntdr
        ' Desarrollo
        .Cells(iR, iC) = tDesarrollo()
        '.Cells(iRenglon, iC) = .Cells(iR, iC)
        cntdr
        ' Novillas
        .Cells(iR, iC) = tNovillas()
        '.Cells(iRenglon, iC) = .Cells(iR, iC)
        cntdr
        ' Vaquillas
        .Cells(iR, iC) = tVaquillas()
        '.Cells(iRenglon, iC) = .Cells(iR, iC)
        cntdr
        ' Vaquillas gestantes
        .Cells(iR, iC) = tVaquillas("P")
        '.Cells(iRenglon, iC) = .Cells(iR, iC)
        cntdr
        ' vaquillas servidas
        .Cells(iR, iC) = tVaquillas("V")
        '.Cells(iRenglon, iC) = .Cells(iR, iC)
        cntdr
        ' Vaquillas problema
        .Cells(iR, iC) = tProblema(2)
        '.Cells(iRenglon, iC) = .Cells(iR, iC)
        cntdr
        ' Vaquillas repetidoras
        .Cells(iR, iC) = tRepetidoras(2)
        '.Cells(iRenglon, iC) = .Cells(iR, iC)
        cntdr
        ' vaqxP30d
        .Cells(iR, iC) = AnimPorParir(30, "R")
        '.Cells(iRenglon, iC) = .Cells(iR, iC)
        cntdr
        ' vaqxP60d
        .Cells(iR, iC) = AnimPorParir(60, "R")
        '.Cells(iRenglon, iC) = .Cells(iR, iC)
        cntdr
        ' vaqxP90d
        .Cells(iR, iC) = AnimPorParir(90, "R")
        '.Cells(iRenglon, iC) = .Cells(iR, iC)
        cntdr
        ' vaqxP120d
        .Cells(iR, iC) = AnimPorParir(120, "R")
        '.Cells(iRenglon, iC) = .Cells(iR, iC)
        cntdr
        ' vaqxP150d
        .Cells(iR, iC) = AnimPorParir(150, "R")
        '.Cells(iRenglon, iC) = .Cells(iR, iC)
        cntdr
        ' vaqxP180d
        .Cells(iR, iC) = AnimPorParir(180, "R")
        '.Cells(iRenglon, iC) = .Cells(iR, iC)
        cntdr
        ' vaqxP210d
        .Cells(iR, iC) = AnimPorParir(210, "R")
        '.Cells(iRenglon, iC) = .Cells(iR, iC)
        cntdr
        ' PromProd30d
        .Cells(iR, iC) = WorksheetFunction.Average(Range("Tabla15[30d]"))
        '.Cells(iRenglon, iC) = .Cells(iR, iC)
        cntdr
        ' PromProd60d
        .Cells(iR, iC) = WorksheetFunction.Average(Range("Tabla15[60d]"))
        '.Cells(iRenglon, iC) = .Cells(iR, iC)
        cntdr
        ' PromProd90d
        .Cells(iR, iC) = WorksheetFunction.Average(Range("Tabla15[90d]"))
        '.Cells(iRenglon, iC) = .Cells(iR, iC)
        cntdr
        ' PromProd120d
        .Cells(iR, iC) = WorksheetFunction.Average(Range("Tabla15[120d]"))
        '.Cells(iRenglon, iC) = .Cells(iR, iC)
        cntdr
        ' PromProd150d
        .Cells(iR, iC) = WorksheetFunction.Average(Range("Tabla15[150d]"))
        '.Cells(iRenglon, iC) = .Cells(iR, iC)
        cntdr
        ' PromProd180d
        .Cells(iR, iC) = WorksheetFunction.Average(Range("Tabla15[180d]"))
        '.Cells(iRenglon, iC) = .Cells(iR, iC)
        cntdr
        ' PromProd210d
        .Cells(iR, iC) = WorksheetFunction.Average(Range("Tabla15[210d]"))
        '.Cells(iRenglon, iC) = .Cells(iR, iC)
        cntdr
        ' PromProd240d
        .Cells(iR, iC) = WorksheetFunction.Average(Range("Tabla15[240d]"))
        '.Cells(iRenglon, iC) = .Cells(iR, iC)
        cntdr
        ' PromProd270d
        .Cells(iR, iC) = WorksheetFunction.Average(Range("Tabla15[270d]"))
        '.Cells(iRenglon, iC) = .Cells(iR, iC)
        cntdr
        ' PromProd300d
        .Cells(iR, iC) = WorksheetFunction.Average(Range("Tabla15[300d]"))
        '.Cells(iRenglon, iC) = .Cells(iR, iC)
        cntdr
        ' PicoProd 1°
        .Cells(iR, iC) = PromParam("PicoProd", 1)
        '.Cells(iRenglon, iC) = .Cells(iR, iC)
        cntdr
        ' PicoProd 2°
        .Cells(iR, iC) = PromParam("PicoProd", 2)
        '.Cells(iRenglon, iC) = .Cells(iR, iC)
        cntdr
        ' PicoProd 3°
        .Cells(iR, iC) = PromParam("PicoProd", 3)
        '.Cells(iRenglon, iC) = .Cells(iR, iC)
        cntdr
        ' PicoProd 4+
        .Cells(iR, iC) = PromParam("PicoProd", 4)
        '.Cells(iRenglon, iC) = .Cells(iR, iC)
        cntdr
        ' DiasPico
        .Cells(iR, iC) = PromParam("DiasPico")
        '.Cells(iRenglon, iC) = .Cells(iR, iC)
        cntdr
        ' DiasPico 1°
        .Cells(iR, iC) = PromParam("DiasPico", 1)
        '.Cells(iRenglon, iC) = .Cells(iR, iC)
        cntdr
        ' DiasPico 2°
        .Cells(iR, iC) = PromParam("DiasPico", 2)
        '.Cells(iRenglon, iC) = .Cells(iR, iC)
        cntdr
        ' DiasPico 3°
        .Cells(iR, iC) = PromParam("DiasPico", 3)
        '.Cells(iRenglon, iC) = .Cells(iR, iC)
        cntdr
        ' DiasPico 4+
        .Cells(iR, iC) = PromParam("DiasPico", 4)
        '.Cells(iRenglon, iC) = .Cells(iR, iC)
        cntdr
        ' Proy305d 1°
        .Cells(iR, iC) = PromParam("Proy305d", 1)
        '.Cells(iRenglon, iC) = .Cells(iR, iC)
        cntdr
        ' Proy305d 2°
        .Cells(iR, iC) = PromParam("Proy305d", 2)
        '.Cells(iRenglon, iC) = .Cells(iR, iC)
        cntdr
        ' Proy305d 3°
        .Cells(iR, iC) = PromParam("Proy305d", 3)
        '.Cells(iRenglon, iC) = .Cells(iR, iC)
        cntdr
        ' Proy305d 4°
        .Cells(iR, iC) = PromParam("Proy305d", 4)
        '.Cells(iRenglon, iC) = .Cells(iR, iC)
        cntdr
        ' tVacasFrescas
        .Cells(iR, iC) = tVacasFrescas()
        '.Cells(iRenglon, iC) = .Cells(iR, iC)
        cntdr
        ' tVacasFrescas
        .Cells(iR, iC) = tVacasFrescas(1)
        '.Cells(iRenglon, iC) = .Cells(iR, iC)
        cntdr
        ' tVacasFrescas
        .Cells(iR, iC) = tVacasFrescas(2)
        '.Cells(iRenglon, iC) = .Cells(iR, iC)
        cntdr
        ' tVacasFrescas
        .Cells(iR, iC) = tVacasFrescas(3)
        '.Cells(iRenglon, iC) = .Cells(iR, iC)
        cntdr
        ' tVacasFrescas
        .Cells(iR, iC) = tVacasFrescas(4)
        '.Cells(iRenglon, iC) = .Cells(iR, iC)
        cntdr
        ' tVacas 1° Parto
        .Cells(iR, iC) = WorksheetFunction. _
          CountIf(Range("tabla1[Parto]"), 1)
        '.Cells(iRenglon, iC) = .Cells(iR, iC)
        cntdr
        ' tVacas 2° Parto
        .Cells(iR, iC) = WorksheetFunction. _
          CountIf(Range("tabla1[Parto]"), 2)
        '.Cells(iRenglon, iC) = .Cells(iR, iC)
        cntdr
        ' tVacas 3° Parto
        .Cells(iR, iC) = WorksheetFunction. _
          CountIf(Range("tabla1[Parto]"), 3)
        '.Cells(iRenglon, iC) = .Cells(iR, iC)
        cntdr
        ' tVacas 4+ Parto
        .Cells(iR, iC) = WorksheetFunction. _
          CountIf(Range("tabla1[Parto]"), ">3")
        '.Cells(iRenglon, iC) = .Cells(iR, iC)
        cntdr
        ' PromProducción
        .Cells(iR, iC) = pProd()
        '.Cells(iRenglon, iC) = .Cells(iR, iC)
        cntdr
        ' PromProducción 1
        .Cells(iR, iC) = pProd(1)
        '.Cells(iRenglon, iC) = .Cells(iR, iC)
        cntdr
        ' PromProducción 2
        .Cells(iR, iC) = pProd(2)
        '.Cells(iRenglon, iC) = .Cells(iR, iC)
        cntdr
        ' PromProducción 3
        .Cells(iR, iC) = pProd(3)
        '.Cells(iRenglon, iC) = .Cells(iR, iC)
        cntdr
        ' PromProducción 4
        .Cells(iR, iC) = pProd(4)
        '.Cells(iRenglon, iC) = .Cells(iR, iC)
        cntdr
        ' Peristencia 1°
        .Cells(iR, iC) = PromParam("Persistencia", 1)
        '.Cells(iRenglon, iC) = .Cells(iR, iC)
        cntdr
        ' Peristencia 2°
        .Cells(iR, iC) = PromParam("Persistencia", 2)
        '.Cells(iRenglon, iC) = .Cells(iR, iC)
        cntdr
        ' Peristencia 3°
        .Cells(iR, iC) = PromParam("Persistencia", 3)
        '.Cells(iRenglon, iC) = .Cells(iR, iC)
        cntdr
        ' Peristencia 4°
        .Cells(iR, iC) = PromParam("Persistencia", 4)
        '.Cells(iRenglon, iC) = .Cells(iR, iC)
        cntdr
        ' PromProd 1°
        .Cells(iR, iC) = pProd(1)
        '.Cells(iRenglon, iC) = .Cells(iR, iC)
        cntdr
        ' PromProd 2°
        .Cells(iR, iC) = pProd(2)
        '.Cells(iRenglon, iC) = .Cells(iR, iC)
        cntdr
        ' PromProd 3°
        .Cells(iR, iC) = pProd(3)
        '.Cells(iRenglon, iC) = .Cells(iR, iC)
        cntdr
        ' PromProd 4+
        .Cells(iR, iC) = pProd(4)
        '.Cells(iRenglon, iC) = .Cells(iR, iC)
        cntdr
        ' promDEL 1°
        .Cells(iR, iC) = pDEL(1)
        '.Cells(iRenglon, iC) = .Cells(iR, iC)
        cntdr
        ' promDEL 2°
        .Cells(iR, iC) = pDEL(2)
        '.Cells(iRenglon, iC) = .Cells(iR, iC)
        cntdr
        ' promDEL 3°
        .Cells(iR, iC) = pDEL(3)
        '.Cells(iRenglon, iC) = .Cells(iR, iC)
        cntdr
        ' promDEL 4+
        .Cells(iR, iC) = pDEL(4)
        '.Cells(iRenglon, iC) = .Cells(iR, iC)
        cntdr
        ' nPartos
        .Cells(iR, iC) = nPartos(Month(Date), Year(Date))
        cntdr
        .Cells(iR, iC) = iC
        cntdr
        .Cells(iR, iC) = nBajas(Month(Date), Year(Date))
        cntdr
        .Cells(iR, iC) = iC
        cntdr
        .Cells(iR, iC) = iC
        cntdr
        .Cells(iR, iC) = iC
        cntdr
        .Cells(iR, iC) = iC
        cntdr
        .Cells(iR, iC) = iC
        cntdr
        .Cells(iR, iC) = iC
    End With
    Application.StatusBar = "Agregando indicadores"
    Application.StatusBar = "Salvando Información"
    ActiveWorkbook.Save
    'OrdenarAcumStat
    Application.StatusBar = False
    Application.ScreenUpdating = True

End Sub

Private Sub cntdr()
    ' Incrementa contador
    iC = iC + 1
    ' muestra avance en barra de estado
    Application.StatusBar = "Calculando indicadores de desempeño... " & _
      Format(iC / 130, "0%")
End Sub

