Attribute VB_Name = "ModParametros"
' Ultima Modificacion: 28-Nov-17
' Adici�n IntervaloCalores
' Correcci�n en tVacasServidas
Option Explicit
Dim rCelda As Range
Dim dDxPct As Double
Dim iC, iR, iRenglon As Long
Dim ws As Worksheet

Function DUE(Arete_Buscado As Variant, _
  Evento_Buscado As String, Col_Buscada As Long)
' (D)ato (U)ltimo (E)vento
' Devuelve el dato del �ltimo evento buscado
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
           ' Si es la �ltima ocurrencia del evento
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
' Devuelve Pico de Producci�n (1) o _
  d�as al pico de producci�n _
  =PicoProd(1025,2) = d�as al pico _
  =PicoProd(1025,1) = pico producci�n
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
              ' Extrae D�as al pico de Metadatos
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
    Range(Range("Estad�sticas!C5"), _
      Range("Estad�sticas!L10")).ClearContents
    Range(Range("Estad�sticas!C16"), _
      Range("Estad�sticas!L21")).ClearContents
    Range(Range("Estad�sticas!C27"), _
      Range("Estad�sticas!K29")).ClearContents
    Range(Range("Estad�sticas!C39"), _
      Range("Estad�sticas!H48")).ClearContents
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
    'Range("Estad�sticas!B1") = Range("Configuracion!C3")
    Range("Estad�sticas!B4") = "Parto"
    Range("Estad�sticas!B5") = 1
    Range("Estad�sticas!B6") = 2
    Range("Estad�sticas!B7") = 3
    Range("Estad�sticas!B8") = 4
    Range("Estad�sticas!B9") = "5+"
    Range("Estad�sticas!B14") = "REEMPLAZOS"
    Range("Estad�sticas!B15") = "Etapa"
    Range("Estad�sticas!B16") = "Lactancia"
    Range("Estad�sticas!B17") = "Desarrollo"
    Range("Estad�sticas!B18") = "Novillas"
    Range("Estad�sticas!B19") = "Vaquillas"
    Range("Estad�sticas!B20") = "Vaq. Gestantes"
    Range("Estad�sticas!B25") = "ANIMALES POR PARIR"
    Range("Estad�sticas!B27") = "Vacas"
    Range("Estad�sticas!B28") = "Reemplazos"
    Range("Estad�sticas!B37") = "ANIMALES POR CORRAL"
    Range("Estad�sticas!B38") = "Corral"
    Range("Estad�sticas!B39") = 1
    Range("Estad�sticas!B40") = 2
    Range("Estad�sticas!B41") = 3
    Range("Estad�sticas!B42") = 4
    Range("Estad�sticas!B43") = 5
    Range("Estad�sticas!B44") = 6
    Range("Estad�sticas!B45") = 7
    Range("Estad�sticas!B46") = 8
    Range("Estad�sticas!B47") = 9
    
    'Columna C
    Range("Estad�sticas!C4") = "N�mero de Animales"
    Range("Estad�sticas!C5") = _
      WorksheetFunction.CountIfs(Range("Tabla1[Parto]"), 1)
    Range("Estad�sticas!C6") = _
      WorksheetFunction.CountIfs(Range("Tabla1[Parto]"), 2)
    Range("Estad�sticas!C7") = _
      WorksheetFunction.CountIfs(Range("Tabla1[Parto]"), 3)
    Range("Estad�sticas!C8") = _
      WorksheetFunction.CountIfs(Range("Tabla1[Parto]"), 4)
    Range("Estad�sticas!C9") = _
      WorksheetFunction.CountIfs(Range("Tabla1[Parto]"), ">=5")
    Range("Estad�sticas!C10") = _
      WorksheetFunction.Sum(Range("Estad�sticas!C5:C9"))
    Range("Estad�sticas!C15") = "N�mero de Animales"
    Range("Estad�sticas!C16") = tLactantes()
    Range("Estad�sticas!C17") = tDesarrollo()
    Range("Estad�sticas!C18") = tNovillas()
    Range("Estad�sticas!C19") = tVaquillas("O")
    Range("Estad�sticas!C20") = tVaquillas("P")
    Range("Estad�sticas!C21") = _
      WorksheetFunction.Sum(Range("Estad�sticas!C16:C20"))
    Range("Estad�sticas!C26") = "30 D�as"
    Range("Estad�sticas!C27") = AnimPorParir(30, "H")
    Range("Estad�sticas!C28") = AnimPorParir(30, "R")
    Range("Estad�sticas!C29") = _
      Range("Estad�sticas!C27") + Range("Estad�sticas!C28")
    Range("Estad�sticas!C38") = "N�mero de Animales"
    Range("Estad�sticas!C39") = _
      WorksheetFunction.CountIfs(Range("Tabla1[Corral]"), 1) + _
      WorksheetFunction.CountIfs(Range("Tabla2[Corral]"), 1)
    Range("Estad�sticas!C40") = WorksheetFunction.CountIfs _
      (Range("Tabla1[Corral]"), 2) + _
      WorksheetFunction.CountIfs(Range("Tabla2[Corral]"), 2)
    Range("Estad�sticas!C41") = _
      WorksheetFunction.CountIfs(Range("Tabla1[Corral]"), 3) + _
      WorksheetFunction.CountIfs(Range("Tabla2[Corral]"), 3)
    Range("Estad�sticas!C42") = WorksheetFunction.CountIfs( _
      Range("Tabla1[Corral]"), 4) + _
      WorksheetFunction.CountIfs(Range("Tabla2[Corral]"), 4)
    Range("Estad�sticas!C43") = WorksheetFunction.CountIfs( _
      Range("Tabla1[Corral]"), 5) + _
      WorksheetFunction.CountIfs(Range("Tabla2[Corral]"), 5)
    Range("Estad�sticas!C44") = WorksheetFunction.CountIfs( _
      Range("Tabla1[Corral]"), 6) + _
      WorksheetFunction.CountIfs(Range("Tabla2[Corral]"), 6)
    Range("Estad�sticas!C45") = WorksheetFunction.CountIfs( _
      Range("Tabla1[Corral]"), 7) + _
      WorksheetFunction.CountIfs(Range("Tabla2[Corral]"), 7)
    Range("Estad�sticas!C46") = WorksheetFunction.CountIfs( _
      Range("Tabla1[Corral]"), 8) + _
      WorksheetFunction.CountIfs(Range("Tabla2[Corral]"), 8)
    Range("Estad�sticas!C47") = WorksheetFunction.CountIfs( _
      Range("Tabla1[Corral]"), 9) + _
      WorksheetFunction.CountIfs(Range("Tabla2[Corral]"), 9)
    Range("Estad�sticas!C48") = WorksheetFunction. _
      Sum(Range("Estad�sticas!C39:C47"))
    
    'Columna D
    Range("Estad�sticas!D4") = "% del Hato"
    Range("Estad�sticas!D5") = Format(Range("Estad�sticas!C5") / _
      Range("Estad�sticas!C10"), "#%")
    Range("Estad�sticas!D6") = Format(Range("Estad�sticas!C6") / _
      Range("Estad�sticas!C10"), "#%")
    Range("Estad�sticas!D7") = Format(Range("Estad�sticas!C7") / _
      Range("Estad�sticas!C10"), "#%")
    Range("Estad�sticas!D8") = Format(Range("Estad�sticas!C8") / _
      Range("Estad�sticas!C10"), "#%")
    Range("Estad�sticas!D9") = Format(Range("Estad�sticas!C9") / _
      Range("Estad�sticas!C10"), "#%")
    Range("Estad�sticas!D10") = Format(WorksheetFunction. _
      Sum(Range("Estad�sticas!D5:D9")), "#%")
    Range("Estad�sticas!D15") = "% de Reemplazos"
    Range("Estad�sticas!D16") = Format(Range("Estad�sticas!C16") / _
      Range("Estad�sticas!C21"), "#.0%")
    Range("Estad�sticas!D17") = Format(Range("Estad�sticas!C17") / _
      Range("Estad�sticas!C21"), "#.0%")
    Range("Estad�sticas!D18") = Format(Range("Estad�sticas!C18") / _
      Range("Estad�sticas!C21"), "#.0%")
    Range("Estad�sticas!D19") = Format(Range("Estad�sticas!C19") / _
      Range("Estad�sticas!C21"), "#.0%")
    Range("Estad�sticas!D20") = Format(Range("Estad�sticas!C20") / _
      Range("Estad�sticas!C21"), "#.0%")
    Range("Estad�sticas!D21") = Format(WorksheetFunction. _
      Sum(Range("Estad�sticas!D16:D20")), "#%")
    Range("Estad�sticas!D26") = "60 D�as"
    Range("Estad�sticas!D27") = AnimPorParir(60, "H")
    Range("Estad�sticas!D28") = AnimPorParir(60, "R")
    Range("Estad�sticas!D29") = Format(Range("Estad�sticas!D27") + _
      Range("Estad�sticas!D28"), "#")
    Range("Estad�sticas!D38") = "%"
    Range("Estad�sticas!D39") = Format(Range("Estad�sticas!C39") / _
      Range("Estad�sticas!C48"), "#%")
    Range("Estad�sticas!D40") = Format(Range("Estad�sticas!C40") / _
      Range("Estad�sticas!C48"), "#%")
    Range("Estad�sticas!D41") = Format(Range("Estad�sticas!C41") / _
      Range("Estad�sticas!C48"), "#%")
    Range("Estad�sticas!D42") = Format(Range("Estad�sticas!C42") / _
      Range("Estad�sticas!C48"), "#%")
    Range("Estad�sticas!D43") = Format(Range("Estad�sticas!C43") / _
      Range("Estad�sticas!C48"), "#%")
    Range("Estad�sticas!D44") = Format(Range("Estad�sticas!C44") / _
      Range("Estad�sticas!C48"), "#%")
    Range("Estad�sticas!D45") = Format(Range("Estad�sticas!C45") / _
      Range("Estad�sticas!C48"), "#%")
    Range("Estad�sticas!D46") = Format(Range("Estad�sticas!C46") / _
      Range("Estad�sticas!C48"), "#%")
    Range("Estad�sticas!D47") = Format(Range("Estad�sticas!C47") / _
      Range("Estad�sticas!C48"), "#%")
    Range("Estad�sticas!D48") = Format(WorksheetFunction. _
      Sum(Range("Estad�sticas!D39:D47")), "#%")

    'Columna E
    Range("Estad�sticas!E4") = "D�as en Leche"
    'E5
    Range("Estad�sticas!E5") = pDEL(1)
    'E6
    Range("Estad�sticas!E6") = pDEL(2)
    'E7
    Range("Estad�sticas!E7") = pDEL(3)
    'E8
    Range("Estad�sticas!E8") = pDEL(4)
    'E9
    Range("Estad�sticas!E9") = pDEL(5)
    'E10
    Range("Estad�sticas!E10") = Format(pDEL(), "#")
    Range("Estad�sticas!E15") = "Edad"
    'E16
    If Not IsError(WorksheetFunction. _
      AverageIfs(Range("Tabla2[Edad2]"), Range("Tabla2[Edad2]"), _
      "<=45")) Then Range("Estad�sticas!E16") = _
      Format(WorksheetFunction.AverageIfs(Range("Tabla2[Edad2]"), _
      Range("Tabla2[Edad2]"), "<=45"), "#") & "d"
    'E17
    If Not IsError(WorksheetFunction. _
      AverageIfs(Range("Tabla2[Edad2]"), Range("Tabla2[Edad2]"), _
      ">45", Range("Tabla2[Edad2]"), "<365", _
      Range("Tabla2[Sexo]"), "H")) Then _
      Range("Estad�sticas!E17") = Format(Int(WorksheetFunction. _
      AverageIfs(Range("Tabla2[Edad2]"), Range("Tabla2[Edad2]"), _
      ">45", Range("Tabla2[Edad2]"), "<365", _
      Range("Tabla2[Sexo]"), "H") / 30.4), "#") & "m"
    'E18
    If Not IsError(WorksheetFunction. _
      AverageIfs(Range("Tabla2[Edad2]"), Range("Tabla2[Edad2]"), _
      ">365", Range("Tabla2[Edad2]"), "<" & _
      Int(Range("Configuracion!C47") * 30.4), _
      Range("Tabla2[Sexo]"), "H")) Then _
      Range("Estad�sticas!E18") = Format(Int(WorksheetFunction. _
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
      Range("Estad�sticas!E19") = Format(Int(WorksheetFunction. _
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
      Range("Estad�sticas!E20") = Format(Int(WorksheetFunction. _
      AverageIfs(Range("Tabla2[Edad2]"), _
      Range("Tabla2[Edad2]"), ">" & _
      Int(Range("Configuracion!C47") * 30.4), _
      Range("Tabla2[Status]"), "P", _
      Range("Tabla2[Sexo]"), "H") / 30.4), "#") & "m"
    Range("Estad�sticas!E26") = "90 D�as"
    Range("Estad�sticas!E27") = AnimPorParir(90, "H")
    Range("Estad�sticas!E28") = AnimPorParir(90, "R")
    Range("Estad�sticas!E29") = Range("Estad�sticas!E27") + _
      Range("Estad�sticas!E28")
    Range("Estad�sticas!E38") = "D�as en Leche"
    'E39
    Range("Estad�sticas!E39") = pDEL(1)
    'E40
    Range("Estad�sticas!E40") = pDEL(2)
    'E41
    Range("Estad�sticas!E41") = pDEL(3)
    'E42
    Range("Estad�sticas!E42") = pDEL(4)
    'E43
    Range("Estad�sticas!E43") = pDEL(5)

    'Columna F
    Range("Estad�sticas!F4") = "Producci�n Promedio"
    'F5
    Range("Estad�sticas!F5") = pProd(1)
    'F6
    Range("Estad�sticas!F6") = pProd(2)
    'F7
    Range("Estad�sticas!F7") = pProd(3)
    'F8
    Range("Estad�sticas!F8") = pProd(4)
    'F9
    Range("Estad�sticas!F9") = pProd(5)
    'F10
    Range("Estad�sticas!F10") = Format(pProd(), "#.#")
    Range("Estad�sticas!F15") = "Peso Promedio"
    'On Error Resume Next
    'F16
    If Not IsError(WorksheetFunction. _
      AverageIfs(Range("Tabla2[PesoCorporal]"), _
      Range("Tabla2[Edad2]"), "<=45")) Then _
      Range("Estad�sticas!F16") = _
      Format(WorksheetFunction. _
      AverageIfs(Range("Tabla2[PesoCorporal]"), _
      Range("Tabla2[Edad2]"), "<=45"), "#")
    'F17
    If Not IsError(WorksheetFunction. _
      AverageIfs(Range("Tabla2[PesoCorporal]"), _
      Range("Tabla2[Edad2]"), ">45", _
      Range("Tabla2[Edad2]"), "<365", _
      Range("Tabla2[Sexo]"), "H")) Then _
      Range("Estad�sticas!F17") = _
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
      Range("Estad�sticas!F18") = _
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
      Range("Estad�sticas!F19") = _
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
      Range("Estad�sticas!F20") = _
      Format(Int(WorksheetFunction. _
      AverageIfs(Range("Tabla2[PesoCorporal]"), _
      Range("Tabla2[Edad2]"), ">" & _
      Int(Range("Configuracion!C47") * 30.4), _
      Range("Tabla2[Status]"), "P", _
      Range("Tabla2[Sexo]"), "H") / 30.4), "#")
    On Error GoTo 0
    Range("Estad�sticas!F26") = "120 D�as"
    Range("Estad�sticas!F27") = AnimPorParir(120, "H")
    Range("Estad�sticas!F28") = AnimPorParir(120, "R")
    Range("Estad�sticas!F29") = _
      Range("Estad�sticas!F27") + Range("Estad�sticas!F28")
    Range("Estad�sticas!F38") = "Producci�n Promedio"
    'F39
    Range("Estad�sticas!F39") = pProd(1)
    'F40
    Range("Estad�sticas!F40") = pProd(2)
    'F41
    Range("Estad�sticas!F41") = pProd(3)
    'F42
    Range("Estad�sticas!F42") = pProd(4)
    'F43
    Range("Estad�sticas!F43") = pProd(5)
   
    'Columna G
    Range("Estad�sticas!G4") = "Prom. Pico de Producci�n"
    Range("Estad�sticas!G26") = "150 D�as"
    Range("Estad�sticas!G27") = AnimPorParir(150, "H")
    Range("Estad�sticas!G28") = AnimPorParir(150, "R")
    Range("Estad�sticas!G29") = _
      Range("Estad�sticas!G27") + Range("Estad�sticas!G28")
    Range("Estad�sticas!G38") = "Lactancia m�nima"
      
    'Columna H
    Range("Estad�sticas!H4") = "Proyecci�n a 305d"
    Range("Estad�sticas!H5") = Format(pProy305d(1), "#,#")
    Range("Estad�sticas!H6") = Format(pProy305d(2), "#,#")
    Range("Estad�sticas!H7") = Format(pProy305d(3), "#,#")
    Range("Estad�sticas!H8") = Format(pProy305d(4), "#,#")
    Range("Estad�sticas!H9") = Format(pProy305d(5), "#,#")
    'H10
    Range("Estad�sticas!H10") = Format(pProy305d(), "#,#")
    Range("Estad�sticas!H26") = "180 D�as"
    Range("Estad�sticas!H27") = AnimPorParir(180, "H")
    Range("Estad�sticas!H28") = AnimPorParir(180, "R")
    Range("Estad�sticas!H29") = _
      Range("Estad�sticas!H27") + Range("Estad�sticas!H28")
    Range("Estad�sticas!H38") = "Lactancia M�xima"

    'Columna I
    Range("Estad�sticas!I4") = "N�mero de Servicios por Vaca"
    'I5
    Range("Estad�sticas!I5") = pServicios(1, , 1)
    'I6
    Range("Estad�sticas!I6") = pServicios(1, , 2)
    'I7
    Range("Estad�sticas!I7") = pServicios(1, , 3)
    'I8
    Range("Estad�sticas!I8") = pServicios(1, , 4)
    'I9
    Range("Estad�sticas!I9") = pServicios(1, , 5)
    'I10
    Range("Estad�sticas!I10") = Format(pServicios(1), "#.0")
    Range("Estad�sticas!I15") = _
      "N�mero de Servicios por Animal"
    'I19
    Range("Estad�sticas!I19") = pServicios(2)
    'I27
    Range("Estad�sticas!I27") = AnimPorParir(210, "H")
    'I28
    Range("Estad�sticas!I28") = AnimPorParir(210, "R")
    'I29
    Range("Estad�sticas!I29") = _
      (Range("Estad�sticas!I27") + Range("Estad�sticas!I28"))
      
    'Columna J
    Range("Estad�sticas!J4") = _
      "N�mero de Servicios por Concepci�n"
    'J5
      Range("Estad�sticas!J5") = pServicios(1, "P", 1)
    'J6
    Range("Estad�sticas!J6") = pServicios(1, "P", 2)
    'J7
    Range("Estad�sticas!J7") = pServicios(1, "P", 3)
    'J8
    Range("Estad�sticas!J8") = pServicios(1, "P", 4)
    'J9
    Range("Estad�sticas!J9") = pServicios(1, "P", 5)
    'J10
    Range("Estad�sticas!J10") = Format(pServicios(1, "P"), "#.0")
    Range("Estad�sticas!J15") = _
      "N�mero de Servicios por Concepci�n"
    'J20
    Range("Estad�sticas!J20") = pServicios(2, "P")
    Range("Estad�sticas!J26") = "240 D�as*"
    'J27
    Range("Estad�sticas!J27") = AnimPorParir(240, "H")
    'J28
    Range("Estad�sticas!J28") = AnimPorParir(240, "R")
    'J29
    Range("Estad�sticas!J29") = _
      (Range("Estad�sticas!J27") + Range("Estad�sticas!J28"))
    
    'Columna K
    'Range("Estad�sticas!K1") = "Situaci�n del Establo al d�a:"
    Range("Estad�sticas!K4") = "Promedio d�as Abiertos"
    Range("Estad�sticas!K5") = Format(pDAb(1), "#")
    Range("Estad�sticas!K6") = Format(pDAb(2), "#")
    Range("Estad�sticas!K7") = Format(pDAb(3), "#")
    Range("Estad�sticas!K8") = Format(pDAb(4), "#")
    Range("Estad�sticas!K9") = Format(pDAb(5), "#")
    Range("Estad�sticas!K10") = Format(pDAb(), "#")
    Range("Estad�sticas!K15") = "Promedio Edad al Parto"
    Range("Estad�sticas!K20") = Format(pEdadAlParto, "#") & "m"
    Range("Estad�sticas!K26") = "270 D�as*"
    Range("Estad�sticas!K27") = AnimPorParir(270, "H")
    Range("Estad�sticas!K28") = AnimPorParir(270, "R")
    Range("Estad�sticas!K29") = _
      (Range("Estad�sticas!K27") + Range("Estad�sticas!K28"))
    Range("Estad�sticas!J30") = _
      "*Incluye animales sin Dx. Gestaci�n"
    'Columna L
    'Range("Estad�sticas!L1") = Format(Date, "dd-mmm-yy")
    Range("Estad�sticas!L4") = "Promedio d�as a 1er. Servicio"
    Range("Estad�sticas!L5") = Format(pD1S(1), "#")
    Range("Estad�sticas!L6") = Format(pD1S(2), "#")
    Range("Estad�sticas!L7") = Format(pD1S(3), "#")
    Range("Estad�sticas!L8") = Format(pD1S(4), "#")
    Range("Estad�sticas!L9") = Format(pD1S(5), "#")
    Range("Estad�sticas!L10") = Format(pD1S(), "#")
    Range("Estad�sticas!L15") = "Promedio Edad a 1er. Servicio"
    If Not pEdad1Serv() = "ND" Then _
      Range("Estad�sticas!L19") = pEdad1Serv() & "m" Else _
      Range("Estad�sticas!L19") = pEdad1Serv()
    If Not pEdad1Serv("P") = "ND" Then _
      Range("Estad�sticas!L20") = pEdad1Serv("P") & "m" Else _
      Range("Estad�sticas!L20") = pEdad1Serv("P")
    LactMinMaxCorral
    'Otros
    On Error GoTo 0
End Sub

Function cvProy305d()
' Calcula el coeficiente de variaci�n de la _
  Proyecci�n a 305 d�as
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
' Calcula el % de Dx Gestantes Positivos en un per�odo _
  por los dias. _
  Si los dias = 0 entonces se calcular� para el a�o anterior.
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
' Contar Parto,Lactancia M�nima y Maxima por Corral
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
    Range("Estad�sticas!G39") = LmMC(1, 2)
    Range("Estad�sticas!H39") = LmMC(1, 3)
    Range("Estad�sticas!G40") = LmMC(2, 2)
    Range("Estad�sticas!H40") = LmMC(2, 3)
    Range("Estad�sticas!G41") = LmMC(3, 2)
    Range("Estad�sticas!H41") = LmMC(3, 3)
    Range("Estad�sticas!G42") = LmMC(4, 2)
    Range("Estad�sticas!H42") = LmMC(4, 3)
    Range("Estad�sticas!G43") = LmMC(5, 2)
    Range("Estad�sticas!H43") = LmMC(5, 3)
End Sub

Function nCaloresPerdidos()
' Calcula el n�mero de Calores Perdidos
' F�rmula desarrollada por JP
' Se agregan 11 d�as por desviaci�n del ciclo estral
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

Function nAbortos(Mes, A�o)
    ' Contabiliza los abortos del mes
    Dim nA As Long
    Dim fDesde, fHasta As Date
    Dim rCelda As Range
    nA = 0
    fDesde = CDate("1," & Mes & "," & A�o)
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

Function nPartos(Mes, A�o)
    ' Contabiliza los partos del mes
    Dim nA As Long
    Dim fDesde, fHasta As Date
    Dim rCelda As Range
    nA = 0
    fDesde = CDate("1," & Mes & "," & A�o)
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

Function nBajas(Mes, A�o)
    ' Contabiliza los abortos del mes
    Dim nA As Long
    Dim fDesde, fHasta As Date
    Dim rCelda As Range
    nA = 0
    fDesde = CDate("1," & Mes & "," & A�o)
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
' Calcula el promedio de los d�as al primer calor de anim. en Hato
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
                        ' Si est�n Gestantes
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
' Calcula promedio de d�as a 1er servicio de anim. en hato.
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
                        ' Si est�n Gestantes
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
' Calcula promedio de d�as abiertos, optativo por parto
    Dim rCelda As Range
    Dim iCuenta, iSuma, iPuntero As Long
    iCuenta = 0: iSuma = 0: 'vRes = "ND"
    On Error Resume Next
' Todos los Animales Servidos y Gestantes
    ' Cualquier Parto y Gestante
    If IsMissing(Parto) Then
            iPuntero = 101
            For Each rCelda In Range("Tabla1[Arete]")
                ' Si est�n Gestantes
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
' Promedio de D�as en Leche por Parto
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
' Promedio de Producci�n por Parto
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
' Promedio de Producci�n en Hato
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
' Promedio de Proyecci�n a 305 d�as en Hato
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
' Extrae la producci�n mensual de la vaca especificada
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
' IntervaloCalores(1) s�lo vacas
' IntervaloCalores(2) s�lo reemplazos
' Debido al loop que ejecuta en Eventos, el rendimiento del _
  sistema se ve dram�ticamente disminuido
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
' Promedio de Proyecci�n a 305 d�as      en Hato
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
            ' S�lo se individualiza hasta parto 3
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
' Extrae la producci�n mensual de la vaca especificada
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
' Totaliza el n�mero de animales por corral _
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
' Fuente: "La Evaluaci�n de tasa de embarazo sobre la _
  fertilidad de la vaca" _
  www.aipl.usda.gov/reference/fertility/DPR-rpt.es.htm
  TasaEmbarazo = 21 / (pDAb() - Range("Configuracion!C6") + 11)
End Function

Function tDesarrollo()
' Totaliza el n�mero de hembras en Desarrollo
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
' Totaliza el n�mero de hembras Lactantes
    Dim rCelda As Range
    tLactantes = 0
    For Each rCelda In Range("Tabla2[Arete]")
        If Date - CDate(rCelda.Offset(0, 4)) <= 45 And _
        rCelda.Offset(0, 13) = "H" Then _
          tLactantes = tLactantes + 1
    Next
End Function

Function tNovillas()
' Totaliza el n�mero de Novillas
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
' Totaliza el n�mero de Vacas en Producci�n
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
' Totaliza el n�mero de Vacas en Producci�n
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
' Totaliza el n�mero de Vacas Secas
    tVacasSecas = WorksheetFunction. _
      CountBlank _
      ( _
      Range("Tabla1[Prod.]") _
      )
End Function

Function tVaquillas(Optional StatusRepro As String)
' Totaliza el n�mero de Vaquillas seg�n estatus reproductivo _
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
    'Devuelve el rengl�n donde se encuentra la fecha
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
' insertar los par�metros del momento
    'Dim ws As Worksheet
    'Dim iR, iRenglon As Integer
    Application.ScreenUpdating = CBool(Range("Desarrollador!B6"))
    Application.DisplayStatusBar = True
    Set ws = Worksheets("AcumStat")
    'iRenglon = Tama�oTabla("Tabla14") + 5
    If IndiceFecha("Tabla14") = 0 Then
            iR = Tama�oTabla("Tabla14") + 2
        Else
            iR = IndiceFecha("Tabla14")
    End If
    'iR = 2
    iC = 1
    With ws
        ' Mes/a�o
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
        ' Vcs Producci�n
        .Cells(iR, iC) = tVacasProd()
        '.Cells(iRenglon, iC) = .Cells(iR, iC)
        cntdr
        ' Vcas Secas
        .Cells(iR, iC) = tVacasSecas()
        '.Cells(iRenglon, iC) = .Cells(iR, iC)
        cntdr
        ' Producci�n
        .Cells(iR, iC) = Int(WorksheetFunction.Sum(Range("Tabla1[Prod.]")))
        '.Cells(iRenglon, iC) = .Cells(iR, iC)
        cntdr
        ' Persistencia
        .Cells(iR, iC) = Format(WorksheetFunction.Average(Range("Tabla15[Persistencia]")), "0.0")
        '.Cells(iRenglon, iC) = .Cells(iR, iC)
        cntdr
        ' Producci�n m�xima
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
        ' d�as en leche
        .Cells(iR, iC) = pDEL()
        '.Cells(iRenglon, iC) = .Cells(iR, iC)
        cntdr
        ' d�as seca
        .Cells(iR, iC) = WorksheetFunction.Average(Range("Tabla4[D�asSeca]"))
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
        ' d�as abiertos
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
        ' % gestantes 1� servicio
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
        ' Edad 1� Serv
        .Cells(iR, iC) = Format(pEdad1Serv(), "0")
        '.Cells(iRenglon, iC) = .Cells(iR, iC)
        cntdr
        ' Edad 1� Serv Gestante
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
        ' PicoProd 1�
        .Cells(iR, iC) = PromParam("PicoProd", 1)
        '.Cells(iRenglon, iC) = .Cells(iR, iC)
        cntdr
        ' PicoProd 2�
        .Cells(iR, iC) = PromParam("PicoProd", 2)
        '.Cells(iRenglon, iC) = .Cells(iR, iC)
        cntdr
        ' PicoProd 3�
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
        ' DiasPico 1�
        .Cells(iR, iC) = PromParam("DiasPico", 1)
        '.Cells(iRenglon, iC) = .Cells(iR, iC)
        cntdr
        ' DiasPico 2�
        .Cells(iR, iC) = PromParam("DiasPico", 2)
        '.Cells(iRenglon, iC) = .Cells(iR, iC)
        cntdr
        ' DiasPico 3�
        .Cells(iR, iC) = PromParam("DiasPico", 3)
        '.Cells(iRenglon, iC) = .Cells(iR, iC)
        cntdr
        ' DiasPico 4+
        .Cells(iR, iC) = PromParam("DiasPico", 4)
        '.Cells(iRenglon, iC) = .Cells(iR, iC)
        cntdr
        ' Proy305d 1�
        .Cells(iR, iC) = PromParam("Proy305d", 1)
        '.Cells(iRenglon, iC) = .Cells(iR, iC)
        cntdr
        ' Proy305d 2�
        .Cells(iR, iC) = PromParam("Proy305d", 2)
        '.Cells(iRenglon, iC) = .Cells(iR, iC)
        cntdr
        ' Proy305d 3�
        .Cells(iR, iC) = PromParam("Proy305d", 3)
        '.Cells(iRenglon, iC) = .Cells(iR, iC)
        cntdr
        ' Proy305d 4�
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
        ' tVacas 1� Parto
        .Cells(iR, iC) = WorksheetFunction. _
          CountIf(Range("tabla1[Parto]"), 1)
        '.Cells(iRenglon, iC) = .Cells(iR, iC)
        cntdr
        ' tVacas 2� Parto
        .Cells(iR, iC) = WorksheetFunction. _
          CountIf(Range("tabla1[Parto]"), 2)
        '.Cells(iRenglon, iC) = .Cells(iR, iC)
        cntdr
        ' tVacas 3� Parto
        .Cells(iR, iC) = WorksheetFunction. _
          CountIf(Range("tabla1[Parto]"), 3)
        '.Cells(iRenglon, iC) = .Cells(iR, iC)
        cntdr
        ' tVacas 4+ Parto
        .Cells(iR, iC) = WorksheetFunction. _
          CountIf(Range("tabla1[Parto]"), ">3")
        '.Cells(iRenglon, iC) = .Cells(iR, iC)
        cntdr
        ' PromProducci�n
        .Cells(iR, iC) = pProd()
        '.Cells(iRenglon, iC) = .Cells(iR, iC)
        cntdr
        ' PromProducci�n 1
        .Cells(iR, iC) = pProd(1)
        '.Cells(iRenglon, iC) = .Cells(iR, iC)
        cntdr
        ' PromProducci�n 2
        .Cells(iR, iC) = pProd(2)
        '.Cells(iRenglon, iC) = .Cells(iR, iC)
        cntdr
        ' PromProducci�n 3
        .Cells(iR, iC) = pProd(3)
        '.Cells(iRenglon, iC) = .Cells(iR, iC)
        cntdr
        ' PromProducci�n 4
        .Cells(iR, iC) = pProd(4)
        '.Cells(iRenglon, iC) = .Cells(iR, iC)
        cntdr
        ' Peristencia 1�
        .Cells(iR, iC) = PromParam("Persistencia", 1)
        '.Cells(iRenglon, iC) = .Cells(iR, iC)
        cntdr
        ' Peristencia 2�
        .Cells(iR, iC) = PromParam("Persistencia", 2)
        '.Cells(iRenglon, iC) = .Cells(iR, iC)
        cntdr
        ' Peristencia 3�
        .Cells(iR, iC) = PromParam("Persistencia", 3)
        '.Cells(iRenglon, iC) = .Cells(iR, iC)
        cntdr
        ' Peristencia 4�
        .Cells(iR, iC) = PromParam("Persistencia", 4)
        '.Cells(iRenglon, iC) = .Cells(iR, iC)
        cntdr
        ' PromProd 1�
        .Cells(iR, iC) = pProd(1)
        '.Cells(iRenglon, iC) = .Cells(iR, iC)
        cntdr
        ' PromProd 2�
        .Cells(iR, iC) = pProd(2)
        '.Cells(iRenglon, iC) = .Cells(iR, iC)
        cntdr
        ' PromProd 3�
        .Cells(iR, iC) = pProd(3)
        '.Cells(iRenglon, iC) = .Cells(iR, iC)
        cntdr
        ' PromProd 4+
        .Cells(iR, iC) = pProd(4)
        '.Cells(iRenglon, iC) = .Cells(iR, iC)
        cntdr
        ' promDEL 1�
        .Cells(iR, iC) = pDEL(1)
        '.Cells(iRenglon, iC) = .Cells(iR, iC)
        cntdr
        ' promDEL 2�
        .Cells(iR, iC) = pDEL(2)
        '.Cells(iRenglon, iC) = .Cells(iR, iC)
        cntdr
        ' promDEL 3�
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
    Application.StatusBar = "Salvando Informaci�n"
    ActiveWorkbook.Save
    'OrdenarAcumStat
    Application.StatusBar = False
    Application.ScreenUpdating = True

End Sub

Private Sub cntdr()
    ' Incrementa contador
    iC = iC + 1
    ' muestra avance en barra de estado
    Application.StatusBar = "Calculando indicadores de desempe�o... " & _
      Format(iC / 130, "0%")
End Sub

