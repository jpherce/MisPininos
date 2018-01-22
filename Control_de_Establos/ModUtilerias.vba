Attribute VB_Name = "ModUtilerias"
' Ultima modificación: 5-oct-2015
' Herramientas del sistema
Option Explicit
Dim rCelda As Object

Private Sub ConvertirFechas()
    Range("Desarrollador!B20") = "T"
    'Hato2
    Application.StatusBar = _
      "Convirtiendo Fechas en Hato2"
    For Each rCelda In Range("Tabla15[F.Secado]")
        ConvertirFechas1
    Next rCelda
    'For Each rCelda In Range("Tabla15[F.Prod]")
    '    ConvertirFechas1
    'Next rCelda
    'For Each rCelda In Range("Tabla15[F.Revision2]")
    '   ConvertirFechas1
    'Next rCelda
    'For Each rCelda In Range("Tabla15[F.Nacim]")
    '    ConvertirFechas1
    'Next rCelda
    'Eventos
    Application.StatusBar = _
      "Convirtiendo Fechas en Eventos"
    For Each rCelda In Range("Tabla6[Fecha]")
        ConvertirFechas1
    Next rCelda
    'BajaReemplazos
    Application.StatusBar = _
      "Convirtiendo Fechas en BajaReemplazos"
    'For Each rCelda In Range("Tabla5[F.Nacim]")
    '    ConvertirFechas1
    'Next rCelda
    For Each rCelda In Range("Tabla5[F.Servicio]")
        ConvertirFechas1
    Next rCelda
    'For Each rCelda In Range("Tabla5[F.Vacuna]")
    '    ConvertirFechas1
    'Next rCelda
    'For Each rCelda In Range("Tabla5[F.Iman]")
    '    ConvertirFechas1
    'Next rCelda
    For Each rCelda In Range("Tabla5[F.Baja]")
        ConvertirFechas1
    Next rCelda
    'LactanciasAnteriores
    Application.StatusBar = _
      "Convirtiendo Fechas en LactanciasAnteriores"
    For Each rCelda In Range("Tabla4[F.Parto]")
        ConvertirFechas1
    Next rCelda
    For Each rCelda In Range("Tabla4[F.Servicio]")
        ConvertirFechas1
    Next rCelda
    For Each rCelda In Range("Tabla4[F.Terminación]")
        ConvertirFechas1
    Next rCelda
    'Reemplazos
    Application.StatusBar = _
      "Convirtiendo Fechas en Reemplazos"
    Worksheets("Reemplazos").Activate
    Application.Run "Desproteger" 'Mód2
    For Each rCelda In Range("Tabla2[F.Nacim]")
        ConvertirFechas1
    Next rCelda
    For Each rCelda In Range("Tabla2[F.Servicio]")
        ConvertirFechas1
    Next rCelda
    For Each rCelda In Range("Tabla2[FxParir]")
        ConvertirFechas1
    Next rCelda
    'For Each rCelda In Range("Tabla2[F.Vacuna]")
    '    ConvertirFechas1
    'Next rCelda
    'For Each rCelda In Range("Tabla2[F.Iman]")
    '    ConvertirFechas1
    'Next rCelda
    'For Each rCelda In Range("Tabla2[F.Revision]")
    '    ConvertirFechas1
    'Next rCelda
    'For Each rCelda In Range("Tabla2[FxParir]")
    '    ConvertirFechas1
    'Next rCelda
    Application.Run "Proteger" 'Mód2
    'Hato
    Application.StatusBar = _
      "Convirtiendo Fechas en Hato"
    Worksheets("Hato").Activate
    Application.Run "Desproteger" 'Mód2
    For Each rCelda In Range("Tabla1[F.Parto]")
        ConvertirFechas1
    Next rCelda
    For Each rCelda In Range("Tabla1[F.Servicio]")
        ConvertirFechas1
    Next rCelda
    For Each rCelda In Range("Tabla1[FxSecar]")
        ConvertirFechas1
    Next rCelda
    For Each rCelda In Range("Tabla1[FxParir]")
        ConvertirFechas1
    Next rCelda
    Application.Run "Proteger" 'Mód2
    Range("Desarrollador!B20").Clear
    Application.StatusBar = False
End Sub

Private Sub ConvertirFechas1()
    If IsEmpty(rCelda.Offset(0, 0)) Then Exit Sub
    If IsDate(rCelda.Offset(0, 0)) Then
        With rCelda.Offset(0, 0)
            .Value = CDate(rCelda.Offset(0, 0))
            .NumberFormat = "d-mmm-yy"
        End With
    End If
End Sub

Private Sub ConvertirNumeros()
    Range("Desarrollador!B20") = "T"
    'Hato2
    Application.StatusBar = _
      "Convirtiendo números en Hato2"
    For Each rCelda In Range("Tabla15[Arete]")
        ConvertirNumeros1
    Next rCelda
    For Each rCelda In Range("Tabla15[d1S]")
        ConvertirNumeros1
    Next rCelda
    For Each rCelda In Range("Tabla15[dAbiertos]")
        ConvertirNumeros1
    Next rCelda
    For Each rCelda In Range("Tabla15[30d]")
        ConvertirNumeros1
    Next rCelda
    For Each rCelda In Range("Tabla15[60d]")
        ConvertirNumeros1
    Next rCelda
    For Each rCelda In Range("Tabla15[90d]")
        ConvertirNumeros1
    Next rCelda
    For Each rCelda In Range("Tabla15[120d]")
        ConvertirNumeros1
    Next rCelda
    For Each rCelda In Range("Tabla15[150d]")
        ConvertirNumeros1
    Next rCelda
    For Each rCelda In Range("Tabla15[180d]")
        ConvertirNumeros1
    Next rCelda
    For Each rCelda In Range("Tabla15[210d]")
        ConvertirNumeros1
    Next rCelda
    For Each rCelda In Range("Tabla15[240d]")
        ConvertirNumeros1
    Next rCelda
    For Each rCelda In Range("Tabla15[270d]")
        ConvertirNumeros1
    Next rCelda
    For Each rCelda In Range("Tabla15[300d]")
        ConvertirNumeros1
    Next rCelda
    For Each rCelda In Range("Tabla15[ProdAcum]")
        ConvertirNumeros1
    Next rCelda
    For Each rCelda In Range("Tabla15[Proy305d]")
        ConvertirNumeros1
    Next rCelda
    'Eventos
    Application.StatusBar = _
      "Convirtiendo números en Eventos"
    For Each rCelda In Range("Tabla6[Arete]")
        ConvertirNumeros1
    Next rCelda
    'BajaReemplazos
    Application.StatusBar = _
      "Convirtiendo números en BajaReemplazos"
    For Each rCelda In Range("Tabla5[Arete]")
        ConvertirNumeros1
    Next rCelda
    For Each rCelda In Range("Tabla5[Peso]")
        ConvertirNumeros1
    Next rCelda
    For Each rCelda In Range("Tabla5[Servicio]")
        ConvertirNumeros1
    Next rCelda
    'For Each rCelda In Range("Tabla5[Edad1Serv]")
    '    ConvertirNumeros1
    'Next rCelda
    'LactAnteriores
    Application.StatusBar = _
      "Convirtiendo números en LactAnteriores"
    For Each rCelda In Range("Tabla4[Arete]")
        ConvertirNumeros1
    Next rCelda
    For Each rCelda In Range("Tabla4[Parto]")
        ConvertirNumeros1
    Next rCelda
    For Each rCelda In Range("Tabla4[Servicio]")
        ConvertirNumeros1
    Next rCelda
    For Each rCelda In Range("Tabla4[DiasLactancia]")
        ConvertirNumeros1
    Next rCelda
    For Each rCelda In Range("Tabla4[DíasSeca]")
        ConvertirNumeros1
    Next rCelda
    For Each rCelda In Range("Tabla4[ProdAcum]")
        ConvertirNumeros1
    Next rCelda
    For Each rCelda In Range("Tabla4[Proy.305d]")
        ConvertirNumeros1
    Next rCelda
    For Each rCelda In Range("Tabla4[Dias1Serv]")
        ConvertirNumeros1
    Next rCelda
    For Each rCelda In Range("Tabla4[DiasAbierta]")
        ConvertirNumeros1
    Next rCelda
    'Reemplazos
    Application.StatusBar = _
      "Convirtiendo números en Reemplazos"
    Worksheets("Reemplazos").Activate
    Application.Run "Desproteger" 'Mód2
    For Each rCelda In Range("Tabla2[Arete]")
        ConvertirNumeros1
    Next rCelda
    For Each rCelda In Range("Tabla2[Servicio]")
        ConvertirNumeros1
    Next rCelda
    For Each rCelda In Range("Tabla2[PesoCorporal]")
        ConvertirNumeros1
    Next rCelda
    'For Each rCelda In Range("Tabla2[Edad1Serv]")
    '    ConvertirNumeros1
    'Next rCelda
    'For Each rCelda In Range("Tabla2[EdadAlParto]")
    '    ConvertirNumeros1
    'Next rCelda
    'For Each rCelda In Range("Tabla2[Edad2]")
    '    ConvertirNumeros1
    'Next rCelda
    'For Each rCelda In Range("Tabla2[dServicio]")
    '    ConvertirNumeros1
    'Next rCelda
    Application.Run "Proteger" 'Modulo2
    'Hato
    Application.StatusBar = _
      "Convirtiendo números en Hato"
    Worksheets("Hato").Activate
    Application.Run "Desproteger" 'Modulo2
    For Each rCelda In Range("Tabla1[Arete]")
        ConvertirNumeros1
    Next rCelda
    For Each rCelda In Range("Tabla1[Prod.]")
        ConvertirNumeros1
    Next rCelda
    For Each rCelda In Range("Tabla1[DEL]")
        ConvertirNumeros1
    Next rCelda
    For Each rCelda In Range("Tabla1[Parto]")
        ConvertirNumeros1
    Next rCelda
    For Each rCelda In Range("Tabla1[Servicio]")
        ConvertirNumeros1
    Next rCelda
    'For Each rCelda In Range("Tabla1[dServicio]")
    '    ConvertirNumeros1
    'Next rCelda
    Application.Run "Proteger" 'Modulo2
    Range("Desarrollador!B20").Clear
    Application.StatusBar = False
End Sub

Private Sub ConvertirNumeros1()
    If IsEmpty(rCelda.Offset(0, 0)) Then Exit Sub
    If IsNumeric(rCelda.Offset(0, 0)) Then
        With rCelda.Offset(0, 0)
            .Value = CDbl(rCelda.Offset(0, 0))
        End With
    End If
End Sub

Private Sub RenombrarCamposBD()
    On Error Resume Next
    Range("Desarrollador!B20") = "T"
    'Hato2
    Application.StatusBar = _
      "Reparando Tabla15"
    Range("Hato2!A1") = "Arete"
    Range("Hato2!B1") = "d1S"
    Range("Hato2!C1") = "dAbiertos"
    Range("Hato2!D1") = "30d"
    Range("Hato2!E1") = "60d"
    Range("Hato2!F1") = "90d"
    Range("Hato2!G1") = "120d"
    Range("Hato2!H1") = "150d"
    Range("Hato2!I1") = "180d"
    Range("Hato2!J1") = "210d"
    Range("Hato2!K1") = "240d"
    Range("Hato2!L1") = "270d"
    Range("Hato2!M1") = "300d"
    Range("Hato2!N1") = "ProdAcum"
    Range("Hato2!O1") = "Proy305d"
    Range("Hato2!P1") = "F.Secado"
    'Eventos
    Application.StatusBar = _
      "Reparando Tabla6"
    Range("Eventos!A1") = "Arete"
    Range("Eventos!B1") = "Fecha"
    Range("Eventos!C1") = "Clave"
    Range("Eventos!D1") = "Observaciones"
    Range("Eventos!E1") = "Responsable"
    Range("Eventos!F1") = "Usuario"
    Range("Eventos!G1") = "F.Captura"
    Range("Eventos!H1") = "H.Captura"
    'BajaReemplazos
    Application.StatusBar = _
      "Reparando Tabla5"
    Range("BajaReemplazos!A1") = "Arete"
    Range("BajaReemplazos!B1") = "Peso"
    Range("BajaReemplazos!C1") = "Edad"
    Range("BajaReemplazos!D1") = "F.Nacim"
    Range("BajaReemplazos!E1") = "Servicio"
    Range("BajaReemplazos!F1") = "F.Servicio"
    Range("BajaReemplazos!G1") = "Semental"
    Range("BajaReemplazos!H1") = "Técnico"
    Range("BajaReemplazos!I1") = "Status"
    Range("BajaReemplazos!J1") = "Clave1"
    Range("BajaReemplazos!K1") = "Clave2"
    Range("BajaReemplazos!L1") = "F.Baja"
    Range("BajaReemplazos!M1") = "CausaBaja"
    'LactanciasAnteriores
    Application.StatusBar = _
      "Reparando Tabla4"
    'Worksheets("LactanciasAnteriores").Activate
    'Application.Run "Desproteger" 'Modulo2
    Range("LactanciasAnteriores!A1") = "Arete"
    Range("LactanciasAnteriores!B1") = "Parto"
    Range("LactanciasAnteriores!C1") = "F.Parto"
    Range("LactanciasAnteriores!D1") = "Servicio"
    Range("LactanciasAnteriores!E1") = "F.Servicio"
    Range("LactanciasAnteriores!F1") = "Semental"
    Range("LactanciasAnteriores!G1") = "Técnico"
    Range("LactanciasAnteriores!H1") = "Status"
    Range("LactanciasAnteriores!I1") = "Clave1"
    Range("LactanciasAnteriores!J1") = "Clave2"
    Range("LactanciasAnteriores!K1") = "DiasLactancia"
    Range("LactanciasAnteriores!L1") = "DíasSeca"
    Range("LactanciasAnteriores!M1") = "ProdAcum"
    Range("LactanciasAnteriores!N1") = "Proy.305d"
    Range("LactanciasAnteriores!O1") = "dias1Serv"
    Range("LactanciasAnteriores!P1") = "diasAbierta"
    Range("LactanciasAnteriores!Q1") = "F.Terminación"
    Range("LactanciasAnteriores!R1") = "TipoTermino"
    Range("LactanciasAnteriores!S1") = "CausaTermino"
    'Reemplazos
    Application.StatusBar = _
      "Reparando Tabla2"
    Worksheets("Reemplazos").Activate
    Application.Run "Desproteger" 'Modulo2
    Range("Reemplazos!A1") = "Arete"
    Range("Reemplazos!B1") = "Corral"
    Range("Reemplazos!C1") = "PesoCorporal"
    Range("Reemplazos!D1") = "Edad"
    Range("Reemplazos!E1") = "F.Nacim"
    Range("Reemplazos!F1") = "Servicio"
    Range("Reemplazos!G1") = "F.Servicio"
    Range("Reemplazos!H1") = "Semental"
    Range("Reemplazos!I1") = "Técnico"
    Range("Reemplazos!J1") = "Status"
    Range("Reemplazos!K1") = "FxParir"
    Range("Reemplazos!L1") = "Clave1"
    Range("Reemplazos!M1") = "Clave2"
    Range("Reemplazos!N1") = "Sexo"
    Application.Run "Proteger" 'Modulo2
    'Hato
    Application.StatusBar = _
      "Reparando Tabla1"
    Worksheets("Hato").Activate
    Application.Run "Desproteger" 'Modulo2
    Range("Hato!A1") = "Arete"
    Range("Hato!B1") = "Corral"
    Range("Hato!C1") = "Prod."
    Range("Hato!D1") = "DEL"
    Range("Hato!E1") = "Parto"
    Range("Hato!F1") = "F.Parto"
    Range("Hato!G1") = "Servicio"
    Range("Hato!H1") = "F.Servicio"
    Range("Hato!I1") = "Semental"
    Range("Hato!J1") = "Técnico"
    Range("Hato!K1") = "Status"
    Range("Hato!L1") = "FxSecar"
    Range("Hato!M1") = "FxParir"
    Range("Hato!N1") = "Clave1"
    Range("Hato!O1") = "Clave2"
    Application.Run "Proteger" 'Modulo2
    Range("Desarrollador!B20").Clear
    On Error GoTo 0
    Application.StatusBar = False
End Sub

Private Sub RepararBaseDatos()
    'RenombrarCamposBD
    ConvertirNumeros
    ConvertirFechas
End Sub
