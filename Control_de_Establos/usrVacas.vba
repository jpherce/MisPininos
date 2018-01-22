VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} usrVacas 
   Caption         =   "Control de Establos"
   ClientHeight    =   5460
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4335
   OleObjectBlob   =   "usrVacas.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "usrVacas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' Ultima modificación: 17.01.2018
' Corr. sMetadatos en ProcesarBajas
' Corr. metadatos en Calores y Servicios
' Añadir DCC en Movimientos
' Corrección de metadatos en calores y servicios 19.11.17
' Sustitución de R por P en Metadatos 18.11.17
' Corrección mensajes del secado 18.11-17
' Application.StatusBar en "Cerrar" 18-11-17
' Incluye Corral a Metadatos en Prod 12.11.17
' Calcular persistencia 1.11.17
' cmboEvento.ListRows = 13 13.10.17
' Se corrigió error en captura por lotes 10.10.17
' Se eliminó condicionante de Consecutivo de eventos 21.04.17
' Se corrigió entrada de arete 10.04.17
' Ver 14.1144 Se incorporó medidas de seguridad en hojas
Dim bFlagError, bAreteNoEncontrado, bAceptar, _
  bMovCorral As Boolean
Dim iAviso, iARowInfoVital, iARH, iARH2, _
  iAR, nMM As Long
Dim sMsjTitulo, sUltimoRegistro, sEnf, sTextoMsj2, _
  sPartoDet, sLocAnimal, sLocPrevia, sEvento, sObserv, _
  sResp, mPadreCria As String
Dim sMetadato As String
Dim sArete, mMadre, mCorral, mCria1, mCria2 As Variant
Dim ws, wsE, wsH, wsR, wsH2, wsIV As Worksheet

Private Sub ActualizarTextBox4()
    Select Case Me.cmboEvento
        Case vbNullString
            ' Parar Ahorrar tiempo de ejecución
        Case "Secar"
            If Cells(iARH, 12) = vbNullString Then
                    Me.TextBox4 = "Vacía"
                Else
                    If IsDate(Cells(iARH, 13)) Then _
                      Me.TextBox4 = _
                      Format(CDate(Cells(iARH, 13)), _
                      "dd-mmm-yy")
            End If
        Case "Movimiento"
            If sLocAnimal = "H" Then
                Me.TextBox4 = Cells(iARH, 2)
            End If
            If sLocAnimal = "R" Then
                Me.TextBox4 = wsR.Cells(iAR, 2)
            End If
        Case "Producción"
            If Cells(iARH, 3) = vbNullString Then
                    Me.TextBox4 = Format(0, "#0.0")
                Else
                    Me.TextBox4 = _
                      Format(Cells(iARH, 3), _
                      "#0.0")
            End If
        Case "Pesaje"
            If wsR.Cells(iAR, 3) = vbNullString Then
                    Me.TextBox4 = Format(0, "#0.0")
                Else
                    Me.TextBox4 = _
                      Format(wsR.Cells(iAR, 3), _
                      "#0.0")
            End If
    End Select
End Sub

Private Sub AdicionarCria1()
    ' Adicionar Reemplazos Cría1
    Dim iRenglon, iNRowInfoVital As Long
    Set ws = Worksheets("Reemplazos")
    Desproteger1
    iRenglon = TamañoTabla("Tabla2") + 2
    If IsNumeric(Me.TextBox6) Then _
      ws.Cells(iRenglon, 1) = mCria1 'Arete
    ws.Cells(iRenglon, 2) = _
      Range("Configuracion!C13") 'Corral
    With ws.Cells(iRenglon, 5) 'F.Nacim
        .Value = CDate(Me.txtFecha)
        .NumberFormat = "d-mmm-yy"
    End With
    ws.Cells(iRenglon, 14) = Left(sPartoDet, 1) 'Sexo
    ' Adicionar InfoVitalicia Cría1
    iNRowInfoVital = TamañoTabla("Tabla8") + 2
    With wsIV
        If IsNumeric(Me.TextBox6) Then _
          .Cells(iNRowInfoVital, 1) = mCria1 'Arete
        With .Cells(iNRowInfoVital, 3)  'F.Nacim
            .Value = CDate(Me.txtFecha)
            .NumberFormat = "d-mmm-yy"
        End With
        .Cells(iNRowInfoVital, 4) = _
          mPadreCria 'Padre
        .Cells(iNRowInfoVital, 5) = mMadre 'Madre
    End With
End Sub

Private Sub AdicionarCria2()
    ' Adicionar Reemplazos Cria2
    Dim iRenglon, iNRowInfoVital As Long
    iRenglon = TamañoTabla("Tabla2") + 2
    If IsNumeric(Me.TextBox7) Then _
      wsR.Cells(iRenglon, 1) = mCria2 'Arete
    wsR.Cells(iRenglon, 2) = _
      Range("Configuracion!C13") 'Corral
    With wsR.Cells(iRenglon, 5) 'F.Nacim
        .Value = CDate(Me.txtFecha)
        .NumberFormat = "d-mmm-yy"
    End With
    wsR.Cells(iRenglon, 14) = Left(sPartoDet, 1) 'Sexo
    ' Adicionar infoVitalicia
    iNRowInfoVital = TamañoTabla("Tabla8") + 2
    With wsIV
        .Cells(iNRowInfoVital, 1) = mCria2  'Arete
        With .Cells(iNRowInfoVital, 3) 'F.Nacim
            .Value = CDate(Me.txtFecha)
            .NumberFormat = "d-mmm-yy"
        End With
        .Cells(iNRowInfoVital, 4) = _
          mPadreCria 'Padre
        .Cells(iNRowInfoVital, 5) = mMadre 'Madre
    End With
End Sub

Private Sub AltaHoja2()
' Alta de animales en Hato2
    Dim iRenglon As Long
    Set ws = Worksheets("Hato2")
    Desproteger
    iRenglon = TamañoTabla("Tabla15") + 2
    ws.Cells(iRenglon, 1) = sArete
    If Not CBool(Range("Desarrollador!B6")) Then _
      ws.Visible = xlSheetVeryHidden
End Sub

Private Sub AltaInfovital()
' Alta de animales en InfoVitalicia
    Dim iRenglon As Long
    Set ws = Worksheets("InfoVitalicia")
    Desproteger
    iRenglon = TamañoTabla("Tabla8") + 2
    With ws
        .Cells(iRenglon, 1) = sArete
        Select Case Me.cmboEvento
            Case Is = "Alta"
                ' Se agregan sólo si son requeridos
                If IsDate(Me.TextBox4) Then
                    With ws.Cells(iRenglon, 3)
                        .Value = CDate(Me.TextBox4)
                        .NumberFormat = "d-mmm-yy"
                    End With
                End If
                .Cells(iRenglon, 4) = UCase(Me.TextBox5) 'Padre
                If IsNumeric(Me.TextBox6) Then 'Madre
                        .Cells(iRenglon, 5) = CDbl(Me.TextBox6)
                    Else
                        .Cells(iRenglon, 5) = Me.TextBox6
                End If
                .Cells(iRenglon, 6) = UCase(Me.TextBox7) 'Raza
        End Select
    End With
    If Not CBool(Range("Desarrollador!B6")) Then _
      ws.Visible = xlSheetVeryHidden
End Sub

Private Sub Avisos()
' Mensajes del formulario
    Dim sTextoMsj, sMensaje As String
    Dim bTipoMensaje As Boolean
    bTipoMensaje = False
    bFlagError = True
    Select Case iAviso
        Case 1 '100
            sTextoMsj = _
              "Fecha del Evento es anterior a la Fecha del " _
                & sTextoMsj2 & " registrado"
        Case 2 '100
            sTextoMsj = _
              "Fecha del Evento es igual o anterior a la Fecha del " _
                & Me.cmboEvento & " registrado"
        Case 3 '300
            sTextoMsj = _
              "Este Evento ya está registrado"
        Case 4
            sTextoMsj = _
              "Este animal no es hembra"
        Case 5  'Aviso Informativo
            bTipoMensaje = True
            bFlagError = False
            sTextoMsj = "Animal previamente reportado como Gestante." _
            & Chr(13) & "Posiblemente tuvo una Reabsorción o un Aborto."
        Case 6
            sTextoMsj = "Animal sin Servicios"
        Case 7 'Aviso Informativo
            bTipoMensaje = True
            bFlagError = False
            sTextoMsj = "Animal con menos de " & _
              Range("Configuracion!C5") & " días de servicio."
        Case 8
            sTextoMsj = "Animal sin Servicios." & Chr(13) & _
              "El último Servicio reportado fue un Calor."
        Case 9 'Aviso Informativo
            bFlagError = False
            sTextoMsj = "Animal previamente reportado como Seca"
        Case 10 'Aviso Informativo
            bTipoMensaje = True
            bFlagError = False
            sTextoMsj = "La lactancia sólo tuvo " & _
              Cells(iARH, 4) & " días de duración."
        Case 11
            sTextoMsj = "Animal previamente reportado en este corral"
        Case 12
            sTextoMsj = sTextoMsj2 & Chr(13) & _
              "El Dato ingresado es un valor Negativo"
        Case 13
        Case 14 'Aviso Informativo
            bTipoMensaje = True
            bFlagError = False
            sTextoMsj = _
              "El intervalo entre Calores o Servicios es mayor a 36 días," _
              & Chr(13) & _
              "esta condición sugiere que se han perdido calores observados." _
              & Chr(13) & _
              "Se recomienda prestar atención a los calores."
        Case 15
            sTextoMsj = sTextoMsj2 & Chr(13) & _
              "El Dato ingresado no es Numérico"
        Case 16 'Aviso Informativo
            bTipoMensaje = True
            bFlagError = False
            sTextoMsj = _
              "El intervalo entre Calores o Servicios es menor a 18 días," _
              & Chr(13) & _
              "esta condición puede deberse a un Quiste Folicular." & Chr(13) & _
              "Se sugiere Reportar al Médico."
        Case 17
        Case 18 '100
            sTextoMsj = sTextoMsj2 & Chr(13) & "No es una Fecha Válida"
        Case 19
            sTextoMsj = "Este animal no Existe o no está Registrada"
        Case 20 '100
            sTextoMsj = "  ¡La Fecha es para el Futuro!"
        Case 21
        Case 22
        Case 23 '200
            sTextoMsj = "El Arete " & sArete & " ya sido utilizado"
        Case 24 'Aviso Informativo
            bTipoMensaje = True
            bFlagError = False
            sTextoMsj = "Este animal no está gestante."
        Case 98 '300
            bFlagError = False
            sTextoMsj = "No está configurado este Evento."
        Case 99
            bFlagError = False
            sTextoMsj = "Código en Construcción"
        Case 100 '100
            sTextoMsj = "Falta Ingresar Fecha"
        Case 101 '200
            sTextoMsj = "Falta Ingresar Arete"
        Case 102 '300
            sTextoMsj = "Falta Ingresar Evento"
        Case 104 '400
            sTextoMsj = "Falta Ingresar " & Me.Label4.Caption
        Case 105 '500
            sTextoMsj = "Falta Ingresar " & Me.Label5.Caption
        Case 106 '600
            sTextoMsj = "Falta Ingresar " & Me.Label6.Caption
        Case 107 '700
            sTextoMsj = "Falta Ingresar " & Me.Label7.Caption
     End Select
     If bTipoMensaje Then
            sMensaje = MsgBox(sTextoMsj _
              & Chr(13) & " (m" & iAviso & ")", _
              vbInformation, sMsjTitulo)
        Else
            sMensaje = MsgBox(sTextoMsj, _
              vbCritical, sMsjTitulo)
    End If
    sTextoMsj2 = vbNullString
End Sub

Private Sub BajaReemplazos()
    ' Registro de Baja como recría
    Dim iRowBR As Long
    Set ws = Worksheets("BajaReemplazos")
    Desproteger
    iRowBR = TamañoTabla("Tabla5") + 2
    ws.Cells(iRowBR, 1) = _
      wsR.Cells(iAR, 1) 'Arete
    ws.Cells(iRowBR, 2) = _
      wsR.Cells(iAR, 3) 'Peso
    ws.Cells(iRowBR, 3) = _
      wsR.Cells(iAR, 4) 'Edad
    ws.Cells(iRowBR, 4) = _
      wsR.Cells(iAR, 5) 'F.Nacim
    ws.Cells(iRowBR, 5) = _
      wsR.Cells(iAR, 6) 'Serv
    ws.Cells(iRowBR, 6) = _
      wsR.Cells(iAR, 7) 'F.Serv
    ws.Cells(iRowBR, 7) = _
      wsR.Cells(iAR, 8) 'Toro
    ws.Cells(iRowBR, 8) = _
      wsR.Cells(iAR, 9) 'Tecnico
    ws.Cells(iRowBR, 9) = _
      wsR.Cells(iAR, 10) 'Status
    ws.Cells(iRowBR, 10) = _
      wsR.Cells(iAR, 12) 'clave1
    ws.Cells(iRowBR, 11) = _
      wsR.Cells(iAR, 13) 'clave2
    With ws.Cells(iRowBR, 12)
        .Value = CDate(Me.txtFecha)
        .NumberFormat = "d-mmm-yy"
    End With
    If Me.cmboEvento = "Parto" Then
            ws.Cells(iRowBR, 13) = "Parto" & " " & sPartoDet
      Else
            ws.Cells(iRowBR, 13) = Me.ComboBox4
    End If
    If Not CBool(Range("Desarrollador!B6")) Then _
      ws.Visible = xlSheetVeryHidden
End Sub

Private Sub BorrarSubformulario()
    With Me
        .ComboBox4.Clear
        .ComboBox4.Visible = False
        .Label4.Visible = False
        .TextBox4 = vbNullString
        .TextBox4.Visible = False
        .TextBox4.Enabled = True
        .ComboBox5.Visible = False
        .Label5.Visible = False
        .TextBox5 = vbNullString
        .TextBox5.Visible = False
    End With
    BorrarSubformulario1
End Sub

Private Sub BorrarSubformulario1()
    With Me
        .Label6.Visible = False
        .TextBox6 = vbNullString
        .TextBox6.Visible = False
        .Label7.Visible = False
        .TextBox7 = vbNullString
        .TextBox7.Visible = False
    End With
End Sub

Private Sub BorrarSubformulario2()
    With Me
        .ComboBox5 = vbNullString
        .ComboBox5.Visible = False
    End With
End Sub

Private Sub CalcularProxRevision()
' Calcula la prox. fecha de revisión en base a las
'   observaciones.
    Dim iDias As Long
    Dim sCadena As String
    Select Case Me.cmboEvento
        Case "Revisión"
            sCadena = UCase(Trim(Me.TextBox4))
        Case "Enfermedad"
            sCadena = UCase(Trim(Me.TextBox5))
    End Select
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
                .Value = CDate(Me.txtFecha) + iDias
                .NumberFormat = "d-mmm-yy"
            End With
        Else
            wsIV.Cells(iARowInfoVital, 16).Clear
     End If
End Sub

Private Sub ChecarFechaParto()
    ' Checar errores
    'ChecarMismoEvento
    If sLocAnimal = "R" Then Exit Sub
    If Not IsDate(Cells(iARH, 6)) _
     And Not IsEmpty(Cells(iARH, 6)) Then
        sTextoMsj2 = "La Fecha del Parto registrado"
        iAviso = 18
        Avisos
        Me.txtFecha.SetFocus
        Exit Sub
    End If
End Sub

Private Sub ChecarFechaServicio()
    ' Checar errores
    Dim sTxt3 As String
    'ChecarMismoEvento
    Select Case sLocAnimal
        Case Is = "H"
            If Not IsDate(Cells(iARH, 8)) _
            And Not IsEmpty(Cells(iARH, 8)) Then
                If Cells(iARH, 9) = "Calor" Then _
                  sTxt3 = "Calor" Else sTxt3 = "Servicio"
                GoTo 1234
            End If
        Case Is = "R"
            If Not IsDate(wsR.Cells(iAR, 7)) And _
              Not IsEmpty(wsR.Cells(iAR, 7)) Then
                If wsR.Cells(iAR, 8) = "Calor" Then _
                  sTxt3 = "Calor" Else sTxt3 = "Servicio"
                 GoTo 1234
            End If
    End Select
    Exit Sub

1234:
    sTextoMsj2 = "La Fecha del " & sTxt3 & " registrado"
    iAviso = 18
    Avisos
    Me.txtFecha.SetFocus
    Exit Sub

End Sub

Private Sub ChecarIntegridad()
    ' Evitar vacíos de información
    If Me.txtFecha = vbNullString Then
        iAviso = 100
        Avisos
        Me.txtFecha.SetFocus
    End If
    If bFlagError Then Exit Sub
    On Error GoTo 6785
    If Year(CDate(Me.txtFecha)) = 1931 Or _
      Year(CDate(Me.txtFecha)) = 1930 Or _
      Year(CDate(Me.txtFecha)) = 1929 Or _
      Year(CDate(Me.txtFecha)) <= 2010 Then
        GoTo 6785
    End If
    If Not IsDate(CDate(Me.txtFecha)) Then ' Fecha inválida
        GoTo 6785
    End If
    If CDate(Me.txtFecha) > Date Then ' Fecha Posterior al Sistema
        iAviso = 20
        Avisos
        Me.txtFecha.SetFocus
    End If
    Select Case sLocAnimal
        Case Is = "H"
            'Fecha anterior al parto
            If CDate(Me.txtFecha) < CDate(Cells(iARH, 6)) Then
                sTextoMsj2 = "Parto"
                iAviso = 1
                Avisos
                Me.txtFecha.SetFocus
            End If
        Case Is = "R"
            'Fecha anterior al nacimiento
            If CDate(Me.txtFecha) < _
              CDate(wsR.Cells(iAR, 5)) Then
                sTextoMsj2 = "Nacimiento"
                iAviso = 1
                Avisos
                Me.txtFecha.SetFocus
        End If
    End Select
    On Error GoTo 0
    If bFlagError Then Exit Sub
    If Me.cmboIdArete = vbNullString Then
        iAviso = 101
        Avisos
        Me.cmboIdArete.SetFocus
    End If
    If bFlagError Then Exit Sub
    If Me.cmboEvento = vbNullString Then
        iAviso = 102
        Avisos
        Me.cmboEvento.SetFocus
    End If
    If bFlagError Then Exit Sub
    Select Case cmboEvento
        Case "Servicio"
            If CBool(Range("Configuracion!C15")) = _
              True Then QTextBox4   'Semental
            If bFlagError Then Exit Sub
            If CBool(Range("Configuracion!C16")) = _
              True Then QTextBox5   'Tecnico
            If bFlagError Then Exit Sub
            nMM = 28
            Factorizacion
            If bFlagError Then Exit Sub
        Case "Calor"
            If CBool(Range("Configuracion!C16")) = _
              True Then QTextBox5   'Tecnico
            If bFlagError Then Exit Sub
            nMM = 16
            Factorizacion
        Case "Revisión"
            ChecarMismoEvento
            QTextBox4   'Observaciones
            If bFlagError Then Exit Sub
            QTextBox5   'Responsable
        Case "Dx Gest."
            ChecarMismoEvento
            QComboBox4  'Resultado
            If bFlagError Then Exit Sub
            QTextBox5   'Responsable
            nMM = 24
            Factorizacion
        Case "Secar"
            ChecarMismoEvento
            nMM = 4
            Factorizacion
        Case "Producción"
            ChecarMismoEvento
            QTextBox5
            If bFlagError Then Exit Sub
            If Not IsNumeric(Me.TextBox5) Then
                iAviso = 15
                Avisos
            End If
            nMM = 4
            Factorizacion
        Case "Movimiento"
            ChecarMismoEvento
            QTextBox5
        Case "Enfermedad"
            QComboBox4  'Dx Enfermedad
            If bFlagError Then Exit Sub
            QTextBox5   'Tratamiento
            If bFlagError Then Exit Sub
            QTextBox6   'Responsable
        Case "Otro"
            QComboBox4
            
        Case "Parto"
            IntegridadParto
            nMM = 6
            Factorizacion
        Case "Baja"
            QComboBox4
            nMM = 6
            Factorizacion
    End Select
    Exit Sub

6785:
    iAviso = 18
    sTextoMsj2 = "La fecha del Evento"
    Avisos
End Sub

Private Sub ChecarMismoEvento()
    If BuscarEvento(Me.cmboIdArete, sEvento, Me.txtFecha) > 0 Then
        iAviso = 3
        Avisos
        Me.txtFecha.SetFocus
        Exit Sub
    End If
End Sub

Private Sub ChecarNumParto()
    ' Checar errores
    If sLocAnimal = "R" Then Exit Sub
    If Not IsNumeric(Cells(iARH, 5)) Then
        sTextoMsj2 = "El Número del Parto registrado"
        iAviso = 15
        Avisos
        Exit Sub
    End If
End Sub

Private Sub ChecarNumServicio()
    ' Checar errores
    If sLocAnimal = "H" Then _
      If Not IsNumeric(Cells(iARH, 7)) Then _
      GoTo 1234
    If sLocAnimal = "R" Then _
      If Not IsNumeric(wsR.Cells(iAR, 6)) Then _
      GoTo 1234
    Exit Sub
    
1234:
    sTextoMsj2 = "El Número del Servicio registrado"
    iAviso = 15
    Avisos
    Exit Sub

End Sub

Private Sub ChecarSexo()
    ' Checar que el reemplazo sea una hembra
    If Not wsR.Cells(iAR, 14) = "H" Then
        iAviso = 4
        Avisos
    End If
End Sub

Private Sub ChecarUnicidad()
    ' Checar que no se repita el numero del animal
    Dim sTabla As String
    Dim i As Long
    For i = 1 To 3
       If i = 3 Then i = 8
       sTabla = "Tabla" & i & "[Arete]"
       If Application.WorksheetFunction. _
         CountIf(Range(sTabla), sArete) >= 1 _
         Then
            iAviso = 23
            Avisos
        End If
    Next i
End Sub

'Private Sub cmboEvento_AfterUpdate()
Private Sub cmboEvento_Change()
    BorrarSubformulario
    Select Case cmboEvento
        Case vbNullString
            ' Para ahorrar tiempo de proceso
            Exit Sub
        Case "Servicio"
            sEvento = "Serv"
            With Me
                If CBool(Range("Configuracion!C15")) Then _
                  .Label4.Caption = "Toro*" Else _
                  .Label4.Caption = "Toro"
                .TextBox4.Visible = True
                If CBool(Range("Configuracion!C16")) Then _
                  .Label5.Caption = "Técnico*" Else _
                  .Label5.Caption = "Técnico"
                .Label5.Visible = True
                .TextBox5.Visible = True
                .TextBox4.SetFocus
            End With
        Case "Calor"
            sEvento = "Calor"
            With Me
                .Label4.Caption = "Toro"
                .TextBox4.Visible = True
                .TextBox4 = "Calor S/Serv."
                .TextBox4.Enabled = False
                If CBool(Range("Configuracion!C16")) Then _
                  .Label5.Caption = "Técnico*" Else _
                  .Label5.Caption = "Técnico"
                .Label5.Visible = True
                .TextBox5.Visible = True
                .TextBox5.SetFocus
            End With
        Case "Revisión"
            sEvento = "Rev"
            With Me
                .Label4.Caption = "Observaciones*"
                .TextBox4.Visible = True
                .TextBox4.SetFocus
                .Label5.Visible = True
                .Label5.Caption = "Responsable*"
                .TextBox5.Visible = True
            End With
        Case "Dx Gest."
            ' Mostrar días último servicio
            sEvento = "DxGst"
            With Me
                .Label4.Caption = "Resultado*"
                With .ComboBox4
                    .AddItem "Gestante"
                    .AddItem "Vacía"
                    .Visible = True
                    .SetFocus
                End With
                .Label5.Visible = True
                .Label5.Caption = "Responsable*"
                .TextBox5.Visible = True
            End With
        Case "Secar"
            sEvento = "Seca"
            ' Mostrar fecha para secar
            ActualizarTextBox4
            With Me
                .Label4.Caption = "Prox. Parto"
                .TextBox4.Enabled = False
                .TextBox4.Visible = True
            End With
            ' Mostrar produccion
            ' Indicar a que corral se mueve
        Case "Producción"
            sEvento = "Prod"
            ActualizarTextBox4
            With Me
                .Label4.Caption = "Prod. Ant."
                .TextBox4.Enabled = False 'Prod Anterior
                .TextBox4.Visible = True
                .Label5.Caption = "Producción*"
                .Label5.Visible = True
                .TextBox5.Enabled = True
                .TextBox5.Visible = True
                .TextBox5.SetFocus
                .Label6.Caption = "Corral"
                .Label6.Visible = True
                .TextBox6.Enabled = True
                .TextBox6.Visible = True
            End With
        Case "Movimiento"
            sEvento = "Mov"
            ActualizarTextBox4
            With Me
                .Label4.Caption = "Corral. Ant."
                .TextBox4.Enabled = False 'Corral Actual
                .TextBox4.Visible = True
                .Label5.Caption = "Corral*"
                .Label5.Visible = True
                .TextBox5.Enabled = True
                .TextBox5.Visible = True
                .TextBox5.SetFocus
            End With
        Case "Enfermedad"
            sEvento = "Enf-"
            With Me
                .Label4.Caption = "Dx.*"
                 With .ComboBox4
                    .AddItem "Mastitis"         'MA
                    .AddItem "Ret. Placenta"    'RP
                    .AddItem "Metritis"         'UM
                    .AddItem "Despl. Abomaso"   'DA
                    .AddItem "Gabarro"          'GA
                    .AddItem "Neumonía"         'NE
                    .AddItem "Diarrea"          'DI
                    .AddItem "Herida"           'HE
                    .AddItem "Otra"             'OT
                    .Visible = True
                    .SetFocus
                End With
                .Label5.Caption = "Tratamiento*"
                .Label5.Visible = True
                .TextBox5.Enabled = True
                .TextBox5.Visible = True
                .Label6.Visible = True
                .Label6.Caption = "Responsable*"
                .TextBox6.Visible = True
            End With
        Case "Imantación"
            sEvento = "Iman"
            With Me
                .Label4.Caption = "Responsable*"
                .TextBox4.Visible = True
                .TextBox4.SetFocus
            End With
        Case "Parto"
            sEvento = "Parto"
            With Me
                .Label4.Caption = "Tipo Parto*"
                With .ComboBox4
                    .AddItem "Natural"
                    .AddItem "Asistido"
                    .AddItem "Distocia"
                    .AddItem "Aborto"
                    .AddItem "Inducido"
                    .Visible = True
                    .SetFocus
                End With
            End With
        Case "Destete"
            ' Mostrar datos del destete
            With Me
                .Label4.Caption = "Corral"
                .Label5.Caption = "Pesaje"
                .Label5.Visible = True
                .TextBox4.Enabled = True
                .TextBox4.Visible = True
                .TextBox4.SetFocus
                .TextBox5.Enabled = True
                .TextBox5.Visible = True
            End With
        Case "Pesaje"
            sEvento = "Pesaje"
            ActualizarTextBox4
            With Me
                .Label4.Caption = "Pesaje. Ant."
                .TextBox4.Enabled = False 'Peso Anterior
                .TextBox4.Visible = True
                .Label5.Caption = "Peso*"
                .Label5.Visible = True
                .TextBox5.Enabled = True
                .TextBox5.Visible = True
                .TextBox5.SetFocus
            End With
        Case "Nota"
            sEvento = "Nota"
            With Me
                .Label4.Caption = "Observaciones"
                .TextBox4.Visible = True
                .Label5.Caption = "Responsable"
                .Label5.Visible = True
                .TextBox5.Visible = True
                .TextBox4.SetFocus
            End With
        Case "Otro"
            Me.Label4.Caption = "Descripcion"
            With Me.ComboBox4
                .AddItem "Vacunación"
                .AddItem "DNB"
                .AddItem "TB+"
                .AddItem "K Caseína"
                .AddItem "Id Adicional"
                .Visible = True
                .SetFocus
            End With
        Case "Baja"
            sEvento = "Baja"
            With Me
                .Label4.Caption = "Causa*"
                With .ComboBox4
                    .Clear
                    .AddItem "Producción"
                    .AddItem "Reproducción"
                    .AddItem "Mastitis"
                    .AddItem "Gabarro"
                    .AddItem "Lesiones"
                    .AddItem "Neumonía"
                    .AddItem "Diarrea"
                    .AddItem "Otra"
                    .Visible = True
                    .SetFocus
                End With
            End With
        Case "Alta"
            sEvento = "Alta"
            With Me
                If CBool(Range("Configuracion!C22")) Then _
                  .Label4.Caption = "F.Nacim*" Else _
                  .Label4.Caption = "F.Nacim"
                .Label4.Visible = True
                If CBool(Range("Configuracion!C19")) Then _
                  .Label5.Caption = "Padre*" Else _
                  .Label5.Caption = "Padre"
                .Label5.Visible = True
                If CBool(Range("Configuracion!C20")) Then _
                  .Label6.Caption = "Madre*" Else _
                  .Label6.Caption = "Madre"
                .Label6.Visible = True
                If CBool(Range("Configuracion!C21")) Then _
                  .Label7.Caption = "Raza*" Else _
                  .Label7.Caption = "Raza"
                .Label7.Visible = True
                .TextBox4.Visible = True
                .TextBox5.Visible = True
                .TextBox6.Visible = True
                .TextBox7.Visible = True
            End With
        Case Else
            ' No mostrar nada
            Exit Sub
    End Select
    Me.Label4.Visible = True
End Sub

Private Sub cmboIdArete_AfterUpdate()
 ' Localiza la clave y toma ciertos valores
    bAreteNoEncontrado = False
    sLocPrevia = sLocAnimal
    sLocAnimal = vbNullString
    If Me.cmboIdArete = vbNullString Then GoTo 4231
    If IsNumeric(Me.cmboIdArete) Then sArete = _
      CDbl(Me.cmboIdArete) Else sArete = Me.cmboIdArete
    ' Buscar y posicionarse
    iARH = IndiceTabla(Me.cmboIdArete, "Tabla1")
    If iARH > 0 Then
            sLocAnimal = "H"
            Cells(iARH, 1).Activate
            Set ws = Worksheets("Hato2")
            Desproteger
1234:
            iARH2 = _
              IndiceTabla(Me.cmboIdArete, "Tabla15")
            If iARH2 = 0 Then GoTo 3412 'Sí no existe
        Else
            iAR = IndiceTabla(Me.cmboIdArete, "Tabla2")
            If iAR = 0 Then GoTo ControlDeErrores 'Sí no existe
            sLocAnimal = "R"
    End If
    Set ws = Worksheets("InfoVitalicia")
    Desproteger
2341:
    iARowInfoVital = IndiceTabla(Me.cmboIdArete, "Tabla8")
    If iARowInfoVital = 0 Then GoTo 4123 'Sí no existe
    Poblar_cmboEvento
    If Not Me.TextBox4 = vbNullString Then ActualizarTextBox4
    Exit Sub

3412:
    AltaHoja2
    GoTo 1234
4123:
    AltaInfovital
    GoTo 2341

ControlDeErrores:
    bAreteNoEncontrado = True
    Range("A2").Select
    iAviso = 19
    Avisos
4231:
    sMetadato = vbNullString
    Poblar_cmboEvento
End Sub

Private Sub cmndAceptar_Click()
' Aceptar
    If bAreteNoEncontrado = True And _
      Not Me.cmboEvento = "Alta" Then
        iAviso = 19
        Avisos
        Me.cmboIdArete.SetFocus
        GoTo CtrlErrs
    End If
    ChecarIntegridad
    If bFlagError Then GoTo CtrlErrs
        If Worksheets("Eventos").Visible = False Then _
          Worksheets("Eventos").Visible = True
        Select Case Me.cmboEvento
            Case "Servicio"
                ProcesarServicios
                If bFlagError Then GoTo CtrlErrs
            Case "Calor"
                ProcesarCalores
                If bFlagError Then GoTo CtrlErrs
            Case "Revisión"
                ProcesarRevision
                If bFlagError Then GoTo CtrlErrs
            Case "Dx Gest."
                ProcesarDxGest
                If bFlagError Then GoTo CtrlErrs
            Case "Secar"
                ProcesarSecar
                If bFlagError Then GoTo CtrlErrs
            Case "Producción"
                ProcesarProduccion
                If bFlagError Then GoTo CtrlErrs
            Case "Movimiento"
                ProcesarMovimientos
                If bFlagError Then GoTo CtrlErrs
            Case "Enfermedad"
                ProcesarEnfermedad
                If bFlagError Then GoTo CtrlErrs
            Case "Parto"
                ProcesarParto
                If bFlagError Then GoTo CtrlErrs
            Case "Imantación"
                ProcesarImantacion
                If bFlagError Then GoTo CtrlErrs
            Case "Nota"
                ProcesarNota
                If bFlagError Then GoTo CtrlErrs
            Case "Otro"
                ProcesarOtro
                If bFlagError Then GoTo CtrlErrs
            Case "Destete"
                ProcesarDestete
                If bFlagError Then GoTo CtrlErrs
            Case "Pesaje"
                ProcesarControlPeso
                If bFlagError Then GoTo CtrlErrs
            Case "Vacuna Brucela"
                ProcesarVacunacion
                If bFlagError Then GoTo CtrlErrs
            Case "Baja"
                ProcesarBaja
                If bFlagError Then GoTo CtrlErrs
            Case "Alta"
                ProcesarAlta
                If bFlagError Then GoTo CtrlErrs
            Case Else
                iAviso = 98
                Avisos
        End Select
        sMetadato = vbNullString
        Worksheets("Hato").Select
        Worksheets("Eventos").Visible = xlSheetVeryHidden
        With Me
            .Label11.Caption = _
            .txtFecha & "|" & _
            .cmboIdArete & "|" & _
            .cmboEvento
            .Label9.Visible = False
            .cmboIdArete = vbNullString
            .cmboEvento = vbNullString
            '+++++++++++++++++++++++
            .ComboBox4 = vbNullString
            .ComboBox5 = vbNullString
            .TextBox4 = vbNullString
            .TextBox5 = vbNullString
            .TextBox6 = vbNullString
            .TextBox7 = vbNullString
            '+++++++++++++++++++++++
            .cmboIdArete.SetFocus
            .Label4.Visible = False
        End With
CtrlErrs:
    bAceptar = True
    bFlagError = False
End Sub

Private Sub ComboBox4_Change()
    Select Case Me.ComboBox4
        Case vbNullString
        ' Para ahorrar tiempo de proceso
        Case "Natural", "Asistido"
            With Me
                .Label5.Caption = "Sexo Cría"
                .Label5.Visible = True
                With .ComboBox5
                    .RowSource = "Tabla11"
                    .Visible = True
                End With
            End With
        Case "Distocia", "Aborto"
            With Me
                .Label5.Visible = False
                .ComboBox5.Visible = False
            End With
            BorrarSubformulario1
        Case "Mastitis"         'MA
            sEvento = "Enf-" & "Ma"
        Case "Ret. Placenta"    'RP
            sEvento = "Enf-" & "RP"
        Case "Piometra"         'PI
            sEvento = "Enf-" & "Pi"
        Case "Despl. Abomaso"   'DA
            sEvento = "Enf-" & "DA"
        Case "Gabarro"          'GA
            sEvento = "Enf-" & "Ga"
        Case "Neumonía"         'NE
            sEvento = "Enf-" & "Ne"
        Case "Diarrea"          'DI
            sEnf = "Enf-" & "Di"
        Case "Herida"           'HE
            sEvento = "Enf-" & "He"
        Case "Otra"             'OT
            sEvento = "Enf-" & "Ot"
        Case "Vacunación"
            With Me
                .Label5.Caption = "Tipo Vacuna"
                .Label5.Visible = True
                With .ComboBox5
                    '.Clear
                    .RowSource = "Tabla10"
                    .Visible = True
                    .SetFocus
                End With
                .Label6.Caption = "Responsable"
                .Label6.Visible = True
                .TextBox6.Visible = True
            End With
        Case "DNB", "TB+"
            BorrarSubformulario2
            With Me
                .Label5.Caption = "Motivo"
                .Label5.Visible = True
                .TextBox5.Visible = True
                .TextBox5.SetFocus
                .Label6.Caption = "Responsable"
                .Label6.Visible = True
                .TextBox6.Visible = True
            End With
        Case "K Caseína"
            BorrarSubformulario2
            With Me
                .Label5.Caption = "Observaciones"
                .Label5.Visible = True
                .TextBox5.Visible = True
                .TextBox5.SetFocus
                .Label6.Caption = "Responsable"
                .Label6.Visible = True
                .TextBox6.Visible = True
            End With
        Case "Id Adicional"
            BorrarSubformulario2
            With Me
                .Label5 = "Id Adicional"
                .Label5.Visible = True
                .TextBox5.Visible = True
                .TextBox5.SetFocus
                .Label6.Caption = "Responsable"
                .Label6.Visible = True
                .TextBox6.Visible = True
            End With
    End Select
End Sub

Private Sub ComboBox5_Change()
    If Not Me.ComboBox5 = vbNullString Then
        BorrarSubformulario1
        Select Case Me.ComboBox5
            Case "Macho", "Hembra"
                HabilitarRegistro6
            Case "T Hembras", "T Machos", "FreeMartin"
                HabilitarRegistro6
                HabilitarRegistro7
            Case Else
                With Me
                    .Label6.Caption = "Responsable"
                    .Label6.Visible = True
                    .TextBox6.Visible = True
                End With
        End Select
    End If
End Sub

Private Sub CommandButton2_Click()
' Borrar
    bFlagError = False
    bAreteNoEncontrado = False
    Me.txtFecha = vbNullString
    Me.cmboIdArete = vbNullString
    Me.cmboEvento = vbNullString
    BorrarSubformulario
    Me.txtFecha = Format(Date, "dd-mmm-yy")
    Me.txtFecha.SetFocus
End Sub

Private Sub CommandButton3_Click()
' Cerrar
    Application.StatusBar = "Grabando información..."
    ActiveWorkbook.Save
    Application.StatusBar = "Ordenando información..."
    Application.Run "OrdenarHojas" 'Módulo2
    Application.Run "MostrarHojas" 'ModSeguridad
    Set ws = Worksheets("Reemplazos")
    Desproteger
    Range("Desarrollador!B20").Clear
    Set ws = Worksheets("Hato")
    'Desproteger
    Range("Desarrollador!B20").Clear
    If bAceptar = True Then
        Application.Run "ADEL"
        Application.Run "FCH"
    End If
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Unload Me
End Sub

Private Sub CommandButton4_Click()
    usrKardex.Show
End Sub

Private Sub ConsecutivoEventos()
    ' Agregar Información al Consecutivo
    Dim iRenglon As Long
    Dim iBus As Long
    Dim vArete, vFecha
    Set ws = Worksheets("Eventos")
    Desproteger
    iRenglon = TamañoTabla("Tabla6") + 2
    vArete = Me.cmboIdArete
    If IsNumeric(Me.cmboIdArete) Then _
      ws.Cells(iRenglon, 1) = CDbl(Me.cmboIdArete) Else _
      ws.Cells(iRenglon, 1) = (Me.cmboIdArete)
    vFecha = Me.txtFecha
    With ws.Cells(iRenglon, 2)
        .Value = CDate(Me.txtFecha)
        .NumberFormat = "d-mmm-yy"
    End With
    Select Case Me.cmboEvento
        Case "Servicio"
            sEvento = "Serv"
            If Me.TextBox4 = vbNullString Then
                    sObserv = "N.D."
                Else
                    sObserv = UCase(Me.TextBox4)
            End If
            If Me.TextBox5 = vbNullString Then
                    sResp = "N.D."
                Else
                    sResp = UCase(Me.TextBox5)
            End If
        Case "Calor"
            sEvento = "Calor"
            sObserv = "Calor"
            sResp = UCase(Me.TextBox5)
        Case "Producción"
            sEvento = "Prod"
            sObserv = Format(CDbl(Me.TextBox5), "0.0")
        Case "Pesaje"
            sEvento = "Pesaje"
            If IsEmpty(Me.TextBox6) Then
                    sObserv = Format(CDbl(Me.TextBox5), "0.0")
                Else
                    sObserv = Format(CDbl(Me.TextBox5), "0.0") _
                     & "-> " & Me.TextBox6
            End If
        Case "Movimiento"
            sEvento = "Mov"
            sObserv = CDbl(Me.TextBox5)
        Case "Enfermedad"
            sObserv = UCase(Trim(Me.TextBox5))
            sResp = UCase(Trim(Me.TextBox6))
        Case "Revisión"
            sEvento = "Rev"
            sObserv = Trim(UCase(Me.TextBox4))
            sResp = Trim(UCase(Me.TextBox5))
        Case "Dx Gest."
            sEvento = "DxGst"
            If Me.ComboBox4 = "Gestante" Then
                    sObserv = "Gest"
                    iBus = BUS(Me.cmboIdArete, "Serv")
                    ws.Cells(iBus, 9) = ws.Cells(iBus, 9) & "-P"
                Else
                    sObserv = "Vacía"
                    iBus = BUS(Me.cmboIdArete, "Serv")
                    ws.Cells(iBus, 9) = ws.Cells(iBus, 9) & "-O"
            End If
            sResp = Trim(UCase(Me.TextBox5))
        Case "Secar"
            sEvento = "Seca"
        Case "Parto"
            sEvento = "Parto"
            If Me.ComboBox4 = "Aborto" Then
                    sEvento = "Aborto"
                Else
                    sEvento = "Parto"
                    sObserv = UCase(sPartoDet)
            End If
        Case "Imantación"
            sEvento = "Iman"
            sResp = UCase(Me.TextBox4)
        Case "Nota"
            sEvento = "Nota"
            sObserv = UCase(Trim(Me.TextBox4))
            sResp = UCase(Trim(Me.TextBox5))
        Case "Otro"
            Select Case Me.ComboBox4
                Case "Vacunación"
                    sObserv = Me.ComboBox5
                    sResp = Me.TextBox6
                Case "DNB"
                    sEvento = "DNB"
                    sObserv = Me.TextBox5 'Me.ComboBox5
                    sResp = Me.TextBox6
                Case "TB+"
                    sEvento = "TB+"
                    sObserv = Me.TextBox5
                    sResp = Me.TextBox6
                Case "K Caseína"
                    sEvento = "KCaseína"
                    sObserv = Me.TextBox5
                    sResp = Me.TextBox6
                Case "Id Adicional"
                    sEvento = "Id+"
                    sObserv = Me.TextBox5
                    sResp = Me.TextBox6
            End Select
            sObserv = UCase(Trim(Me.TextBox5))
            sResp = UCase(Trim(Me.TextBox6))
        Case "Baja"
            sEvento = "Baja"
            sObserv = Me.ComboBox4
        Case "Alta"
            sEvento = "Alta"
    End Select
    With ws
        .Cells(iRenglon, 3) = sEvento
        .Cells(iRenglon, 4) = sObserv
        .Cells(iRenglon, 5) = sResp
        .Cells(iRenglon, 6) = Range("Configuracion!C49")
        .Cells(iRenglon, 7) = Format(Date, "d-mmm-yy")
        .Cells(iRenglon, 8) = Format(Time, "hh:mm")
        .Cells(iRenglon, 9) = sMetadato
        Select Case Me.cmboEvento
            Case "Secar"
                ' No agregar si mismo corral
                If bMovCorral Then _
                  .Cells(iRenglon + 1, 4) = Range("Configuracion!C9")
            Case "Parto"
                ' No agregar si mismo corral
                If wsH.Cells(iARH, 5) > 1 And bMovCorral Then _
                  .Cells(iRenglon + 1, 4) = Range("Configuracion!C11")
                If wsH.Cells(iARH, 5) = 1 And bMovCorral Then _
                  .Cells(iRenglon + 1, 4) = Range("Configuracion!C12")
            Case "Producción"
                ' No agregar si mismo corral
                'If Me.cmboEvento = "Producción" And _
                  (Not Me.TextBox6 = vbNullString Or _
                   Val(Me.TextBox6) = Val(wsH.Cells(iARH, 2))) Then _
                   .Cells(iRenglon + 1, 4) = Val(Me.TextBox6)
                If bMovCorral = True Then _
                  .Cells(iRenglon + 1, 4) = Val(Me.TextBox6)
        End Select
        If bMovCorral = True Then
            .Cells(iRenglon + 1, 1) = vArete
            .Cells(iRenglon + 1, 2) = CDate(vFecha)
            .Cells(iRenglon + 1, 3) = "Mov"
            .Cells(iRenglon + 1, 5) = sResp
            .Cells(iRenglon + 1, 6) = Range("Configuracion!C49")
            .Cells(iRenglon + 1, 7) = Format(Date, "d-mmm-yy")
            .Cells(iRenglon + 1, 8) = Format(Time, "hh:mm")
        End If
    End With
    bMovCorral = False
    If CBool(Range("Configuracion!C25")) Then _
      LoginRecord
    If Not CBool(Range("Desarrollador!B6")) Then _
      ws.Visible = xlSheetVeryHidden
End Sub

Private Sub Desproteger()
' Desprotege la hoja activa
    If Not ws.Visible Then ws.Visible = True
    Application.Run "Desproteger" 'Módulo2
End Sub

Private Sub Desproteger1()
' Desprotege la hoja activa
    If ws.ProtectContents Then _
      ws.Unprotect Password:="0246813579"
End Sub

Sub Factorizacion()
    Dim n As Long
    n = nMM
    If n >= 2 ^ 4 Then '16
        ChecarFechaServicio
        If bFlagError Then Exit Sub
        If n - 2 ^ 4 >= 0 Then n = n - 2 ^ 4
    End If
    If n >= 2 ^ 3 Then '8
        ChecarNumServicio
        If bFlagError Then Exit Sub
        If n - 2 ^ 3 >= 0 Then n = n - 2 ^ 3
    End If
    If n >= 2 ^ 2 Then '4
        ChecarFechaParto
        If bFlagError Then Exit Sub
        If n - 2 ^ 2 >= 0 Then n = n - 2 ^ 2
    End If
    If n >= 2 Then
        ChecarNumParto
        If bFlagError Then Exit Sub
        If n - 2 ^ 1 >= 0 Then n = n - 2 ^ 1
    End If
End Sub

Private Sub IntegridadParto()
    QComboBox4  'Tipo Parto
    If Me.ComboBox4 = "Natural" Or Me.ComboBox4 = "Asistido" Then
        If CBool(Range("Configuracion!C30")) Then
            ' Parto sencillo
            If Me.ComboBox5 = "Macho" Or Me.ComboBox5 = _
              "Hembra" Then
                    QTextBox6   'Cría 1
                    sArete = Me.TextBox6
                    ChecarUnicidad
                    If bFlagError Then Exit Sub
                Else
                    ' Parto Gemelar
                    QTextBox6   'Cría 1
                    If bFlagError Then Exit Sub
                    sArete = Me.TextBox6
                    ChecarUnicidad
                    If bFlagError Then Exit Sub
                    QTextBox7   'Cría 2
                    sArete = Me.TextBox7
                    ChecarUnicidad
                    If bFlagError Then Exit Sub
            End If
        End If
    End If
End Sub

Private Sub HabilitarRegistro6()
    With Me
        .Label6.Caption = "Id. Cría 1"
        .Label6.Visible = True
        .TextBox6.Visible = True
        .TextBox6.Enabled = True
        .TextBox6.SetFocus
    End With
    If CBool(Range("Configuracion!C30")) Then _
      Me.TextBox6 = _
      Application.WorksheetFunction.Max(Range("Tabla2[[Arete]]")) + 1
End Sub

Private Sub HabilitarRegistro7()
    With Me
        .Label7.Caption = "Id. Cría 2"
        .Label7.Visible = True
        .TextBox7.Visible = True
        .TextBox7.Enabled = True
    End With
    If CBool(Range("Configuracion!C30")) Then _
      Me.TextBox7 = _
      Application.WorksheetFunction.Max(Range("Tabla2[[Arete]]")) + 2
End Sub

Private Sub LactAnteriores()
    Dim iRenglon As Long
    Set ws = Worksheets("LactanciasAnteriores")
    Desproteger
    iRenglon = TamañoTabla("Tabla4") + 2
    With ws
        .Cells(iRenglon, 1) = _
          CDbl(Cells(iARH, 1)) 'Arete
        .Cells(iRenglon, 2) = _
          Cells(iARH, 5) 'Parto
        With .Cells(iRenglon, 3) 'F.Parto
            .Value = CDate((Cells(iARH, 6)))
            .NumberFormat = "d-mmm-yy"
        End With
        .Cells(iRenglon, 4) = _
          Cells(iARH, 7) 'Servicio
        With .Cells(iRenglon, 5) 'F.Servicio
            .Value = CDate((Cells(iARH, 8)))
            .NumberFormat = "d-mmm-yy"
        End With
        .Cells(iRenglon, 6) = _
          Cells(iARH, 9) 'Toro
        .Cells(iRenglon, 7) = _
          Cells(iARH, 10) 'Tecnico
        .Cells(iRenglon, 8) = _
          Cells(iARH, 11) 'Status
        .Cells(iRenglon, 9) = _
          Cells(iARH, 14) 'Clave1
        .Cells(iRenglon, 10) = _
          Cells(iARH, 15) 'Clave2
        If IsEmpty(Sheets("Hato2").Cells(iARH2, 16)) Then
                ' Si la vaca no se secó
                .Cells(iRenglon, 11) = CDate(Me.txtFecha) - _
                  CDate(Cells(iARH, 6)) 'DiasLact
                .Cells(iRenglon, 12) = 0 'DiasSeca
            Else
                ' Si la vaca seca
                .Cells(iRenglon, 11) = _
                  CDate(Sheets("Hato2").Cells(iARH2, 16)) - _
                  CDate(Cells(iARH, 6)) 'DiasLact
                .Cells(iRenglon, 12) = CDate(Me.txtFecha) - _
                  CDate(Sheets("Hato2").Cells(iARH2, 16)) 'DiasSeca
        End If
        .Cells(iRenglon, 13) = _
          Sheets("Hato2").Cells(iARH2, 14) 'ProdAcum
        .Cells(iRenglon, 14) = _
          Sheets("Hato2").Cells(iARH2, 15) 'Proy305d
        .Cells(iRenglon, 15) = _
          Sheets("Hato2").Cells(iARH2, 2) 'Dias1Serv
        .Cells(iRenglon, 16) = _
          Sheets("Hato2").Cells(iARH2, 3) 'DiasAbiertos
        With .Cells(iRenglon, 17) 'F.Terminación
            .Value = CDate(Me.txtFecha)
            .NumberFormat = "d-mmm-yy"
        End With
        .Cells(iRenglon, 18) = Me.cmboEvento 'TipoTerminación
        .Cells(iRenglon, 19) = Me.ComboBox4 'CausaTerminación
    End With
    If Not CBool(Range("Desarrollador!B6")) Then _
      ws.Visible = xlSheetVeryHidden
End Sub

Private Sub LoginRecord()
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
    Write #1, Range("Configuracion!D3"), Me.cmboIdArete, Me.txtFecha, _
      sEvento, sObserv, sResp, Range("Configuracion!C49"), Date, _
      Format(Time, "h:mm")
    Close #1
    Exit Sub
100
    Open "Log101.txt" For Append As #1
    GoTo 200
End Sub

Private Sub Poblar_cmboEvento()
    If sLocAnimal = sLocPrevia And _
     Not sLocAnimal = vbNullString Then Exit Sub
    With Me
        With .cmboEvento
            .Clear
            Select Case sLocAnimal
             Case Is = "H"
                .AddItem "Servicio"
                .AddItem "Calor"
                .AddItem "Producción"
                .AddItem "Movimiento"
                .AddItem "Enfermedad"
                .AddItem "Revisión"
                .AddItem "Dx Gest."
                .AddItem "Secar"
                .AddItem "Nota"
                .AddItem "Parto"
                .AddItem "Imantación"
                .AddItem "Otro"
                .AddItem "Baja"
             Case Is = "R"
                .AddItem "Servicio"
                .AddItem "Calor"
                .AddItem "Revisión"
                .AddItem "Dx Gest."
                .AddItem "Movimiento"
                .AddItem "Enfermedad"
                .AddItem "Pesaje"
                .AddItem "Parto"
                .AddItem "Destete"
                .AddItem "Imantación"
                .AddItem "Otro"
                .AddItem "Nota"
                .AddItem "Baja"
             Case Is = vbNullString
                .AddItem vbNullString
                .AddItem "Alta"
            End Select
        End With
        .cmboIdArete.SetFocus
    End With
End Sub

Private Sub ProcesarAlta()
' Alta de animales
    Dim iRenglon As Long
    ChecarUnicidad ' Comprobar que no esté repetido
    If bFlagError Then Exit Sub
    If CBool(Range("Configuracion!C22")) Then QTextBox4
    If bFlagError Then Exit Sub
    If CBool(Range("Configuracion!C19")) Then QTextBox5
    If bFlagError Then Exit Sub
    If CBool(Range("Configuracion!C20")) Then QTextBox6
    If bFlagError Then Exit Sub
    If CBool(Range("Configuracion!C21")) Then QTextBox7
    If bFlagError Then Exit Sub
    ' Agregar en Info Vitalicia
    AltaInfovital
    ' Agregar en Hato2
    AltaHoja2
    ' Agregar en hoja Hato
    iRenglon = TamañoTabla("Tabla1") + 2 'Encabezado más renglón
    Cells(iRenglon, 1) = _
      CDbl(Me.cmboIdArete)
    ConsecutivoEventos
End Sub

Private Sub ProcesarBaja()
    Set ws = Worksheets("InfoVitalicia")
    'Edad-Lactancia-DEL
    sMetadato = _
      Format(CDate(Me.txtFecha) - CDate(ws.Cells(iARowInfoVital, 3)), "0000")
    If sLocAnimal = "H" Then
            LactAnteriores
            sMetadato = sMetadato _
              & "-" & Format(wsH.Cells(iARH, 5), "00") & "-" & _
              Format(CDate(Me.txtFecha) - CDate(wsH.Cells(iARH, 6)), "000")
        Else
            sMetadato = sMetadato & "-00-000"
    End If
    ConsecutivoEventos
    ' Borrar renglón activo
    'Set ws = Worksheets("InfoVitalicia")
    Desproteger
    With ws.Cells(iARowInfoVital, 14) 'F.Baja
        .Value = CDate(Me.txtFecha)
        .NumberFormat = "d-mmm-yy"
    End With
    ws.Cells(iARowInfoVital, 15) = Me.ComboBox4 'Causa
    If Not CBool(Range("Desarrollador!B6")) Then _
      ws.Visible = xlSheetVeryHidden
    If sLocAnimal = "H" Then
        Set ws = Worksheets("Hato2")
        Desproteger
        ws.Rows(iARH2).Delete Shift:=xlShiftUp
        If Not CBool(Range("Desarrollador!B6")) Then _
          ws.Visible = xlSheetVeryHidden
        Set ws = Worksheets("Hato")
        ws.Select
        ws.Rows(iARH).Delete Shift:=xlShiftUp
    End If
    If sLocAnimal = "R" Then
        BajaReemplazos
        wsR.Rows(iAR).Delete Shift:=xlShiftUp
    End If
End Sub

Private Sub ProcesarCalores()
    Dim iBus As Long
    Set ws = Worksheets("Eventos")
    If sLocAnimal = "H" Then
          ' Último Servicio o calor
        If CDate(Me.txtFecha) <= CDate(Cells(iARH, 8)) Then
            iAviso = 2
            Avisos
            Exit Sub
        End If
        If CDate(txtFecha) - CDate(Cells(iARH, 8)) <= 18 And _
          Not IsEmpty(Cells(iARH, 8)) Then
            iAviso = 16
            Avisos
        End If
        If CDate(txtFecha) - CDate(Cells(iARH, 8)) >= 36 And _
          Not IsEmpty(Cells(iARH, 8)) Then
            iAviso = 14
            Avisos
        End If
        ' Escribir Datos
        ' 0-DíasÚltimoServicio-DEL
        If IsEmpty(Cells(iARH, 7)) Then _
          sMetadato = "00-" _
          Else _
          sMetadato = _
          Format(Val(Cells(iARH, 7)), "00") & "-"
        If IsEmpty(Cells(iARH, 8)) Then _
          sMetadato = _
          sMetadato & "000-" _
          Else _
          sMetadato = _
          sMetadato & Format(CDate(txtFecha) - _
          CDate(Cells(iARH, 8)), "000")
        ' Añadir DEL
        sMetadato = sMetadato & "-" & _
          Format(CDate(txtFecha) - CDate(Cells(iARH, 6)), "000")
'        If IsEmpty(Cells(iARH, 7)) Then sMetadato = "00-000" Else _
          sMetadato = "00-" & Format(CDate(txtFecha) - _
          CDate(Cells(iARH, 8)), "000")
'        sMetadato = sMetadato & "-" & _
          Format(CDate(txtFecha) - CDate(Cells(iARH, 6)), "000")
        If Cells(iARH, 8) = vbNullString Then
            With wsH2.Cells(iARH2, 17) 'd1Calor
                .Value = CDate(Me.txtFecha) - CDate(Cells(iARH, 6))
            End With
        End If
        If Cells(iARH, 11) = "P" Then
            iBus = BUS(Me.cmboIdArete, "Serv")
            ws.Cells(iBus, 9) = Left(ws.Cells(iBus, 9), 11) & "R"
            Cells(iARH, 12).Clear 'FxSecar
            Cells(iARH, 13).Clear 'FxParir
            Cells(iARH, 14) = "pAb" 'Clave1
            wsH2.Cells(iARH2, 3).Clear 'DAbiertos
            iAviso = 5
            Avisos
            ' Exit Sub
        End If
        With Cells(iARH, 8)    'F.Servicio
            .Value = CDate(Me.txtFecha)
            .NumberFormat = "d-mmm-yy"
        End With
        Cells(iARH, 9) = "Calor" 'Semental
        Cells(iARH, 10) = UCase(Me.TextBox5) 'Tecnico
        Cells(iARH, 11).Clear 'Status
    End If
    If sLocAnimal = "R" Then
        ' Checar errores
        ' Último Servicio o calor
        If CDate(Me.txtFecha) <= CDate(wsR.Cells(iAR, 7)) Then
            iAviso = 2
            Avisos
            Exit Sub
        End If
        If CDate(txtFecha) - CDate(wsR.Cells(iAR, 7)) <= 18 And _
          Not IsEmpty(wsR.Cells(iAR, 7)) Then
            iAviso = 16
            Avisos
        End If
        If CDate(txtFecha) - CDate(wsR.Cells(iAR, 7)) >= 36 And _
          Not IsEmpty(wsR.Cells(iAR, 7)) Then
            iAviso = 14
            Avisos
        End If
        ChecarSexo
        ' Escribir Datos
        ' 0-DíasÜltimoServicio-DEL
         If IsEmpty(wsR.Cells(iAR, 6)) Then _
          sMetadato = "00-" _
          Else _
          sMetadato = _
          Format(Val(wsR.Cells(iAR, 6)), "00") & "-"
        If IsEmpty(wsR.Cells(iAR, 7)) Then _
          sMetadato = _
          sMetadato & "000-" _
          Else _
          sMetadato = _
          sMetadato & Format(CDate(txtFecha) - _
          CDate(wsR.Cells(iAR, 7)), "000")
        ' Añadir DEL
        sMetadato = sMetadato & "-" & _
          Format(CDate(txtFecha) - CDate(wsR.Cells(iAR, 5)), "000")
'        If IsEmpty(wsR.Cells(iAR, 6)) Then sMetadato = "00-000" Else _
          sMetadato = "00-" & _
          Format(CDate(txtFecha) - CDate(wsR.Cells(iAR, 7)), "000")
'        sMetadato = sMetadato & "-" & _
          Format(CDate(txtFecha) - CDate(wsR.Cells(iAR, 5)), "000")
        If wsR.Cells(iAR, 10) = "P" Then
            iBus = BUS(Me.cmboIdArete, "Serv")
            ws.Cells(iBus, 9) = Left(ws.Cells(iBus, 9), 11) & "R"
            wsR.Cells(iAR, 11).Clear 'FxParir
            wsR.Cells(iAR, 12) = "pAb" 'Clave1
            wsIV.Cells(iARowInfoVital, 11).Clear 'Edad1Parto
            iAviso = 5
            Avisos
            ' Exit Sub
        End If
        With wsR.Cells(iAR, 7)  'F.Servicio
            .Value = CDate(txtFecha)
            .NumberFormat = "d-mmm-yy"
        End With
        wsR.Cells(iAR, 8) = "Calor"   'Semental
        wsR.Cells(iAR, 9) = UCase(Me.TextBox5)  'Tecnico
        wsR.Cells(iAR, 10).Clear  'Status
    End If
    ConsecutivoEventos
End Sub

Private Sub ProcesarControlPeso()
    If bFlagError Then Exit Sub
    If CBool(Range("Configuracion!C35")) Then QTextBox5
    If bFlagError Then Exit Sub
    If CBool(Range("Configuracion!C35")) = True And _
      IsNumeric(TextBox5) = False Then
        iAviso = 15
        Avisos
    End If
    If CDate(Me.txtFecha) <= CDate(BuscarUltimoEvento(Me.cmboIdArete, _
      "Pesaje")) Then
        iAviso = 2
        Avisos
        Exit Sub
    End If
    With wsR.Cells(iAR, 3)
        .Value = CDbl(Me.TextBox5)
        .NumberFormat = "0.0"
    End With
    ConsecutivoEventos
End Sub

Private Sub ProcesarDestete()
    ' Vaquilla ya destetada
    If wsR.Cells(iAR, 2) <> _
      Range("Configuracion!C13") Then
        iAviso = 9
        Avisos
        Exit Sub
    End If
    ' Mover al mismo corral
    If wsR.Cells(iAR, 2) = Val(Me.TextBox4) Then
        iAviso = 11
        Avisos
        Exit Sub
    End If
    ' Lactancia menor a edad para destetar
    If Date - CDate(wsR.Cells(iAR, 5)) <= _
      Range("Configuracion!C34") Then
        iAviso = 10
        Avisos
        Exit Sub
    End If
    wsR.Cells(iAR, 2) = Me.TextBox4
    With wsR.Cells(iAR, 3)
        .Value = CDbl(Me.TextBox5)
        .NumberFormat = "0.0"
    End With
    ConsecutivoEventos
End Sub

Private Sub ProcesarDxGest()
    ' Mínimo un servicio
    Dim sDx As String
    If Me.ComboBox4 = "Gestante" Then sDx = "P" Else sDx = "O"
    If sLocAnimal = "H" Then
        If Me.ComboBox4 = "Gestante" Then
                ' Sin servicios
                If Cells(iARH, 7) < 1 Then
                    iAviso = 6
                    Avisos
                    Exit Sub
                End If
                ' Mínimo con 45 días post servicio
                If CDate(Me.txtFecha) - CDate(Cells(iARH, 8)) _
                  < Range("Configuracion!C5") Then
                    iAviso = 7
                    Avisos
                    'Exit Sub
                End If
                ' Último servicio
                If UCase(Cells(iARH, 9)) = "CALOR" Then
                    iAviso = 8
                    Avisos
                    Exit Sub
                End If
                ' Escribir Datos
                'Servicio-DíasCarga 00-000
                sMetadato = Format(Cells(iARH, 7), "00") & "-" _
                  & Format(CDate(txtFecha) - _
                  CDate(Cells(iARH, 8)), "000")
                Cells(iARH, 11) = sDx 'Status
                On Error Resume Next
                With Cells(iARH, 12)
                    .Value = CDate(Cells(iARH, 8)) + 213 'FxSecar
                    .NumberFormat = "d-mmm-yy"
                End With
                With Cells(iARH, 13)
                    .Value = CDate(Cells(iARH, 8)) + 273 'FxParir
                    .NumberFormat = "d-mmm-yy"
                End With
                If Range(iARH, 14) = "pAb" Then _
                  Range(iARH, 14).Clear 'Clave1
                wsH2.Cells(iARH2, 3) = _
                  CDate(Cells(iARH, 8)) - _
                  CDate(Cells(iARH, 6)) 'DAbiertos
                On Error GoTo 0
            Else
                ' Previamente Gestante
                If Cells(iARH, 11) = "P" Then
                    ' Avisos informativo
                    iAviso = 5
                    Avisos
                    ' Registrar Dato
                    Cells(iARH, 11) = sDx 'Status
                    Cells(iARH, 12).Clear 'FxSecar
                    Cells(iARH, 13).Clear 'FxParir
                    wsH2.Cells(iARH2, 3).Clear 'DAbiertos
                End If
        End If
    End If
    If sLocAnimal = "R" Then
        ChecarSexo
         ' Mínimo un servicio
        If Me.ComboBox4 = "Gestante" Then
                ' Sin servicios
                If wsR.Cells(iAR, 6) < 1 Then
                    iAviso = 6
                    Avisos
                    Exit Sub
                End If
                ' Mínimo con 45 días post servicio
                If CDate(Me.txtFecha) - _
                  CDate(wsR.Cells(iAR, 7)) < _
                  Range("Configuracion!C5") Then
                    iAviso = 7
                    Avisos
                    ' Exit Sub
                End If
                ' Último servicio
                If UCase(wsR.Cells(iAR, 8)) = UCase("Calor") Then
                    iAviso = 8
                    Avisos
                    Exit Sub
                End If
                ' Escribir Datos
                ' Servicio-DíasCarga 00-000
                sMetadato = Format(wsR.Cells(iAR, 6), "00") & "-" _
                  & Format(CDate(txtFecha) - _
                  CDate(wsR.Cells(iAR, 7)), "000")
                wsR.Cells(iAR, 10) = sDx  'Status
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
                    iAviso = 5
                    Avisos
                    ' Registrar Dato
                    wsR.Cells(iAR, 10).Clear 'Status
                    wsR.Cells(iAR, 11).Clear 'FxParir
                    If wsR.Cells(iAR, 12) = "pAb" Then _
                      wsR.Cells(iAR, 13).Clear 'Clave1
                    wsIV.Cells(iARowInfoVital, 11).Clear 'EdadAlParto
                End If
        End If
    End If
    
    ConsecutivoEventos
End Sub

Private Sub ProcesarEnfermedad()
    ' Misma Enfermedad
    Dim dFecha
    ChecarMismoEvento
    If bFlagError Then Exit Sub
    dFecha = _
      CDate(BuscarUltimoEvento(Me.cmboIdArete, (sEvento)))
    If CDate(Me.txtFecha) = CDate(dFecha) Then
        iAviso = 3
        Avisos
        Exit Sub
    End If
    ' Escribir Datos
    If sLocAnimal = "H" Then _
      Cells(iARH, 15) = sEvento 'Enf
    If sLocAnimal = "R" Then _
      wsR.Cells(iAR, 13) = sEvento 'Enf
    CalcularProxRevision
    ConsecutivoEventos
End Sub

Private Sub ProcesarImantacion()
    Dim sRespuesta As String
    Set ws = Worksheets("InfoVitalicia")
    Desproteger
    If CDate(Me.txtFecha) = _
      CDate(BuscarUltimoEvento(Me.cmboIdArete, "Imantación")) Then
        iAviso = 3
        sTextoMsj2 = "Imantado"
        Avisos
        Exit Sub
    End If
    If Not ws.Cells(iARowInfoVital, 9) = vbNullString Then
        iAviso = 3
        sTextoMsj2 = "Imantado"
        Avisos
        Exit Sub
    End If
    QTextBox4
    If bFlagError Then Exit Sub 'GoTo ControlErrores
    sRespuesta = MsgBox("Confirmar la Imantación del Animal", _
      vbYesNo + vbDefaultButton2 + vbQuestion, sMsjTitulo)
    If sRespuesta = vbYes Then
        With ws.Cells(iARowInfoVital, 9) 'F.Imán
            .Value = CDate(Me.txtFecha)
            .NumberFormat = "d-mmm-yy"
        End With
    End If
    ConsecutivoEventos
    If Not CBool(Range("Desarrollador!B6")) Then _
      ws.Visible = xlSheetVeryHidden
End Sub

Private Sub ProcesarMovimientos()
    If sLocAnimal = "H" Then
        ' Mismo corral
        If Cells(iARH, 2) = CDbl(Me.TextBox5) Then
            iAviso = 11
            Avisos
            Exit Sub
        End If
        ' Reportada previamente como seca
        If Cells(iARH, 2) = _
          Range("Configuracion!C9") And CDbl(Me.TextBox5) <> _
          Range("Configuracion!C9") Then
            iAviso = 9
            Avisos
        End If
        ' Escribir Datos
        DCC
        Cells(iARH, 2) = CDbl(Me.TextBox5) 'Corral
    End If
    If sLocAnimal = "R" Then
        ' Mismo corral
        If wsR.Cells(iAR, 2) = CDbl(Me.TextBox5) Then
            iAviso = 11
            Avisos
            Exit Sub
        End If
        ' Reportada previamente como seca
        If wsR.Cells(iAR, 2) = _
          Range("Configuracion!C9") And CDbl(Me.TextBox5) <> _
          Range("Configuracion!C9") Then
            iAviso = 9
            Avisos
        End If
        ' Escribir Datos
        DCC
        wsR.Cells(iAR, 2) = CDbl(Me.TextBox5) 'Corral
    End If
    ConsecutivoEventos
End Sub

Private Sub DCC()
' Calcula los Días Cabeza Corral
    Dim iBus As Long
    Set ws = Worksheets("Eventos")
    iBus = BUS(Me.cmboIdArete, "Mov")
    If iBus > 0 Then _
      ws.Cells(iBus, 9) = _
      Format(CDate(Me.txtFecha) - CDate(ws.Cells(iBus, 2)), "000")
End Sub

Private Sub ProcesarNota()
    QTextBox4
    If bFlagError Then Exit Sub
    QTextBox5
    If bFlagError Then Exit Sub
    ConsecutivoEventos
End Sub

Private Sub ProcesarOtro()
    QComboBox4
    If bFlagError Then Exit Sub
    If Not Me.ComboBox4 = "Vacunación" Then _
      QTextBox5 Else QComboBox5
    If bFlagError Then Exit Sub
    QTextBox6
    If bFlagError Then Exit Sub
    Set ws = Worksheets("InfoVitalicia")
    Desproteger
    Select Case Me.ComboBox4
        Case "Vacunación"
            ProcesarVacunacion
        Case "DNB"
            sEvento = "DNB"
            ChecarMismoEvento
            Cells(iARH, 14) = Me.ComboBox4
        Case "TB+"
            sEvento = "TB+"
            ChecarMismoEvento
            Cells(iARH, 14) = Me.ComboBox4
        Case "K Caseína"
            sEvento = "KCaseína"
            ChecarMismoEvento
            wsIV.Cells(iARowInfoVital, 6) = _
              Me.TextBox5
        Case "Id Adicional"
            sEvento = "ID+"
            ChecarMismoEvento
            wsIV.Cells(iARowInfoVital, 1) = _
              Me.TextBox5
        Case Else
            iAviso = 98
            Avisos
            Exit Sub
    End Select
    ConsecutivoEventos
    If Not CBool(Range("Desarrollador!B6")) Then _
      ws.Visible = xlSheetVeryHidden
End Sub

Private Sub ProcesarParto()
    Dim ws2 As Worksheet
    Dim col, iRenglon, iNRowR As Long
    Dim bFlagBaja As Boolean
    sPartoDet = ""
    Select Case Me.ComboBox4
        Case "Natural"
        Case "Asistido"
            sPartoDet = "AS"
        Case "Distocia"
            sPartoDet = "DS"
            GoTo 1234
        Case "Aborto"
            sPartoDet = "Ab"
            GoTo 1234
        Case "Inducido"
            sPartoDet = "In"
            GoTo 1234
    End Select
    ' Checar que el arete no esté repetido
    If CBool(Range("Configuracion!C32")) Then
        If Not Me.TextBox6 = vbNullString Then
            sArete = Me.TextBox6
            ChecarUnicidad
            Me.TextBox6.SetFocus
            If bFlagError Then Exit Sub
        End If
        If Not Me.TextBox7 = vbNullString Then
            sArete = Me.TextBox7
            ChecarUnicidad
            Me.TextBox7.SetFocus
            If bFlagError Then Exit Sub
        End If
    End If
    If IsNumeric(Me.cmboIdArete) Then _
      mMadre = CDbl(Me.cmboIdArete) Else _
      mMadre = Me.cmboIdArete
    If IsNumeric(Me.TextBox6) Then _
      mCria1 = CDbl(Me.TextBox6) Else _
      mCria1 = Me.TextBox6
    If IsNumeric(Me.TextBox7) Then _
      mCria2 = CDbl(Me.TextBox7) Else _
      mCria2 = Me.TextBox7
    Select Case Me.ComboBox5
        Case "Macho"
            sPartoDet = "M" & Me.TextBox6 & " " & sPartoDet
        Case "Hembra"
            sPartoDet = "H" & Me.TextBox6 & " " & sPartoDet
        Case "T Hembras"
            sPartoDet = "H" & Me.TextBox6 & ", H" & _
              Me.TextBox7 & " " & sPartoDet
        Case "T Machos"
            sPartoDet = "M" & Me.TextBox6 & ", M" & _
              Me.TextBox7 & " " & sPartoDet
        Case "FreeMartin"
            sPartoDet = "FM" & Me.TextBox6 & ", FM" & _
              Me.TextBox7 & " " & sPartoDet
    End Select
1234:
    If sLocAnimal = "H" Then
       ' sPartoDet = vbnullstring
        ' Mismo Parto
        If CDate(Me.txtFecha) = CDate(Cells(iARH, 6)) Then
            iAviso = 3
            Avisos
            Exit Sub
        End If
        mCorral = Range("Configuracion!B11")
        mPadreCria = UCase(Cells(iARH, 9))
        ' Registrar LactanciasAnteriores
        If Not IsEmpty(Cells(iARH, 5)) Then LactAnteriores
        ' Limpiar registros en Hato2
        Set ws = Worksheets("Hato2")
        Desproteger
        Range(ws.Cells(iARH2, 2), ws.Cells(iARH2, 19)).Clear
        ws.Cells(iARH2, 18) = Me.ComboBox4 'TipoParto
        If Not CBool(Range("Desarrollador!B6")) Then _
          ws.Visible = xlSheetVeryHidden
        ' Adicionar Reemplazos
        If Me.ComboBox4 = "Distocia" Or Me.ComboBox4 = "Aborto" Or _
          Me.ComboBox4 = "Inducido" Then _
          GoTo 3412
        If CBool(Range("Configuracion!C32")) Then 'ControlDeReemplazos
            Set ws2 = Worksheets("InfoVitalicia")
            Desproteger
            Set ws = Worksheets("Reemplazos")
            Desproteger
            ' Adicionat Cría1
            If Me.TextBox6 = vbNullString Then GoTo 2341
            If InStr(UCase(Me.TextBox6), "M") Then GoTo 2341
            If InStr(UCase(Me.TextBox6), "R") Then GoTo 2341
            If InStr(UCase(Me.TextBox6), "V") Then GoTo 2341
            AdicionarCria1
2341:
            ' Adicionat Cría2
            If Me.TextBox7 = vbNullString Then GoTo 3412
            If InStr(UCase(Me.TextBox7), "M") Then GoTo 3412
            If InStr(UCase(Me.TextBox7), "R") Then GoTo 3412
            If InStr(UCase(Me.TextBox7), "V") Then GoTo 3412
            AdicionarCria2
3412:
            If Not CBool(Range("Desarrollador!B6")) Then _
              wsIV.Visible = xlSheetVeryHidden
        End If
        Set ws = Worksheets("Hato")
        If sPartoDet = "Ab" Then
                'Parto-DiasCarga
                If IsEmpty(Cells(iARH, 5)) Then sMetadato = "00-000" Else _
                  sMetadato = Format(Cells(iARH, 5), "00") & "-" & _
                  Format(CDate(txtFecha) - CDate(Cells(iARH, 8)), "000")
            Else
                On Error GoTo 4312
                'Parto-DiasSeca
                If IsEmpty(Cells(iARH, 5)) Then sMetadato = "01-000" Else _
                  sMetadato = Format(Cells(iARH, 5) + 1, "00") & "-" & _
                  Format(CDate(txtFecha) - CDate(wsH2.Cells(iARH2, 16)), "000")
                GoTo 4321
4312:
                'Vaquilla
                sMetadato = Format(Cells(iARH, 5) + 1, "00") & "-000"
4321:
                On Error GoTo 0
        End If
        With ws
            If .Cells(iARH, 5) + 1 > 1 Then 'Corral
                    If Not .Cells(iARH, 2) = Range("Configuracion!C11") Then
                        DCC
                        bMovCorral = True
                    End If
                    .Cells(iARH, 2) = Range("Configuracion!C11")
                Else
                    If Not .Cells(iARH, 2) = Range("Configuracion!C12") Then
                        DCC
                        bMovCorral = True
                    End If
                    .Cells(iARH, 2) = Range("Configuracion!C12")
            End If
            .Cells(iARH, 3) = _
              Format(Range("Configuracion!C24"), "0.0") ' Prod
            ' si aborto mas de 152 días then parto +1
            .Cells(iARH, 5) = .Cells(iARH, 5) + 1 'No.Parto
            With .Cells(iARH, 6)    ' F.Parto
                .Value = CDate(Me.txtFecha)
                .NumberFormat = "d-mmm-yy"
            End With
            Cells(iARH, 15).Clear
            Select Case Right(sPartoDet, 2)
                Case "DS", "Ab", "AS", "In"
                    .Cells(iARH, 15) = Right(sPartoDet, 2)
            End Select
            ' Limpiar Celdas de columnas G,H,I,J,K,L,M,N
            Range(Cells(iARH, 7), Cells(iARH, 14)).Clear
        End With
    End If
    '******************************
    If sLocAnimal = "R" Then
        ChecarSexo
        If bFlagError = True Then Exit Sub
        bFlagBaja = False
        ' Mismo Parto
        If CDate(Me.txtFecha) = CDate(wsR.Cells(iAR, 5)) Then
            iAviso = 3
            Avisos
            Exit Sub
        End If
        'On Error GoTo 0
        mPadreCria = UCase(wsR.Cells(iAR, 8))
        If sPartoDet = "Ab" Then
                ' Parto_DíasCarga
                sMetadato = "00-" & _
                Format(CDate(txtFecha) - CDate(wsR.Cells(iAR, 7)), "000")
            Else
                sMetadato = "01-000"
        End If
        BajaReemplazos
        If CBool(Range("Configuracion!C32")) Then 'ControlDeReemplazos
            ' Adicionar Reemplazos
            Set ws = Worksheets("Reemplazos")
            Desproteger
            If Me.TextBox6 = vbNullString Then GoTo 4123
            If InStr(UCase(Me.TextBox6), "M") Then GoTo 4123
            If InStr(UCase(Me.TextBox6), "R") Then GoTo 4123
            If InStr(UCase(Me.TextBox6), "V") Then GoTo 4123
            ' Adicionar Reemplazos Cría1
            AdicionarCria1
4123:
            ' Adicionar Cria2
            If Me.TextBox7 = vbNullString Then GoTo 4231
            If InStr(UCase(Me.TextBox7), "M") Then GoTo 4231
            If InStr(UCase(Me.TextBox7), "R") Then GoTo 4231
            If InStr(UCase(Me.TextBox7), "V") Then GoTo 4231
            AdicionarCria2
4231:
            If Not CBool(Range("Desarrollador!B6")) Then _
              wsIV.Visible = xlSheetVeryHidden
        End If
        ' Alta Hato
        iRenglon = TamañoTabla("Tabla1") + 2
        With wsH
            .Cells(iRenglon, 1) = CDbl(Me.cmboIdArete) 'Arete
            DCC
            .Cells(iRenglon, 2) = Range("Configuracion!C12") 'Corral
            .Cells(iRenglon, 3) = _
              Format(Range("Configuracion!C24"), "0.0") 'Prod
            ' si aborto mas de 152 días then parto +1
            .Cells(iRenglon, 5) = 1  'Num. Parto
            With .Cells(iRenglon, 6) 'F.Parto
                .Value = CDate(Me.txtFecha)
                .NumberFormat = "d-mmm-yy"
            End With
            .Cells(iRenglon, 15).Clear 'Clave2
            Select Case Right(sPartoDet, 2)
                Case "AS", "Ab", "DS", "In"
                    .Cells(iRenglon, 15) = Right(sPartoDet, 2)
                End Select
        End With
        AltaHoja2
        wsIV.Cells(iARowInfoVital, 11) = _
          Int((CDate(Me.txtFecha) - _
          wsR.Cells(iAR, 5)) / 30.4) 'Edad1Ser
        wsR.Select
        Desproteger
        wsR.Rows(iAR).Delete Shift:=xlShiftUp
        wsH.Select
        sMetadato = "01-000"
    End If
    ConsecutivoEventos
End Sub

Private Sub ProcesarRevision()
' Ultima revisión
    If CDate(Me.txtFecha) = _
      CDate(BuscarUltimoEvento(Me.cmboIdArete, _
      "Rev")) Then
        iAviso = 14
        Avisos
        Exit Sub
    End If
    CalcularProxRevision
    If InStr(UCase(TextBox4), "OE") Or (InStr(UCase(TextBox4), "ODE") _
      And InStr(UCase(TextBox4), "OIE")) Then sMetadato = "Anestro"
    ConsecutivoEventos
End Sub

Private Sub ProcesarSecar()
    ' Vaca ya Seca
    If Cells(iARH, 2) = _
      Range("Configuracion!C9") Then
        iAviso = 9
        Avisos
        'Exit Sub
    End If
    ' Menos de 152 dias en leche
    If CDate(Me.txtFecha) - _
      CDate(Cells(iARH, 5)) <= 152 Then _
        iAviso = 10
    ' Animal no gestante
    If Not Cells(iARH, 11) = "P" Then
        iAviso = 24
        Avisos
    End If
    Set ws = Worksheets("Hato2")
    Desproteger
    ' Escribir Datos
    'Parto-DEL
    sMetadato = Format(Cells(iARH, 5), "00") & "-" & _
      Format(CDate(txtFecha) - _
      CDate(Cells(iARH, 6)), "000")
    If Not Cells(iARH, 2) = _
      Range("Configuracion!C9") Then bMovCorral = True
    Cells(iARH, 2) = _
      Range("Configuracion!C9") 'Corral
    Cells(iARH, 3).Clear 'Produccion
    Cells(iARH, 4).Clear 'del
    Cells(iARH, 12) = "**SECA**" 'FxSecar
    With ws.Cells(iARH2, 16) 'F.Secado
        .Value = CDate(Me.txtFecha)
        .NumberFormat = "d-mmm-yy"
    End With
    If Not CBool(Range("Desarrollador!B6")) Then _
      ws.Visible = xlSheetVeryHidden
    ConsecutivoEventos
End Sub

Private Sub ProcesarServicios()
    Dim iD1S, iBus As Long
    Set ws = Worksheets("Eventos")
    If sLocAnimal = "H" Then
        ChecarMismoEvento
        If bFlagError Then Exit Sub
        ' Último Servicio
        If CDate(txtFecha) <= CDate(Cells(iARH, 8)) Then
            iAviso = 2
            Avisos
            Exit Sub
        End If
        If CDate(txtFecha) - CDate(Cells(iARH, 8)) <= 18 And _
          Not IsEmpty(Cells(iARH, 8)) Then
            iAviso = 16
            Avisos
        End If
        If CDate(txtFecha) - CDate(Cells(iARH, 8)) >= 36 And _
          Not IsEmpty(Cells(iARH, 8)) Then
            iAviso = 14
            Avisos
        End If
        ' Escribir Datos
        'Servicio-DíasÚltimoServicio-DEL
        If IsEmpty(Cells(iARH, 7)) Then _
          sMetadato = "01-" Else _
          sMetadato = Format(Val(Cells(iARH, 7)) + 1, "00") & "-"
        If IsEmpty(Cells(iARH, 8)) Then _
          sMetadato = sMetadato & "000" _
          Else _
          sMetadato = _
          sMetadato & Format(CDate(txtFecha) - _
          CDate(Cells(iARH, 8)), "000")
        ' Añadir DEL
        sMetadato = sMetadato & "-" & _
          Format(CDate(txtFecha) - CDate(Cells(iARH, 6)), "000")
        If Cells(iARH, 11) = "P" Then
            iBus = BUS(Me.cmboIdArete, "Serv")
            ws.Cells(iBus, 9) = Left(ws.Cells(iBus, 9), 11) & "R"
            Cells(iARH, 11).Clear 'Status
            Cells(iARH, 12).Clear 'FxSecar
            Cells(iARH, 13).Clear 'FxParir
            Cells(iARH, 14) = "pAb" 'Clave1
            Set ws = Worksheets("Hato2")
            Desproteger
            ws.Cells(iARH2, 3).Clear 'dAbiertos
            iAviso = 5
            Avisos
        End If
        '1er Servicio
        If Cells(iARH, 7) + 1 = 1 Then
            wsH2.Cells(iARH2, 2) = _
              CDate(Me.txtFecha) - _
              CDate((Cells(iARH, 6))) 'Dias1Serv
            If wsH2.Cells(iARH2, 17) = vbNullString Then _
              wsH2.Cells(iARH2, 17) = _
                CDate(Me.txtFecha) - CDate(Cells(iARH, 6)) 'd1Calor
        End If
        With Cells(iARH, 8) 'F.Servicio
            .Value = CDate(txtFecha)
            .NumberFormat = "d-mmm-yy"
        End With
        Cells(iARH, 7) = Cells(iARH, 7) + 1 'Servicio
        If Me.TextBox4 = vbNullString Then
                Cells(iARH, 9) = "N.D."
            Else
                'Semental
                Cells(iARH, 9) = UCase(Me.TextBox4)
        End If
        If Me.TextBox5 = vbNullString Then
                Cells(iARH, 10) = "N.D."
            Else
                'Tecnico
                Cells(iARH, 10) = UCase(Me.TextBox5)
        End If
        Cells(iARH, 11).Clear 'Status
    End If
    If sLocAnimal = "R" Then
        ' Último Servicio
        If CDate(txtFecha) <= CDate(wsR.Cells(iAR, 7)) Then
            iAviso = 2
            Avisos
            Exit Sub
        End If
        If CDate(txtFecha) - CDate(wsR.Cells(iAR, 7)) <= 18 And _
          Not IsEmpty(wsR.Cells(iAR, 7)) Then
            iAviso = 16
            Avisos
        End If
        If CDate(txtFecha) - CDate(wsR.Cells(iAR, 7)) >= 36 And _
          Not IsEmpty(wsR.Cells(iAR, 7)) Then
            iAviso = 14
            Avisos
        End If
        ChecarSexo
        ' Escribir Datos
        'Servicio-DíasÚltimoServicio-DEL
         If IsEmpty(wsR.Cells(iAR, 6)) Then _
          sMetadato = "01-" _
          Else _
          sMetadato = _
          Format(Val(wsR.Cells(iAR, 6)) + 1, "00") & "-"
        If IsEmpty(wsR.Cells(iAR, 7)) Then _
          sMetadato = sMetadato & "000" _
          Else _
          sMetadato = _
          sMetadato & Format(CDate(txtFecha) - _
          CDate(wsR.Cells(iAR, 7)), "000")
        ' Añadir DEL
        sMetadato = sMetadato & "-" & _
          Format(CDate(txtFecha) - CDate(wsR.Cells(iAR, 5)), "000")
        'If IsEmpty(wsR.Cells(iAR, 6)) Then sMetadato = "01-000" Else _
          sMetadato = Format(Val(wsR.Cells(iAR, 6)) + 1, "00") _
          & "-" & Format(CDate(txtFecha) - CDate(wsR.Cells(iAR, 7)), "000")
        'sMetadato = sMetadato & "-" & _
          Format(CDate(txtFecha) - CDate(wsR.Cells(iAR, 5)), "000")
        If wsR.Cells(iAR, 10) = "P" Then
            iBus = BUS(Me.cmboIdArete, "Serv")
            ws.Cells(iBus, 9) = Left(ws.Cells(iBus, 9), 11) & "R"
            wsR.Cells(iAR, 10).Clear 'Status
            wsR.Cells(iAR, 11).Clear 'FxParir
            wsR.Cells(iAR, 12) = "pAb" 'Clave1
            wsIV.Cells(iARowInfoVital, 11).Clear 'EdadAlParto
            iAviso = 5
            Avisos
        End If
        wsR.Cells(iAR, 6) = _
          wsR.Cells(iAR, 6) + 1 'Servicio
        With wsR.Cells(iAR, 7) 'F.Servicio
            .Value = CDate(Me.txtFecha)
            .NumberFormat = "d-mmm-yy"
        End With
        If Me.TextBox4 = vbNullString Then
                wsR.Cells(iAR, 8) = "N.D."
            Else
                wsR.Cells(iAR, 8) = _
                  UCase(Me.TextBox4) 'Semental
        End If
        If Me.TextBox5 = vbNullString Then
                wsR.Cells(iAR, 9) = "N.D."
            Else
                wsR.Cells(iAR, 9) = _
                  UCase(Me.TextBox5) 'Tecnico
        End If
        If wsR.Cells(iAR, 6) = 1 Then _
          wsIV.Cells(iARowInfoVital, 10) = _
          Int((CDate(wsR.Cells(iAR, 7)) _
          - CDate(wsR.Cells(iAR, 5))) / 30.4) 'Edad1Serv
    End If
    If Not CBool(Range("Desarrollador!B6")) Then _
      wsH2.Visible = xlSheetVeryHidden
    ConsecutivoEventos
End Sub

Private Sub ProcesarProduccion()
    Dim mProdAcum, mDiasProd, mProdMax, _
      mPicoProd, mProy305, mPersist As Double
    Dim iCol, iParto, i As Long
    Dim dFParto As Date
    Dim bTest As Boolean
    ' Ingreso numérico
    If Not IsNumeric(Me.TextBox5) Then
        iAviso = 15
        Avisos
        Me.TextBox5.SetFocus
        Exit Sub
    End If
    If Not Me.TextBox6 = vbNullString And _
      Not IsNumeric(Me.TextBox6) Then
        iAviso = 15
        Avisos
      Me.TextBox6.SetFocus
      Exit Sub
    End If
    ' Producciones negativas
    If Val(TextBox5) < 0 Then
        iAviso = 12
        Avisos
        Me.TextBox5.SetFocus
        Exit Sub
    End If
    ChecarFechaParto
    If bFlagError = True Then Exit Sub
    ' Mismo Pesaje
    If CDate(Me.txtFecha) = _
      CDate(BuscarUltimoEvento(Me.cmboIdArete, _
      "Prod")) Then
        iAviso = 3
        Avisos
        Exit Sub
    End If
    iParto = Cells(iARH, 5)
    dFParto = CDate(Cells(iARH, 6))
    Set ws = Worksheets("Hato2")
    Desproteger
    ws.Visible = xlSheetVisible
    ' Vaca Seca
    If Cells(iARH, 2) = Range("Configuracion!C9") Or _
      Cells(iARH, 2) = Range("Configuracion!C10 ") Then
        iAviso = 9
        Avisos
    End If
    ' Escribir Datos
    mDiasProd = CDate(Me.txtFecha) - dFParto
    If Cells(iARH, 3) > 0 Then _
      mPersist = Int(CDbl(Me.TextBox5) / Cells(iARH, 3) * 100) _
      Else mPersist = 0
    ' DEL-Persistencia-Parto-Corral
    sMetadato = Format(mDiasProd, "000") & "-" & _
      Format(mPersist, "000") & "-" & _
      Format(wsH.Cells(iARH, 5), "00") & "-" & _
      Format(wsH.Cells(iARH, 2), "00")
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
        .Value = CDbl(Me.TextBox5)
        .NumberFormat = "0.00"
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
          Val(ws.Cells(iARH2, iCol)) * 30
    Next iCol
    mProdAcum = Int(mProdAcum / 10) * 10
    ws.Cells(iARH2, 14) = mProdAcum 'Prod.Acum
    ' Proyectar producción a 305d
    If Not IsEmpty(wsIV.Cells(iARowInfoVital, 6)) Then _
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
    Range("CurvaLact!N10").GoalSeek Goal:=Val(Me.TextBox5), _
      ChangingCell:=Range("CurvaLact!N4")
    'On Error GoTo 0
    mProy305 = Range("CurvaLact!N11")
    ws.Cells(iARH2, 15) = mProy305 'Proy.305d
    'Calcular peristencia
    If mPersist > 0 Then ws.Cells(iARH2, 19) = mPersist
    With Cells(iARH, 3) 'Prod
        .Value = CDbl(Me.TextBox5)
        .NumberFormat = "0.00"
        If CBool(Range("Configuracion!B65")) Then _
          Cells(iARH, 16) = Format(mProy305, "#,#")
    End With
    ' Escribir fecha
    If Not Me.TextBox6 = vbNullString And _
      Not Me.TextBox6 = Cells(iARH, 2) Then
        bMovCorral = True
        DCC
        Cells(iARH, 2) = CDbl(Me.TextBox6)
    End If
    If Not CBool(Range("Desarrollador!B6")) Then _
      ws.Visible = xlSheetVeryHidden
    ConsecutivoEventos
End Sub

Private Sub ProcesarVacunacion()
    Dim sRespuesta As String
    Select Case Me.ComboBox5
        Case "Brucela"
            sEvento = "Vac-Br"
        Case "Triple"
            sEvento = "Vac-III"
        Case "Cuádruple"
            sEvento = "Vac-IV"
        Case "Leptospira"
            sEvento = "Vac-Lepto"
        Case "DBV"
            sEvento = "Vac-DBV"
    End Select
    ChecarMismoEvento
    If bFlagError Then Exit Sub
    If Me.ComboBox5 = "Brucela" Then
        sRespuesta = _
          MsgBox("Confirmar la Vacunación del Animal", _
          vbYesNo + vbDefaultButton2 + vbQuestion, sMsjTitulo)
        If sRespuesta = vbYes Then
            Set ws = Worksheets("InfoVitalicia")
            Desproteger
            With ws.Cells(iARowInfoVital, 8)
                .Value = CDate(Me.txtFecha)
                .NumberFormat = "d-mmm-yy"
            End With
        End If
    End If
End Sub

Private Sub QComboBox4()
    ' Comprueba si ComboBox4 está vacío
    If Me.ComboBox4 = vbNullString Then
        iAviso = 104
        Avisos
        Me.ComboBox4.SetFocus
    End If
End Sub

Private Sub QComboBox5()
    ' Comprueba si ComboBox5 está vacío
    If Me.ComboBox5 = vbNullString Then
        iAviso = 105
        Avisos
        Me.ComboBox5.SetFocus
    End If
End Sub

Private Sub QTextBox4()
    ' Comprueba si TextBox4 está vacío
    If Me.TextBox4 = vbNullString Then
        iAviso = 104
        Avisos
        Me.TextBox4.SetFocus
    End If
End Sub

Private Sub QTextBox5()
    ' Comprueba si TextBox5 está vacío
    If Me.TextBox5 = vbNullString Then
        iAviso = 105
        Avisos
        Me.TextBox5.SetFocus
    End If
End Sub

Private Sub QTextBox6()
    ' Comprueba si TextBox6 está vacío
    If Me.TextBox6 = vbNullString Then
        iAviso = 106
        Avisos
        Me.TextBox6.SetFocus
    End If
End Sub

Private Sub QTextBox7()
    ' Comprueba si TextBox7 está vacío
    If Me.TextBox7 = vbNullString Then
        iAviso = 107
        Avisos
        Me.TextBox7.SetFocus
    End If
End Sub

Private Sub TextBox4_AfterUpdate()
    If Me.cmboEvento = "Alta" Then
        On Error Resume Next
        Me.TextBox4 = Format(CDate(Me.TextBox4), "d-mmm-yy")
        On Error GoTo 0
    End If
End Sub

Private Sub TextBox5_AfterUpdate()
    If Me.TextBox5 <> vbNullString Then
        If Me.cmboEvento = "Producción" Then
            If Not IsNumeric(Me.TextBox5) Then GoTo ErrorNumerico
            If Me.TextBox5 >= 100 Then
                On Error GoTo ErrorNumerico
                Me.TextBox5 = Val(Me.TextBox5) / 10
            End If
            Me.TextBox5 = Format(Val(Me.TextBox5), "0.0")
        End If
    End If
    Exit Sub

ErrorNumerico:
    iAviso = 15
    Avisos
End Sub

Private Sub txtFecha_AfterUpdate()
    On Error Resume Next
    Me.txtFecha = Format(CDate(Me.txtFecha), "d-mmm-yy")
    On Error GoTo 0
End Sub

Private Sub UserForm_Initialize()
    Dim mRenglones As Double
    Application.ScreenUpdating = CBool(Range("Desarrollador!B6"))
    Me.Caption = Range("Configuracion!C3")
    Me.Label9.Caption = "" ' Indicar Versión del módulo
    bAceptar = False
    bFlagError = False
    sMsjTitulo = "Registro de eventos"
    Set wsH = Worksheets("Hato")
    Set wsH2 = Worksheets("Hato2")
    Set wsIV = Worksheets("InfoVitalicia")
    Set wsR = Worksheets("Reemplazos")
    Range("Desarrollador!B20") = "T" 'Detener cálculos en la hoja
    Set ws = Worksheets("Reemplazos")
    ws.DisplayPageBreaks = False
    ws.Select
    Desproteger
    With Range("Tabla2")
        .AutoFilter
        .AutoFilter
    End With
    Set ws = Worksheets("Hato")
    ws.DisplayPageBreaks = False
    ws.Select
    Desproteger
    With Range("Tabla1")
        .AutoFilter
        .AutoFilter
    End With
    If Not ActiveSheet.Name = "Hato" Then Worksheets("Hato").Activate
    Me.cmboIdArete.RowSource = "Tabla1[[Arete]]"
    Poblar_cmboEvento
    Me.txtFecha = Format(Date, "dd-mmm-yy")
    UltimoEventoCapturado
End Sub

Private Sub UserForm_QueryClose(mCancel As Integer, _
                                mCloseMode As Integer)
' Deshabilitar el Botón X para cerrar el cuadro de diálogo
    If mCloseMode <> vbFormCode Then
        MsgBox _
          "Utiliza 'Cerrar' para salir de este formulario.", _
          vbExclamation, _
          "JP's Automatización de Aplicaciones"
        mCancel = True
    End If
End Sub

Private Function BUS(Arete_Buscado As Variant, _
  Evento_Buscado As String)
' Devuelve el renglón del último evento buscado
' Arete_Buscado
' Evento_Buscado:
' Ejemplo: =BuscarUltimoEvento(1084,"Serv")
    Dim cont, i As Long
    Dim Ocurrencia As Double
    Dim rCelda As Range
    ' Contar las ocurrencias de estos eventos
    Ocurrencia = WorksheetFunction. _
      CountIfs(Range("Tabla6[Arete]"), Arete_Buscado, _
        Range("Tabla6[Clave]"), Evento_Buscado)
    BUS = 0
    For Each rCelda In Range("Tabla6[Arete]")
       If Val(rCelda.Offset(i, 0)) = _
         Arete_Buscado And _
         rCelda.Offset(i, 2) = Evento_Buscado Then
           cont = cont + 1
           ' Si es la última ocurrencia del evento
           If cont = Ocurrencia Then
               ' Devuelve fecha del ultimo evento
               BUS = rCelda.Offset.Row
               GoTo 100
           End If
       End If
    Next
100:
End Function

Private Sub UltimoEventoCapturado()
' Muestra el último evento
    Me.Label11.Caption = _
      Format(BuscarValorInverso( _
      Application.WorksheetFunction.Max(Range("Tabla6[Indice]")), _
      Range("Tabla6[Indice]"), -8), "dd-mmm-yy") & "|" & _
      BuscarValorInverso( _
      Application.WorksheetFunction.Max(Range("Tabla6[Indice]")), _
      Range("Tabla6[Indice]"), -9) & "|" & _
      BuscarValorInverso( _
      Application.WorksheetFunction.Max(Range("Tabla6[Indice]")), _
      Range("Tabla6[Indice]"), -7)
End Sub
