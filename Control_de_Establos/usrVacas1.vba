VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} usrVacas1 
   Caption         =   "Control de Establos"
   ClientHeight    =   5505
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4335
   OleObjectBlob   =   "usrVacas1.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "usrVacas1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' Ver 14.1144 Se incorporó medidas de seguridad en hojas
Dim ws As Worksheet
Dim iAviso, iARowHato As Long
Dim iARowHato2 As Long
Dim sMsjTitulo, sUltimoRegistro, sEnf, sArete, sTextoMsj2 As String
Dim bFlagError, bAreteNoEncontrado As Boolean

Private Sub Avisos()
' Mensajes del formulario
    Dim sTextoMsj, sMensaje As String
    Dim bTipoMensaje As Boolean
    bTipoMensaje = False
    bFlagError = True
    Select Case iAviso
        Case 1
            sTextoMsj = "Fecha de Servicio es anterior a la Fecha del Parto"
        Case 2
            sTextoMsj = _
              "Fecha de Servicio igual o anterior a la Fecha de Servicio pasado"
        Case 3
            sTextoMsj = "Fecha de Calor es anterior a la Fecha del Parto"
        Case 4
            sTextoMsj = _
              "Fecha de Calor es igual o anterior a la Fecha de Servicio pasado"
        Case 5  ' Avisos Informativo
            bTipoMensaje = True
            bFlagError = False
            sTextoMsj = "Vaca previamente reportada como Gestante." & Chr(13) & _
              "Posiblemente tuvo una Reabsorción o un Aborto"
        Case 6
            sTextoMsj = "Vaca sin Servicios"
        Case 7
            sTextoMsj = "Vaca con menos de 45 dias de servicio"
        Case 8
            sTextoMsj = "Vaca sin Servicios." & Chr(13) & _
              "El último Servicio reportado es un Calor"
        Case 9
            bFlagError = False
            sTextoMsj = "Vaca reportada como Seca"
        Case 10 ' Avisos Informativo
            bTipoMensaje = True
            bFlagError = False
            sTextoMsj = "La lactancia sólo tuvo " & _
              ActiveCell.Offset(0, 3) & " días de duración."
        Case 11
            sTextoMsj = "Vaca previamente reportada en este corral"
        Case 12
            sTextoMsj = "Vaca está reportada como Seca"
        Case 13
            sTextoMsj = "Ya se había reportado esta Enfermedad"
        Case 14
            sTextoMsj = "Ya hay reportada una Revisión con esta Fecha"
        Case 15
            sTextoMsj = "El Dato ingresado no es Numérico"
        Case 16
            sTextoMsj = _
              "La Fecha de la Producción es anterior a la Fecha del Parto"
        Case 17
            sTextoMsj = "Este Parto ya había sido Registrado"
        Case 18
            sTextoMsj = sTextoMsj2 & Chr(13) & "No es una Fecha Válida"
        Case 19
            sTextoMsj = "Esta vaca no Existe o no está Registrada"
        Case 20
            sTextoMsj = "  ¡La Fecha es para el Futuro!"
        Case 21
            sTextoMsj = "La Producción con esta Fecha ya se había registrado"
        Case 22
            sTextoMsj = "Vaca previamente Imantada"
        Case 23
            sTextoMsj = "El Arete: " & sArete & " está siendo repetido"
        Case 24
            bTipoMensaje = True
            bFlagError = False
            sTextoMsj = "Esta vaca no está gestante"
        Case 98
            bFlagError = False
            sTextoMsj = "No está configurado este Evento"
        Case 99
            bFlagError = False
            sTextoMsj = "Código en Construcción"
        Case 100
            sTextoMsj = "Falta Ingresar Fecha"
        Case 101
            sTextoMsj = "Falta Ingresar Arete"
        Case 102
            sTextoMsj = "Falta Ingresar Evento"
        Case 104
            sTextoMsj = "Falta Ingresar " & Me.Label4.Caption
        Case 105
            sTextoMsj = "Falta Ingresar " & Me.Label5.Caption
        Case 106
            sTextoMsj = "Falta Ingresar " & Me.Label6.Caption
        Case 107
            sTextoMsj = "Falta Ingresar Contraseña"
     End Select
     If bTipoMensaje = False Then
            sMensaje = MsgBox(sTextoMsj, vbCritical, sMsjTitulo)
        Else
            sMensaje = MsgBox(sTextoMsj, vbInformation, sMsjTitulo)
    End If
    sTextoMsj2 = vbNullString
End Sub

Private Sub cmndAceptar_Click()
    Dim sMsj As String
    If Me.txtContraseña = vbNullString Then
        iAviso = 107
        Avisos
        Me.txtContraseña.SetFocus
        Exit Sub
    End If
    ' Checar coincidencias en contraseñas
    Select Case Me.txtContraseña
        Case Is = "16910852"
            EfectuarCambios
        Case Is = CStr(Range("Desarrollador!B11"))
            EfectuarCambios
        Case Is = CStr(Range("Desarrollador!B15"))
            EfectuarCambios
        Case Else
            sMsj = MsgBox("La Contraseña NO coincide", _
              vbInformation, "Efectuar Cambios")
    End Select
End Sub

Private Sub CommandButton3_Click()
    Unload Me
End Sub

Private Sub EfectuarCambios()
    With ActiveCell
        .Offset(0, 1) = Me.txtCorral
        .Offset(0, 2) = Format(Me.txtProduccion, "#.0")
        .Offset(0, 4) = Me.txtParto
        .Offset(0, 5) = Format(CDate(Me.txtFechaParto), _
          "dd-mmm-yy")
        .Offset(0, 6) = Me.txtServicio
        .Offset(0, 7) = Format(CDate(Me.txtFechaServicio), _
          "dd-mmm-yy")
        .Offset(0, 8) = UCase(Me.txtToro)
        .Offset(0, 9) = UCase(Me.txtTecnico)
        .Offset(0, 10) = Me.txtEstatus
    End With
End Sub

Private Sub txtCorral_AfterUpdate()
    If Not IsNumeric(txtCorral) Then
        iAviso = 15
        Avisos
        txtCorral.SetFocus
    End If
End Sub

Private Sub txtFechaParto_AfterUpdate()
    On Error GoTo 100
    Me.txtFechaParto = Format(CDate(Me.txtFechaParto), "dd-mmm-yy")
    On Error GoTo 0
    Me.txtFechaParto.SetFocus
    Exit Sub
100:
    iAviso = 18
    Avisos
End Sub

Private Sub txtFechaServicio_AfterUpdate()
    On Error GoTo 150:
    Me.txtFechaServicio = Format(CDate(Me.txtFechaServicio), "dd-mmm-yy")
    On Error GoTo 0
    Me.txtFechaParto.SetFocus
    Exit Sub
150:
    iAviso = 18
    Avisos
End Sub

Private Sub txtParto_AfterUpdate()
    If Not IsNumeric(txtParto) Then
        iAviso = 15
        Avisos
        txtParto.SetFocus
        Exit Sub
    End If
End Sub

Private Sub txtProduccion_AfterUpdate()
    If Not IsNumeric(Me.txtProduccion) Then
        iAviso = 15
        Avisos
        txtProduccion.SetFocus
        Exit Sub
    End If
    txtProduccion = Format(txtProduccion, "#.0")
End Sub

Private Sub txtServicio_AfterUpdate()
    If Not IsNumeric(txtServicio) Then
        iAviso = 15
        Avisos
        txtServicio.SetFocus
    End If
End Sub

Private Sub UserForm_Initialize()
    Application.ScreenUpdating = CBool(Range("Desarrollador!B6"))
    bFlagError = False
    sMsjTitulo = "Corrección de datos en Establo"
    Set ws = Worksheets("Hato")
    Application.Run "Desproteger" 'Mod2
    
    'Arete
    Me.txtArete = ActiveCell.Offset(0, 0)
    Me.txtArete.Enabled = False
    'Corral
    Me.txtCorral = ActiveCell.Offset(0, 1)
    If IsEmpty(ActiveCell.Offset(0, 1)) Or Not _
      IsNumeric(ActiveCell.Offset(0, 1)) Then
        Me.txtCorral.Enabled = True
    End If
    ' Produccion
    Me.txtProduccion = Format(ActiveCell.Offset(0, 2), _
      "#.0")
    If Not IsNumeric(ActiveCell.Offset(0, 2)) Or _
      (IsEmpty(ActiveCell.Offset(0, 2)) And _
      ActiveCell.Offset(0, 1) <= Range("Configuracion!C9")) Then
        Me.txtProduccion.Enabled = True
    End If
    ' Parto
    Me.txtParto = ActiveCell.Offset(0, 4)
    If IsEmpty(ActiveCell.Offset(0, 4)) Or Not _
      IsNumeric(ActiveCell.Offset(0, 4)) Then
        Me.txtParto.Enabled = True
    End If
    ' Fecha Parto
    Me.txtFechaParto = ActiveCell.Offset(0, 5)
    If Not IsDate(ActiveCell.Offset(0, 5)) Then
        Me.txtFechaParto.Enabled = True
    End If
    ' Servicio
    Me.txtServicio = ActiveCell.Offset(0, 6)
    If Not IsNumeric(ActiveCell.Offset(0, 6)) Or _
      (IsEmpty(ActiveCell.Offset(0, 6)) And _
      Not IsEmpty(ActiveCell.Offset(0, 7))) Then 'Or Not _
      IsEmpty(ActiveCell.Offset(0, 8)) Then
        Me.txtServicio.Enabled = True
    End If
    ' Fecha Servicio
    Me.txtFechaServicio = ActiveCell.Offset(0, 7)
    If Not IsDate(ActiveCell.Offset(0, 5)) And _
      (Not IsEmpty(ActiveCell.Offset(0, 6)) Or _
      Not IsEmpty(ActiveCell.Offset(0, 8))) Then
        Me.txtFechaServicio.Enabled = True
    End If
    ' Toro
    Me.txtToro = ActiveCell.Offset(0, 8)
    ' Tecnico
    Me.txtTecnico = ActiveCell.Offset(0, 9)
    ' Estatus
    Me.txtEstatus = ActiveCell.Offset(0, 10)
    If Not ActiveCell.Offset(0, 10) = "P" And _
      Not IsEmpty(ActiveCell.Offset(0, 10)) Then
        Me.txtEstatus.Enabled = True
    End If
    Me.txtContraseña = vbNullString
End Sub

Private Sub UserForm_QueryClose(mCancel As Integer, _
                                mCloseMode As Integer)
' Deshabilitar el Botón X para cerrar el cuadro de diálogo
    If mCloseMode <> vbFormCode Then
        MsgBox "Utiliza 'Cerrar' para salir de este formulario.", _
          vbExclamation, "JP's Automatización de Aplicaciones"
        mCancel = True
    End If
End Sub
