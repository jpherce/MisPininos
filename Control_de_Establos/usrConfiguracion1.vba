VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} usrConfiguracion1 
   Caption         =   "Control de Establos"
   ClientHeight    =   6495
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8040
   OleObjectBlob   =   "usrConfiguracion1.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "usrConfiguracion1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim bCambios, bProteccion, bInitialFlag, bBorrarTodo As Boolean

Private Sub AA()
' Se ha efectuado un cambio en algún objeto del formulario
    bCambios = True
    Me.CommandButton2.Enabled = True
End Sub

Private Sub boxCambiarContraseña_AfterUpdate()
' Habilitar el cambio de contraseñas
    Me.Label25.Visible = Me.boxCambiarContraseña
    Me.txtPW2Config.Visible = Me.boxCambiarContraseña
End Sub

Private Sub boxEditarTablas_Click()
    AA
End Sub

Private Sub boxHato_Click()
    AA
End Sub

Private Sub boxLactAnteriores_Click()
   AA
End Sub

Private Sub boxReemplazos_Click()
    AA
End Sub

Private Sub boxReqRespaldo_Click()
    AA
End Sub

Private Sub boxSemen_Click()
    AA
End Sub

Private Sub boxUsuario_Click()
    AA
End Sub

Private Sub boxVersionDemo_Click()
    If bInitialFlag = True Then Exit Sub
    If Not Me.boxVersionDemo Then
        If Range("Desarrollador!B14").Text = vbNullString Then
            Range("Desarrollador!B14") = Date
          Else
            MsgBox _
              "Este sistema ya ha sido licenciado como Demo", _
              vbCritical, _
              "JP's Automatización de Aplicaciones"
            Application.Run "QuitarBanderaDemo" 'ModSeguridad
            Me.boxVersionDemo = _
              Range("Desarrollador!B13")
        End If
    End If
    AA
End Sub

Private Sub cmndBorrarInfo_Click()
    Dim Msg, Style, Title, Help, Ctxt, Response, MyString
    Msg = "Se borrará toda la informacion" _
      & Chr(13) & "Desea Continuar ?"
    Style = vbYesNo + vbCritical + vbDefaultButton2
    Title = "Borrar Toda la Información"
    'Ctxt = 1000    ' Define topic
        ' context.
        ' Display message.
    Response = MsgBox(Msg, Style, Title, Help, Ctxt)
    If Response = vbYes Then bBorrarTodo = True
End Sub

Private Sub cmndRecuperarInfo_Click()
    Application.Run "UnderConstruction"
End Sub

Private Sub CommandButton1_Click()
' Cerrar formulario
    'Application.Run "MostrarHojas" 'ModSeguridad
    Application.ScreenUpdating = True
    Unload Me
End Sub

Private Sub CommandButton2_Click()
' Guardar cambios
    Dim sMsj As String
    If Me.boxCambiarContraseña Then
        ' checar q contrseñas no estén en blanco
        If Me.txtPW1Config = vbNullString And _
          Me.txtPW2Config = vbNullString Then
            sMsj = MsgBox( _
              "Las Contraseñas están en Blanco", _
              vbInformation, _
              "Cambiar Contraseñas")
            Exit Sub
        End If
        ' Cambiar contraseñas
        Range("Desarrollador!B11") = _
          Me.txtPW1Config.Text
    End If
    ' Checar existencia de contraseña
    If Me.txtPW1Config = vbNullString Then
        sMsj = MsgBox("Ingresar Contraseña", _
          vbInformation, _
          "Efectuar Cambios")
        Me.txtPW1Config.SetFocus
        Exit Sub
    End If
    
    ' Checar coincidencias en contraseñas
    Select Case Me.txtPW1Config.Text
        Case Is = "16910852"
            EfectuarCambios
        Case Is = Range("Desarrollador!B11").Text
            EfectuarCambios
        'Case Is = Range("Desarrollador!B15").Text
            'EfectuarCambios
        Case Else
            sMsj = MsgBox("La Contraseña NO coincide", _
              vbInformation, "Efectuar Cambios")
    End Select
End Sub

Private Sub EfectuarCambios()
    On Error GoTo 120
100:
    Range("Configuracion!C39") = _
      Me.boxEditarTablas
    Range("Configuracion!B40") = _
      Me.boxHato
    Range("Configuracion!B43") = _
      Me.boxLactAnteriores
    Range("Configuracion!B41") = _
      Me.boxReemplazos
    Range("Configuracion!C25") = _
      Me.boxReqRespaldo
    Range("Configuracion!B42") = _
      Me.boxSemen
    Range("Desarrollador!B12") = _
      Me.boxUsuario
    'Range("Desarrollador!B13") = _
      Me.boxVersionDemo
    Range("Configuracion!C3") = _
      Me.txtNomEstablo
    If bBorrarTodo = True Then _
      Application.Run "BorrarTodo" 'ModSeguridad
    With Me
        .txtPW1Config = vbNullString
        .txtPW2Config = vbNullString
        .CommandButton2.Enabled = False
    End With
    On Error GoTo 0
    bBorrarTodo = False
    bCambios = False
    Exit Sub
120:
HabilitarHojas
GoTo 100
End Sub

Private Sub CommandButton5_Click()
' Otras Configuraciones
    Dim sMsj As String
    If bCambios Then
        sMsj = MsgBox( _
          "Hay cambios sin guardar", _
          vbInformation, _
          "Efectuar Cambios")
        Exit Sub
    End If
    usrConfiguracion.Show
    Unload Me
End Sub

Private Sub HabilitarHojas()
    With Sheets("Desarrollador")
        .Unprotect Password:="0246813579"
    End With
    With Sheets("Configuracion")
        .Unprotect Password:="0246813579"
    End With
    bProteccion = True
End Sub

Private Sub txtNomEstablo_Change()
    AA
End Sub

Private Sub txtPW2Config_AfterUpdate()
    ' Comprobar coincidencias en las contraseñas
    Dim sMsj As String
    If Not Me.txtPW1Config = _
      Me.txtPW2Config Then _
      sMsj = MsgBox( _
        "Las Contraseñas no coinciden", _
        vbInformation, _
        "Cambiar Contraseñas")
    AA
End Sub

Private Sub UserForm_Initialize()
    bInitialFlag = True
    Application.ScreenUpdating = _
      CBool(Range("Desarrollador!B6"))
    With Me
        ' Tomar Valores
        .boxEditarTablas = _
          CBool(Range("Configuracion!C39"))
        .boxHato = _
          CBool(Range("Configuracion!B40"))
        .boxLactAnteriores = _
          CBool(Range("Configuracion!B43"))
        .boxReemplazos = _
          CBool(Range("Configuracion!B41"))
        .boxReqRespaldo = _
          CBool(Range("Configuracion!C25"))
        .boxSemen = _
          CBool(Range("Configuracion!B42"))
        .boxUsuario = _
          CBool(Range("Desarrollador!B12"))
        .boxVersionDemo = _
          CBool(Range("Desarrollador!B13"))

        .txtNomEstablo = _
          Range("Configuracion!C3")
        .txtPW1Config = vbNullString
        .txtPW2Config = vbNullString
        ' Mostrar u Ocultar Controles
        .txtPW2Config.Visible = False
        .Label25.Visible = False
        .CommandButton2.Enabled = False
    End With
    bInitialFlag = False
    bBorrarTodo = False
    bCambios = False
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
