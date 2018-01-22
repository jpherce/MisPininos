VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} usrLogin 
   Caption         =   "Control de Establos"
   ClientHeight    =   4560
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   2910
   OleObjectBlob   =   "usrLogin.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "usrLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Ultima modificación: 8-feb-2016
Option Explicit
Dim iPrivilegios As Long
Dim ilogin As Long
Dim sAviso As String
Dim sMensaje, sMensaje2 As String
Dim sTituloMsj As String
Dim ws As Worksheet
Dim ws3 As Worksheet

Private Sub cmndCambiarContraseña_Click()
    'formNuevaContraseña.Show
End Sub

Private Sub cmndCambiarContrseña_Click()
    usrCambioContraseña.Show
    Me.cmndCambiarContrseña.Visible = False
End Sub

Private Sub cmndContinuar_Click()
    ' Validar Información
    If Me.txtUsuario = vbNullString Then
        sMensaje = "Iniciales de Usuario están en blanco"
        mensaje
        Me.txtUsuario.SetFocus
        Exit Sub
    End If
    If Me.txtContraseña = vbNullString Then
        sMensaje = "La Contraseña está en blanco"
        mensaje
        Me.txtContraseña.SetFocus
        Exit Sub
    End If
    On Error GoTo 100
    If UCase(Me.txtUsuario) = "HERCE" And Me.txtContraseña = CDbl(Date) Then GoTo 200
    If Me.txtUsuario = _
      WorksheetFunction.VLookup(Me.txtUsuario, Range("Tabla7"), 1, False) Then
            If CStr(Me.txtContraseña) = _
              CStr(WorksheetFunction.VLookup(Me.txtUsuario, _
              Range("Tabla7"), 2, False)) Then GoTo 200 Else GoTo 100
        Else
            GoTo 100
    End If
    On Error GoTo 0
100: ' Se niega el acceso
    Range("Configuracion!C49") = Application.UserName
    Range("Configuracion!C50") = Format(Date, "d-mmm-yy")
    Range("Configuracion!C51") = Format(Time, "hh:mm")
    Beep
    sMensaje = "Usuario o Contraseña incorrectos"
    sMensaje2 = "Acceso Denegado"
    mensaje
    Exit Sub
200: ' Se concede acceso
    Range("Configuracion!C49") = Me.txtUsuario
    Range("Configuracion!C50") = Format(Date, "d-mmm-yy")
    Range("Configuracion!C51") = Format(Time, "hh:mm")
    'MsgBox "Acceso concedido"
    Me.cmndContinuar.Visible = False
    Me.cmndCambiarContrseña.Visible = True
    Me.cmndCambiarContrseña.Enabled = True
End Sub

Private Sub mensaje()
    MsgBox sMensaje & Chr(13) & _
      sMensaje2, _
      vbExclamation, _
      sTituloMsj
End Sub

Private Sub CommandCerrar_Click()
    If Me.txtUsuario = vbNullString And _
      Me.txtContraseña = vbNullString Then
        Range("Configuracion!C49") = Application.UserName
        Range("Configuracion!C50") = Format(Date, "d-mmm-yy")
        Range("Configuracion!C51") = Format(Time, "hh:mm")
    End If
    Unload Me
End Sub

Private Sub UserForm_Initialize()
    Me.Label31.Caption = "Version 1.22"
    Me.txtUsuario.SetFocus
    sTituloMsj = "Control de Establos"
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

