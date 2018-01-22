VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} usrCambioContrase�a 
   Caption         =   "Cambio de Contrase�a"
   ClientHeight    =   1635
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5625
   OleObjectBlob   =   "usrCambioContrase�a.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "usrCambioContrase�a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Ultima modificaci�n: 22-Ene-16
Option Explicit

Private Sub CommandButton1_Click()
    Dim rCelda As Object
    Dim iRgln As Long
    ' Localizar usuario
    If Me.TextBox1 = Me.TextBox2 Then
            ' Salir si usuario no existe
            If Application.WorksheetFunction.CountIf(Range("Tabla7"), _
              Range("Configuracion!C49")) = 0 Then
                MsgBox "Usuario no registrado", vbCritical, Me.Caption
                Exit Sub
            End If
            For Each rCelda In Range("Tabla7[Usuario]")
                If rCelda = Range("Configuracion!C49") Then
                    iRgln = rCelda.Offset.Row
                End If
            Next rCelda
            Worksheets("Colaboradores").Cells(iRgln, 2) = _
              Me.TextBox1
            MsgBox "Contrase�a actualizada", vbInformation, Me.Caption
            Unload Me
        Else
            MsgBox "Hay diferencia en las contrase�as propuestas", _
              vbCritical, Me.Caption
            Me.TextBox1.SetFocus
    End If
End Sub

Private Sub CommandButton2_Click()
    Unload Me
End Sub

Private Sub UserForm_Initialize()
    ' Nombre de la empresa
    Me.Caption = "Cambiar Contrase�as"
End Sub

Private Sub UserForm_QueryClose(mCancel As Integer, _
                                mCloseMode As Integer)
' Deshabilitar el Bot�n X para cerrar el cuadro de di�logo
    If mCloseMode <> vbFormCode Then
        MsgBox "Utiliza 'Cerrar' para salir de este formulario.", _
          vbExclamation, "JP's Automatizaci�n de Aplicaciones"
        mCancel = True
    End If
End Sub

