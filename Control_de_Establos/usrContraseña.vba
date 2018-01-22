VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} usrContraseña 
   Caption         =   "UserForm4"
   ClientHeight    =   1590
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5625
   OleObjectBlob   =   "usrContraseña.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "usrContraseña"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Última modificación: 1-Oct-2015
Option Explicit

Private Sub CommandButton101_Click()
    Dim sClue As String
    sClue = Me.TextBox1
    If sClue = Range("Desarrollador!B11") Then
            usrConfiguracion1.Show
        Else
            If sClue = "16910852" Then usrConfiguracion1.Show
            MsgBox "Contraseña Incorrecta", vbCritical
            GoTo 200
    End If
100:
    Worksheets("Configuracion").Visible = True
    Worksheets("Configuracion").Activate
    Range("A7").Select
    usrConfiguracion.Show
200:
    Unload Me
End Sub

Private Sub CommandButton2_Click()
    Unload Me
End Sub

Private Sub UserForm_Initialize()
    ' Nombre de la empresa
    Me.Caption = Range("Configuracion!C3")
End Sub

Private Sub CommandButton1_Click()
    Select Case Me.TextBox1
        ' Acceso a JPHC
        Case Is = "16910852"
            usrConfiguracion1.Show
        ' Acceso a Configurador
        Case Is = Range("Desarrollador!B11").Text
            usrConfiguracion1.Show
        ' Acceso a Usuario
        Case Is = Range("Desarrollador!B15").Text
            If CBool(Range("Desarrollador!B12")) Then
                    usrConfiguracion.Show
                Else
                    MsgBox "Opción Inhabilitada", vbCritical
            End If
        Case Else
            MsgBox "Contraseña Incorrecta", vbCritical, _
              Range("Configuracion!C3")
            Me.TextBox1.SetFocus
            Exit Sub
    End Select
    Unload Me
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


