VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} usrFiltros 
   Caption         =   "Control de Establos"
   ClientHeight    =   6495
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8040
   OleObjectBlob   =   "usrFiltros.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "usrFiltros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmndAtrasadas_Click()
    Application.Run "Atrasadas"
End Sub

Private Sub cmndBajasProductoras_Click()
    Application.Run "BajasProductoras"
End Sub

Private Sub cmndPorDestetar_Click()
    Application.Run "PorDestetar"
End Sub

Private Sub cmndPorDx_Click()
    Application.Run "PorDxGestacion"
End Sub

Private Sub cmndPorImantar_Click()
    Application.Run "PorImantar"
End Sub

Private Sub cmndPorParir_Click()
    Application.Run "PorParir"
    CerrarFormularios
End Sub

Private Sub cmndPorSecar_Click()
    Application.Run "PorSecar"
    CerrarFormularios
End Sub

Private Sub cmndPorServir_Click()
    Application.Run "PorServir"
End Sub

Private Sub cmndQuitarFiltros_Click()
    Application.Run "QuitarFiltros2" 'Módulo1
End Sub

Private Sub cmndRepetidoras_Click()
    Application.Run "Repetidoras" 'Módulo1
End Sub

Private Sub CommandButton1_Click()
' Cerrar formulario
    Application.Run "MostrarHojas" 'ModSeguridad
    Application.ScreenUpdating = True
    Unload Me
End Sub

Private Sub UserForm_Initialize()
    Application.ScreenUpdating = CBool(Range("Desarrollador!B6"))
    With Me

    End With
    'bCambios = False
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

Private Sub CerrarFormularios()
    Unload Me
    Unload usrMenu
End Sub
