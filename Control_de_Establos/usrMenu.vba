VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} usrMenu 
   Caption         =   "Control de Establos"
   ClientHeight    =   5025
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   2865
   OleObjectBlob   =   "usrMenu.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "usrMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Ultima modificación: 2-Oct-2015
Option Explicit
Public ws2 As Sheets

Private Sub cmndCerrar_Click()
' Cerrar
    Application.Run "Proteger" 'Módulo2
    Application.ScreenUpdating = True
    Unload Me
End Sub

Private Sub cmndConfigurar_Click()
    usrContraseña.Show
End Sub

Private Sub cmndEstadisticas_Click()
    usrEstadísticas.Show
End Sub

Private Sub cmndHato_Click()
    usrVacas.Show
End Sub

Private Sub cmndHerramientas_Click()
    Application.Run "UnderConstruction"
End Sub

Private Sub cmndInformes_Click()
    Application.Run "UnderConstruction"
    'usrFiltros.Show
End Sub

Private Sub cmndReemplazos_Click()
    'usrReemplazos.Show
End Sub

Private Sub UserForm_Initialize()
    Application.ScreenUpdating = _
      CBool(Range("Desarrollador!B6"))
    If CBool(Range("Configuracion!C27")) Then 'Contraseña requerida
        If Range("Configuracion!C49") = _
          WorksheetFunction. _
          VLookup(Range("Configuracion!C49"), _
          Range("Tabla7"), 1, False) Then
            If CBool(WorksheetFunction. _
              VLookup(Range("Configuracion!C49"), _
              Range("Tabla7"), 5, False)) Then
                With Me
                    .cmndHato.Enabled = True
                    '.cmndReemplazos.Enabled = True
                End With
            End If
            If CBool(WorksheetFunction. _
              VLookup(Range("Configuracion!C49"), _
              Range("Tabla7"), 6, False)) Then
                With Me
                    .cmndConfigurar.Enabled = True
                    '.cmndHerramientas.Enabled = True
                End With
            End If
        End If
    Else
        With Me 'Contraseña no requerida
            .cmndHato.Enabled = True
            '.cmndReemplazos.Enabled = True
            '.cmndConfigurar.Enabled = True
            '.cmndHerramientas.Enabled = True
        End With
    End If
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

